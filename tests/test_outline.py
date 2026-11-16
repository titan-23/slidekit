"""吹き出しの輪郭（`outline.tail_outline`・`tail_from_point`）を確かめる：辺ごとの頂点と範囲、角丸の弧、付け根の寄せ方、
長さ 0 の線を除く規則、幅の境界、文字領域、接続点の番号。"""

import math

import pytest

from slidekit.layout import Rect
from slidekit.measure import text_area
from slidekit.outline import (
    SITE_NUMBERS,
    TAIL_AT,
    TAIL_LENGTH_PT,
    TAIL_WIDTH_PT,
    TailPointError,
    TailWidthError,
    corner_radius,
    tail_from_point,
    tail_outline,
)

AREA = Rect(420, 150, 240, 70)


@pytest.mark.parametrize(
    ("side", "tail", "actual", "outside"),
    [
        ("left", [(420, 159), (412, 164), (420, 169)], (412, 150, 248, 70), (412, 159, 8, 10)),
        ("right", [(660, 159), (668, 164), (660, 169)], (420, 150, 248, 70), (660, 159, 8, 10)),
        ("top", [(463, 150), (468, 142), (473, 150)], (420, 142, 240, 78), (463, 142, 10, 8)),
        ("bottom", [(463, 220), (468, 228), (473, 220)], (420, 150, 240, 78), (463, 220, 10, 8)),
    ],
)
def test_each_side_has_vertices_area_and_outside_part(side, tail, actual, outside):
    outline = tail_outline(AREA, "rectangle", side=side, at=0.2)
    assert [tuple(point) for point in outline.tail] == tail
    assert tuple(outline.actual_area) == actual
    assert [(name, tuple(rect)) for name, rect in outline.outside] == [("tail", outside)]
    assert outline.at == pytest.approx(0.2)
    # 輪郭はしっぽの 3 頂点を、進む向き（時計回り）の順に含む。
    points = [segment[1:] for segment in outline.segments if segment[0] in ("move", "line")]
    order = [points.index(point) for point in tail]
    expected = sorted(order) if side in ("top", "right") else sorted(order, reverse=True)
    assert order == expected
    assert outline.segments[-1] == ("close",)


def test_rectangle_outline_is_clockwise_without_arcs():
    outline = tail_outline(AREA, "rectangle", side="left", at=0.2)
    assert outline.segments == (
        ("move", 420, 150),
        ("line", 660, 150),
        ("line", 660, 220),
        ("line", 420, 220),
        ("line", 420, 169),
        ("line", 412, 164),
        ("line", 420, 159),
        ("line", 420, 150),
        ("close",),
    )


def test_rounded_rectangle_arcs_start_and_end_on_corners():
    outline = tail_outline(Rect(100, 80, 200, 50), "rounded_rectangle", side="bottom", corner_radius_pt=6)
    assert outline.segments == (
        ("move", 106, 80),
        ("line", 294, 80),
        ("arc", 6, 6, 270, 90),
        ("line", 300, 124),
        ("arc", 6, 6, 0, 90),
        ("line", 205, 130),
        ("line", 200, 138),
        ("line", 195, 130),
        ("line", 106, 130),
        ("arc", 6, 6, 90, 90),
        ("line", 100, 86),
        ("arc", 6, 6, 180, 90),
        ("close",),
    )
    # 弧の始点と終点：a:arcTo と同じく、直前の点を開始角の位置とする円の上を回転角だけ進む。始点と終点は角の接点になる。
    current, ends = None, []
    for segment in outline.segments:
        if segment[0] in ("move", "line"):
            current = segment[1:]
        elif segment[0] == "arc":
            _kind, rx, _ry, start, sweep = segment
            center = (current[0] - rx * math.cos(math.radians(start)), current[1] - rx * math.sin(math.radians(start)))
            end = math.radians(start + sweep)
            current = (center[0] + rx * math.cos(end), center[1] + rx * math.sin(end))
            ends.append(pytest.approx(current, abs=1e-9))
    assert ends == [(300, 86), (294, 130), (100, 124), (106, 80)]


@pytest.mark.parametrize(("at", "center"), [(0.0, 11), (1.0, 189), (0.5, 100)])
def test_base_is_moved_off_the_arcs(at, center):
    # 角丸の半径 6、付け根の幅 10 の上辺：中心は 6 + 5 = 11 から 200 - 11 = 189 までに寄る。
    outline = tail_outline(Rect(0, 0, 200, 50), "rounded_rectangle", side="top", at=at, corner_radius_pt=6)
    assert outline.tail[1][0] == pytest.approx(center)
    assert outline.at == pytest.approx(center / 200)


def test_rectangle_corner_tail_has_no_zero_length_line():
    outline = tail_outline(Rect(0, 0, 100, 40), "rectangle", side="top", at=0)
    # 付け根の左端が角の頂点（移動の点）と同じなので、長さ 0 の線を入れない。
    assert outline.segments[:3] == (("move", 0, 0), ("line", 5, -8), ("line", 10, 0))
    previous = None
    for segment in outline.segments:
        if segment[0] in ("move", "line"):
            if previous is not None:
                assert math.dist(previous, segment[1:]) > 1e-9
            previous = segment[1:]
    left = tail_outline(Rect(0, 0, 100, 40), "rectangle", side="left", at=0)
    assert left.segments[-3:] == (("line", -8, 5), ("line", 0, 0), ("close",))


def test_width_equal_to_straight_part_is_accepted():
    # 直線部分 = 50 - 2 × 6 = 38。幅 38 は収まり、付け根の中心は r + d / 2 に固定する。
    outline = tail_outline(
        Rect(0, 0, 50, 50), "rounded_rectangle", side="right", at=0.9, width_pt=38, corner_radius_pt=6
    )
    assert outline.tail[0][1] == pytest.approx(6) and outline.tail[2][1] == pytest.approx(44)
    # 半径が小数になる寸法（既定の比 0.16667）でも、丸め誤差でエラーにしない。
    area = Rect(0, 0, 37.3, 23.9)
    radius = corner_radius(area, "rounded_rectangle")
    tail_outline(area, "rounded_rectangle", side="left", width_pt=23.9 - 2 * radius)
    with pytest.raises(TailWidthError) as raised:
        tail_outline(Rect(0, 0, 50, 50), "rounded_rectangle", side="right", width_pt=38.01, corner_radius_pt=6)
    assert (raised.value.side_length, raised.value.radius, raised.value.max_width) == (50, 6, 38)


def test_text_area_and_sites_follow_the_body():
    for shape, radius in (("rectangle", None), ("rounded_rectangle", None), ("rounded_rectangle", 9)):
        outline = tail_outline(AREA, shape, side="bottom", corner_radius_pt=radius)
        assert outline.text_area == text_area(AREA, shape, radius)
    outline = tail_outline(AREA, "rectangle", side="left")
    assert dict(SITE_NUMBERS) == {"top": 0, "left": 1, "bottom": 2, "right": 3}
    assert outline.sites == {
        "top": (0, (540, 150)),
        "left": (1, (420, 185)),
        "bottom": (2, (540, 220)),
        "right": (3, (660, 185)),
    }


def test_defaults():
    assert (TAIL_AT, TAIL_WIDTH_PT, TAIL_LENGTH_PT) == (0.5, 10.0, 8.0)
    outline = tail_outline(AREA, "rectangle", side="right")
    assert outline.tail[1] == (668, 185)


def test_tail_from_point_chooses_the_farther_side():
    area = Rect(100, 100, 100, 50)
    assert tail_from_point(area, "rectangle", (40, 120)) == ("left", pytest.approx(0.4))
    assert tail_from_point(area, "rectangle", (150, 300)) == ("bottom", pytest.approx(0.5))
    # 横と縦の距離が等しいときは左右の辺。
    assert tail_from_point(area, "rectangle", (90, 90))[0] == "left"
    # 付け根は直線部分へ寄せる。
    assert tail_from_point(area, "rectangle", (300, 0)) == ("right", pytest.approx(5 / 50))
    # 辺を固定すると、その辺の外側の点だけを受け付ける。
    assert tail_from_point(area, "rectangle", (150, 0), side="top") == ("top", pytest.approx(0.5))
    with pytest.raises(TailPointError):
        tail_from_point(area, "rectangle", (150, 120))
    with pytest.raises(TailPointError) as raised:
        tail_from_point(area, "rectangle", (150, 0), side="left")
    assert raised.value.side == "left"
