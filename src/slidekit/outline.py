"""図形の輪郭。吹き出し（長方形か角丸の長方形の辺に三角形が付いた図形）の形を決める。

吹き出しの三角形を「しっぽ」、長方形の部分を「本体」、しっぽのとがった頂点を「先端」、本体の辺の上にある 2 頂点を
「付け根」と呼ぶ。要素の `area` は本体で、しっぽは `area` の外へ出る。
出力（`writer`）、配置の検証（場面の `actual_area`・`outside_parts`）、試験は、形を `tail_outline` の結果だけから求める。
座標はすべてページ上の pt である。
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
import math
from types import MappingProxyType

from .layout import Rect, as_rect
from .measure import _ROUNDED_ADJUST, text_area

# しっぽの寸法の既定値。書式にない項目だけを補う（styles.LIBRARY_DEFAULTS には置かない）。
TAIL_AT = 0.5
TAIL_WIDTH_PT = 10.0
TAIL_LENGTH_PT = 8.0
# しっぽを付けることができる辺と図形。
TAIL_SIDES = ("left", "right", "top", "bottom")
TAIL_SHAPES = ("rectangle", "rounded_rectangle")
# 接続点の番号（rectangle の既定図形と同じ）。しっぽの辺や位置で番号は変わらない。
SITE_NUMBERS = MappingProxyType({"top": 0, "left": 1, "bottom": 2, "right": 3})
# 比較の許容（model._EPSILON と同じ値）。
EPSILON = 1e-9


class TailWidthError(ValueError):
    """付け根の幅が辺の直線部分に収まらない。side_length は辺の長さ、radius は角丸の半径、max_width は指定できる最大の幅（pt）。"""

    def __init__(self, side: str, width: float, side_length: float, radius: float):
        self.side, self.width, self.side_length, self.radius = side, width, side_length, radius
        self.max_width = max(0.0, side_length - 2 * radius)
        super().__init__(
            f"しっぽの付け根の幅 {width:g} pt が {side} の辺の直線部分（{self.max_width:g} pt）を超えています"
        )


class TailPointError(ValueError):
    """指す点からしっぽの辺を決めることができない（点が本体の内側にあるか、固定した辺の外側にない）。"""

    def __init__(self, message: str, point, side: str | None):
        self.point, self.side = (float(point[0]), float(point[1])), side
        super().__init__(message)


@dataclass(frozen=True)
class Outline:
    """図形の輪郭。

    segments は輪郭で、`("move", x, y)`・`("line", x, y)`・`("arc", rx, ry, 開始角, 回転角)`・`("close",)` の列である。
    角度は `a:arcTo` と同じく右向きを 0°、時計回りを正とする（度）。actual_area は輪郭の外接長方形、text_area は文字領域、
    sites は接続点（辺の名前から (番号, 点)）、outside は本体の外へ出る部分の (名前, 外接長方形) の列である。
    tail（しっぽの 3 頂点。付け根・先端・付け根）と at（寄せた後の付け根の中心の位置、0〜1）は吹き出しだけが持つ。
    """

    segments: tuple
    actual_area: Rect
    text_area: Rect
    sites: Mapping = field(default_factory=dict)
    outside: tuple = ()
    tail: tuple | None = None
    at: float | None = None


def corner_radius(area, shape: str, corner_radius_pt: float | None = None) -> float:
    """角丸の半径（pt）。書き出しと計測と同じく min(0.5, corner_radius_pt / min(w, h)) × min(w, h)、
    指定がなければ 0.16667 × min(w, h)。rectangle は 0。"""
    area = as_rect(area)
    if shape == "rectangle":
        return 0.0
    if shape != "rounded_rectangle":
        raise ValueError(f"しっぽを付けることができない図形です: {shape}")
    shortest = min(area.w, area.h)
    if shortest <= 0:
        return 0.0
    adjust = _ROUNDED_ADJUST if corner_radius_pt is None else min(0.5, max(0.0, corner_radius_pt / shortest))
    return adjust * shortest


def _positive(value, name: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value <= 0:
        raise ValueError(f"{name} は正の有限の数にする: {value!r}")
    return float(value)


def tail_outline(
    area,
    shape: str,
    *,
    side: str,
    at: float = TAIL_AT,
    width_pt: float = TAIL_WIDTH_PT,
    length_pt: float = TAIL_LENGTH_PT,
    corner_radius_pt: float | None = None,
    tip_shift_pt: float = 0.0,
) -> Outline:
    """本体 area（x, y, w, h）の辺 side にしっぽを付けた吹き出しの輪郭。

    付け根の中心は辺の始点（左右の辺は上端、上下の辺は左端）から c = min(max(at × L, r + d/2), L − r − d/2) の位置で、
    角丸の弧にかからない位置へ寄せる（L は辺の長さ、r は角丸の半径、d は付け根の幅）。
    d が L − 2r を ε より大きく超えると `TailWidthError` にし、差が ε 以下なら c = r + d/2 にする。
    先端は付け根の中心から辺に垂直に length_pt 離れた点を、辺に沿って tip_shift_pt だけずらした点である
    （左右の辺は下向き、上下の辺は右向きが正。指す点へ向けるときに `tail_from_point` が決める）。
    輪郭は本体を時計回りに上辺・右辺・下辺・左辺の順にたどり、しっぽの辺では「付け根・先端・付け根」を直線で挿入する。
    直前の点との距離が ε 以下の直線は輪郭に入れない。
    """
    area = as_rect(area)
    if area.w is None or area.h is None or area.w <= 0 or area.h <= 0:
        raise ValueError(f"本体の幅と高さは正にする: {list(area)}")
    if side not in TAIL_SIDES:
        raise ValueError(f"しっぽの辺は {'、'.join(TAIL_SIDES)} のどれかにする: {side!r}")
    if isinstance(at, bool) or not isinstance(at, (int, float)) or not 0 <= at <= 1:
        raise ValueError(f"しっぽの位置は 0 以上 1 以下にする: {at!r}")
    d = _positive(width_pt, "width_pt")
    length = _positive(length_pt, "length_pt")
    if isinstance(tip_shift_pt, bool) or not isinstance(tip_shift_pt, (int, float)) or not math.isfinite(tip_shift_pt):
        raise ValueError(f"tip_shift_pt は有限の数にする: {tip_shift_pt!r}")
    s = float(tip_shift_pt)
    r = corner_radius(area, shape, corner_radius_pt)
    x, y, w, h = area
    edge = h if side in ("left", "right") else w
    straight = edge - 2 * r
    if d > straight + EPSILON:
        raise TailWidthError(side, d, edge, r)
    if abs(d - straight) <= EPSILON:
        c = r + d / 2
    else:
        c = min(max(at * edge, r + d / 2), edge - r - d / 2)
    # しっぽの頂点（左右の辺は上から下、上下の辺は左から右）。
    if side == "left":
        tail = ((x, y + c - d / 2), (x - length, y + c + s), (x, y + c + d / 2))
    elif side == "right":
        tail = ((x + w, y + c - d / 2), (x + w + length, y + c + s), (x + w, y + c + d / 2))
    elif side == "top":
        tail = ((x + c - d / 2, y), (x + c + s, y - length), (x + c + d / 2, y))
    else:
        tail = ((x + c - d / 2, y + h), (x + c + s, y + h + length), (x + c + d / 2, y + h))
    inserted = {side: tail if side in ("top", "right") else tuple(reversed(tail))}

    segments: list[tuple] = [("move", x + r, y)]
    current = (x + r, y)

    def line(point) -> None:
        nonlocal current
        if math.hypot(point[0] - current[0], point[1] - current[1]) > EPSILON:
            segments.append(("line", *point))
            current = point

    def arc(start: float, end_point) -> None:
        nonlocal current
        if r > 0:
            segments.append(("arc", r, r, start, 90.0))
            current = end_point

    for point in inserted.get("top", ()):
        line(point)
    line((x + w - r, y))
    arc(270.0, (x + w, y + r))
    for point in inserted.get("right", ()):
        line(point)
    line((x + w, y + h - r))
    arc(0.0, (x + w - r, y + h))
    for point in inserted.get("bottom", ()):
        line(point)
    line((x + r, y + h))
    arc(90.0, (x, y + h - r))
    for point in inserted.get("left", ()):
        line(point)
    line((x, y + r))
    arc(180.0, (x + r, y))
    segments.append(("close",))

    xs = [x, x + w, *(point[0] for point in tail)]
    ys = [y, y + h, *(point[1] for point in tail)]
    actual = Rect(min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))
    tail_xs, tail_ys = [point[0] for point in tail], [point[1] for point in tail]
    outside = (
        (
            "tail",
            Rect(min(tail_xs), min(tail_ys), max(tail_xs) - min(tail_xs), max(tail_ys) - min(tail_ys)),
        ),
    )
    sites = {
        "top": (SITE_NUMBERS["top"], (x + w / 2, y)),
        "left": (SITE_NUMBERS["left"], (x, y + h / 2)),
        "bottom": (SITE_NUMBERS["bottom"], (x + w / 2, y + h)),
        "right": (SITE_NUMBERS["right"], (x + w, y + h / 2)),
    }
    return Outline(
        segments=tuple(segments),
        actual_area=actual,
        text_area=text_area(area, shape, corner_radius_pt),
        sites=sites,
        outside=outside,
        tail=tail,
        at=c / edge,
    )


def tail_from_point(
    area,
    shape: str,
    point,
    *,
    side: str | None = None,
    corner_radius_pt=None,
    width_pt=TAIL_WIDTH_PT,
    length_pt=TAIL_LENGTH_PT,
) -> tuple[str, float, float]:
    """指す点 point から、しっぽの辺と付け根の位置（0〜1）と先端のずらし（pt）を決める。(辺, 位置, ずらし) を返す。

    side を渡すとその辺に固定し、point がその辺の外側になければ `TailPointError` にする。渡さなければ、point が本体の外へ
    出ている距離を横と縦で求め、大きい方の辺にする（等しいときは左右の辺）。point が本体の内側なら `TailPointError` にする。
    付け根の中心は point を辺へ垂直に下ろした位置で、`tail_outline` と同じく直線部分へ寄せる（幅が収まらなければ
    `TailWidthError`）。
    寄せたために point が付け根の正面にないときは、先端が付け根の中心から point へ向かう直線の上に来るよう、辺に沿って
    ずらす。ずらしは ±length_pt まで（辺の法線から 45° まで）に抑える。point が付け根の正面にあればずらしは 0 である。
    """
    area = as_rect(area)
    x, y, w, h = area
    px, py = point
    if side is None:
        horizontal = x - px if px < x else px - (x + w) if px > x + w else 0.0
        vertical = y - py if py < y else py - (y + h) if py > y + h else 0.0
        if horizontal <= 0 and vertical <= 0:
            raise TailPointError("指す点が本体の内側にあります", point, None)
        if horizontal >= vertical:
            side = "left" if px < x else "right"
        else:
            side = "top" if py < y else "bottom"
    else:
        outside = {"left": px < x, "right": px > x + w, "top": py < y, "bottom": py > y + h}
        if side not in outside:
            raise ValueError(f"しっぽの辺は {'、'.join(TAIL_SIDES)} のどれかにする: {side!r}")
        if not outside[side]:
            raise TailPointError(f"指す点が {side} の辺の外側にありません", point, side)
    edge = h if side in ("left", "right") else w
    foot = py - y if side in ("left", "right") else px - x
    r = corner_radius(area, shape, corner_radius_pt)
    d = _positive(width_pt, "width_pt")
    length = _positive(length_pt, "length_pt")
    straight = edge - 2 * r
    if d > straight + EPSILON:
        raise TailWidthError(side, d, edge, r)
    if abs(d - straight) <= EPSILON:
        c = r + d / 2
    else:
        c = min(max(foot, r + d / 2), edge - r - d / 2)
    # 辺から point までの距離（法線の向き）と、付け根の中心から point までの辺に沿った距離。
    normal = {"left": x - px, "right": px - (x + w), "top": y - py, "bottom": py - (y + h)}[side]
    along = foot - c
    shift = 0.0 if abs(along) <= EPSILON else min(length, max(-length, length * along / normal))
    return side, c / edge, shift
