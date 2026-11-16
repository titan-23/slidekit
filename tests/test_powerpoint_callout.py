"""SLIDEKIT_TEST_POWERPOINT=1 を指定したときだけ、吹き出し（しっぽのある box）を実際の PowerPoint の変換で確かめる。

確かめること：文字がしっぽのない図形と同じ位置に描かれること（a:rect の文字領域）、本体の角がしっぽのない図形と同じ
見た目になること（角丸の a:arcTo と、長方形の角の結合 a:round）、輪郭が 1 つの描画で付け根に継ぎ目がないこと、枠線が
actual_area から線幅の半分より外へ出ないこと、線が接続点に結ばれて書いた端点のまま描かれること、数式を含む吹き出しの
互換画像がしっぽまでを含むこと。資料は作成 API で構築し（フィクスチャ windows_fonts_environment で Windows の既定フォントを
使う）、PowerPoint で PDF にする。直列に実行する（`-p no:xdist`）。
"""

import json
import os

from pdf_tools import line_segments, open_pdf, render_region
import pytest

from test_powerpoint_build import build_deck, contains, dark_pixels
from test_powerpoint_shapes import construct_scene, elements, write_and_render

pytestmark = [
    pytest.mark.powerpoint,
    pytest.mark.skipif(
        os.environ.get("SLIDEKIT_TEST_POWERPOINT") != "1", reason="実機 PowerPoint を使う任意の検査である。"
    ),
]

# 文字の位置（pt）、描画の範囲（pt）、線の端点（pt）の許容幅。
TEXT_TOLERANCE_PT = 0.05
PATH_TOLERANCE_PT = 0.1
LINE_TOLERANCE_PT = 0.6
# 本体の角を比べる範囲（pt）と、画像の倍率。
CORNER_PT = 16
SCALE = 4
LINE_WIDTH = 3.0
SIDES = ("left", "right", "top", "bottom")

# 吹き出しと、同じ本体のしっぽのない図形を上下に並べる。枠線は太くし、先端の結合の違いが画素に表れるようにする。
CALLOUTS = """from slidekit import Deck, Rect

TEXT = "吹き出しの文字領域は本体と同じ。Callout text."

def build(context):
    deck = Deck(context)
    slide = deck.slide(id="callouts", kind="blank")
    style = dict(size_pt=14, valign="middle", padding_pt=6, line_color="183153", fill="FFF4D6", line_width_pt=3)
    for row, (shape, radius) in enumerate((("rectangle", None), ("rounded_rectangle", 12))):
        extra = {} if radius is None else {"corner_radius_pt": radius}
        for column, side in enumerate(("left", "right", "top", "bottom")):
            x, y = 40 + column * 230, 20 + row * 260
            slide.box(id=f"{shape}_{side}", area=Rect(x + 12, y + 14, 180, 90), text=TEXT, shape=shape,
                      tail_side=side, tail_at=0.3, **style, **extra)
            slide.box(id=f"{shape}_{side}_plain", area=Rect(x + 12, y + 134, 180, 90), text=TEXT, shape=shape,
                      **style, **extra)
    return deck
"""

# 線を接続した吹き出し。しっぽのない辺に、辺の中央（gap 0、at 0.5）と、ずらした位置（gap 6、at 0.25）で結ぶ。
LINKS = """from slidekit import Deck, Rect

def build(context):
    deck = Deck(context)
    slide = deck.slide(id="links", kind="blank")
    targets = [
        slide.box(id="top", area=Rect(300, 60, 140, 50), text="上", line_color="183153", tail_side="top"),
        slide.box(id="right", area=Rect(300, 220, 140, 50), shape="rounded_rectangle", text="右",
                  line_color="183153", corner_radius_pt=8, tail_side="right", tail_at=0.8),
        slide.box(id="toward", area=Rect(300, 380, 140, 50), text="指す", line_color="183153",
                  tail_toward=(560, 360)),
    ]
    for index, target in enumerate(targets):
        x, y = target.area.x, target.area.y
        slide.line(id=f"{target.id}_left", start=(x - 180, y + 10), end=target.anchor("left"),
                   line_color=COLORS[2 * index], line_width_pt=1.5)
        slide.line(id=f"{target.id}_bottom", start=target.anchor("bottom", gap=6, at=0.25), end=(x - 60, y + 100),
                   line_color=COLORS[2 * index + 1], line_width_pt=1.5)
    return deck

COLORS = ["8A1F1F", "1F6F8A", "6F8A1F", "8A6F1F", "1F8A4D", "4D1F8A"]
"""

# 数式を含む吹き出し（互換画像を付ける資料）。しっぽの側に別の文字を置く。
MATH = """from slidekit import Deck, Math, Rect, Text

def build(context):
    deck = Deck(context)
    slide = deck.slide(id="math", kind="content", title="吹き出し", notes="説明します。")
    slide.text(id="beside", area=Rect(60, 160, 222, 40), text="左の本文", align="right")
    slide.box(id="note", area=Rect(300, 150, 260, 70), text=[Text("ここで "), Math("x_i"), Text(" と比べる")],
              line_color="1F2937", line_width_pt=2, tail_side="left", tail_at=0.4)
    slide.box(id="below", area=Rect(300, 300, 260, 70), shape="rounded_rectangle", corner_radius_pt=8,
              text=[Math("\\\\frac{a}{b}")], fill="FFF4D6", line_color="1F2937", tail_side="top", align="center")
    return deck
"""


def body_characters(page, area):
    """本体の範囲に中心がある見える文字の (文字, 原点の本体の左上からの位置) の列（描画順）。"""
    x, y, w, h = area
    found = []
    for span in page.text_spans():
        for char in span.chars:
            cx, cy = (char.bbox[0] + char.bbox[2]) / 2, (char.bbox[1] + char.bbox[3]) / 2
            if x <= cx <= x + w and y <= cy <= y + h and chr(char.unicode).strip():
                found.append((chr(char.unicode), (char.origin[0] - x, char.origin[1] - y)))
    return found


def corner_difference(page, first, second):
    """2 つの範囲（pt）を画像にし、明度の差が 64 を超える画素の数。"""
    a, b = render_region(page, SCALE, first), render_region(page, SCALE, second)
    assert a.size == b.size
    pixels = zip(a.convert("L").tobytes(), b.convert("L").tobytes())
    return sum(1 for left, right in pixels if abs(left - right) > 64)


def test_real_powerpoint_draws_callouts_like_boxes(tmp_path, write_deck, windows_fonts_environment):
    root = tmp_path / "deck"
    project, scene = construct_scene(root, write_deck, CALLOUTS)
    _result, pdf = write_and_render(project, scene, "callouts")
    found = elements(scene)
    with open_pdf(pdf) as document:
        page = document[0]
        drawings = page.drawings()
        for shape in ("rectangle", "rounded_rectangle"):
            for side in SIDES:
                callout, plain = found[f"{shape}_{side}"], found[f"{shape}_{side}_plain"]
                name = callout["id"]
                # 1. 文字は、同じ本体のしっぽのない図形と同じ位置に描かれる（文字領域が本体の形で決まる）。
                text, expected = body_characters(page, callout["area"]), body_characters(page, plain["area"])
                assert [char for char, _ in text] == [char for char, _ in expected] and text, name
                for (char, a), (_, b) in zip(text, expected):
                    assert all(abs(p - q) <= TEXT_TOLERANCE_PT for p, q in zip(a, b)), (name, char, a, b)
                # 3. 輪郭は actual_area の範囲の描画だけで、付け根の継ぎ目などの別の描画はない。描画の数はしっぽのない図形と同じ。
                ax, ay, aw, ah = callout["actual_area"]
                own = matching(drawings, (ax, ay, ax + aw, ay + ah))
                near = [d for d in drawings if contains((ax - 1, ay - 1, ax + aw + 1, ay + ah + 1), d.rect, 0)]
                px, py, pw, ph = plain["area"]
                reference = matching(drawings, (px, py, px + pw, py + ph))
                assert own and len(own) == len(near) == len(reference), (name, own, near, reference)
                # 4. 枠線は actual_area から線幅の半分（と画素の境界の分）より外へ出ない（先端と角の結合は a:round）。
                margin = LINE_WIDTH / 2 + 0.5
                outside = dark_pixels_beyond(page, (ax, ay, ax + aw, ay + ah), margin)
                assert outside == 0, (name, outside)
                # 2・4. しっぽのない辺の角は、しっぽのない図形の角と同じに見える（角丸の弧と、長方形の角の結合）。
                corner = {"left": "right_top", "right": "left_top", "top": "left_bottom", "bottom": "left_top"}[side]
                difference = corner_difference(
                    page, corner_box(callout["area"], corner), corner_box(plain["area"], corner)
                )
                assert difference <= 12, (name, corner, difference)


def matching(drawings, box):
    """外接矩形が box（x0, y0, x1, y1）と一致する描画。"""
    return [d for d in drawings if all(abs(p - q) <= PATH_TOLERANCE_PT for p, q in zip(d.rect, box))]


def corner_box(area, corner):
    """本体の角を含む CORNER_PT 四方の範囲（x0, y0, x1, y1）。枠線の外側の半分を含める。"""
    x, y, w, h = area
    pad = LINE_WIDTH
    left = x - pad if corner.startswith("left") else x + w - CORNER_PT + pad
    top = y - pad if corner.endswith("top") else y + h - CORNER_PT + pad
    return (left, top, left + CORNER_PT, top + CORNER_PT)


def dark_pixels_beyond(page, box, margin):
    """box を margin だけ広げた範囲の外側（幅 4 pt の帯）にある、暗い画素の数。"""
    x0, y0, x1, y1 = box[0] - margin, box[1] - margin, box[2] + margin, box[3] + margin
    bands = [(x0 - 4, y0 - 4, x0, y1 + 4), (x1, y0 - 4, x1 + 4, y1 + 4), (x0, y0 - 4, x1, y0), (x0, y1, x1, y1 + 4)]
    return sum(dark_pixels(page, band) for band in bands)


def test_real_powerpoint_connects_lines_to_callouts(tmp_path, write_deck, windows_fonts_environment):
    root = tmp_path / "deck"
    project, scene = construct_scene(root, write_deck, LINKS)
    result, pdf = write_and_render(project, scene, "links")
    assert result["connections"] == 6
    found = elements(scene)
    assert found["toward"]["tail"]["side"] == "right"
    colors = ["8A1F1F", "1F6F8A", "6F8A1F", "8A6F1F", "1F8A4D", "4D1F8A"]
    lines = line_segments(pdf, 0)
    names = [f"{target}_{end}" for target in ("top", "right", "toward") for end in ("left", "bottom")]
    for name, color in zip(names, colors):
        element = found[name]
        expected = tuple(int(color[i : i + 2], 16) / 255 for i in (0, 2, 4))
        segments = [(s, e) for drawn, s, e in lines if all(abs(a - b) <= 2 / 255 for a, b in zip(drawn, expected))]
        assert len(segments) == 1, (name, segments)
        # PowerPoint は接続した線を開いたときに接続点へ引き直さず、書いた端点のまま描く。
        for actual, wanted in zip(sorted(segments[0]), sorted([tuple(element["start"]), tuple(element["end"])])):
            assert all(abs(a - b) <= LINE_TOLERANCE_PT for a, b in zip(actual, wanted)), (name, actual, wanted)


def test_real_fallback_crop_includes_the_tail(tmp_path, write_deck, windows_fonts_environment):
    result, run, record = build_deck(tmp_path / "deck", write_deck, MATH)
    scene = json.loads((run / "deck.json").read_text(encoding="utf-8"))
    found = elements(scene)
    crops = {crop["element_id"]: crop for crop in record["operations"]["math_fallback"]["crops"]}
    assert set(crops) == {"note", "below"}
    codes = {(issue["code"], issue.get("element_id")) for issue in result["issues"]}
    assert not {code for code, _ in codes} & {"fallback_crop_clipped", "pdf_element_text_overflow", "element_overlap"}
    step = 1 / record["operations"]["math_fallback"]["image_scale"]
    for element_id, crop in crops.items():
        x, y, w, h = found[element_id]["actual_area"]
        half = crop["outline_pt"] / 2
        # 互換画像の図形は a:xfrm（しっぽまでの範囲）で、切り出しはしっぽと枠線の外側の半分を含む。
        assert crop["rectangle_pt"] == pytest.approx([x, y, x + w, y + h], abs=0.01)
        assert contains(crop["crop_pt"], [x - half, y - half, x + w + half, y + h + half], tolerance=step)
        assert not crop["clipped"] and crop["nearby_elements"] == []
        # 互換表示でも、しっぽが標準表示と同じだけの暗い画素で描かれる。
        [tail] = found[element_id]["outside_parts"]
        tx, ty, tw, th = tail["area"]
        area = [tx - half, ty - half, tx + tw + half, ty + th + half]
        with open_pdf(run / "renders/powerpoint/standard/deck.pdf") as document:
            standard = dark_pixels(document[0], area)
        with open_pdf(run / "renders/powerpoint/fallback/deck.pdf") as document:
            shown = dark_pixels(document[0], area)
        assert standard > 0 and abs(shown - standard) <= max(6, standard // 5), (element_id, standard, shown)
