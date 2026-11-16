"""互換画像の埋め込み、最終形の PPTX、互換表示の確認用複製を確かめる。"""

from contextlib import contextmanager
import io
import math
from pathlib import Path
import pickle
import tempfile
from types import SimpleNamespace
import zipfile

from lxml import etree
from pdf_tools import Document, render_region
from PIL import Image
from pptx import Presentation
from pptx.util import Pt
import pytest

from slidekit.fallback import (
    ALPHA_FLOOR,
    MAX_MARGIN_PT,
    _crop_png,
    _math_fonts,
    _Region,
    _shape_regions,
    _wrap_math,
    compare_package_structure,
    crop_rectangle,
    fallback_required,
    finalize_math,
    keep_slides,
    outline_width_pt,
    prepare_fallback,
    prepare_final_form,
)
from slidekit.ooxml import NS, read_package, write_package
from slidekit.pdf import Box, open_pdf
from slidekit.writer import write_pptx


def project(tmp_path, **changes):
    config = {
        "metadata": {"title": "出力の検査", "author": "", "last_modified_by": ""},
        "math": {"functions": ["Lookup"]},
    }
    config.update(changes)
    return SimpleNamespace(config=config, root=tmp_path, resolve_input=lambda value: (tmp_path / value).resolve())


def element(identifier="label", kind="text", content=None, area=(30, 40, 200, 70), style=None, **attributes):
    return {
        "id": identifier,
        "kind": kind,
        "area": list(area),
        "style": {
            "size_pt": 24,
            "min_size_pt": 18,
            "font_japanese": "Yu Gothic UI",
            "font_latin": "Arial",
            "font_math": "Cambria Math",
            "color": "17212D",
            "padding_pt": 0,
            "line_spacing": 1.15,
            **(style or {}),
        },
        "content": content or [{"kind": "text", "text": "長い文章も縮めない\n改行"}],
        **attributes,
    }


def scene(elements, *pages):
    """1 ページ目の要素と、続くページの (識別子, 要素) から場面を作る。"""
    slides = [{"id": "first", "kind": "content", "title": "検査", "notes": "説明します。", "elements": elements}]
    slides += [
        {"id": identifier, "kind": "content", "title": identifier, "notes": "", "elements": items}
        for identifier, items in pages
    ]
    return {
        "schema_version": 1,
        "size": [720, 540],
        "metadata": {"title": "出力の検査", "author": "", "last_modified_by": ""},
        "slides": slides,
    }


def math_element(identifier="label", expression="x_i", **changes):
    return element(identifier, content=[{"kind": "math", "text": expression}], **changes)


def xml_part(path, name="ppt/slides/slide1.xml"):
    with zipfile.ZipFile(path) as archive:
        return etree.fromstring(archive.read(name))


def pdf_pages(path, *drawers, size=(720, 540)):
    """size（既定 720×540）のページを並べた PDF。各ページは描く関数（None なら空白）で作る。"""
    with Document() as document:
        for draw in drawers:
            page = document.new_page(*size)
            if draw:
                draw(page)
        document.save(path)
    return path


@contextmanager
def drawn_page(draw, size=(720, 540)):
    """draw で描いた 1 ページの PDF を一時ファイルに書き、`slidekit.pdf` で開いたページ。"""
    with tempfile.TemporaryDirectory() as directory:
        with open_pdf(pdf_pages(Path(directory) / "page.pdf", draw, size=size)) as document:
            yield document[0]


def own_math(page):
    """矩形 (30, 40, 230, 110) の数式。描画順で後の塗りが下へ 8 pt はみ出す。"""
    page.insert_text((35, 80), "x", fontsize=24)
    page.draw_rect((150, 60, 160, 118), color=None, fill=(0, 0, 0))


def crop_of(draw, rectangle, *, outline_pt=0.0, **options):
    with drawn_page(draw) as page:
        return crop_rectangle(
            page,
            Box.of(rectangle),
            outline_pt=outline_pt,
            traces=page.spans(),
            drawings=page.drawings(),
            scale=4,
            **options,
        )


def image_of(path, name):
    with zipfile.ZipFile(path) as archive:
        image = Image.open(io.BytesIO(archive.read(name)))
        image.load()
    return image


def test_final_math_preserves_native_and_embeds_pdf_crop(tmp_path):
    data = scene([element(content=[{"kind": "math", "text": "Lookup(x_i)"}])])
    native, pdf, final, fallback = [
        tmp_path / name for name in ("native.pptx", "native.pdf", "deck.pptx", "fallback.pptx")
    ]
    write_pptx(data, project(tmp_path), native)
    pdf_pages(pdf, lambda page: page.insert_text((35, 60), "Lookup(x)", fontsize=24))
    before = native.read_bytes()
    result = finalize_math(native, pdf, final)
    root = xml_part(final)
    assert native.read_bytes() == before
    assert result["fallback_shapes"] == 1
    assert result["crops"][0]["element_id"] == "label"
    assert result["crops"][0]["rectangle_pt"] == [30, 40, 230, 110]
    choice = root.find(".//mc:Choice", NS)
    assert choice.get("Requires") == "a14"
    assert choice.find(".//m:oMath", NS) is not None
    assert root.find(".//mc:Fallback//a:blip", NS) is not None
    assert root.find(".//mc:Fallback//p:cNvPr", NS).get("name") == "label"
    cloned = prepare_fallback(final, fallback)
    assert cloned["alternatives_replaced"] == 1
    clone = xml_part(fallback)
    assert not clone.findall(".//m:oMath", NS)
    assert len(clone.findall(".//a:blip", NS)) == 1
    assert root.find(".//m:oMath", NS) is not None


def test_bad_pdf_dimensions_and_unfinished_fallback_fail(tmp_path):
    data = scene([element(content=[{"kind": "math", "text": "x_i"}])])
    native = tmp_path / "native.pptx"
    write_pptx(data, project(tmp_path), native)
    pdf_pages(tmp_path / "wrong.pdf", None, size=(100, 100))
    with pytest.raises(ValueError, match="寸法"):
        finalize_math(native, tmp_path / "wrong.pdf", tmp_path / "final.pptx")
    assert not (tmp_path / "final.pptx").exists()
    with pytest.raises(ValueError, match="互換画像のない数式"):
        prepare_fallback(native, tmp_path / "fallback.pptx")


def test_crop_includes_overflowing_own_glyph_and_fill():
    def draw(page):
        own_math(page)
        page.insert_text((40, 100), "y", fontsize=20)
        page.insert_text((232, 80), "z", fontsize=12)

    # 基点が枠内の文字の外接矩形と、その後に描いた塗りを含める。右端の文字は基点が 0.5 pt の余裕の外なので含めない。
    crop, details = crop_of(draw, (30, 40, 230, 110))
    assert list(crop) == [30, 40, 230, 118]
    assert details == {"margin_pt": [0, 0, 0, 8], "clipped": False, "outline_pt": 0.0, "nearby_elements": []}
    glyph, _ = crop_of(lambda page: page.insert_text((35, 108), "g", fontsize=24), (30, 40, 230, 110))
    assert glyph.y1 > 110


def test_crop_ignores_neighbor_glyph_and_later_fill():
    def draw(page):
        page.insert_text((35, 80), "x", fontsize=24)
        page.insert_text((35, 150), "N", fontsize=60)  # 隣の文字：基点は枠外、外接矩形は枠に掛かる
        page.draw_rect((100, 90, 120, 125), color=None, fill=(1, 0, 0))  # 隣の文字の後に描いた塗り

    crop, details = crop_of(draw, (30, 40, 230, 110))
    assert list(crop) == [30, 40, 230, 110]
    assert details["margin_pt"] == [0, 0, 0, 0] and not details["clipped"]


def nested_math(page):
    """矩形 (100, 100, 300, 124) の図形に描いた、枠からはみ出す入れ子の数式（PowerPoint の描画順）。

    前の要素の文字、図形の塗りと枠線、括弧（矩形に掛かり、下へ 12 pt はみ出す）、分数線（矩形の下）、根号、
    基点が矩形の中の文字 x、基点が矩形の下の文字 y（はみ出した分母）の順に描く。
    """
    page.insert_text((400, 60), "N", fontsize=20)
    page.draw_rect((100, 100, 300, 124), color=(0, 0, 0), fill=(1, 1, 1))
    page.draw_rect((102, 104, 140, 136), color=None, fill=(0, 0, 0))
    page.draw_rect((110, 126, 130, 127), color=None, fill=(0, 0, 0))
    page.draw_rect((110, 104, 130, 106), color=None, fill=(0, 0, 0))
    page.insert_text((112, 120), "x", fontsize=14)
    page.insert_text((112, 135), "y", fontsize=14)


def test_crop_follows_math_lines_drawn_before_own_glyphs():
    # 括弧と根号は最初の字形より前に描かれる。図形の枠線で止まるまでさかのぼって含める（前の要素の文字は含めない）。
    crop, details = crop_of(nested_math, (100, 100, 300, 124))
    assert list(crop) == [100, 100, 300, 136] and not details["clipped"]
    # 数式の書体を渡すと、自身の文字に続く同じ書体の文字（はみ出した分母）も含める。
    with drawn_page(nested_math) as page:
        (y,) = [span for span in page.spans() if span.chars[0].unicode == ord("y")]
        crop, details = crop_rectangle(
            page,
            Box(100, 100, 300, 124),
            outline_pt=0,
            scale=4,
            traces=page.spans(),
            drawings=page.drawings(),
            math_fonts=["helvetica"],
        )
    assert crop.y1 >= y.bbox.y1 and crop.y1 - y.bbox.y1 < 0.25 and not details["clipped"]


def test_crop_ignores_decoration_drawn_just_before_math():
    # 枠線のない数式の図形の直前に描いた、別の図形の飾りの線（自身の文字と同じ色で、矩形の下に接する）は含めない。
    def draw(page):
        page.insert_text((400, 60), "N", fontsize=20)
        page.draw_rect((100, 126, 300, 130), color=None, fill=(0, 0, 0))
        page.insert_text((112, 120), "x", fontsize=14)

    crop, details = crop_of(draw, (100, 100, 300, 124))
    assert list(crop) == [100, 100, 300, 124] and not details["clipped"]


def test_crop_ignores_neighbor_starting_at_right_edge():
    # 間 0 で右に隣り合う文字枠の最初の文字は、基点が 0.5 pt の余裕の中でも、送り幅の中央が外にあるので含めない。
    def draw(page):
        page.insert_text((35, 80), "x", fontsize=24)
        page.insert_text((230.07, 80), "M", fontsize=24)

    crop, details = crop_of(draw, (30, 40, 230, 110))
    assert list(crop) == [30, 40, 230, 110] and not details["clipped"]


def test_crop_continuation_stops_at_other_math_and_other_font():
    def draw(page):
        page.insert_text((35, 80), "x", fontsize=24)
        page.insert_text((236, 80), "z", fontsize=24)  # ほかの数式の図形の中の文字
        page.insert_text((40, 120), "w", fontsize=24, fontname="tiro")  # 文字の書体の文字

    crop, details = crop_of(draw, (30, 40, 230, 110), math_fonts=["helvetica"], others=[(234, 40, 300, 110)])
    assert list(crop) == [30, 40, 230, 110] and not details["clipped"]
    crop, _ = crop_of(draw, (30, 40, 230, 110), math_fonts=["helvetica"])
    assert crop.x1 > 240


def test_finalize_reads_math_fonts_from_pptx(tmp_path):
    # 数式の書体名は PPTX の数式の文字から読み、はみ出した分母を切り出しに含める。`\text` の文字の書体は使わない。
    shape = math_element(expression=r"x + \text{a}", area=(100, 100, 200, 24), style={"font_math": "Helvetica"})
    native, pdf, final = tmp_path / "native.pptx", tmp_path / "native.pdf", tmp_path / "deck.pptx"
    write_pptx(scene([shape]), project(tmp_path), native)
    assert _math_fonts(xml_part(native)) == ["helvetica"]
    crop = finalize_math(native, pdf_pages(pdf, nested_math), final)["crops"][0]
    assert crop["crop_pt"][3] > 136 and not crop["clipped"]


def test_crop_includes_outline_and_snaps_to_pixels():
    def draw(page):
        page.insert_text((35, 80), "x", fontsize=24)

    crop, details = crop_of(draw, (30, 40, 230, 110), outline_pt=1.0)
    assert list(crop) == [29.25, 39.25, 230.75, 110.75]
    assert details["margin_pt"] == [0.75] * 4 and details["outline_pt"] == 1.0
    crop, _ = crop_of(draw, (30.1, 40.1, 230.1, 109.9))
    assert list(crop) == [30, 40, 230.25, 110]
    assert all((value * 4).is_integer() for value in crop)


def test_crop_is_clipped_to_limit_and_page():
    def far(page):
        page.insert_text((110, 130), "x", fontsize=20)
        page.draw_rect((150, 95, 160, 175), color=None, fill=(0, 0, 0))  # 下へ 35 pt 出る

    crop, details = crop_of(far, (100, 100, 300, 140))
    assert list(crop) == [100, 95, 300, 156]
    assert details["margin_pt"] == [0, 5, 0, 16] and details["clipped"]

    def edge(page):
        page.insert_text((110, 525), "x", fontsize=20)
        page.draw_rect((150, 505, 160, 560), color=None, fill=(0, 0, 0))  # ページの下端を越える

    crop, details = crop_of(edge, (100, 500, 300, 535))
    assert list(crop) == [100, 500, 300, 540] and details["clipped"]


def test_outline_width_comes_from_visible_line():
    def shape(line=""):
        return etree.fromstring(f'<p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><p:spPr>{line}</p:spPr></p:sp>')

    assert outline_width_pt(shape()) == 0
    assert outline_width_pt(shape("<a:ln w='25400'><a:noFill/></a:ln>")) == 0
    assert outline_width_pt(shape("<a:ln w='25400'><a:solidFill/></a:ln>")) == 2
    assert outline_width_pt(shape("<a:ln/>")) == 1


def test_fallback_shape_uses_crop_geometry(tmp_path):
    outlined = math_element(style={"line_color": "17212D", "line_width_pt": 1})
    native, pdf, final = tmp_path / "native.pptx", tmp_path / "native.pdf", tmp_path / "deck.pptx"
    write_pptx(scene([outlined]), project(tmp_path), native)
    result = finalize_math(native, pdf_pages(pdf, own_math), final)
    crop = result["crops"][0]
    assert crop["rectangle_pt"] == [30, 40, 230, 110]
    assert crop["frame_pt"] == [14, 24, 246, 126]
    assert crop["crop_pt"] == [29.25, 39.25, 230.75, 118]
    assert crop["margin_pt"] == [0.75, 0.75, 0.75, 8] and crop["outline_pt"] == 1.0 and not crop["clipped"]
    root = xml_part(final)
    choice = root.find(".//mc:Choice/p:sp/p:spPr/a:xfrm", NS)
    assert [choice.find("a:off", NS).get(key) for key in ("x", "y")] == [str(30 * 12700), str(40 * 12700)]
    assert [choice.find("a:ext", NS).get(key) for key in ("cx", "cy")] == [str(200 * 12700), str(70 * 12700)]
    picture = root.find(".//mc:Fallback/p:sp/p:spPr", NS)
    assert [picture.find("a:xfrm/a:off", NS).get(key) for key in ("x", "y")] == [str(14 * 12700), str(24 * 12700)]
    assert [picture.find("a:xfrm/a:ext", NS).get(key) for key in ("cx", "cy")] == [str(232 * 12700), str(102 * 12700)]
    assert picture.find("a:ln/a:noFill", NS) is not None
    # 画像は図形の範囲を 4 倍で描き、切り出す範囲だけが不透明になる。その画素は PDF の同じ範囲と一致する。
    image = image_of(final, crop["image"])
    assert image.size == (928, 408) and image.mode == "RGBA"
    alpha = image.getchannel("A")
    opaque = alpha.point(lambda value: 255 if value == 255 else 0).getbbox()
    assert opaque == (61, 61, 867, 376)
    with open_pdf(pdf) as document:
        expected = render_region(document[0], 4, crop["crop_pt"])
    assert image.crop(opaque).convert("RGB").tobytes() == expected.tobytes()
    # 切り出す範囲のすぐ外側の 2 画素は透明度を段階的に下げ、その外はほぼ透明（不透明度 ALPHA_FLOOR）にする。
    x0, y0, x1, y1 = opaque
    assert [alpha.getpixel((x, 200)) for x in range(x0 - 3, x0 + 1)] == [ALPHA_FLOOR, 85, 170, 255]
    assert [alpha.getpixel((400, y)) for y in range(y1 - 1, y1 + 3)] == [255, 170, 85, ALPHA_FLOOR]
    assert alpha.point(lambda value: 255 if value > ALPHA_FLOOR else 0).getbbox() == (x0 - 2, y0 - 2, x1 + 2, y1 + 2)
    assert alpha.getextrema() == (ALPHA_FLOOR, 255)
    # 透明な画素の色は、不透明な範囲の端の色を外へ延ばしたもの（黒にすると、拡大・縮小で境界が灰色の線になる）。
    rgb = image.convert("RGB")
    for x, y in ((0, 0), (x0 - 1, 100), (x1, 300), (927, 407), (400, 0), (500, 407), (0, 407), (927, 0)):
        assert rgb.getpixel((x, y)) == rgb.getpixel((min(max(x, x0), x1 - 1), min(max(y, y0), y1 - 1)))


def test_clipped_crop_is_recorded(tmp_path):
    def far(page):
        page.insert_text((110, 130), "x", fontsize=20)
        page.draw_rect((150, 95, 160, 175), color=None, fill=(0, 0, 0))

    native, final = tmp_path / "native.pptx", tmp_path / "deck.pptx"
    write_pptx(scene([math_element(area=(100, 100, 200, 40))]), project(tmp_path), native)
    crop = finalize_math(native, pdf_pages(tmp_path / "native.pdf", far), final)["crops"][0]
    assert crop["clipped"] and crop["crop_pt"] == [100, 95, 300, 156] and crop["margin_pt"] == [0, 5, 0, 16]
    # 超えた部分に掛かるほかの要素はない（数式自身のはみ出し）。
    assert crop["nearby_elements"] == []


# ---------------------------------------------------------------- ほかの要素の描画を切り出しに含めない

RED = (0.75, 0, 0)
# 矢印の線幅（pt）と、PowerPoint の矢印の範囲の余裕（`_REGION_MARGIN_PT` + 線幅の `_LINE_END_FACTOR` 倍）。
ARROW_WIDTH = 1.8
ARROW = {"id": "branch", "box_pt": [183, 245, 228, 272], "segment_pt": [183, 245, 228, 272], "tolerance_pt": 6.4}
# 矢印の近くの数式 u15 の図形の矩形。
LABEL = (140, 250, 188, 280)


def arrow_outline(start, end, width=ARROW_WIDTH):
    """線と先端（幅と長さは線幅の 3 倍。先端は終点）を 1 つの多角形にした矢印の輪郭。PowerPoint は矢印をこの形の塗りで描く。"""
    (x0, y0), (x1, y1) = start, end
    length = math.hypot(x1 - x0, y1 - y0)
    ux, uy = (x1 - x0) / length, (y1 - y0) / length
    nx, ny = -uy, ux
    half, head = width / 2, width * 3
    bx, by = x1 - ux * head, y1 - uy * head
    return [
        (x0 + nx * half, y0 + ny * half),
        (bx + nx * half, by + ny * half),
        (bx + nx * head / 2, by + ny * head / 2),
        (x1, y1),
        (bx - nx * head / 2, by - ny * head / 2),
        (bx - nx * half, by - ny * half),
        (x0 - nx * half, y0 - ny * half),
    ]


def fill_polygon(page, points, fill=RED):
    shape = page.new_shape()
    shape.draw_polyline(points)
    shape.finish(color=None, fill=fill, closePath=True)
    shape.commit()


def branch(page, start=(183, 245), end=(228, 272)):
    """矢印（既定は (183, 245) から (228, 272)）を描いた直後に、矩形 LABEL の数式 u15 を矢印と同じ色で描く（利用者の資料の配置）。

    矢印の始点は数式の矩形の中にあり、終点は矩形の右から 40 pt 離れる（切り出す範囲の上限は 16 pt）。
    """
    fill_polygon(page, arrow_outline(start, end))
    page.insert_text((150, 271), "u", fontsize=20, color=RED)
    page.insert_text((161, 276), "15", fontsize=14, color=RED)


def test_crop_excludes_arrow_drawn_just_before_math():
    # 矢印は数式の文字と同じ色で、数式の直前に 1 つの塗りとして描かれる。数式の線（括弧・根号）と同じ条件に合うので、
    # ほかの要素の範囲を渡さないと矢印まで含め、右の上限で切り詰める。
    crop, details = crop_of(branch, LABEL)
    assert details["clipped"] and details["margin_pt"][2] == 16 and details["nearby_elements"] == []
    # 矢印の範囲を渡すと、塗りの外接矩形が矢印の範囲とほぼ一致するので矢印の描画とみなし、含めない。
    crop, details = crop_of(branch, LABEL, neighbors=[ARROW])
    assert list(crop) == list(LABEL) and not details["clipped"] and details["nearby_elements"] == []


def test_crop_reports_elements_in_clipped_part():
    # 矢印を区別できない（範囲を知らない）まま切り詰めたときは、上限の外の部分に範囲が掛かる要素を示す。
    # 数式を載せた箱（矩形を含む）と、離れた要素は示さない。
    neighbors = [
        {"id": "panel", "box_pt": [100, 200, 400, 350], "tolerance_pt": 1.5},
        {"id": "tag", "box_pt": [206, 240, 240, 262], "tolerance_pt": 1},
        {"id": "far", "box_pt": [400, 400, 500, 450], "tolerance_pt": 1},
    ]
    _, details = crop_of(branch, LABEL, neighbors=neighbors)
    assert details["clipped"] and details["nearby_elements"] == ["tag"]


def test_crop_excludes_separately_drawn_arrow_head():
    # 数式を指す矢印の線と先端を別の塗りで描くと、先端（数式の矩形に掛かる）だけでは範囲と一致しない。
    # 先端が数式の矩形の外へ出ていて、4 隅が矢印の線分から余裕の内にあれば、矢印の描画とみなす。
    def pointing(page):
        page.draw_rect((191.4, 264.1, 260, 265.9), color=None, fill=RED)
        fill_polygon(page, [(186, 265), (191.4, 262.3), (191.4, 267.7)])
        page.insert_text((150, 271), "u", fontsize=20, color=RED)

    arrow = {"id": "pointer", "box_pt": [186, 265, 260, 265], "segment_pt": [260, 265, 186, 265], "tolerance_pt": 6.4}
    crop, _ = crop_of(pointing, LABEL)
    assert crop.x1 > 191
    crop, details = crop_of(pointing, LABEL, neighbors=[arrow])
    assert list(crop) == list(LABEL) and not details["clipped"]
    # 数式の矩形（0.5 pt の余裕を含む）に収まる描画は、線分に近くても数式の線と区別できないので、矢印の描画とみなさない。
    region, probe = _Region(arrow), Box(*LABEL).expanded(0.5)
    assert region.drew(Box(186, 262.3, 191.4, 267.7), probe)
    assert not region.drew(Box(186, 262.3, 188.2, 267.7), probe)


def test_crop_keeps_own_lines_near_other_elements():
    # 数式を載せた箱と、矩形のすぐ下の横線があっても、矩形の下へはみ出した括弧（数式の線）は含める。
    # 括弧の外接矩形は箱の範囲と一致せず、4 隅が横線から離れている。
    neighbors = [
        {"id": "panel", "box_pt": [80, 80, 400, 200], "tolerance_pt": 1.5},
        {"id": "rule", "box_pt": [100, 140, 300, 140], "segment_pt": [100, 140, 300, 140], "tolerance_pt": 4},
    ]
    crop, details = crop_of(nested_math, (100, 100, 300, 124), neighbors=neighbors)
    assert list(crop) == [100, 100, 300, 136] and not details["clipped"]


def arrow_element(identifier, start, end, width=ARROW_WIDTH):
    (x0, y0), (x1, y1) = start, end
    return {
        "id": identifier,
        "kind": "arrow",
        "area": [min(x0, x1), min(y0, y1), abs(x1 - x0), abs(y1 - y0)],
        "start": list(start),
        "end": list(end),
        "style": {"line_color": "C00000", "line_width_pt": width},
    }


def test_finalize_excludes_arrow_from_slide(tmp_path):
    # 最終形の PPTX のスライドから矢印の範囲を読み、切り出しに矢印を含めない。向きが逆の矢印（反転した図形）は、
    # 外接矩形が同じで、線分の始点と終点が入れ替わる。
    label = math_element("branch_label", "u_{15}", area=(140, 250, 48, 30))
    for name, start, end in (("forward", (183, 245), (228, 272)), ("reverse", (228, 272), (183, 245))):
        data = scene([arrow_element("branch", start, end), label])
        (tmp_path / name).mkdir()
        paths, _, final = build_final_form(
            tmp_path / name, data, lambda page, start=start, end=end: branch(page, start, end)
        )
        (crop,) = final["crops"]
        assert crop["crop_pt"] == list(LABEL) and not crop["clipped"] and crop["nearby_elements"] == []
        (shape,) = _wrap_math(read_package(paths["native.pptx"]), ["ppt/slides/slide1.xml"], (720, 540), 4)
        assert shape["neighbors"] == [{**ARROW, "segment_pt": [*start, *end]}]


def test_shape_regions_follow_group_transform_and_rotation():
    p, a = NS["p"], NS["a"]

    def shape(name, x, y, width, height, extra=""):
        return (
            f'<p:sp><p:nvSpPr><p:cNvPr id="1" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>'
            f'<a:xfrm {extra}><a:off x="{x * 12700}" y="{y * 12700}"/><a:ext cx="{width * 12700}" cy="{height * 12700}"/>'
            '</a:xfrm><a:prstGeom prst="rect"/><a:ln w="25400"><a:solidFill/></a:ln></p:spPr></p:sp>'
        )

    tree = etree.fromstring(
        f'<p:spTree xmlns:p="{p}" xmlns:a="{a}"><p:nvGrpSpPr/><p:grpSpPr/>'
        + shape("turned", 100, 100, 40, 20, 'rot="5400000"')
        + '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="2" name="group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm>'
        f'<a:off x="{200 * 12700}" y="{50 * 12700}"/><a:ext cx="{100 * 12700}" cy="{100 * 12700}"/>'
        f'<a:chOff x="0" y="0"/><a:chExt cx="{50 * 12700}" cy="{50 * 12700}"/></a:xfrm></p:grpSpPr>'
        + shape("inner", 10, 10, 20, 10)
        + "</p:grpSp></p:spTree>"
    )
    regions = {region["id"]: region for _, region in _shape_regions(tree)}
    # 90 度回した図形は中心のまわりで縦長になる。グループの子はグループの拡大（2 倍）と位置を当てる。枠線の外側の半分を余裕に含める。
    assert regions == {
        "turned": {"id": "turned", "box_pt": [110, 90, 130, 130], "tolerance_pt": 2},
        "inner": {"id": "inner", "box_pt": [220, 70, 260, 90], "tolerance_pt": 2},
    }


def test_clipped_warning_names_nearby_elements():
    from slidekit.pipeline import _Build

    crop = {
        "slide_id": "first",
        "element_id": "label",
        "rectangle_pt": [140, 250, 188, 280],
        "frame_pt": [124, 234, 204, 296],
        "crop_pt": [140, 244, 204, 280],
        "margin_pt": [0, 6, 16, 0],
        "outline_pt": 0.0,
        "clipped": True,
        "nearby_elements": ["branch"],
    }
    build = SimpleNamespace(scene=scene([math_element()]))
    (warning,) = _Build.clipped_crops(build, [crop])
    assert warning["observations"][0]["nearby_elements"] == ["branch"] and "branch" in warning["message"]
    # 掛かる要素がなければ、観測値は前と同じ項目にする（許容の記録の証拠を変えない）。
    (warning,) = _Build.clipped_crops(build, [{**crop, "nearby_elements": []}])
    assert "nearby_elements" not in warning["observations"][0] and "文字枠を広げる" in warning["message"]


def exported(png: bytes) -> bytes:
    """PowerPoint の PDF の書き出しと同じく、不透明度 0 の画素の色を黒にした PNG。"""
    image = Image.open(io.BytesIO(png)).convert("RGBA")
    black = Image.new("RGBA", image.size, (0, 0, 0, 0))
    black.paste(image, mask=image.getchannel("A").point(lambda value: 255 if value else 0))
    buffer = io.BytesIO()
    black.save(buffer, "PNG")
    return buffer.getvalue()


def ellipse_points(box, count=360):
    """box に内接する楕円を近似する多角形の頂点（pdf_tools は楕円を描かないので、多角形で描く）。"""
    x0, y0, x1, y1 = box
    center_x, center_y, radius_x, radius_y = (x0 + x1) / 2, (y0 + y1) / 2, (x1 - x0) / 2, (y1 - y0) / 2
    return [
        (center_x + radius_x * math.cos(angle), center_y + radius_y * math.sin(angle))
        for angle in (2 * math.pi * number / count for number in range(count))
    ]


def test_crop_boundary_has_no_gray_line_after_export(tmp_path):
    # 白い背景の上の塗りのある楕円（外接矩形の角は背景の白）を切り出し、PowerPoint の書き出しと同じく不透明度 0 の画素を黒にして
    # 白いページに置き、縮小して描く。切り出す範囲の境界（楕円の外接矩形の角の付近）が灰色にならない。
    def oval(page):
        shape = page.new_shape()
        shape.draw_polyline(ellipse_points((60, 50, 240, 150)))
        shape.finish(color=(0.12, 0.15, 0.2), fill=(0.95, 0.96, 0.98), width=3, closePath=True)
        shape.commit()

    crop = Box(58, 48, 242, 152)
    frame = crop.expanded(MAX_MARGIN_PT)
    with open_pdf(pdf_pages(tmp_path / "oval.pdf", oval, size=(300, 200))) as document:
        page = document[0]
        png = _crop_png(page.render(4), page.rect, 4, frame, crop)
    assert Image.open(io.BytesIO(png)).getchannel("A").getextrema()[0] == ALPHA_FLOOR

    def placed(page):
        page.draw_rect(page.rect, color=None, fill=(1, 1, 1))
        page.insert_image(frame, stream=exported(png))

    with open_pdf(pdf_pages(tmp_path / "placed.pdf", placed, size=(300, 200))) as document:
        for scale in (2, 1.2, 1, 0.75):
            image = document[0].render(scale).convert("L")
            corner = image.crop(
                tuple(round(value * scale) for value in (crop.x0 - 6, crop.y0 - 6, crop.x0 + 14, crop.y0 + 14))
            )
            assert corner.getextrema()[0] >= 250, scale


def test_known_input_hashes_are_not_computed_again(tmp_path, monkeypatch):
    from slidekit import fallback

    data = scene([element("intro")], ("formula", [math_element()]))
    native, form, pdf, final, clone = (
        tmp_path / name for name in ("native.pptx", "final_form.pptx", "final_form.pdf", "deck.pptx", "fallback.pptx")
    )
    write_pptx(data, project(tmp_path), native)
    pdf_pages(pdf, None, own_math)
    hashed = []
    original = fallback.file_hash
    monkeypatch.setattr(fallback, "file_hash", lambda path: hashed.append(Path(path).name) or original(path))
    prepared = prepare_final_form(native, form, source_sha256="n" * 64)
    finalized = finalize_math(form, pdf, final, source_pptx_sha256=prepared["sha256"], source_pdf_sha256="p" * 64)
    cloned = prepare_fallback(final, clone, math_only=True, source_sha256=finalized["sha256"])
    # 書いたファイルだけをハッシュする（入力は渡した値を記録する）。
    assert hashed == ["final_form.pptx", "deck.pptx", "fallback.pptx"]
    assert (
        prepared["source_sha256"],
        finalized["source_pptx_sha256"],
        finalized["source_pdf_sha256"],
        cloned["source_sha256"],
    ) == ("n" * 64, prepared["sha256"], "p" * 64, finalized["sha256"])
    # 渡さなければ、これまでどおり入力をハッシュする。
    assert prepare_fallback(final, tmp_path / "again.pptx")["source_sha256"] == original(final)


def build_final_form(tmp_path, data=None, draw=own_math):
    """2 ページ目に数式のある資料の中間 PPTX、最終形の PPTX、その PDF、最終 PPTX を作る。PDF は最後のページだけを draw で描く。"""
    data = data or scene([element("intro")], ("formula", [math_element()]))
    paths = {name: tmp_path / name for name in ("native.pptx", "final_form.pptx", "final_form.pdf", "deck.pptx")}
    write_pptx(data, project(tmp_path), paths["native.pptx"])
    form = prepare_final_form(paths["native.pptx"], paths["final_form.pptx"])
    pdf_pages(paths["final_form.pdf"], *([None] * (len(data["slides"]) - 1)), draw)
    final = finalize_math(paths["final_form.pptx"], paths["final_form.pdf"], paths["deck.pptx"])
    return paths, form, final


def test_final_form_differs_from_final_only_in_fallback_images(tmp_path):
    paths, form, final = build_final_form(tmp_path)
    native_before = paths["native.pptx"].read_bytes()
    assert form["fallback_shapes"] == 1 and form["shapes"][0]["frame_pt"] == [14, 24, 246, 126]
    assert final["input_form"] == "final_form" and final["source_pptx_sha256"] == form["sha256"]
    first, second = read_package(paths["final_form.pptx"]), read_package(paths["deck.pptx"])
    images = {crop["image"] for crop in final["crops"]}
    assert images == {shape["image"] for shape in form["shapes"]} == {"ppt/media/slidekit_math_0002_0001.png"}
    assert set(first) == set(second)
    assert [name for name in first if first[name] != second[name]] == sorted(images)
    assert compare_package_structure(paths["final_form.pptx"], paths["deck.pptx"]) == {
        "matched": True,
        "parts": len(first),
        "fallback_images": 1,
        "differences": [],
    }
    # 仮の画像は最終の互換画像と同じ寸法の 1 色の PNG である。
    placeholder, image = image_of(paths["final_form.pptx"], *images), image_of(paths["deck.pptx"], *images)
    assert placeholder.size == image.size == (928, 408)
    assert placeholder.convert("RGB").getcolors() == [(928 * 408, (255, 0, 255))]

    # mc:Choice の図形は中間 PPTX の数式の図形のままである。
    def canonical(node):
        return etree.tostring(node, method="c14n", exclusive=True)

    native_shape = xml_part(paths["native.pptx"], "ppt/slides/slide2.xml").find("p:cSld/p:spTree/p:sp", NS)
    choice_shape = xml_part(paths["deck.pptx"], "ppt/slides/slide2.xml").find(".//mc:Choice/p:sp", NS)
    assert canonical(native_shape) == canonical(choice_shape)
    assert paths["native.pptx"].read_bytes() == native_before
    # 互換画像のないページの部品は、中間 PPTX と同じである。
    native = read_package(paths["native.pptx"])
    assert second["ppt/slides/slide1.xml"] == native["ppt/slides/slide1.xml"]


def test_native_input_gives_same_final_as_final_form(tmp_path):
    paths, _, final = build_final_form(tmp_path)
    direct = tmp_path / "direct.pptx"
    result = finalize_math(paths["native.pptx"], paths["final_form.pdf"], direct)
    assert result["input_form"] == "native"
    assert read_package(direct) == read_package(paths["deck.pptx"])
    assert [crop["image_sha256"] for crop in result["crops"]] == [crop["image_sha256"] for crop in final["crops"]]
    with pytest.raises(ValueError, match="中間 PPTX"):
        prepare_final_form(paths["final_form.pptx"], tmp_path / "again.pptx")


def test_compare_package_structure_reports_changed_parts(tmp_path):
    photo = tmp_path / "photo.png"
    Image.new("RGB", (40, 30), "#3366aa").save(photo)
    picture = {
        "id": "photo",
        "kind": "image",
        "area": [300, 200, 100, 80],
        "path": str(photo),
        "fit": "contain",
        "style": {},
    }
    paths, _, final = build_final_form(tmp_path, scene([element("intro"), picture], ("formula", [math_element()])))
    image = final["crops"][0]["image"]
    parts = read_package(paths["deck.pptx"])
    user_media = next(name for name in parts if name.startswith("ppt/media/") and name != image)

    def changed(name, **replace):
        path = tmp_path / name
        write_package({key: value for key, value in {**parts, **replace}.items() if value is not None}, path)
        return compare_package_structure(paths["final_form.pptx"], path)

    assert changed("fallback.pptx", **{image: parts[user_media]})["matched"]
    slide = parts["ppt/slides/slide2.xml"].replace(b'Requires="a14"', b'Requires="a14 "')
    assert changed("slide.pptx", **{"ppt/slides/slide2.xml": slide})["differences"] == [
        {"part": "ppt/slides/slide2.xml", "reason": "content"}
    ]
    assert changed("user.pptx", **{user_media: parts[image]})["differences"] == [
        {"part": user_media, "reason": "content"}
    ]
    assert changed("missing.pptx", **{image: None})["differences"] == [{"part": image, "reason": "only_in_first"}]


class PickledWorkers:
    """関数と引数と結果を pickle で受け渡し、子プロセスと同じ条件で順に実行する。"""

    def __init__(self):
        self.calls = []

    def map(self, function, arguments):
        function, arguments = pickle.loads(pickle.dumps((function, list(arguments))))
        self.calls.append([item[1] for item in arguments])
        return pickle.loads(pickle.dumps([function(*item) for item in arguments]))


def test_workers_give_same_images_as_serial(tmp_path):
    data = scene(
        [math_element("a")], ("plain", [element("b")]), ("second", [math_element("c", area=(100, 300, 300, 60))])
    )
    native, pdf = tmp_path / "native.pptx", tmp_path / "native.pdf"
    write_pptx(data, project(tmp_path), native)

    def second(page):
        page.insert_text((110, 340), "y", fontsize=24)

    pdf_pages(pdf, own_math, None, second)
    serial = finalize_math(native, pdf, tmp_path / "serial.pptx")
    workers = PickledWorkers()
    parallel = finalize_math(native, pdf, tmp_path / "parallel.pptx", workers=workers)
    assert workers.calls == [[0, 2]]
    assert parallel["crops"] == serial["crops"]
    assert read_package(tmp_path / "parallel.pptx") == read_package(tmp_path / "serial.pptx")


def picture_deck(tmp_path, count=3):
    """ページごとに別の画像とノートを持つ PPTX の部品。"""
    presentation = Presentation()
    for number in range(count):
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        picture = tmp_path / f"picture{number}.png"
        Image.new("RGB", (20, 10), (40 * number, 80, 120)).save(picture)
        slide.shapes.add_picture(str(picture), Pt(10), Pt(10))
        slide.notes_slide.notes_text_frame.text = f"ノート {number}"
    path = tmp_path / "pictures.pptx"
    presentation.save(path)
    return read_package(path)


def assert_package_is_consistent(parts):
    from slidekit.fallback import _part_target, _relationships_source

    for name, data in parts.items():
        if name.endswith(".rels"):
            for relationship in etree.fromstring(data):
                if relationship.get("TargetMode") != "External":
                    assert _part_target(_relationships_source(name), relationship.get("Target")) in parts, (
                        name,
                        relationship.get("Target"),
                    )
    overrides = [
        node.get("PartName").lstrip("/")
        for node in etree.fromstring(parts["[Content_Types].xml"])
        if node.tag.endswith("Override")
    ]
    assert set(overrides) <= set(parts)


def slide_media(parts, slide):
    from slidekit.fallback import _part_target

    relationships = etree.fromstring(parts[slide.replace("slides/", "slides/_rels/") + ".rels"])
    return {_part_target(slide, node.get("Target")) for node in relationships if node.get("Type").endswith("/image")}


def test_keep_slides_removes_unreachable_parts(tmp_path):
    parts = picture_deck(tmp_path)
    before = dict(parts)
    kept = keep_slides(parts, ["ppt/slides/slide2.xml"])
    assert parts == before and len(parts) > len(kept)
    for name in (
        "ppt/slides/slide1.xml",
        "ppt/slides/slide3.xml",
        "ppt/slides/_rels/slide1.xml.rels",
        "ppt/notesSlides/notesSlide1.xml",
        "ppt/notesSlides/notesSlide3.xml",
    ):
        assert name in parts and name not in kept
    assert len([name for name in parts if name.startswith("ppt/media/")]) == 3
    assert {name for name in kept if name.startswith("ppt/media/")} == slide_media(parts, "ppt/slides/slide2.xml")
    overrides = etree.fromstring(kept["[Content_Types].xml"])
    assert not [
        node
        for node in overrides
        if "slide1.xml" in node.get("PartName", "") or "slide3.xml" in node.get("PartName", "")
    ]
    assert_package_is_consistent(kept)
    path = tmp_path / "kept.pptx"
    write_package(kept, path)
    presentation = Presentation(path)
    assert len(presentation.slides) == 1
    assert presentation.slides[0].notes_slide.notes_text_frame.text == "ノート 1"


def test_keep_slides_removes_section_and_custom_show_references(tmp_path):
    parts = picture_deck(tmp_path)
    presentation = etree.fromstring(parts["ppt/presentation.xml"])
    identifiers = [
        (node.get("id"), node.get(f"{{{NS['r']}}}id")) for node in presentation.findall("p:sldIdLst/p:sldId", NS)
    ]
    p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
    shows = etree.fromstring(
        f'<p:custShowLst xmlns:p="{NS["p"]}" xmlns:r="{NS["r"]}"><p:custShow name="短縮" id="0"><p:sldLst>'
        + "".join(f'<p:sld r:id="{relation}"/>' for _, relation in identifiers)
        + "</p:sldLst></p:custShow></p:custShowLst>"
    )
    presentation.find("p:notesSz", NS).addnext(shows)
    sections = etree.fromstring(
        f'<p:extLst xmlns:p="{NS["p"]}"><p:ext uri="{{521415D9-36F7-43E2-AB2F-B90AF26B5E84}}"><p14:sectionLst xmlns:p14="{p14}">'
        f'<p14:section name="前半" id="{{11111111-1111-1111-1111-111111111111}}"><p14:sldIdLst>'
        f'<p14:sldId id="{identifiers[0][0]}"/><p14:sldId id="{identifiers[1][0]}"/></p14:sldIdLst></p14:section>'
        f'<p14:section name="後半" id="{{22222222-2222-2222-2222-222222222222}}"><p14:sldIdLst>'
        f'<p14:sldId id="{identifiers[2][0]}"/></p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst>'
    )
    if presentation.find("p:extLst", NS) is None:
        presentation.append(sections)
    else:
        presentation.find("p:extLst", NS).append(sections[0])
    parts["ppt/presentation.xml"] = etree.tostring(
        presentation, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    kept = keep_slides(parts, ["ppt/slides/slide2.xml"])
    result = etree.fromstring(kept["ppt/presentation.xml"])
    assert [node.get("id") for node in result.findall("p:sldIdLst/p:sldId", NS)] == [identifiers[1][0]]
    assert [
        [node.get("id") for node in section.iter(f"{{{p14}}}sldId")] for section in result.iter(f"{{{p14}}}section")
    ] == [[identifiers[1][0]], []]
    assert [node.get(f"{{{NS['r']}}}id") for node in result.iter(f"{{{NS['p']}}}sld")] == [identifiers[1][1]]
    relationships = {node.get("Id") for node in etree.fromstring(kept["ppt/_rels/presentation.xml.rels"])}
    assert identifiers[1][1] in relationships and not {identifiers[0][1], identifiers[2][1]} & relationships
    assert_package_is_consistent(kept)
    path = tmp_path / "kept.pptx"
    write_package(kept, path)
    assert len(Presentation(path).slides) == 1


def test_fallback_clone_with_math_slides_only(tmp_path):
    data = scene([element("intro")], ("formula", [math_element()]), ("summary", [element("closing")]))
    native, pdf, final = tmp_path / "native.pptx", tmp_path / "native.pdf", tmp_path / "deck.pptx"
    write_pptx(data, project(tmp_path), native)
    finalize_math(native, pdf_pages(pdf, None, own_math, None), final)
    result = prepare_fallback(final, tmp_path / "fallback.pptx", math_only=True)
    assert result["slides"] == [{"number": 2, "id": "formula"}]
    assert result["kept_slides"] == [{"number": 1, "id": "formula"}]
    clone = read_package(tmp_path / "fallback.pptx")
    assert_package_is_consistent(clone)
    presentation = Presentation(tmp_path / "fallback.pptx")
    assert [slide.element.find("p:cSld", NS).get("name") for slide in presentation.slides] == ["formula"]
    assert not presentation.slides[0].element.findall(".//m:oMath", NS)
    assert presentation.slides[0].element.findall(".//a:blip", NS)
    full = prepare_fallback(final, tmp_path / "full.pptx")
    assert full["kept_slides"] == [
        {"number": 1, "id": "first"},
        {"number": 2, "id": "formula"},
        {"number": 3, "id": "summary"},
    ]
    assert len(Presentation(tmp_path / "full.pptx").slides) == 3
    plain = tmp_path / "plain.pptx"
    write_pptx(scene([element("intro")]), project(tmp_path), plain)
    with pytest.raises(ValueError, match="数式のページがない"):
        prepare_fallback(plain, tmp_path / "empty.pptx", math_only=True)


def test_fallback_required_follows_settings_and_math():
    with_math = scene([math_element()])
    without_math = scene([element(), {"id": "old", "kind": "text", "content": "x_i"}])

    def settings(fallback, renderers=("powerpoint",)):
        return SimpleNamespace(config={"math": {"fallback": fallback}, "render": {"required": list(renderers)}})

    assert fallback_required(with_math, settings("required"))
    assert fallback_required(with_math, settings("auto", ("powerpoint", "libreoffice")))
    assert not fallback_required(with_math, settings("auto"))
    assert not fallback_required(with_math, settings("off", ("libreoffice",)))
    assert not fallback_required(without_math, settings("required"))
    assert not fallback_required(with_math, SimpleNamespace(config={}))


def test_math_in_group_is_rejected(tmp_path):
    native = tmp_path / "native.pptx"
    write_pptx(scene([math_element()]), project(tmp_path), native)
    parts = read_package(native)
    root = etree.fromstring(parts["ppt/slides/slide1.xml"])
    shape = root.find("p:cSld/p:spTree/p:sp", NS)
    group = etree.SubElement(root.find("p:cSld/p:spTree", NS), f"{{{NS['p']}}}grpSp")
    group.append(shape)
    parts["ppt/slides/slide1.xml"] = etree.tostring(root)
    grouped = tmp_path / "grouped.pptx"
    write_package(parts, grouped)
    with pytest.raises(ValueError, match="スライドの直下"):
        prepare_final_form(grouped, tmp_path / "final_form.pptx")
    with pytest.raises(ValueError, match="スライドの直下"):
        finalize_math(grouped, pdf_pages(tmp_path / "grouped.pdf", own_math), tmp_path / "deck.pptx")


# ---- 吹き出し ----


def callout_element(identifier="note", area=(100, 100, 160, 60), side="left", content=None, **style):
    """しっぽのある box（作成 API と同じく outline.tail_outline から作る）。"""
    from slidekit.outline import tail_outline

    outline = tail_outline(area, "rectangle", side=side)
    return element(
        identifier,
        kind="box",
        area=area,
        content=content,
        style=style,
        shape="rectangle",
        actual_area=list(outline.actual_area),
        outside_parts=[{"part": name, "area": list(rect)} for name, rect in outline.outside],
        tail={"side": side, "at": outline.at, "width_pt": 10.0, "length_pt": 8.0},
    )


def test_shape_region_of_callout_includes_the_tail(tmp_path):
    path = tmp_path / "deck.pptx"
    write_pptx(
        scene([callout_element(content=[{"kind": "text", "text": "補足"}], line_color="183153")]),
        project(tmp_path),
        path,
    )
    tree = xml_part(path).find("p:cSld/p:spTree", NS)
    [(_node, region)] = _shape_regions(tree)
    # 既定図形でない図形は a:xfrm の外接矩形（しっぽまでの範囲）。枠線の外側の半分を余裕に含める。
    assert region == {"id": "note", "box_pt": [92, 100, 260, 160], "tolerance_pt": 1.5}


def test_callout_crop_uses_body_for_own_glyphs_and_keeps_the_tail(tmp_path):
    note = callout_element(content=[{"kind": "math", "text": "x"}])

    def draw(page):
        # しっぽの側の帯（本体の左、a:xfrm の中）に基点がある隣の文字。下へ本体の外まで広がる。
        page.insert_text((93, 158), "N", fontsize=40)
        page.insert_text((110, 140), "x", fontsize=24)

    data = scene([element("intro")], ("formula", [element("beside", area=(40, 120, 52, 40)), note]))
    paths, form, final = build_final_form(tmp_path, data, draw)
    (shape,) = form["shapes"]
    assert shape["rectangle_pt"] == [92, 100, 260, 160]
    (crop,) = final["crops"]
    # 切り出しはしっぽまでの範囲（a:xfrm）を含み、隣の文字を吹き出しの文字として含めない。
    assert crop["crop_pt"] == [92, 100, 260, 160] and not crop["clipped"]
    # 中間 PPTX から切り出しても同じ扱いになる。
    again = finalize_math(paths["native.pptx"], paths["final_form.pdf"], tmp_path / "again.pptx")
    assert again["input_form"] == "native" and again["crops"][0]["crop_pt"] == crop["crop_pt"]
