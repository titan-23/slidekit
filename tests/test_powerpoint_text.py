"""SLIDEKIT_TEST_POWERPOINT=1 を指定したときだけ、文字の書き出しを実際の PowerPoint の変換で確かめる。

確かめること：リンクの文字が文字の色で表示されること、1 段落の中の断片ごとの大きさと太字、数式だけの段落の揃え
（独立した数式と文中の数式）。資料は作成 API で構築し（フィクスチャ windows_fonts_environment で Windows の既定フォントを使う）、
計測と同じフォントを渡して write_pptx で書き出し、PowerPoint で PDF にする。直列に実行する（`-p no:xdist`）。
"""

import os
from pathlib import Path

import pytest

pytestmark = [
    pytest.mark.powerpoint,
    pytest.mark.skipif(
        os.environ.get("SLIDEKIT_TEST_POWERPOINT") != "1", reason="実機 PowerPoint を使う任意の検査である。"
    ),
]

LINK = "https://example.com/guide"
# 数式の揃えの枠：幅 300 pt・高さ 60 pt・余白 0 の枠。式の字形の左右の余白を補正して、揃えの基準から 0.75 pt 以内を合格とする。
ALIGNS = ("left", "center", "right")
EXPRESSIONS = {"0123": ("0", "3"), "x+1": ("\U0001d465", "1")}  # 先頭と末尾の字形（英字は数式用の斜体で描かれる）
SIZES = (20, 32)
TOLERANCE_PT = 0.75
SCALE = 8


def build(root, write_deck, slides):
    """資料を構築して書き出し、PowerPoint で変換する。戻り値は（場面、資料、PDF のパス）。PPTX の検査も通ることを確かめる。"""
    from slidekit import renderers
    from slidekit.config import load_project
    from slidekit.fonts import FontRegistry
    from slidekit.pipeline import construct
    from slidekit.pptx_validation import validate_pptx
    from slidekit.writer import write_pptx

    write_deck(root, slides=slides)
    project = load_project(root)
    scene = construct(project)
    pptx, pdf = root / "deck.pptx", root / "deck.pdf"
    with FontRegistry(project) as fonts:
        write_pptx(scene, project, pptx, fonts=fonts)
        assert [
            issue for issue in validate_pptx(pptx, scene, project, fonts=fonts) if issue["severity"] == "error"
        ] == []
    renderers.render_pdf(pptx, pdf, {"fonts": {}})
    return scene, project, pdf


def characters(page):
    """PDF のページ（`slidekit.pdf` のページ）の見える文字を (文字, 文字サイズ, 書体名, 色, 外接矩形) の列で返す。

    色は 0xRRGGBB の整数、外接矩形は `Box`（横は送り幅、縦はアセンダー・ディセンダー）。
    """
    from slidekit.pdf import Box

    result = []
    for span in page.text_spans():
        color = sum(round(value * 255) << shift for value, shift in zip(span.color, (16, 8, 0)))
        for char in span.chars:
            if chr(char.unicode).strip():
                result.append((chr(char.unicode), span.size, span.font, color, Box.of(char.bbox)))
    return result


def is_bold(font):
    """太字の書体か（PowerPoint の PDF の書体名は Arial-BoldMT・YuGothicUI-Bold のように Bold を含む）。"""
    return "bold" in font.lower()


def element_style(scene, identifier):
    return next(element for slide in scene["slides"] for element in slide["elements"] if element["id"] == identifier)[
        "style"
    ]


def test_real_powerpoint_link_uses_run_color(tmp_path, write_deck, windows_fonts_environment):
    from pdf_tools import links, open_pdf

    slides = f"""from slidekit import Deck, Rect, Text

def build(context):
    deck = Deck(context)
    slide = deck.slide(id="link", kind="blank")
    slide.text(id="line", area=Rect(40, 40, 800, 60), size_pt=28, text=[
        Text("詳細は "), Text("利用ガイド", color="183153", underline=True, link="{LINK}"),
        Text(" と "), Text("手順書", color="A23B2A", link="{LINK}#steps")])
    return deck
"""
    scene, _, pdf = build(tmp_path / "deck", write_deck, slides)
    text_color = int(element_style(scene, "line")["color"], 16)
    assert set(links(pdf, 0)) >= {LINK, LINK + "#steps"}
    with open_pdf(pdf) as document:
        page = document[0]
        found = characters(page)
        # hlinkClr="tx" の効果で、リンクの文字はテーマのリンクの色でなく、断片の色で描かれる。
        for text, color in (("利用ガイド", 0x183153), ("手順書", 0xA23B2A)):
            colors = {item[3] for item in found if item[0] in text}
            assert colors == {color}, (text, colors)
        assert {item[3] for item in found if item[0] in "詳細はと"} == {text_color}
        # 下線の有無を記録する（underline=True の「利用ガイド」と、指定しない「手順書」）。
        drawings = page.drawings()
        for text in ("利用ガイド", "手順書"):
            boxes = [item[4] for item in found if item[0] in text]
            area = boxes[0]
            for box in boxes[1:]:
                area |= box
            lines = [
                drawing
                for drawing in drawings
                if drawing.rect.height <= 3
                and drawing.rect.x0 < area.x1
                and drawing.rect.x1 > area.x0
                and area.y0 + area.height / 2 < drawing.rect.y0 < area.y1 + 6
            ]
            print(f"リンクの下線（{text}）: {'あり' if lines else 'なし'}")


def test_real_powerpoint_fragment_sizes_and_bold(tmp_path, write_deck, windows_fonts_environment):
    from pdf_tools import open_pdf

    # 断片ごとに、ほかの断片と共有しない文字を使う（PDF の文字から断片を特定するため）。
    fragments = [
        ("ちいさい", 14, False),
        ("ふつう", 20, False),
        ("おおきな", 28, False),
        ("太字", 20, True),
        ("Tiny", 14, False),
        ("BOLD", 28, True),
    ]
    listed = ", ".join(
        f'Text("{text}", size_pt={size}' + (", bold=True)" if bold else ")") for text, size, bold in fragments
    )
    slides = f"""from slidekit import Deck, Rect, Text

def build(context):
    deck = Deck(context)
    slide = deck.slide(id="sizes", kind="blank")
    slide.text(id="mixed", area=Rect(40, 40, 880, 80), size_pt=20, text=[{listed}])
    return deck
"""
    _, _, pdf = build(tmp_path / "deck", write_deck, slides)
    with open_pdf(pdf) as document:
        found = characters(document[0])
    for text, size, bold in fragments:
        chars = [item for item in found if item[0] in text]
        assert len(chars) == len(text), (text, chars)
        # PDF の文字サイズは丸められる（20 pt は 20.04 など）。
        assert all(abs(item[1] - size) <= 0.15 for item in chars), (text, [item[1] for item in chars])
        assert {is_bold(item[2]) for item in chars} == {bold}, (text, {item[2] for item in chars})


def bearings(path, index, characters):
    """数式の書体の字形の左右の余白（em）。"""
    from fontTools.ttLib import TTFont

    with TTFont(path, fontNumber=index if Path(path).suffix.lower() in {".ttc", ".otc"} else -1) as face:
        units = face["head"].unitsPerEm
        cmap = face.getBestCmap()
        result = {}
        for character in characters:
            name = cmap[ord(character)]
            glyph = face["glyf"][name]
            result[character] = (glyph.xMin / units, (face["hmtx"][name][0] - glyph.xMax) / units)
        return result


def ink_box(rendered, rect):
    """rect（pt）の中で背景と違う画素の範囲 (左, 上, 右, 下)（pt）。rendered はページを 8 px/pt で画像化した画像。"""
    from PIL import Image, ImageChops

    from slidekit.pdf import Box

    left, top, right, bottom = Box.of(rect).pixels(SCALE)
    image = rendered.crop((left, top, right, bottom))
    background = max(image.getcolors(maxcolors=1 << 24))[1]
    difference = ImageChops.difference(image, Image.new("RGB", image.size, background)).convert("L")
    box = difference.point(lambda value: 255 if value > 40 else 0).getbbox()
    if box is None:
        return None
    return tuple((origin + value) / SCALE for origin, value in zip((left, top, left, top), box))


def test_real_powerpoint_math_alignment(tmp_path, write_deck, windows_fonts_environment):
    from pdf_tools import open_pdf
    from slidekit.fonts import default_font

    frames = []
    for page, kind in enumerate(("equation", "text")):
        for row, (expression, size) in enumerate((expression, size) for expression in EXPRESSIONS for size in SIZES):
            for column, align in enumerate(ALIGNS):
                identifier = f"{kind}_{'digits' if expression == '0123' else 'sum'}_{size}_{align}"
                frames.append(
                    {
                        "id": identifier,
                        "page": page,
                        "kind": kind,
                        "expression": expression,
                        "size": size,
                        "align": align,
                        "x": 20 + 310 * column,
                        "y": 30 + 120 * row,
                    }
                )
    calls = []
    for frame in frames:
        area = f'Rect({frame["x"]}, {frame["y"]}, 300, 60)'
        common = f'size_pt={frame["size"]}, align="{frame["align"]}", padding_pt=0'
        if frame["kind"] == "equation":
            calls.append(
                f'    slides[{frame["page"]}].equation(id="{frame["id"]}", area={area}, '
                f'expression="{frame["expression"]}", {common})'
            )
        else:
            calls.append(
                f'    slides[{frame["page"]}].text(id="{frame["id"]}", area={area}, '
                f'text=[Math("{frame["expression"]}")], {common})'
            )
    body = "\n".join(calls)
    slides = f"""from slidekit import Deck, Math, Rect

def build(context):
    deck = Deck(context)
    slides = [deck.slide(id="display", kind="blank"), deck.slide(id="inline", kind="blank")]
{body}
    return deck
"""
    scene, project, pdf = build(tmp_path / "deck", write_deck, slides)
    font = default_font(element_style(scene, frames[0]["id"])["font_math"], project.environment)
    margins = bearings(font["path"], font["index"], {char for pair in EXPRESSIONS.values() for char in pair})
    results = []
    with open_pdf(pdf) as document:
        for number, page in enumerate(document):
            rendered = page.render(SCALE)
            for frame in (frame for frame in frames if frame["page"] == number):
                x, y, size = frame["x"], frame["y"], frame["size"]
                box = ink_box(rendered, (x - 4, y - 2, x + 304, y + 62))
                assert box is not None, frame["id"]
                first, last = EXPRESSIONS[frame["expression"]]
                left = box[0] - margins[first][0] * size
                right = box[2] + margins[last][1] * size
                offset = {"left": left - x, "center": (left + right) / 2 - (x + 150), "right": (x + 300) - right}[
                    frame["align"]
                ]
                results.append((frame["id"], round(offset, 3)))
    print("数式の揃えのずれ（pt）:", results)
    assert len(results) == len(frames)
    assert [item for item in results if abs(item[1]) > TOLERANCE_PT] == []
