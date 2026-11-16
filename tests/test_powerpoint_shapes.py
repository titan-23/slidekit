"""SLIDEKIT_TEST_POWERPOINT=1 を指定したときだけ、まとまり・接続・図の書き出しを実際の PowerPoint の変換で確かめる。

確かめること：anchor() で結んだ線（PowerPoint の接続）の端点が書いた位置のまま描かれること、PowerPoint のグループにした
まとまりが展開した場合と同じ位置に描かれること、SVG でも埋め込んだ画像が線や面として描かれ、PNG の場合と同じ範囲に
置かれること（PDF の素材は PNG だけにし、vector を指定すると誤りになること）。資料は作成 API で構築し（フィクスチャ
windows_fonts_environment で Windows の既定フォントを使う）、計測と同じフォントを渡して write_pptx で書き出し、
PowerPoint で PDF にする。直列に実行する（`-p no:xdist`）。

場面の pptx_group と画像の vector は、構築した場面に直接書く（作成 API の判定によらず、両方の書き方を比べるため）。
"""

import copy
import os

from pdf_tools import Document, image_boxes, line_segments, open_pdf
import pytest

pytestmark = [
    pytest.mark.powerpoint,
    pytest.mark.skipif(
        os.environ.get("SLIDEKIT_TEST_POWERPOINT") != "1", reason="実機 PowerPoint を使う任意の検査である。"
    ),
]

# 線の端点の許容幅（pt）と、グループの有無で比べる位置の許容幅（pt）。
LINE_TOLERANCE_PT = 0.6
GROUP_TOLERANCE_PT = 0.2
# 接続先の図形。名前、作る式（slides.py の中の式）。
TARGETS = {
    "rect": 'slide.box(id="rect", area=Rect(X, Y, 120, 44), text="長方形", line_color="183153")',
    "round": 'slide.box(id="round", area=Rect(X, Y, 120, 44), shape="rounded_rectangle", text="角丸", line_color="183153")',
    "oval": 'slide.box(id="oval", area=Rect(X, Y, 120, 44), shape="ellipse", text="楕円", line_color="183153")',
    "diamond": 'slide.box(id="diamond", area=Rect(X, Y, 120, 44), shape="diamond", text="ひし形", line_color="183153")',
    "label": 'slide.text(id="label", area=Rect(X, Y, 120, 44), text="文字枠")',
    "formula": 'slide.text(id="formula", area=Rect(X, Y, 120, 44), text=[Text("式 "), Math("x+1")])',
    "picture": 'slide.image(id="picture", path="picture.png", area=Rect(X, Y, 120, 44), fit="cover")',
    "turned": 'slide.box(id="turned", area=Rect(X, Y, 120, 44), text="回転", line_color="183153", rotation=90)',
}


def line_color(index):
    """線ごとに別の色（PDF の描画を線と対応づけるため）。"""
    return f"{(37 * index + 40) % 200:02X}{(91 * index + 20) % 200:02X}{(53 * index + 90) % 200:02X}"


def construct_scene(root, write_deck, slides):
    from slidekit.config import load_project
    from slidekit.pipeline import construct

    write_deck(root, slides=slides)
    project = load_project(root)
    return project, construct(project)


def write_and_render(project, scene, name, *, assets=False):
    """場面を書き出して PowerPoint で PDF にする。PPTX の検査にエラーがないことも確かめる。戻り値は（書き出しの記録、PDF）。"""
    from slidekit import renderers
    from slidekit.assets import prepare_assets
    from slidekit.fonts import FontRegistry
    from slidekit.pptx_validation import validate_pptx
    from slidekit.writer import write_pptx

    root = project.root
    if assets:
        prepare_assets(scene, project, root / f"{name}_assets")
    pptx, pdf = root / f"{name}.pptx", root / f"{name}.pdf"
    with FontRegistry(project) as fonts:
        result = write_pptx(scene, project, pptx, fonts=fonts)
        errors = [issue for issue in validate_pptx(pptx, scene, project, fonts=fonts) if issue["severity"] == "error"]
    assert errors == []
    renderers.render_pdf(pptx, pdf, "powerpoint", {"fonts": {}})
    return result, pdf


def elements(scene):
    return {element["id"]: element for slide in scene["slides"] for element in slide["elements"]}


def test_real_powerpoint_keeps_connected_line_positions(tmp_path, write_deck, windows_fonts_environment):
    from PIL import Image

    root = tmp_path / "deck"
    root.mkdir()
    Image.new("RGB", (240, 88), (200, 220, 240)).save(root / "picture.png")
    lines = []
    for index, (name, creator) in enumerate(TARGETS.items()):
        x, y = (300 if index < 4 else 720), 50 + 120 * (index % 4)
        lines.append(f"    target = {creator.replace('X', str(x)).replace('Y', str(y))}")
        # 左の辺の中央（gap 0、at 0.5）へ結ぶ線と、下の辺の左寄り（gap 6、at 0.25）から出る線。
        lines.append(
            f'    slide.line(id="{name}_left", start=({x - 180}, {y + 10}), end=target.anchor("left"), '
            f'line_color="{line_color(2 * index)}", line_width_pt=1.5)'
        )
        lines.append(
            f'    slide.line(id="{name}_bottom", start=target.anchor("bottom", gap=6, at=0.25), '
            f'end=({x - 60}, {y + 100}), line_color="{line_color(2 * index + 1)}", line_width_pt=1.5)'
        )
    slides = (
        "from slidekit import Deck, Math, Rect, Text\n\ndef build(context):\n    deck = Deck(context)\n"
        '    slide = deck.slide(id="links", kind="blank")\n' + "\n".join(lines) + "\n    return deck\n"
    )
    project, scene = construct_scene(root, write_deck, slides)
    result, pdf = write_and_render(project, scene, "links", assets=True)
    # 回転した図形を含め、すべての端が接続になる。
    assert result["connections"] == 2 * len(TARGETS)
    found = elements(scene)
    lines = line_segments(pdf, 0)
    for index, name in enumerate(f"{target}_{end}" for target in TARGETS for end in ("left", "bottom")):
        element = found[name]
        color = line_color(index)
        expected = tuple(int(color[i : i + 2], 16) / 255 for i in (0, 2, 4))
        segments = [
            (start, end) for drawn, start, end in lines if all(abs(a - b) <= 2 / 255 for a, b in zip(drawn, expected))
        ]
        assert len(segments) == 1, (name, segments)
        points = sorted(segments[0])
        # PowerPoint は接続した線を開いたときに接続点へ引き直さず、書いた端点のまま描く。
        for actual, wanted in zip(points, sorted([tuple(element["start"]), tuple(element["end"])])):
            assert (
                abs(actual[0] - wanted[0]) <= LINE_TOLERANCE_PT and abs(actual[1] - wanted[1]) <= LINE_TOLERANCE_PT
            ), (name, actual, wanted)


GROUP_SLIDES = """from slidekit import Deck, Rect

def build(context):
    deck = Deck(context)
    slide = deck.slide(id="groups", kind="blank")
    slide.box(id="before", area=Rect(40, 40, 120, 40), text="前", fill="DDE7F0")
    outer = slide.group(id="outer", area=Rect(80, 120, 600, 300))
    outer.box(id="first", area=Rect(0, 0, 160, 60), text="外側の図形", fill="F4E3C1", line_color="8A5A00")
    outer.arrow(id="link", start=(160, 30), end=(240, 30), line_color="183153")
    inner = outer.group(id="inner", area=Rect(240, 0, 360, 200))
    inner.box(id="second", area=Rect(0, 0, 160, 60), shape="ellipse", text="内側の楕円", fill="E2F0D9")
    inner.text(id="note", area=Rect(0, 100, 300, 60), text="内側のまとまりの説明文")
    inner.box(id="turned", area=Rect(200, 0, 120, 60), text="回転", rotation=30, line_color="183153", dashed=True)
    # グループの最後の子の吹き出し。しっぽ（右へ 8 pt）がグループの範囲を広げる。
    outer.box(id="talk", area=Rect(440, 230, 152, 50), text="吹き出し", line_color="183153", tail_side="right")
    slide.box(id="after", area=Rect(40, 460, 120, 40), text="後", fill="DDE7F0")
    return deck
"""


def page_marks(pdf):
    """ページの描画（範囲・塗り・線の色）と見える文字（文字・外接矩形）を、比べやすい順に並べて返す。"""
    with open_pdf(pdf) as document:
        page = document[0]
        drawings = sorted(
            (tuple(drawing.rect), tuple(drawing.fill or ()), tuple(drawing.color or ())) for drawing in page.drawings()
        )
        characters = sorted(
            (chr(char.unicode), tuple(char.bbox))
            for span in page.text_spans()
            for char in span.chars
            if chr(char.unicode).strip()
        )
    return drawings, characters


def test_real_powerpoint_renders_grouped_shapes_like_flat_shapes(tmp_path, write_deck, windows_fonts_environment):
    root = tmp_path / "deck"
    project, scene = construct_scene(root, write_deck, GROUP_SLIDES)
    variants = {}
    for name, grouped in (("grouped", True), ("flat", False)):
        data = copy.deepcopy(scene)
        for element in elements(data).values():
            if element["kind"] == "group":
                element["pptx_group"] = grouped
        result, pdf = write_and_render(project, data, name)
        assert result["groups"] == (2 if grouped else 0)
        variants[name] = page_marks(pdf)
    # グループの範囲は、最後に加えた吹き出しのしっぽまでを含む（PowerPoint で開いたときのグループの枠）。
    from pptx import Presentation

    outer = next(shape for shape in Presentation(root / "grouped.pptx").slides[0].shapes if shape.name == "outer")
    x, y, w, h = elements(scene)["outer/talk"]["actual_area"]
    assert outer.left + outer.width >= round((x + w) * 12700) and outer.top + outer.height >= round((y + h) * 12700)
    (grouped_drawings, grouped_characters), (flat_drawings, flat_characters) = variants["grouped"], variants["flat"]
    assert len(grouped_drawings) == len(flat_drawings) and grouped_drawings
    for a, b in zip(grouped_drawings, flat_drawings):
        assert all(abs(x - y) <= GROUP_TOLERANCE_PT for x, y in zip(a[0], b[0])) and a[1:] == b[1:], (a, b)
    assert [char for char, _ in grouped_characters] == [char for char, _ in flat_characters] and grouped_characters
    for (char, a), (_, b) in zip(grouped_characters, flat_characters):
        assert all(abs(x - y) <= GROUP_TOLERANCE_PT for x, y in zip(a, b)), (char, a, b)


VECTOR_SLIDES = """from slidekit import Deck, Rect

def build(context):
    deck = Deck(context)
    slide = deck.slide(id="figures", kind="blank")
    # cover の図は配置領域の左右へはみ出して切り取られるので、ほかの図と範囲が重ならないように行を分ける。
    slide.image(id="svg_contain", path="figure.svg", area=Rect(60, 60, 160, 160), fit="contain")
    slide.image(id="pdf_contain", path="figure.pdf", page=1, area=Rect(460, 60, 160, 160), fit="contain")
    slide.image(id="svg_cover", path="figure.svg", area=Rect(260, 300, 160, 160), fit="cover")
    slide.image(id="pdf_cover", path="figure.pdf", page=1, area=Rect(660, 300, 160, 160), fit="cover")
    return deck
"""
RED, BLUE = (1.0, 0.0, 0.0), (0.0, 0.0, 1.0)


def painted(drawing, colors=(RED, BLUE)):
    """塗りの色が colors のどれかに近ければ、その色を返す。"""
    fill = drawing.fill
    return next((color for color in colors if fill and all(abs(a - b) <= 0.02 for a, b in zip(fill, color))), None)


def make_figures(root):
    """左半分が赤、右半分が青の 2:1 の図を、SVG と PDF で作る。"""
    root.mkdir(parents=True, exist_ok=True)
    (root / "figure.svg").write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100">'
        '<rect width="100" height="100" fill="#FF0000"/>'
        '<rect x="100" width="100" height="100" fill="#0000FF"/></svg>',
        encoding="utf-8",
    )
    with Document() as document:
        page = document.new_page(width=200, height=100)
        page.draw_rect((0, 0, 100, 100), color=RED, fill=RED)
        page.draw_rect((100, 0, 200, 100), color=BLUE, fill=BLUE)
        document.save(root / "figure.pdf")


def test_real_powerpoint_svg_image_is_vector(tmp_path, write_deck, windows_fonts_environment):
    from slidekit.assets import PDF_VECTOR_MESSAGE, prepare_assets
    from slidekit.errors import InputError
    from slidekit.pdf import Box

    root = tmp_path / "deck"
    make_figures(root)
    project, scene = construct_scene(root, write_deck, VECTOR_SLIDES)
    # PDF の素材は PNG だけにする。vector を指定した PDF の素材は、変換の前に誤りになる。
    rejected = copy.deepcopy(scene)
    elements(rejected)["pdf_contain"]["vector"] = True
    with pytest.raises(InputError) as raised:
        prepare_assets(rejected, project, root / "rejected_assets")
    assert PDF_VECTOR_MESSAGE in str(raised.value)
    assert (raised.value.code, raised.value.element_id) == ("invalid_image_option", "pdf_contain")
    pages = {}
    for name, vector in (("vector", True), ("raster", False)):
        data = copy.deepcopy(scene)
        for identifier, element in elements(data).items():
            if element["kind"] == "image":
                element["vector"] = vector and identifier.startswith("svg_")
        result, pdf = write_and_render(project, data, name, assets=True)
        assert result["vector_images"] == (2 if vector else 0)
        pages[name] = (pdf, elements(data))
    for name, (pdf, found) in pages.items():
        with open_pdf(pdf) as document:
            drawings = document[0].drawings()
            pixels = document[0].render(2)
        for identifier, element in found.items():
            if element["kind"] != "image":
                continue
            x, y, w, h = element["area"]
            colors = {painted(drawing) for drawing in drawings if drawing.rect.intersects(Box(x, y, x + w, y + h))}
            # SVG でも埋め込んだ画像は、PowerPoint が線や面として描く。PNG だけの画像（PDF の素材を含む）は画像として描く。
            expected = {RED, BLUE} if name == "vector" and identifier.startswith("svg_") else set()
            assert colors - {None} == expected, (name, identifier, colors)
            # 表示の範囲と cover の切り取りは PNG の場合と同じ（赤と青の位置、contain の余白）。
            samples = {(0.25, 0.5): (255, 0, 0), (0.75, 0.5): (0, 0, 255)}
            if element["fit"] == "contain":
                samples[(0.5, 0.1)] = (255, 255, 255)
            for (u, v), color in samples.items():
                actual = pixels.getpixel((round((x + w * u) * 2), round((y + h * v) * 2)))
                assert all(abs(a - b) <= 8 for a, b in zip(actual, color)), (name, identifier, (u, v), actual)
    # contain の SVG の描画の範囲は PNG の画像の範囲と一致する。
    raster = image_boxes(pages["raster"][0], 0)
    with open_pdf(pages["vector"][0]) as document:
        drawings = document[0].drawings()
    x, y, w, h = pages["vector"][1]["svg_contain"]["area"]
    rects = [
        drawing.rect for drawing in drawings if painted(drawing) and drawing.rect.intersects(Box(x, y, x + w, y + h))
    ]
    union = Box(min(r.x0 for r in rects), min(r.y0 for r in rects), max(r.x1 for r in rects), max(r.y1 for r in rects))
    center = (x + w / 2, y + h / 2)
    image = min(
        raster, key=lambda rect: abs((rect.x0 + rect.x1) / 2 - center[0]) + abs((rect.y0 + rect.y1) / 2 - center[1])
    )
    assert all(abs(a - b) <= LINE_TOLERANCE_PT for a, b in zip(union, image)), ("svg_contain", union, image)
