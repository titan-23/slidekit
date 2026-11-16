"""まとまり・接続・図の書き出しを確かめる：PowerPoint のグループ、線の端の接続、SVG でも埋め込む画像。
中央揃えの独立した数式の左の余白（末尾の添字の後の空き）と、数式だけの行の U+200B の変種も確かめる。

場面の pptx_group は作成 API が決めるので、ここでは場面の辞書を直接作る。
"""

from types import SimpleNamespace
import zipfile

from fontTools.ttLib import TTFont
from lxml import etree
from PIL import Image
from pptx import Presentation
import pytest

from fake_fonts import FakeFonts
from slidekit.fonts import FontRegistry
from slidekit.math.metrics import Box
from slidekit.pptx_validation import validate_pptx
from slidekit.writer import (
    NS,
    SVG_BLIP_URI,
    SVG_CONTENT_TYPE,
    SVG_NS,
    ZERO_WIDTH_SPACE,
    anchored_site,
    display_indent,
    display_trailing_space,
    paragraph_content,
    pptx_group_of,
    write_pptx,
)

EMU = 12700
STYLE = {
    "size_pt": 20,
    "font_japanese": "Yu Gothic UI",
    "font_latin": "Arial",
    "font_math": "Cambria Math",
    "color": "17212D",
    "bold": False,
    "align": "left",
    "valign": "top",
    "line_spacing": 1.15,
    "padding_pt": 0,
    "word_wrap": True,
    "fill": None,
    "line_color": "183153",
    "line_width_pt": 1.0,
}


def project(tmp_path):
    return SimpleNamespace(
        config={"metadata": {}},
        root=tmp_path,
        resolve_input=lambda value: (tmp_path / value).resolve(),
    )


def shape(identifier, kind="box", area=(40, 40, 100, 50), *, text="図形", parent=None, **attributes):
    element = {
        "id": identifier,
        "kind": kind,
        "area": list(area),
        "style": dict(STYLE),
        "content": [{"kind": "text", "text": text}] if text else [],
        **attributes,
    }
    if kind == "box":
        element.setdefault("shape", "rectangle")
    if parent is not None:
        element["parent_id"] = parent
    return element


def group(identifier, *, parent=None, pptx_group=True, description=None):
    element = {"id": identifier, "kind": "group", "area": [0, 0, 0, 0], "style": {}, "content": []}
    if pptx_group is not None:
        element["pptx_group"] = pptx_group
    if parent is not None:
        element["parent_id"] = parent
    if description:
        element["description"] = description
    return element


def line(identifier, start, end, *, start_ref=None, end_ref=None, kind="line", parent=None):
    element = {
        "id": identifier,
        "kind": kind,
        "area": [0, 0, 0, 0],
        "style": dict(STYLE),
        "content": [],
        "start": list(start),
        "end": list(end),
    }
    for key, value in (("start_ref", start_ref), ("end_ref", end_ref)):
        if value is not None:
            element[key] = value
    if parent is not None:
        element["parent_id"] = parent
    return element


def ref(identifier, side, **values):
    return {"element_id": identifier, "side": side, **values}


def write(tmp_path, elements, *, name="deck.pptx"):
    scene = {
        "schema_version": 1,
        "size": [720, 540],
        "metadata": {},
        "slides": [{"id": "page", "kind": "content", "notes": "", "elements": elements}],
    }
    result = write_pptx(scene, project(tmp_path), tmp_path / name)
    with zipfile.ZipFile(tmp_path / name) as archive:
        root = etree.fromstring(archive.read("ppt/slides/slide1.xml"))
    return result, root


def named(root):
    """図形とグループを名前で引く辞書（p:cNvPr の親の親が図形かグループ）。"""
    return {
        node.get("name"): node.getparent().getparent() for node in root.iter(f"{{{NS['p']}}}cNvPr") if node.get("name")
    }


def groups_of(node):
    """図形が入っている p:grpSp の名前の列（外側から）。"""
    names = [
        ancestor.find("p:nvGrpSpPr/p:cNvPr", NS).get("name") for ancestor in node.iterancestors(f"{{{NS['p']}}}grpSp")
    ]
    return list(reversed(names))


def transform(node, path):
    values = node.find(path, NS)
    return {etree.QName(child).localname: {key: int(value) for key, value in child.attrib.items()} for child in values}


def test_group_is_written_as_powerpoint_group_with_identity_transform(tmp_path):
    elements = [
        group("diagram", description="処理の流れ"),
        shape("diagram/first", area=(40, 60, 100, 50), parent="diagram"),
        shape("diagram/second", "text", area=(200, 80, 120, 40), parent="diagram"),
        line("diagram/link", (140, 85), (200, 100), parent="diagram", kind="arrow"),
    ]
    result, root = write(tmp_path, elements)
    shapes = named(root)
    assert result["groups"] == 1 and result["shapes"] == 3
    grouped = shapes["diagram"]
    assert grouped.tag == f"{{{NS['p']}}}grpSp" and grouped.getparent().tag == f"{{{NS['p']}}}spTree"
    assert grouped.find("p:nvGrpSpPr/p:cNvPr", NS).get("descr") == "処理の流れ"
    for name in ("diagram/first", "diagram/second", "diagram/link"):
        assert groups_of(shapes[name]) == ["diagram"], name
    # 座標変換は恒等で、範囲は子の範囲を囲む。子の座標はページ上の座標のまま。
    values = transform(grouped, "p:grpSpPr/a:xfrm")
    assert values["off"] == values["chOff"] == {"x": 40 * EMU, "y": 60 * EMU}
    assert values["ext"] == values["chExt"] == {"cx": 280 * EMU, "cy": 60 * EMU}
    assert transform(shapes["diagram/second"], "p:spPr/a:xfrm")["off"] == {"x": 200 * EMU, "y": 80 * EMU}
    # 再読み込みでき、グループの子として読むことができる。
    group_shape = next(item for item in Presentation(tmp_path / "deck.pptx").slides[0].shapes if item.name == "diagram")
    assert [item.name for item in group_shape.shapes] == ["diagram/first", "diagram/second", "diagram/link"]


def test_nested_and_flat_groups_follow_scene(tmp_path):
    elements = [
        shape("before", area=(10, 10, 40, 20)),
        group("outer"),
        shape("outer/a", area=(40, 60, 60, 30), parent="outer"),
        group("outer/inner", parent="outer"),
        shape("outer/inner/b", area=(120, 60, 60, 30), parent="outer/inner"),
        group("outer/inner/flat", parent="outer/inner", pptx_group=False),
        shape("outer/inner/flat/c", area=(200, 60, 60, 30), parent="outer/inner/flat"),
        shape("outer/d", area=(280, 60, 60, 30), parent="outer"),
        group("plain", pptx_group=False),
        shape("plain/e", area=(40, 200, 60, 30), parent="plain"),
        group("legacy", pptx_group=None),
        shape("legacy/f", area=(120, 200, 60, 30), parent="legacy"),
        group("inside_flat", pptx_group=False),
        group("inside_flat/kept", parent="inside_flat"),
        shape("inside_flat/kept/g", area=(200, 200, 60, 30), parent="inside_flat/kept"),
        group("empty"),
        shape("after", area=(10, 300, 40, 20)),
    ]
    result, root = write(tmp_path, elements)
    shapes = named(root)
    expected = {
        "before": [],
        "outer/a": ["outer"],
        "outer/inner/b": ["outer", "outer/inner"],
        # 書き出さないまとまりの子は、最も近い書き出すまとまりに入る。
        "outer/inner/flat/c": ["outer", "outer/inner"],
        "outer/d": ["outer"],
        "plain/e": [],
        "legacy/f": [],
        "inside_flat/kept/g": ["inside_flat/kept"],
        "after": [],
    }
    assert {name: groups_of(shapes[name]) for name in expected} == expected
    # 書き出すまとまりだけが p:grpSp になる。描く子孫のないまとまりはグループにしない。
    grouped = sorted(node.find("p:nvGrpSpPr/p:cNvPr", NS).get("name") for node in root.iter(f"{{{NS['p']}}}grpSp"))
    assert grouped == ["inside_flat/kept", "outer", "outer/inner"] and result["groups"] == 3
    # 重なり順は場面の順のまま。
    tree = root.find("p:cSld/p:spTree", NS)
    top = [node.find(".//p:cNvPr", NS).get("name") for node in tree if node.find(".//p:cNvPr", NS) is not None][1:]
    assert top == ["before", "outer", "plain/e", "legacy/f", "inside_flat/kept", "after"]
    assert [
        node.find(".//p:cNvPr", NS).get("name")
        for node in shapes["outer"]
        if node.tag != f"{{{NS['p']}}}nvGrpSpPr" and node.find(".//p:cNvPr", NS) is not None
    ] == ["outer/a", "outer/inner", "outer/d"]
    by_id = {element["id"]: element for element in elements}
    assert pptx_group_of(by_id["outer/inner/flat/c"], by_id) == "outer/inner"
    assert pptx_group_of(by_id["outer/inner"], by_id) == "outer"
    assert pptx_group_of(by_id["legacy/f"], by_id) is None


def test_math_in_powerpoint_group_is_rejected(tmp_path):
    # 数式を含むまとまりは PowerPoint のグループにしない（作成 API が pptx_group を真にしない）。場面が求めても書き出さない。
    elements = [
        group("diagram"),
        shape("diagram/formula", parent="diagram", text=None, content=[{"kind": "math", "text": "x+1"}]),
    ]
    with pytest.raises(ValueError, match="数式を含む要素は PowerPoint のグループに入れられない"):
        write(tmp_path, elements)
    elements[0]["pptx_group"] = False
    _, root = write(tmp_path, elements, name="flat.pptx")
    assert root.find(".//p:grpSp", NS) is None and root.find(".//m:oMath", NS) is not None


def connections(root, name):
    connector = named(root)[name]
    ids = {node.get("id"): node.get("name") for node in root.iter(f"{{{NS['p']}}}cNvPr")}
    result = []
    for key in ("stCxn", "endCxn"):
        node = connector.find(f"p:nvCxnSpPr/p:cNvCxnSpPr/a:{key}", NS)
        result.append(None if node is None else (ids[node.get("id")], int(node.get("idx"))))
    return result


def test_anchored_ends_are_connected_to_sites(tmp_path):
    picture = tmp_path / "picture.png"
    Image.new("RGB", (200, 100), "white").save(picture)
    targets = [
        shape("rect", area=(300, 40, 100, 50)),
        shape("round", area=(300, 110, 100, 50), shape="rounded_rectangle"),
        shape("oval", area=(300, 180, 100, 50), shape="ellipse"),
        shape("diamond", area=(300, 250, 100, 50), shape="diamond"),
        shape("label", "text", area=(300, 320, 100, 50)),
        shape("formula", "equation", area=(300, 390, 100, 50), text=None, content=[{"kind": "math", "text": "x"}]),
        shape("turned", area=(500, 40, 100, 50), rotation=90),
        {
            "id": "cover",
            "kind": "image",
            "area": [500, 200, 100, 50],
            "path": str(picture),
            "fit": "cover",
            "actual_area": [500, 200, 100, 50],
            "style": {},
            "content": [],
        },
        {
            "id": "boxed",
            "kind": "image",
            "area": [500, 300, 100, 100],
            "path": str(picture),
            "fit": "contain",
            "actual_area": [500, 325, 100, 50],
            "style": {},
            "content": [],
        },
    ]
    lines = [
        line("to_rect", (100, 65), (300, 65), end_ref=ref("rect", "left")),
        line("from_round", (350, 160), (350, 175), start_ref=ref("round", "bottom"), end_ref=ref("oval", "top")),
        line("oval_sides", (300, 205), (400, 205), start_ref=ref("oval", "left"), end_ref=ref("oval", "right")),
        line("oval_bottom", (350, 230), (350, 250), start_ref=ref("oval", "bottom"), end_ref=ref("diamond", "top")),
        line("diamond_right", (400, 275), (450, 275), start_ref=ref("diamond", "right")),
        # gap と at は位置だけを変え、接続点の番号は辺で決まる。
        line("spaced", (100, 345), (294, 332.5), end_ref=ref("label", "left", gap=6, at=0.25)),
        line("formula_top", (350, 360), (350, 390), end_ref=ref("formula", "top")),
        line("center", (100, 100), (350, 65), end_ref=ref("rect", "center")),
        # 回転した図形も、図形自身の向きの辺の番号を使う。
        line("turned_top", (620, 65), (575, 65), end_ref=ref("turned", "top")),
        line("to_cover", (450, 225), (500, 225), end_ref=ref("cover", "left")),
        line("to_boxed", (450, 350), (500, 350), end_ref=ref("boxed", "left")),
        line("missing", (10, 10), (20, 20), end_ref=ref("absent", "left")),
    ]
    result, root = write(tmp_path, targets + lines)
    assert connections(root, "to_rect") == [None, ("rect", 1)]
    assert connections(root, "from_round") == [("round", 2), ("oval", 0)]
    # 楕円は 8 点：上 0、左 2、下 4、右 6。
    assert connections(root, "oval_sides") == [("oval", 2), ("oval", 6)]
    assert connections(root, "oval_bottom") == [("oval", 4), ("diamond", 0)]
    assert connections(root, "diamond_right") == [("diamond", 3), None]
    assert connections(root, "spaced") == [None, ("label", 1)]
    assert connections(root, "formula_top") == [None, ("formula", 0)]
    assert connections(root, "turned_top") == [None, ("turned", 0)]
    assert connections(root, "to_cover") == [None, ("cover", 1)]
    # 中心、表示範囲が配置領域と違う contain の画像、場面にない接続先は接続しない。
    for name in ("center", "to_boxed", "missing"):
        assert connections(root, name) == [None, None], name
    assert result["connections"] == 12
    # 端点の座標は場面の値のまま（接続点へ動かさない）。
    spaced = named(root)["spaced"].find("p:spPr/a:xfrm", NS)
    assert int(spaced.find("a:off", NS).get("x")) == 100 * EMU and int(spaced.find("a:ext", NS).get("cx")) == 194 * EMU
    # 接続の要素は ECMA-376 の順（a:stCxn、a:endCxn）に並ぶ。
    properties = named(root)["from_round"].find("p:nvCxnSpPr/p:cNvCxnSpPr", NS)
    assert [etree.QName(child).localname for child in properties] == ["stCxn", "endCxn"]


def test_anchored_site_ignores_gap_and_position():
    # 描いた端が接続点にない端（gap が 0 でないか at が 0.5 でない端）も接続する。端か接続先がなければ接続しない。
    target = shape("rect", area=(300, 40, 100, 50))
    assert anchored_site(ref("rect", "left"), target) == anchored_site(ref("rect", "left", gap=6, at=0.25), target) == 1
    assert anchored_site(None, target) is None and anchored_site(ref("rect", "left"), None) is None


def svg_file(path, color="#C43D3D"):
    path.write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100">'
        f'<rect width="200" height="100" fill="{color}"/></svg>',
        encoding="utf-8",
    )
    return path


def test_vector_image_has_svg_blip_and_png_fallback(tmp_path):
    png = tmp_path / "picture.png"
    Image.new("RGB", (200, 100), "white").save(png)
    vector, other = svg_file(tmp_path / "picture.svg"), svg_file(tmp_path / "other.svg", "#183153")

    def image(identifier, svg, y):
        element = {
            "id": identifier,
            "kind": "image",
            "area": [40, y, 200, 100],
            "path": str(png),
            "fit": "contain",
            "actual_area": [40, y, 200, 100],
            "style": {},
            "content": [],
        }
        if svg is not None:
            element["vector_path"] = str(svg)
        return element

    elements = [
        image("vector", vector, 20),
        image("again", vector, 140),
        image("other", other, 260),
        image("raster", None, 380),
    ]
    result, root = write(tmp_path, elements)
    assert result["vector_images"] == 3
    shapes = named(root)
    with zipfile.ZipFile(tmp_path / "deck.pptx") as archive:
        relationships = {
            node.get("Id"): node for node in etree.fromstring(archive.read("ppt/slides/_rels/slide1.xml.rels"))
        }
        content_types = etree.fromstring(archive.read("[Content_Types].xml"))
        targets = {}
        for name in ("vector", "again", "other"):
            blip = shapes[name].find("p:blipFill/a:blip", NS)
            # a:blip の PNG（互換用）はそのまま残し、拡張に SVG の参照を加える。
            png_target = relationships[blip.get(f"{{{NS['r']}}}embed")].get("Target")
            assert png_target.endswith(".png")
            extension = blip.find("a:extLst/a:ext", NS)
            assert extension.get("uri") == SVG_BLIP_URI
            relationship = relationships[extension.find(f"{{{SVG_NS}}}svgBlip").get(f"{{{NS['r']}}}embed")]
            assert relationship.get("Type").endswith("/image")
            targets[name] = "ppt/" + relationship.get("Target").removeprefix("../")
        # 同じ中身の SVG は 1 つの部品にする。
        assert targets["vector"] == targets["again"] != targets["other"]
        assert archive.read(targets["vector"]) == vector.read_bytes()
        assert archive.read(targets["other"]) == other.read_bytes()
        declared = {node.get("PartName"): node.get("ContentType") for node in content_types if node.get("PartName")}
        defaults = {node.get("Extension"): node.get("ContentType") for node in content_types if node.get("Extension")}
        for target in set(targets.values()):
            assert declared.get("/" + target, defaults.get("svg")) == SVG_CONTENT_TYPE
    assert shapes["raster"].find("p:blipFill/a:blip/a:extLst", NS) is None
    # python-pptx で読み直すことができる（SVG の部品を含む資料を破損させない）。
    assert [item.name for item in Presentation(tmp_path / "deck.pptx").slides[0].shapes] == [
        "vector",
        "again",
        "other",
        "raster",
    ]


def test_dashed_box_outline(tmp_path):
    def styled(identifier, kind="box", **style):
        element = shape(identifier, kind)
        element["style"].update(style)
        return element

    elements = [
        styled("dashed", dashed=True),
        styled("solid", dashed=False),
        styled("text", "text", dashed=True),
        styled("hidden", dashed=True, line_color=None),
    ]
    _, root = write(tmp_path, elements)
    shapes = named(root)

    def outline(name):
        node = shapes[name].find("p:spPr/a:ln", NS)
        dash = node.find("a:prstDash", NS)
        return (node.find("a:noFill", NS) is not None, None if dash is None else dash.get("val"))

    # 枠線のある箱と文字枠は、dashed なら破線。枠線のない図形は線を描かない。
    assert outline("dashed") == (False, "dash")
    assert outline("text") == (False, "dash")
    assert outline("solid") == (False, None)
    assert outline("hidden") == (True, None)


def test_math_targets_are_connected(tmp_path):
    # 数式を含む接続先も接続する（互換画像を付ける資料でも同じ判定）。
    formula = shape("formula", "equation", area=(300, 40, 100, 50), text=None, content=[{"kind": "math", "text": "x"}])
    mixed = shape(
        "mixed",
        area=(300, 120, 100, 50),
        shape="ellipse",
        text=None,
        content=[{"kind": "text", "text": "長さ "}, {"kind": "math", "text": "n"}],
    )
    plain = shape("plain", area=(300, 200, 100, 50))
    lines = [
        line("to_formula", (100, 65), (300, 65), end_ref=ref("formula", "left")),
        line("to_mixed", (350, 90), (350, 120), start_ref=ref("formula", "bottom"), end_ref=ref("mixed", "top")),
        line("to_plain", (100, 225), (300, 225), end_ref=ref("plain", "left")),
    ]
    _, root = write(tmp_path, [formula, mixed, plain, *lines])
    assert [connections(root, name) for name in ("to_formula", "to_mixed", "to_plain")] == [
        [None, ("formula", 1)],
        [("formula", 2), ("mixed", 0)],
        [None, ("plain", 1)],
    ]


class LayoutFonts(FakeFonts):
    """数式の箱（`math.metrics.Box`）と MATH 表の SpaceAfterScript（0.05 em）を返す偽のフォント。

    boxes は式から (送り幅, 描画の右端, 末尾の斜体補正) の em への辞書。
    """

    functions = ()

    def __init__(self, boxes):
        super().__init__()
        self.boxes = boxes
        self.layouts = []

    def math_layout(self, expr, size, bold=False, *, family=None, display=False, japanese=None, latin=None):
        self.layouts.append({"expression": expr, "size": size, "bold": bold, "family": family, "display": display})
        width, right, italic = self.boxes[expr]
        return Box(width * size, 0.7 * size, 0.1 * size, (0.0, right * size, 0.7 * size, 0.1 * size), italic * size)

    def math_constant(self, name, *, family=None):
        return {"SpaceAfterScript": 0.05}[name]


def display(identifier, expression, *, align="center", **style):
    element = shape(
        identifier, "equation", area=(40, 40, 400, 60), text=None, content=[{"kind": "math", "text": expression}]
    )
    element["style"].update(align=align, **style)
    return element


def test_centered_display_math_adds_trailing_script_space(tmp_path):
    fonts = LayoutFonts(
        {
            "x_i": (1.0, 0.9, 0.025),
            r"\sum_{i=1}^{n} a_i": (3.0, 2.95, 0.025),
            "(x)": (1.5, 1.4, 0.0),
            r"\sum_{i} a": (1.2, 1.1, 0.03),
            r"\max_i": (1.6, 1.55, 0.0),
            r"f_{i}\,": (1.2, 1.0, 0.0),
            "y^2": (1.0, 0.8, -0.01),
        }
    )
    elements = [
        display("script", "x_i"),
        display("operand", r"\sum_{i=1}^{n} a_i", bold=True),
        display("paren", "(x)"),
        display("operator", r"\sum_{i} a"),
        display("name", r"\max_i"),
        display("spaced", r"f_{i}\,"),
        display("upright", "y^2"),
        display("left", "x_i", align="left"),
    ]
    # 中央揃えの独立した数式の左の余白は、PowerPoint が式の後ろに加える幅（空白の送り幅 0.5 em × 20 pt）と、末尾が添字の式の
    # 箱の末尾の空き（SpaceAfterScript 0.05 em × 20 pt と、末尾の添字の正の斜体補正）の和。末尾が添字でない式（括弧、
    # 引数の最後が添字でない大型演算子、空白）は加えない。
    expected = {
        "script": 11.5,
        "operand": 11.5,
        "paren": 10,
        "operator": 10,
        "name": 11,
        "spaced": 10,
        "upright": 11,
        "left": 0,
    }
    assert {element["id"]: round(display_indent(element, fonts), 6) for element in elements} == expected
    assert display_trailing_space(elements[0], fonts) == pytest.approx(1.5)
    assert fonts.layouts[0] == {
        "expression": "x_i",
        "size": 20,
        "bold": False,
        "family": "Cambria Math",
        "display": True,
    }
    assert {
        "expression": r"\sum_{i=1}^{n} a_i",
        "size": 20,
        "bold": True,
        "family": "Cambria Math",
        "display": True,
    } in fonts.layouts
    # 数式の箱を返さない計測器（math_layout のない FontSource）と、組めない式は加えない。
    assert display_trailing_space(elements[0], FakeFonts()) == 0 and display_indent(elements[0], FakeFonts()) == 10
    assert display_trailing_space(display("bad", r"x_{"), fonts) == 0
    # 計測器の中の誤り（MATH 表の定数を読む先の形が違う）は、0 にして知らせずに補正をやめることはせず、例外にする。
    changed = LayoutFonts(fonts.boxes)
    changed.math_constant = lambda name, *, family=None: SimpleNamespace().font.constant(name)
    with pytest.raises(AttributeError):
        display_trailing_space(elements[0], changed)
    # 書き出しと PPTX の検査は同じ余白を使う。
    data = {
        "schema_version": 1,
        "size": [720, 540],
        "metadata": {},
        "slides": [{"id": "page", "kind": "content", "notes": "", "elements": elements}],
    }
    path = tmp_path / "deck.pptx"
    write_pptx(data, project(tmp_path), path, fonts=fonts)
    root = etree.fromstring(zipfile.ZipFile(path).read("ppt/slides/slide1.xml"))
    margins = {name: int(node.find(".//a:pPr", NS).get("marL", "0")) / EMU for name, node in named(root).items()}
    assert margins == expected
    assert [
        issue
        for issue in validate_pptx(path, data, project(tmp_path), fonts=fonts)
        if issue["code"] == "math_layout_mismatch"
    ] == []


def test_trailing_script_space_reads_math_table_of_real_font(tmp_path, dejavu):
    # 実際の FontRegistry の処理（数式の箱と、math_constant が読む MATH 表の SpaceAfterScript）を通す。期待値の定数は
    # 計測器を通さず、フォントの MATH 表から直接読む。
    family = dejavu.family
    font = TTFont(dejavu.path)
    try:
        space = font["MATH"].table.MathConstants.SpaceAfterScript.Value / font["head"].unitsPerEm
    finally:
        font.close()
    assert space > 0
    project_fonts = SimpleNamespace(
        theme={"fonts": {"japanese": family, "latin": family, "math": family}},
        environment={"fonts": {family: {"path": str(dejavu.path), "bold_path": str(dejavu.bold_path)}}},
        config={"math": {}},
    )
    fonts_style = {"font_math": family, "font_latin": family, "font_japanese": family}
    elements = [display("script", "x_i", **fonts_style), display("paren", "(x)", **fonts_style)]
    with FontRegistry(project_fonts) as fonts:
        assert fonts.math_constant("SpaceAfterScript", family=family) == pytest.approx(space)
        with pytest.raises(KeyError):
            fonts.math_constant("NoSuchConstant")
        box = fonts.math_layout("x_i", 20, display=True, family=family, japanese=family, latin=family)
        expected = space * 20 + max(0.0, box.italic)
        assert display_trailing_space(elements[0], fonts) == pytest.approx(expected)
        assert display_trailing_space(elements[1], fonts) == 0
        # 末尾が添字の式だけ、左の余白が末尾の空きの分だけ広い（後ろに加える幅は同じ書体・文字サイズなので等しい）。
        indents = {element["id"]: display_indent(element, fonts) for element in elements}
        assert indents["script"] - indents["paren"] == pytest.approx(expected)
        # 書き出しと PPTX の検査も同じ処理で同じ余白を使う。
        data = {
            "schema_version": 1,
            "size": [720, 540],
            "metadata": {},
            "slides": [{"id": "page", "kind": "content", "notes": "", "elements": elements}],
        }
        path = tmp_path / "deck.pptx"
        write_pptx(data, project(tmp_path), path, fonts=fonts)
        root = etree.fromstring(zipfile.ZipFile(path).read("ppt/slides/slide1.xml"))
        margins = {name: int(node.find(".//a:pPr", NS).get("marL", "0")) / EMU for name, node in named(root).items()}
        assert margins == pytest.approx(indents, abs=1 / EMU)
        assert [
            issue
            for issue in validate_pptx(path, data, project(tmp_path), fonts=fonts)
            if issue["code"] == "math_layout_mismatch"
        ] == []


def test_math_only_lines_get_separators(tmp_path):
    two = shape(
        "two",
        text=None,
        content=[{"kind": "math", "text": "a"}, {"kind": "text", "text": "\n"}, {"kind": "math", "text": "b"}],
    )
    mixed = shape("mixed", text=None, content=[{"kind": "text", "text": "説明\n"}, {"kind": "math", "text": "x"}])
    inline = shape(
        "inline",
        text=None,
        content=[{"kind": "math", "text": "a"}, {"kind": "text", "text": " と "}, {"kind": "math", "text": "b"}],
    )

    def kinds(element):
        return [(kind, text if kind != "math" else "") for kind, text, _, _ in paragraph_content(element)]

    # 明示改行で分けた数式だけの行ごとに、その行の後に U+200B を加える（通常文字のある行には加えない）。区切りのない数式だけの
    # 行は、PowerPoint が揃えによらず枠の中央に置く（PowerPoint 16.0、build 20430 で確かめた。
    # 中央揃えの枠で +10.4 pt、左揃えの広い枠で +117.6 pt ずれた）。
    assert kinds(two) == [
        ("math", ""),
        ("run", ZERO_WIDTH_SPACE),
        ("break", "\n"),
        ("math", ""),
        ("run", ZERO_WIDTH_SPACE),
    ]
    assert kinds(mixed) == [("run", "説明"), ("break", "\n"), ("math", ""), ("run", ZERO_WIDTH_SPACE)]
    assert kinds(inline) == [("math", ""), ("run", " と "), ("math", "")]
    assert ZERO_WIDTH_SPACE not in [text for _, text, _, _ in paragraph_content(display("solo", "x"))]
    # 書き出しと PPTX の検査は同じ区切りを使う。
    data = {
        "schema_version": 1,
        "size": [720, 540],
        "metadata": {},
        "slides": [{"id": "page", "kind": "content", "notes": "", "elements": [two, mixed, inline]}],
    }
    path = tmp_path / "deck.pptx"
    write_pptx(data, project(tmp_path), path)
    paragraph = named(etree.fromstring(zipfile.ZipFile(path).read("ppt/slides/slide1.xml")))["two"].find(".//a:p", NS)
    assert [etree.QName(child).localname for child in paragraph][1:] == ["m", "r", "br", "m", "r", "endParaRPr"]
    assert [
        issue["code"]
        for issue in validate_pptx(path, data, project(tmp_path))
        if issue["code"] in {"font_mismatch", "text_mismatch", "math_layout_mismatch"}
    ] == []


# ---- 吹き出し ----


def callout(identifier, area=(40, 40, 100, 50), *, side="left", at=0.5, shape_name="rectangle", **attributes):
    """しっぽのある box の場面の要素（作成 API と同じく outline.tail_outline から作る）。"""
    from slidekit.outline import tail_outline

    element = shape(identifier, area=area, shape=shape_name, **attributes)
    outline = tail_outline(
        area, shape_name, side=side, at=at, corner_radius_pt=element["style"].get("corner_radius_pt")
    )
    element["actual_area"] = list(outline.actual_area)
    element["outside_parts"] = [{"part": name, "area": list(rect)} for name, rect in outline.outside]
    element["tail"] = {
        "side": side,
        "at": outline.at,
        "width_pt": 10.0,
        "length_pt": 8.0,
        "points": [list(point) for point in outline.tail],
    }
    return element


def test_callout_is_written_as_custom_geometry(tmp_path):
    from slidekit.outline import tail_outline

    rounded = callout("rounded", (300.25, 80.5, 120.5, 60.25), side="bottom", at=0.3, shape_name="rounded_rectangle")
    rounded["style"]["corner_radius_pt"] = 6.0
    elements = [
        callout("note", (420, 150, 240, 70), side="left", at=0.2),
        shape("plain", area=(40, 300, 100, 40)),
        rounded,
    ]
    _, root = write(tmp_path, elements)
    shapes = named(root)
    # しっぽのない box は既定図形のまま。
    assert shapes["plain"].find("p:spPr/a:prstGeom", NS).get("prst") == "rect"
    note = shapes["note"]
    assert note.find("p:spPr/a:prstGeom", NS) is None
    geometry = note.find("p:spPr/a:custGeom", NS)
    assert [etree.QName(child).localname for child in geometry] == [
        "avLst",
        "gdLst",
        "ahLst",
        "cxnLst",
        "rect",
        "pathLst",
    ]
    # a:xfrm はしっぽまでの範囲（actual_area）、a:path の座標はその左上を原点とする EMU。
    assert transform(note, "p:spPr/a:xfrm") == {
        "off": {"x": 412 * EMU, "y": 150 * EMU},
        "ext": {"cx": 248 * EMU, "cy": 70 * EMU},
    }
    path = geometry.find("a:pathLst/a:path", NS)
    assert (int(path.get("w")), int(path.get("h"))) == (248 * EMU, 70 * EMU)
    points = [
        (etree.QName(node).localname, int(node[0].get("x")) / EMU, int(node[0].get("y")) / EMU)
        for node in path
        if len(node)
    ]
    assert points == [
        ("moveTo", 8, 0),
        ("lnTo", 248, 0),
        ("lnTo", 248, 70),
        ("lnTo", 8, 70),
        ("lnTo", 8, 19),
        ("lnTo", 0, 14),
        ("lnTo", 8, 9),
        ("lnTo", 8, 0),
    ]
    assert etree.QName(path[-1]).localname == "close"
    guides = {node.get("name"): node.get("fmla") for node in geometry.find("a:gdLst", NS)}
    assert guides["bL"] == f"*/ w {8 * EMU} {248 * EMU}" and guides["bB"] == f"*/ h {70 * EMU} {70 * EMU}"
    assert geometry.find("a:rect", NS).attrib == {"l": "tL", "t": "tT", "r": "tR", "b": "tB"}
    # 接続点は本体の 4 辺の中点で、番号は rectangle と同じ（top=0・left=1・bottom=2・right=3）。
    connections = [(node.get("ang"), node[0].get("x"), node[0].get("y")) for node in geometry.find("a:cxnLst", NS)]
    assert connections == [
        ("16200000", "bHC", "bT"),
        ("10800000", "bL", "bVC"),
        ("5400000", "bHC", "bB"),
        ("0", "bR", "bVC"),
    ]
    # 小数の pt の座標でも、a:xfrm は actual_area と 0.1 pt 以内で一致し、角丸は a:arcTo で書く。
    outline = tail_outline(
        rounded["area"], "rounded_rectangle", side="bottom", at=rounded["tail"]["at"], corner_radius_pt=6
    )
    offset = transform(shapes["rounded"], "p:spPr/a:xfrm")
    written = [offset["off"]["x"] / EMU, offset["off"]["y"] / EMU, offset["ext"]["cx"] / EMU, offset["ext"]["cy"] / EMU]
    assert written == pytest.approx(list(outline.actual_area), abs=1e-4)
    arcs = shapes["rounded"].findall(".//a:arcTo", NS)
    assert [(arc.get("wR"), arc.get("stAng"), arc.get("swAng")) for arc in arcs] == [
        (str(6 * EMU), str(angle), "5400000") for angle in (16200000, 0, 5400000, 10800000)
    ]
    # 文字領域は本体の角丸の内側（measure.text_area と同じ）。
    text = outline.text_area
    rounded_guides = {node.get("name"): node.get("fmla") for node in shapes["rounded"].find(".//a:gdLst", NS)}
    width = offset["ext"]["cx"]
    assert rounded_guides["tL"] == f"*/ w {round((text.x - outline.actual_area.x) * EMU)} {width}"
    # 再読み込みでき、PowerPoint の自由図形として読む。
    reloaded = {item.name: item for item in Presentation(tmp_path / "deck.pptx").slides[0].shapes}
    assert reloaded["note"].shape_type is not None and reloaded["note"].text_frame.text == "図形"


def test_callout_outline_uses_round_join_only_with_a_line(tmp_path):
    framed = callout("framed")
    framed["style"]["dashed"] = True
    bare = callout("bare", (200, 40, 100, 50))
    bare["style"]["line_color"] = None
    _, root = write(tmp_path, [framed, bare])
    shapes = named(root)
    line = shapes["framed"].find("p:spPr/a:ln", NS)
    assert [etree.QName(child).localname for child in line] == ["solidFill", "prstDash", "round"]
    assert shapes["bare"].find("p:spPr/a:ln/a:round", NS) is None


def test_lines_connect_to_callout_sites(tmp_path):
    elements = [
        callout("note", (300, 100, 120, 60), side="left"),
        line("link", (100, 130), (300, 130), end_ref=ref("note", "left")),
        line("down", (360, 160), (360, 260), start_ref=ref("note", "bottom")),
    ]
    result, root = write(tmp_path, elements)
    shapes = named(root)
    assert result["connections"] == 2
    assert shapes["link"].find(".//a:endCxn", NS).get("idx") == "1"
    assert shapes["down"].find(".//a:stCxn", NS).get("idx") == "2"


def test_group_extents_include_a_callout_written_last(tmp_path):
    elements = [
        group("diagram"),
        shape("diagram/first", area=(40, 60, 100, 50), parent="diagram"),
        callout("diagram/note", (200, 60, 100, 50), side="right", parent="diagram"),
    ]
    _, root = write(tmp_path, elements)
    values = transform(named(root)["diagram"], "p:grpSpPr/a:xfrm")
    # しっぽ（右へ 8 pt）を含む範囲。
    assert values["off"] == values["chOff"] == {"x": 40 * EMU, "y": 60 * EMU}
    assert values["ext"] == values["chExt"] == {"cx": 268 * EMU, "cy": 50 * EMU}


def test_callout_with_math_and_validation(tmp_path):
    fonts = FakeFonts()
    note = callout("note", (300, 100, 160, 60), side="left", at=0.3)
    note["content"] = [{"kind": "text", "text": "比べる "}, {"kind": "math", "text": "x_i"}]
    elements = [
        group("diagram"),
        shape("diagram/first", area=(40, 300, 100, 50), parent="diagram"),
        callout("diagram/second", (200, 300, 100, 50), side="top", parent="diagram"),
        note,
        line("link", (100, 130), (300, 130), end_ref=ref("note", "left")),
    ]
    elements[-1]["area"] = [100, 130, 200, 0]
    data = {
        "schema_version": 1,
        "size": [720, 540],
        "metadata": {},
        "slides": [{"id": "page", "kind": "content", "notes": "", "elements": elements}],
    }
    path = tmp_path / "deck.pptx"
    write_pptx(data, project(tmp_path), path, fonts=fonts)
    root = etree.fromstring(zipfile.ZipFile(path).read("ppt/slides/slide1.xml"))
    assert named(root)["note"].find(".//m:oMath", NS) is not None
    codes = {issue["code"] for issue in validate_pptx(path, data, project(tmp_path), fonts=fonts)}
    assert not codes & {"geometry_mismatch", "connection_mismatch", "group_mismatch", "element_missing"}
