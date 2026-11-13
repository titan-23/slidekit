"""PPTX の編集可能性と雛形の保護、文字枠・断片・数式・禁則の一覧の書き出しを確かめる。

まとまり・接続・図の書き出しは test_writer_shapes.py にある。
"""

from copy import deepcopy
import os
from types import SimpleNamespace
import zipfile

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.util import Pt
import pytest

from fake_fonts import FakeFonts
from slidekit.math.office_math import display_trailing_width
from slidekit.ooxml import read_package, write_package
from slidekit.writer import (
    HYPERLINK_COLOR_NS,
    HYPERLINK_COLOR_URI,
    NS,
    ZERO_WIDTH_SPACE,
    connection_site,
    paragraph_content,
    write_pptx,
)

EMU = 12700
STYLE = {
    "size_pt": 24,
    "min_size_pt": 18,
    "font_japanese": "Yu Gothic UI",
    "font_latin": "Arial",
    "font_math": "Cambria Math",
    "color": "17212D",
    "emphasis_color": "C43D3D",
    "bold": False,
    "align": "left",
    "valign": "top",
    "line_spacing": 1.15,
    "padding_pt": 0,
    "word_wrap": True,
    "fill": None,
    "line_color": None,
    "line_width_pt": 1.0,
}


def project(tmp_path, **changes):
    config = {
        "metadata": {"title": "出力の検査", "author": "", "last_modified_by": ""},
        "math": {"functions": ["Lookup"]},
    }
    config.update(changes)
    return SimpleNamespace(config=config, root=tmp_path, resolve_input=lambda value: (tmp_path / value).resolve())


def element(identifier="label", kind="text", content=None, style=None, **attributes):
    """作成 API が作る形の要素（書式は解決済みの値をすべて持つ）。"""
    return {
        "id": identifier,
        "kind": kind,
        "area": [30, 40, 200, 70],
        "style": {**STYLE, **(style or {})},
        "content": content or [{"kind": "text", "text": "長い文章も縮めない\n改行"}],
        **attributes,
    }


def math(expression="x+1", **fragment):
    return {"kind": "math", "text": expression, **fragment}


def scene(elements):
    return {
        "schema_version": 1,
        "size": [720, 540],
        "metadata": {"title": "出力の検査", "author": "", "last_modified_by": ""},
        "slides": [{"id": "first", "kind": "content", "title": "検査", "notes": "説明します。", "elements": elements}],
    }


def xml_part(path, name="ppt/slides/slide1.xml"):
    with zipfile.ZipFile(path) as archive:
        return etree.fromstring(archive.read(name))


def written(tmp_path, elements, *, fonts=None, name="deck.pptx"):
    """要素を書き出し、図形の名前から p:sp への辞書を返す。"""
    write_pptx(scene(elements), project(tmp_path), tmp_path / name, fonts=fonts)
    root = xml_part(tmp_path / name)
    return {node.find(".//p:cNvPr", NS).get("name"): node for node in root.iter(f"{{{NS['p']}}}sp")}


def text_runs(shape):
    """通常文字の a:r（数式の中の m:r を除く）。"""
    return shape.findall(".//a:p/a:r", NS)


def run_fonts(shape):
    """通常文字の a:r ごとの（文字、a:latin、a:ea、a:cs の書体）。"""
    return [
        (
            run.findtext("a:t", namespaces=NS),
            *(run.find(f"a:rPr/a:{kind}", NS).get("typeface") for kind in ("latin", "ea", "cs")),
        )
        for run in text_runs(shape)
    ]


def test_editable_shapes_and_exact_text_size(tmp_path):
    contents = [
        {"kind": "text", "text": "式 "},
        {"kind": "math", "text": r"\frac{x_i}{2}"},
        {"kind": "emphasis", "text": "\n強調"},
    ]
    elements = [element(content=contents), {"id": "diagram", "kind": "group"}]
    elements += [
        element(f"box_{shape}", "box", shape=shape, parent_id="diagram")
        for shape in ("rectangle", "rounded_rectangle", "ellipse", "diamond")
    ]
    elements += [element("fixed", style={"word_wrap": False})]
    elements += [element("arrow", "arrow", start=[20, 20], end=[100, 100])]
    data = scene(elements)
    before = deepcopy(data)
    result = write_pptx(data, project(tmp_path), tmp_path / "native.pptx")
    root = xml_part(tmp_path / "native.pptx")
    assert data == before
    assert result["native_math"] == 1
    assert root.find("p:cSld", NS).get("name") == "first"
    assert not root.findall(".//p:grpSp", NS)
    assert len(root.findall(".//m:oMath", NS)) == 1
    assert root.find(".//mc:AlternateContent", NS) is None
    assert {node.get("prst") for node in root.findall(".//a:prstGeom", NS)} >= {
        "rect",
        "roundRect",
        "ellipse",
        "diamond",
    }
    assert root.find(".//a:tailEnd", NS).get("type") == "triangle"
    shapes = {node.find(".//p:cNvPr", NS).get("name"): node for node in root.iter(f"{{{NS['p']}}}sp")}
    # 折り返しは既定で有効（wrap="square"）、word_wrap=False の要素だけ wrap="none"。自動の縮小はしない。
    for name, shape in shapes.items():
        body = shape.find(".//a:bodyPr", NS)
        assert body.get("wrap") == ("none" if name == "fixed" else "square"), name
        assert body.find("a:noAutofit", NS) is not None
    assert {node.get("sz") for node in root.findall(".//a:rPr", NS)} == {"2400"}
    assert root.find(".//a:br", NS) is not None
    assert root.find(".//a:ea", NS).get("typeface") == "Yu Gothic UI"
    output = Presentation(tmp_path / "native.pptx")
    assert output.core_properties.author == ""
    assert output.core_properties.last_modified_by == ""
    assert output.core_properties.title == "出力の検査"
    assert output.slides[0].notes_slide.notes_text_frame.text == "説明します。"


def test_scene_without_word_wrap_is_not_wrapped(tmp_path):
    # 0.1 で作った場面（書式に word_wrap がない）は、計測・検査と同じく折り返さない。
    old = element()
    del old["style"]["word_wrap"]
    shapes = written(tmp_path, [old])
    assert shapes["label"].find(".//a:bodyPr", NS).get("wrap") == "none"


def test_paragraph_line_breaking_attributes_are_explicit(tmp_path):
    contents = [{"kind": "text", "text": "文章 text"}, math("y"), {"kind": "text", "text": "\n次の行"}]
    shapes = written(tmp_path, [element(content=contents), element("eq", "equation", [math()])])
    for shape in shapes.values():
        for paragraph in shape.findall(".//a:p", NS):
            properties = paragraph.find("a:pPr", NS)
            assert (properties.get("eaLnBrk"), properties.get("latinLnBrk"), properties.get("hangingPunct")) == (
                "1",
                "0",
                "0",
            )
    runs = text_runs(shapes["label"])
    assert runs and {run.find("a:rPr", NS).get("kern") for run in runs} == {"1200"}
    # 数式の中の文字（m:r の a:rPr）には kern を書かない。
    assert {node.get("kern") for node in shapes["label"].findall(".//m:r/a:rPr", NS)} == {None}


def test_side_padding_is_written(tmp_path):
    sides = {"left": 12, "top": 3, "right": 6, "bottom": 9}
    shapes = written(
        tmp_path, [element("sides", style={"padding_pt": sides}), element("even", "box", style={"padding_pt": 5})]
    )
    body = shapes["sides"].find(".//a:bodyPr", NS)
    assert [int(body.get(key)) / EMU for key in ("lIns", "tIns", "rIns", "bIns")] == [12, 3, 6, 9]
    body = shapes["even"].find(".//a:bodyPr", NS)
    assert [int(body.get(key)) / EMU for key in ("lIns", "tIns", "rIns", "bIns")] == [5, 5, 5, 5]


def test_fallback_symbol_run_names_both_fonts(tmp_path):
    fonts = FakeFonts(missing={"Arial": {"※"}})
    content = [{"kind": "text", "text": "注※記 ±→A"}]
    shapes = written(tmp_path, [element(content=content)], fonts=fonts)
    # 欧文の書体にない ※ は a:latin と a:ea の両方を和文の書体に、欧文の書体にある幅の曖昧な → は両方を欧文の書体にする。
    # PowerPoint が既定で和文の書体で描く ± は、欧文の書体にもあるが両方を和文の書体にする。
    # forced でない区切り（書式の書体のまま）は隣どうしをまとめる。a:cs は常に欧文の書体。
    assert run_fonts(shapes["label"]) == [
        ("注", "Arial", "Yu Gothic UI", "Arial"),
        ("※", "Yu Gothic UI", "Yu Gothic UI", "Arial"),
        ("記 ", "Arial", "Yu Gothic UI", "Arial"),
        ("±", "Yu Gothic UI", "Yu Gothic UI", "Arial"),
        ("→", "Arial", "Arial", "Arial"),
        ("A", "Arial", "Yu Gothic UI", "Arial"),
    ]
    # fonts がなければ 0.1 と同じく区切らない。
    plain = written(tmp_path, [element(content=content)], name="plain.pptx")
    assert [run.findtext("a:t", namespaces=NS) for run in text_runs(plain["label"])] == ["注※記 ±→A"]


def test_symbols_powerpoint_draws_in_japanese_font_are_forced_to_japanese_font(tmp_path):
    # PowerPoint は既定（lang="ja-JP"）で ○ × を和文の書体で、→ — を欧文の書体で描く（PowerPoint 16.0、build 20430）。
    # 計測と同じく、○ × は a:latin と a:ea の両方を和文の書体に、→ — は両方を欧文の書体にする（0.1 の見た目のまま）。
    content = [{"kind": "text", "text": "可○ 否× 例→ 間—A"}, {"kind": "emphasis", "text": "○×"}]
    shapes = written(tmp_path, [element(content=content)], fonts=FakeFonts())
    assert run_fonts(shapes["label"]) == [
        ("可", "Arial", "Yu Gothic UI", "Arial"),
        ("○", "Yu Gothic UI", "Yu Gothic UI", "Arial"),
        (" 否", "Arial", "Yu Gothic UI", "Arial"),
        ("×", "Yu Gothic UI", "Yu Gothic UI", "Arial"),
        (" 例", "Arial", "Yu Gothic UI", "Arial"),
        ("→", "Arial", "Arial", "Arial"),
        (" 間", "Arial", "Yu Gothic UI", "Arial"),
        ("—", "Arial", "Arial", "Arial"),
        ("A", "Arial", "Yu Gothic UI", "Arial"),
        ("○×", "Yu Gothic UI", "Yu Gothic UI", "Arial"),
    ]
    # 太字の断片も同じ書体の区切りになる。
    assert [run.find("a:rPr", NS).get("b") for run in text_runs(shapes["label"])][-1] == "1"
    # 和文の書体に字形がなければ、計測と同じく欧文の書体で計測し、両方を欧文の書体にする。
    missing = written(
        tmp_path,
        [element(content=[{"kind": "text", "text": "可○×"}])],
        fonts=FakeFonts(missing={"Yu Gothic UI": {"○"}}),
        name="missing.pptx",
    )
    assert run_fonts(missing["label"]) == [
        ("可", "Arial", "Yu Gothic UI", "Arial"),
        ("○", "Arial", "Arial", "Arial"),
        ("×", "Yu Gothic UI", "Yu Gothic UI", "Arial"),
    ]


def test_glyph_missing_in_both_fonts_is_written_with_style_fonts(tmp_path):
    # どちらの書体にも字形がない文字（配置の検査が報告する）でも、書き出しは止めない。
    fonts = FakeFonts(missing={"🙂"})
    shapes = written(tmp_path, [element(content=[{"kind": "text", "text": "a🙂※"}])], fonts=fonts)
    runs = [
        (run.findtext("a:t", namespaces=NS), run.find("a:rPr/a:ea", NS).get("typeface"))
        for run in text_runs(shapes["label"])
    ]
    assert runs == [("a🙂", "Yu Gothic UI"), ("※", "Arial")]


def test_fragment_runs_have_own_size_bold_color_underline_and_link(tmp_path):
    content = [
        {"kind": "text", "text": "詳細は "},
        {
            "kind": "text",
            "text": "利用ガイド",
            "style": {"color": "183153", "underline": True},
            "link": "https://example.com/guide",
        },
        {"kind": "text", "text": "（2026 年版）", "style": {"size_pt": 14, "bold": True}},
        {"kind": "emphasis", "text": "強調"},
    ]
    shapes = written(tmp_path, [element(content=content)])
    runs = text_runs(shapes["label"])
    properties = [run.find("a:rPr", NS) for run in runs]
    assert [run.findtext("a:t", namespaces=NS) for run in runs] == ["詳細は ", "利用ガイド", "（2026 年版）", "強調"]
    assert [node.get("sz") for node in properties] == ["2400", "2400", "1400", "2400"]
    assert [node.get("b") for node in properties] == ["0", "0", "1", "1"]
    assert [node.get("u") for node in properties] == [None, "sng", None, None]
    assert [node.find("a:solidFill/a:srgbClr", NS).get("val") for node in properties] == [
        "17212D",
        "183153",
        "17212D",
        "C43D3D",
    ]
    clicks = [node.find("a:hlinkClick", NS) for node in properties]
    assert [click is not None for click in clicks] == [False, True, False, False]
    extension = clicks[1].find("a:extLst/a:ext", NS)
    assert extension.get("uri") == HYPERLINK_COLOR_URI
    assert extension.find(f"{{{HYPERLINK_COLOR_NS}}}hlinkClr").get("val") == "tx"
    # a:hlinkClick は書体の要素の後に置く（ECMA-376 の順）。
    assert [etree.QName(child).localname for child in properties[1]] == ["solidFill", "latin", "ea", "cs", "hlinkClick"]
    relations = xml_part(tmp_path / "deck.pptx", "ppt/slides/_rels/slide1.xml.rels")
    link = next(node for node in relations if node.get("Id") == clicks[1].get(f"{{{NS['r']}}}id"))
    assert (link.get("Target"), link.get("TargetMode")) == ("https://example.com/guide", "External")
    assert link.get("Type").endswith("/hyperlink")


def test_superscript_fragment_writes_baseline_offset(tmp_path):
    content = [
        {"kind": "text", "text": "脚注"},
        {"kind": "text", "text": "1", "style": {"superscript": True}},
    ]
    shapes = written(tmp_path, [element(content=content)])
    properties = [run.find("a:rPr", NS) for run in text_runs(shapes["label"])]
    # PowerPoint の上付きの既定と同じ 30 % だけ基準線を上げる（大きさは PowerPoint が 3 分の 2 に縮めて描く）。
    assert [node.get("baseline") for node in properties] == [None, "30000"]
    assert [node.get("sz") for node in properties] == ["2400", "2400"]


def test_line_break_uses_fragment_properties(tmp_path):
    # a:br に書式がないと、PowerPoint は空の行を雛形の既定の文字サイズ（18 pt）で置く。改行がある断片の実効書式を書く。
    content = [
        {"kind": "text", "text": "一行目\n"},
        {"kind": "text", "text": "\n大きい", "style": {"size_pt": 32, "bold": True}},
        {"kind": "text", "text": "リンク\n", "link": "https://example.com/"},
    ]
    shapes = written(tmp_path, [element(content=content)])
    breaks = shapes["label"].findall(".//a:p/a:br", NS)
    assert [(node.find("a:rPr", NS).get("sz"), node.find("a:rPr", NS).get("b")) for node in breaks] == [
        ("2400", "0"),
        ("3200", "1"),
        ("2400", "0"),
    ]
    properties = breaks[1].find("a:rPr", NS)
    assert (properties.find("a:latin", NS).get("typeface"), properties.find("a:ea", NS).get("typeface")) == (
        "Arial",
        "Yu Gothic UI",
    )
    # リンクは改行の書式に写さない。
    assert breaks[2].find("a:rPr/a:hlinkClick", NS) is None
    assert [item[0] for item in paragraph_content(element(content=content))] == [
        "run",
        "break",
        "break",
        "run",
        "run",
        "break",
    ]


def kinsoku_nodes(path):
    root = xml_part(path, "ppt/presentation.xml")
    children = [etree.QName(child).localname for child in root]
    return root.findall("p:kinsoku", NS), children


def test_kinsoku_uses_measure_tables_and_replaces_template_value(tmp_path):
    from slidekit.measure import NO_LINE_END, NO_LINE_START

    write_pptx(scene([element()]), project(tmp_path), tmp_path / "deck.pptx")
    nodes, children = kinsoku_nodes(tmp_path / "deck.pptx")
    # 計測と同じ行頭・行末の禁止文字を、p:defaultTextStyle の前に 1 つだけ書く。
    assert len(nodes) == 1 and children.index("kinsoku") + 1 == children.index("defaultTextStyle")
    assert nodes[0].get("lang") == "ja-JP"
    assert set(nodes[0].get("invalStChars")) == set(NO_LINE_START) and len(nodes[0].get("invalStChars")) == len(
        NO_LINE_START
    )
    assert set(nodes[0].get("invalEndChars")) == set(NO_LINE_END) and len(nodes[0].get("invalEndChars")) == len(
        NO_LINE_END
    )
    # 雛形の p:kinsoku（スキーマでは 1 つだけ置くことができる）は、言語によらず置き換える。
    template = tmp_path / "template.pptx"
    Presentation().save(template)
    with zipfile.ZipFile(template) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    root = etree.fromstring(parts["ppt/presentation.xml"])
    for lang in ("ja-JP", "zh-CN"):
        node = etree.Element(f"{{{NS['p']}}}kinsoku", lang=lang, invalStChars="ー", invalEndChars="（")
        root.find("p:defaultTextStyle", NS).addprevious(node)
    parts["ppt/presentation.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    write_package(parts, template)
    write_pptx(
        scene([element()]),
        project(tmp_path, template={"path": "template.pptx", "layout_index": 6}),
        tmp_path / "templated.pptx",
    )
    templated, children = kinsoku_nodes(tmp_path / "templated.pptx")
    assert len(templated) == 1 and dict(templated[0].attrib) == dict(nodes[0].attrib)
    assert children.index("kinsoku") + 1 == children.index("defaultTextStyle")


def test_inline_math_is_kept_on_one_line(tmp_path):
    # PowerPoint は行の残りに入らない文中の数式を演算子の後で改行する。計測と同じく式ごと送るように、文中の数式は
    # m:box（m:noBreak）で包む。独立した数式は包まない。
    elements = [
        element(content=[{"kind": "text", "text": "式 "}, math("a+b")]),
        element("only", content=[math("c=d")]),
        element("display", "equation", [math("a+b")]),
    ]
    shapes = written(tmp_path, elements)
    for name in ("label", "only"):
        equation = shapes[name].find(".//a14:m/m:oMath", NS)
        assert [etree.QName(child).localname for child in equation] == ["box"], name
        assert equation.find("m:box/m:boxPr/m:noBreak", NS).get(f"{{{NS['m']}}}val") == "1"
        assert equation.find("m:box/m:e", NS) is not None
    assert shapes["display"].find(".//m:oMathPara/m:oMath/m:box", NS) is None


def test_paragraph_end_uses_last_fragment_properties(tmp_path):
    # 段落の末尾の書式（a:endParaRPr）がないと、PowerPoint は段落の最後の断片のリンクを PDF のリンクにしない。
    # 最後の断片の実効書式（リンクは除く）を書き、断片のない段落は要素の書式にする。
    content = [
        {"kind": "text", "text": "詳細は "},
        {
            "kind": "text",
            "text": "手順書",
            "style": {"color": "A23B2A", "size_pt": 28, "bold": True},
            "link": "https://example.com/guide#steps",
        },
    ]
    empty = element("empty", "box", shape="rectangle")
    empty["content"] = []
    elements = [
        element("link", content=content),
        empty,
        element("math", content=[{"kind": "text", "text": "式 "}, math(style={"size_pt": 32})]),
    ]
    shapes = written(tmp_path, elements)
    ends = {}
    for name in ("link", "empty", "math"):
        paragraph = shapes[name].find(".//a:p", NS)
        ends[name] = paragraph[-1]
        assert ends[name].tag == f"{{{NS['a']}}}endParaRPr", name
    link = ends["link"]
    assert (link.get("sz"), link.get("b"), link.find("a:solidFill/a:srgbClr", NS).get("val")) == ("2800", "1", "A23B2A")
    assert link.find("a:hlinkClick", NS) is None
    assert (link.find("a:latin", NS).get("typeface"), link.find("a:ea", NS).get("typeface")) == (
        "Arial",
        "Yu Gothic UI",
    )
    assert (ends["empty"].get("sz"), ends["empty"].get("b")) == ("2400", "0")
    assert ends["math"].get("sz") == "3200"


def test_math_fragment_uses_its_own_style(tmp_path):
    content = [
        math("a", style={"size_pt": 32, "color": "183153", "bold": True}),
        {"kind": "text", "text": " と "},
        math("b", emphasis=True),
    ]
    shapes = written(tmp_path, [element(content=content)])
    first, second = shapes["label"].findall(".//m:oMath", NS)
    properties = first.find(".//m:r/a:rPr", NS)
    assert (properties.get("sz"), properties.get("b"), properties.find("a:solidFill/a:srgbClr", NS).get("val")) == (
        "3200",
        "1",
        "183153",
    )
    properties = second.find(".//m:r/a:rPr", NS)
    assert (properties.get("sz"), properties.get("b"), properties.find("a:solidFill/a:srgbClr", NS).get("val")) == (
        "2400",
        "1",
        "C43D3D",
    )


def test_equation_is_display_math_with_alignment(tmp_path):
    aligns = {"left": ("left", "l"), "center": ("center", "ctr"), "right": ("right", "r"), "justify": ("left", "just")}
    shapes = written(tmp_path, [element(align, "equation", [math()], style={"align": align}) for align in aligns])
    for align, (jc, algn) in aligns.items():
        shape = shapes[align]
        assert shape.find(".//a14:m/m:oMathPara/m:oMathParaPr/m:jc", NS).get(f"{{{NS['m']}}}val") == jc
        assert shape.find(".//a:pPr", NS).get("algn") == algn
        # 独立した数式の段落には通常文字を置かず、段落の末尾の書式に要素の書体を書く。
        assert not text_runs(shape)
        assert shape.find(".//a:p/a:endParaRPr/a:latin", NS).get("typeface") == "Arial"


def test_math_only_text_gets_zero_width_run(tmp_path):
    elements = [
        element("text", content=[math()]),
        element("box", "box", [math("y")], shape="rectangle"),
        element("two", content=[math("a"), {"kind": "text", "text": "\n"}, math("b")]),
        element("mixed", content=[{"kind": "text", "text": "式 "}, math()]),
        element("spaced", "equation", [{"kind": "text", "text": " "}, math()]),
    ]
    shapes = written(tmp_path, elements)
    for name in ("text", "box", "two"):
        shape = shapes[name]
        assert shape.find(".//m:oMathPara", NS) is None
        runs = text_runs(shape)
        # 明示改行で分けた数式だけの行ごとに、その行の後に置く。
        assert [run.findtext("a:t", namespaces=NS) for run in runs] == [ZERO_WIDTH_SPACE] * (
            2 if name == "two" else 1
        ), name
        # 段落の最後（段落の末尾の書式の前）に、要素の文字の書式で置く。
        paragraph = shape.find(".//a:p", NS)
        assert paragraph[-2] is runs[-1] and paragraph[-1].tag == f"{{{NS['a']}}}endParaRPr"
        assert all(run.find("a:rPr", NS).get("sz") == "2400" for run in runs)
    two = shapes["two"].find(".//a:p", NS)
    assert [etree.QName(child).localname for child in two if etree.QName(child).localname != "pPr"] == [
        "m",
        "r",
        "br",
        "m",
        "r",
        "endParaRPr",
    ]
    # 通常の文字がある段落（空白だけでも）には加えない。
    assert [run.findtext("a:t", namespaces=NS) for run in text_runs(shapes["mixed"])] == ["式 "]
    assert [run.findtext("a:t", namespaces=NS) for run in text_runs(shapes["spaced"])] == [" "]
    assert shapes["spaced"].find(".//m:oMathPara", NS) is None
    assert ZERO_WIDTH_SPACE not in "".join(
        item[1] for item in paragraph_content(element(content=[{"kind": "text", "text": "x"}]))
    )


def test_centered_display_math_is_shifted_by_trailing_space(tmp_path):
    fonts = FakeFonts()  # 空白の送り幅は 0.5 em
    elements = [
        element("center", "equation", [math()], style={"align": "center"}),
        element("large", "equation", [math(style={"size_pt": 32})], style={"align": "center"}),
        element("left", "equation", [math()], style={"align": "left"}),
        element("right", "equation", [math()], style={"align": "right"}),
        element("inline", content=[math()], style={"align": "center"}),
    ]
    shapes = written(tmp_path, elements, fonts=fonts)

    def indent(name):
        value = shapes[name].find(".//a:pPr", NS).get("marL")
        return None if value is None else int(value) / EMU

    # PowerPoint が式の後ろに加える幅（空白の送り幅 × 数式の文字サイズ）を左の余白にする。中央揃えの独立した数式だけ。
    assert (indent("center"), indent("large")) == (12, 16)
    assert (indent("left"), indent("right"), indent("inline")) == (None, None, None)
    # fonts がなければ空白の送り幅が分からないので補正しない。
    plain = written(tmp_path, elements[:1], name="plain.pptx")
    assert plain["center"].find(".//a:pPr", NS).get("marL") is None
    assert plain["center"].find(".//a:endParaRPr/a:latin", NS).get("typeface") == "Arial"
    # 幅の規則：空白の送り幅（em）× 数式の文字サイズ。
    assert display_trailing_width(0.2778, 20) == pytest.approx(5.556)
    for advance, size in ((-0.1, 20), (0.25, 0), (float("nan"), 20), (True, 20)):
        with pytest.raises(ValueError):
            display_trailing_width(advance, size)


def test_line_style_and_description_use_scene_keys(tmp_path):
    elements = [
        element("plain", "line", start=[10, 10], end=[100, 10], style={"line_width_pt": 2.5}),
        element("dashed", "line", start=[10, 30], end=[100, 30], style={"dashed": True, "line_color": "183153"}),
        element("old", "line", start=[10, 50], end=[100, 50], dashed=True, alt_text="旧い説明"),
        element("described", description="図の説明"),
    ]
    root_shapes = written(tmp_path, elements)
    connectors = {
        node.find(".//p:cNvPr", NS).get("name"): node
        for node in xml_part(tmp_path / "deck.pptx").iter(f"{{{NS['p']}}}cxnSp")
    }
    line = connectors["plain"].find(".//a:ln", NS)
    assert (int(line.get("w")) / EMU, line.find("a:solidFill/a:srgbClr", NS).get("val")) == (2.5, "17212D")
    assert connectors["dashed"].find(".//a:ln/a:prstDash", NS).get("val") == "dash"
    assert connectors["dashed"].find(".//a:ln/a:solidFill/a:srgbClr", NS).get("val") == "183153"
    # 要素の直下の旧い別名（dashed・alt_text）は読まない。
    assert connectors["old"].find(".//a:ln/a:prstDash", NS) is None
    assert connectors["old"].find(".//p:cNvPr", NS).get("descr") is None
    assert root_shapes["described"].find(".//p:cNvPr", NS).get("descr") == "図の説明"


def test_connection_site_numbers_follow_preset_shapes():
    sides = ("top", "left", "bottom", "right")
    for target in (
        {"kind": "text"},
        {"kind": "equation"},
        {"kind": "box", "shape": "rectangle"},
        {"kind": "box", "shape": "rounded_rectangle"},
        {"kind": "box", "shape": "diamond"},
        {"kind": "box", "shape": "rectangle", "rotation": 90, "flip_h": True},
    ):
        assert [connection_site(target, side) for side in sides] == [0, 1, 2, 3], target
    assert [connection_site({"kind": "box", "shape": "ellipse"}, side) for side in sides] == [0, 2, 4, 6]
    assert connection_site({"kind": "box", "shape": "rectangle"}, "center") is None
    cover = {"kind": "image", "fit": "cover", "area": [0, 0, 100, 50], "actual_area": [0, 0, 100, 50]}
    matching = {"kind": "image", "fit": "contain", "area": [0, 0, 100, 50], "actual_area": [0, 0, 100, 50]}
    letterbox = {"kind": "image", "fit": "contain", "area": [0, 0, 100, 50], "actual_area": [25, 0, 50, 50]}
    assert [connection_site(image, "left") for image in (cover, matching, letterbox)] == [1, 1, None]
    assert connection_site({"kind": "image", "fit": "contain", "area": [0, 0, 100, 50]}, "left") is None
    assert connection_site({"kind": "group"}, "top") is None
    assert connection_site({"kind": "line"}, "top") is None


def test_template_keeps_theme_but_removes_old_slides_and_metadata(tmp_path):
    template = Presentation()
    old = template.slides.add_slide(template.slide_layouts[6])
    old.shapes.add_textbox(Pt(10), Pt(10), Pt(100), Pt(40)).text = "旧スライド"
    old.notes_slide.notes_text_frame.text = "旧ノート"
    template.core_properties.author = "雛形の著者"
    template.core_properties.last_modified_by = "Steve Canny"
    path = tmp_path / "template.pptx"
    template.save(path)
    source = path.read_bytes()
    result = write_pptx(
        scene([element()]),
        project(tmp_path, template={"path": "template.pptx", "layout_index": 6}),
        tmp_path / "deck.pptx",
    )
    assert path.read_bytes() == source
    assert result["template"]["removed_slides"] == 1
    with zipfile.ZipFile(path) as original, zipfile.ZipFile(tmp_path / "deck.pptx") as output:
        assert original.read("ppt/theme/theme1.xml") == output.read("ppt/theme/theme1.xml")
        all_xml = b"".join(output.read(name) for name in output.namelist() if name.endswith(".xml"))
        assert "旧スライド".encode() not in all_xml
        assert "旧ノート".encode() not in all_xml
        assert b"Steve Canny" not in output.read("docProps/core.xml")
    assert len(Presentation(tmp_path / "deck.pptx").slides) == 1
    with pytest.raises(ValueError, match="空白レイアウト"):
        write_pptx(
            scene([element()]),
            project(tmp_path, template={"path": "template.pptx", "layout_index": 0}),
            tmp_path / "invalid.pptx",
        )


def test_template_inspection_from_config_is_not_repeated(tmp_path):
    # 資料の読み込み（config）が空白レイアウトを確かめた結果（template["inspection"]）があれば、書き出しは調べ直さない。
    Presentation().save(tmp_path / "template.pptx")
    inspected = {
        "path": "template.pptx",
        "layout_index": 0,
        "inspection": {
            "layout_index": 0,
            "layout_name": "Title Slide",
            "original_slide_count": 0,
            "blank_layout": True,
        },
    }
    result = write_pptx(scene([element()]), project(tmp_path, template=inspected), tmp_path / "deck.pptx")
    assert result["template"]["layout_index"] == 0
    # 検査の結果が別のレイアウトのものなら、書き出しが確かめる。
    inspected["inspection"] = {**inspected["inspection"], "layout_index": 6}
    with pytest.raises(ValueError, match="空白レイアウト"):
        write_pptx(scene([element()]), project(tmp_path, template=inspected), tmp_path / "other.pptx")


def test_image_contain_and_cover_keep_aspect_ratio(tmp_path):
    source = tmp_path / "picture.png"
    Image.new("RGB", (400, 200), "white").save(source)
    data = scene(
        [
            element("contain", "image", path=str(source), fit="contain", area=[0, 0, 100, 100]),
            element("cover", "image", path=str(source), fit="cover", area=[200, 0, 100, 100]),
        ]
    )
    write_pptx(data, project(tmp_path), tmp_path / "images.pptx")
    pictures = Presentation(tmp_path / "images.pptx").slides[0].shapes
    assert tuple(value.pt for value in (pictures[0].left, pictures[0].top, pictures[0].width, pictures[0].height)) == (
        0,
        25,
        100,
        50,
    )
    assert pictures[1].crop_left == pytest.approx(0.25)
    assert pictures[1].crop_right == pytest.approx(0.25)


def test_math_rotation_is_rejected(tmp_path):
    with pytest.raises(ValueError, match="回転・反転"):
        write_pptx(
            scene([element(content=[{"kind": "math", "text": "x"}], rotation=20)]),
            project(tmp_path),
            tmp_path / "rotated.pptx",
        )


def test_pptx_outputs_follow_umask_and_leave_no_temporary_file(tmp_path):
    output, copy = tmp_path / "out" / "deck.pptx", tmp_path / "out" / "copy.pptx"
    previous = os.umask(0o027)
    try:
        write_pptx(scene([element()]), project(tmp_path), output)
        write_package(read_package(output), copy)
    finally:
        os.umask(previous)
    assert sorted(item.name for item in output.parent.iterdir()) == ["copy.pptx", "deck.pptx"]
    assert read_package(copy) == read_package(output)
    if os.name != "nt":
        assert {path.stat().st_mode & 0o777 for path in (output, copy)} == {0o640}
