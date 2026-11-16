"""PPTX の検査が、参照の欠落、文字・ノート・数式の不一致、文字の縮小、書き出しの規則との違いを取りこぼさないことを確かめる。"""

from types import SimpleNamespace
from zipfile import ZipFile

from lxml import etree
import pytest

from fake_fonts import FakeFonts
from slidekit.ooxml import NS
from slidekit.pptx_validation import _normalize_text, validate_pptx

STYLE = {
    "size_pt": 20,
    "min_size_pt": 18,
    "font_latin": "Arial",
    "font_japanese": "Yu Gothic UI",
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
SLIDE = "ppt/slides/slide1.xml"


@pytest.fixture
def project(tmp_path):
    return SimpleNamespace(
        root=tmp_path,
        config={
            "metadata": {"title": "Example", "author": "", "last_modified_by": ""},
            "inputs": {"documents": [{"id": "paper"}]},
            "math": {"fallback": "off"},
        },
        theme={"roles": {"body": {"size_pt": 20, "min_size_pt": 18}}},
        rules={},
        environment={},
    )


def element(id="body", *, area=(20, 20, 180, 40), text="Sample", **extra):
    return {
        "id": id,
        "kind": "text",
        "area": list(area),
        "role": "body",
        "style": {
            "size_pt": 20,
            "min_size_pt": 18,
            "font_latin": "Arial",
            "font_japanese": "Arial",
            "font_math": "Cambria Math",
            "padding_pt": 0,
            "line_spacing": 1.0,
        },
        "content": [{"kind": "text", "text": text}],
        **extra,
    }


def styled(id="body", kind="text", content=None, *, style=None, area=(20, 20, 300, 60), **extra):
    """作成 API が作る形の要素（書式は解決済みの値をすべて持つ）。"""
    return {
        "id": id,
        "kind": kind,
        "area": list(area),
        "role": "body",
        "style": {**STYLE, **(style or {})},
        "content": content or [{"kind": "text", "text": "Sample"}],
        **extra,
    }


def math(expression="x+1", **fragment):
    return {"kind": "math", "text": expression, **fragment}


def scene(*elements, notes="説明します。"):
    return {
        "schema_version": 1,
        "size": [720, 540],
        "metadata": {"title": "Example", "author": "", "last_modified_by": ""},
        "slides": [{"id": "first", "kind": "content", "notes": notes, "elements": list(elements)}],
    }


def codes(issues):
    return {issue["code"] for issue in issues}


def located(issues, code):
    """その識別子の診断の要素の識別子。"""
    return sorted(issue["element_id"] for issue in issues if issue["code"] == code)


def pptx_file(tmp_path, scene_data):
    from pptx import Presentation
    from pptx.util import Pt
    from pptx.oxml.xmlchemy import OxmlElement
    from pptx.enum.text import MSO_AUTO_SIZE

    prs = Presentation()
    prs.slide_width, prs.slide_height = [Pt(value) for value in scene_data["size"]]
    for field in (
        "title",
        "subject",
        "author",
        "keywords",
        "comments",
        "last_modified_by",
        "category",
        "content_status",
    ):
        setattr(prs.core_properties, field, scene_data["metadata"].get(field, ""))
    for source in scene_data["slides"]:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide._element.cSld.set("name", source["id"])
        for item in source["elements"]:
            shape = slide.shapes.add_textbox(*[Pt(value) for value in item["area"]])
            shape.name = item["id"]
            shape.text_frame.word_wrap = False
            shape.text_frame.auto_size = MSO_AUTO_SIZE.NONE
            run = shape.text_frame.paragraphs[0].add_run()
            run.text = "".join(part["text"] for part in item["content"] if part["kind"] != "math")
            run.font.name = item["style"]["font_latin"]
            run.font.size = Pt(item["style"]["size_pt"])
            east_asian = OxmlElement("a:ea")
            east_asian.set("typeface", item["style"]["font_japanese"])
            run._r.get_or_add_rPr().append(east_asian)
        slide.notes_slide.notes_text_frame.text = source["notes"]
    output = tmp_path / "deck.pptx"
    prs.save(output)
    return output


def written(tmp_path, project, data, *, fonts=None, name="deck.pptx"):
    from slidekit.writer import write_pptx

    path = tmp_path / name
    write_pptx(data, project, path, fonts=fonts)
    return path


def rewrite_zip(path, replacements):
    with ZipFile(path) as source:
        contents = {name: source.read(name) for name in source.namelist()}
    contents.update(replacements)
    with ZipFile(path, "w") as target:
        for name, content in contents.items():
            target.writestr(name, content)


def edit_slide(path, change):
    """スライドの XML を lxml で読み、change(根, 図形の名前から p:sp への辞書) で書き換えて保存する。"""
    with ZipFile(path) as source:
        root = etree.fromstring(source.read(SLIDE))
    change(root, {node.find(".//p:cNvPr", NS).get("name"): node for node in root.iter(f"{{{NS['p']}}}sp")})
    rewrite_zip(path, {SLIDE: etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)})


def test_valid_pptx_and_metadata_contamination(project, tmp_path):
    data = scene(element())
    path = pptx_file(tmp_path, data)
    assert validate_pptx(path, data, project) == []
    project.config["metadata"]["last_modified_by"] = "正しい更新者"
    issues = validate_pptx(path, data, project)
    assert "metadata_mismatch" in codes(issues)


def test_pptx_detects_broken_relationship_notes_order_and_editable_math(project, tmp_path):
    data = scene(element())
    path = pptx_file(tmp_path, data)
    with ZipFile(path) as source:
        rels = source.read("ppt/slides/_rels/slide1.xml.rels").replace(
            b"../notesSlides/notesSlide1.xml", b"../notesSlides/missing.xml"
        )
        slide_xml = source.read("ppt/slides/slide1.xml").replace(b'name="first"', b'name="other"')
    rewrite_zip(path, {"ppt/slides/_rels/slide1.xml.rels": rels, "ppt/slides/slide1.xml": slide_xml})
    data["slides"][0]["elements"][0]["content"].append({"kind": "math", "text": "x"})
    project.config["math"]["fallback"] = "required"
    issues = validate_pptx(path, data, project)
    assert {
        "package_reference_missing",
        "notes_mismatch",
        "slide_order_mismatch",
        "editable_math_missing",
        "math_fallback_missing",
    } <= codes(issues)


def test_pptx_text_is_not_silently_lost(project, tmp_path):
    data = scene(element())
    path = pptx_file(tmp_path, data)
    data["slides"][0]["elements"][0]["content"][0]["text"] = "different text"
    assert "text_mismatch" in codes(validate_pptx(path, data, project))


def test_pptx_cannot_silently_shrink_text(project, tmp_path):
    data = scene(element())
    path = pptx_file(tmp_path, data)
    with ZipFile(path) as source:
        slide_xml = source.read("ppt/slides/slide1.xml").replace(b'sz="2000"', b'sz="1200"')
    rewrite_zip(path, {"ppt/slides/slide1.xml": slide_xml})
    assert "font_size_mismatch" in codes(validate_pptx(path, data, project))


def test_writer_final_math_is_validated_and_changed_equation_is_detected(project, tmp_path):
    from pdf_tools import Document
    from slidekit.fallback import finalize_math
    from slidekit.writer import write_pptx

    text = element()
    text["content"].append({"kind": "math", "text": r"\frac{x_i}{2}"})
    data = scene(text)
    native, pdf, final = [tmp_path / name for name in ("native.pptx", "native.pdf", "final.pptx")]
    write_pptx(data, project, native)
    with Document() as document:
        document.new_page(width=720, height=540)
        document.save(pdf)
    finalize_math(native, pdf, final)
    project.config["math"]["fallback"] = "required"
    assert validate_pptx(final, data, project) == []
    data["slides"][0]["elements"][0]["content"][-1]["text"] = r"\frac{y_i}{2}"
    assert "math_content_mismatch" in codes(validate_pptx(final, data, project))


def test_pptx_wrap_mismatch(project, tmp_path):
    data = scene(styled("wrapped"), styled("fixed", style={"word_wrap": False}, area=(20, 100, 300, 60)))
    path = written(tmp_path, project, data)
    assert validate_pptx(path, data, project) == []

    def unwrap(root, shapes):
        shapes["wrapped"].find(".//a:bodyPr", NS).set("wrap", "none")
        del shapes["fixed"].find(".//a:bodyPr", NS).attrib["wrap"]  # wrap がなければ square とみなす

    edit_slide(path, unwrap)
    issues = validate_pptx(path, data, project)
    assert codes(issues) == {"text_wrap_mismatch"} and located(issues, "text_wrap_mismatch") == ["fixed", "wrapped"]
    observed = next(issue for issue in issues if issue["element_id"] == "fixed")["observations"]
    assert observed == [{"expected": "none", "actual": "square"}]


def test_automatic_text_adjustment_requires_no_autofit(project, tmp_path):
    data = scene(styled())
    path = written(tmp_path, project, data)

    def autofit(root, shapes):
        body = shapes["body"].find(".//a:bodyPr", NS)
        body.remove(body.find("a:noAutofit", NS))
        etree.SubElement(body, f"{{{NS['a']}}}normAutofit")

    edit_slide(path, autofit)
    issues = validate_pptx(path, data, project)
    assert codes(issues) == {"automatic_text_adjustment"} and located(issues, "automatic_text_adjustment") == ["body"]


def test_geometry_mismatch_is_detected(project, tmp_path):
    data = scene(styled())
    path = written(tmp_path, project, data)
    edit_slide(path, lambda root, shapes: shapes["body"].find(".//a:xfrm/a:off", NS).set("x", str(25 * 12700)))
    issues = validate_pptx(path, data, project)
    assert codes(issues) == {"geometry_mismatch"}
    assert issues[0]["observations"] == [{"expected_area": [20, 20, 300, 60], "actual_area": [25, 20, 300, 60]}]


def test_page_size_mismatch_is_detected(project, tmp_path):
    data = scene(styled())
    path = written(tmp_path, project, data)
    with ZipFile(path) as source:
        presentation = source.read("ppt/presentation.xml").replace(b'cx="9144000"', b'cx="9906000"')
    rewrite_zip(path, {"ppt/presentation.xml": presentation})
    assert codes(validate_pptx(path, data, project)) == {"page_size_mismatch"}


def test_slide_count_mismatch_is_detected(project, tmp_path):
    data = scene(styled())
    path = written(tmp_path, project, data)
    data["slides"].append({"id": "second", "kind": "content", "notes": "", "elements": []})
    issues = validate_pptx(path, data, project)
    assert codes(issues) == {"slide_count_mismatch"}
    assert issues[0]["observations"] == [{"expected": 2, "actual": 1}]


def test_font_check_accepts_forced_runs(project, tmp_path):
    fonts = FakeFonts(missing={"Arial": {"※"}})
    data = scene(styled(content=[{"kind": "text", "text": "注※記"}]))
    path = written(tmp_path, project, data, fonts=fonts)
    # 欧文の書体にない ※ は、a:latin と a:ea の両方が和文の書体の区切りになる。同じ fonts で照合すれば受け付ける。
    assert validate_pptx(path, data, project, fonts=fonts) == []
    # fonts がなければ、すべての区切りが要素の書体であることを求める（0.1 の規則）。
    assert codes(validate_pptx(path, data, project)) == {"font_mismatch"}

    def latin(root, shapes):
        forced = shapes["body"].findall(".//a:p/a:r", NS)[1]
        forced.find("a:rPr/a:latin", NS).set("typeface", "Arial")

    edit_slide(path, latin)
    issues = validate_pptx(path, data, project, fonts=fonts)
    assert codes(issues) == {"font_mismatch"}
    assert issues[0]["observations"] == [
        {"run": 2, "text": "※", "field": "font_latin", "expected": "Yu Gothic UI", "actual": "Arial"}
    ]

    def shift(root, shapes):
        # 区切りの数と書式は変えず、※ を欧文の書体の区切りへ移す（forced の区切りには「記」が入る）。
        for run, text in zip(shapes["body"].findall(".//a:p/a:r", NS), ("注※", "記", "")):
            run.find("a:t", NS).text = text

    shifted = written(tmp_path, project, data, fonts=fonts, name="shifted.pptx")
    edit_slide(shifted, shift)
    issues = validate_pptx(shifted, data, project, fonts=fonts)
    # 全体の文字は同じなので text_mismatch にはならない。区切りごとの文字の違いとして見つける。
    assert codes(issues) == {"font_mismatch"}
    assert sorted(issues[0]["observations"], key=lambda item: item["run"]) == [
        {"run": 1, "field": "text", "expected": "注", "actual": "注※"},
        {"run": 2, "field": "text", "expected": "※", "actual": "記"},
        {"run": 3, "field": "text", "expected": "記", "actual": ""},
    ]


def test_font_check_expects_japanese_font_for_symbols_powerpoint_draws_in_japanese_font(project, tmp_path):
    fonts = FakeFonts()
    data = scene(styled(content=[{"kind": "text", "text": "可○ 否× 例→ 間—"}]))
    path = written(tmp_path, project, data, fonts=fonts)
    # 書き出しと同じ判定（measure.font_runs）で期待値を作るので、同じ fonts で照合すれば受け付ける。
    assert validate_pptx(path, data, project, fonts=fonts) == []

    def previous_rule(root, shapes):
        # ○ × の区切りを、改修の途中の書き出し（幅の曖昧な文字は欧文の書体）と同じく欧文の書体にする。
        # → の区切りは和文の書体にする（→ — は欧文の書体のままであることを求める）。
        runs = {run.findtext("a:t", namespaces=NS): run for run in shapes["body"].findall(".//a:p/a:r", NS)}
        for text, family in (("○", "Arial"), ("×", "Arial"), ("→", "Yu Gothic UI")):
            for kind in ("latin", "ea"):
                runs[text].find(f"a:rPr/a:{kind}", NS).set("typeface", family)

    edit_slide(path, previous_rule)
    issues = validate_pptx(path, data, project, fonts=fonts)
    assert codes(issues) == {"font_mismatch"}
    assert sorted(issues[0]["observations"], key=lambda item: (item["run"], item["field"])) == [
        {"run": 2, "text": "○", "field": "font_japanese", "expected": "Yu Gothic UI", "actual": "Arial"},
        {"run": 2, "text": "○", "field": "font_latin", "expected": "Yu Gothic UI", "actual": "Arial"},
        {"run": 4, "text": "×", "field": "font_japanese", "expected": "Yu Gothic UI", "actual": "Arial"},
        {"run": 4, "text": "×", "field": "font_latin", "expected": "Yu Gothic UI", "actual": "Arial"},
        {"run": 6, "text": "→", "field": "font_japanese", "expected": "Arial", "actual": "Yu Gothic UI"},
        {"run": 6, "text": "→", "field": "font_latin", "expected": "Arial", "actual": "Yu Gothic UI"},
    ]


def test_font_size_checked_per_fragment(project, tmp_path):
    content = [
        {"kind": "text", "text": "通常の大きさ "},
        {"kind": "text", "text": "小さい注記", "style": {"size_pt": 14}},
    ]
    data = scene(styled(content=content))
    path = written(tmp_path, project, data)
    # 断片の大きさ（14 pt）は要素の大きさ（20 pt）と違っても受け付ける。
    assert validate_pptx(path, data, project) == []
    edit_slide(
        path, lambda root, shapes: shapes["body"].findall(".//a:p/a:r", NS)[1].find("a:rPr", NS).set("sz", "2000")
    )
    issues = validate_pptx(path, data, project)
    assert codes(issues) == {"font_size_mismatch"}
    assert issues[0]["observations"] == [
        {"run": 2, "text": "小さい注記", "expected_size_pt": 14, "actual_size_pt": 20.0}
    ]

    def shift(root, shapes):
        # 区切りの大きさは変えず、14 pt の断片の先頭の 4 文字を 20 pt の区切りへ移す。
        for run, text in zip(shapes["body"].findall(".//a:p/a:r", NS), ("通常の大きさ 小さい注", "記")):
            run.find("a:t", NS).text = text

    shifted = written(tmp_path, project, data, name="shifted.pptx")
    edit_slide(shifted, shift)
    issues = validate_pptx(shifted, data, project)
    assert codes(issues) == {"font_mismatch"}
    assert sorted(issues[0]["observations"], key=lambda item: item["run"]) == [
        {"run": 1, "field": "text", "expected": "通常の大きさ ", "actual": "通常の大きさ 小さい注"},
        {"run": 2, "field": "text", "expected": "小さい注記", "actual": "記"},
    ]


def test_run_style_and_split_are_checked(project, tmp_path):
    content = [
        {"kind": "text", "text": "強調 "},
        {"kind": "emphasis", "text": "語"},
        {
            "kind": "text",
            "text": "リンク",
            "style": {"underline": True, "color": "183153"},
            "link": "https://example.com/",
        },
    ]
    data = scene(styled(content=content))
    path = written(tmp_path, project, data)
    assert validate_pptx(path, data, project) == []

    def restyle(root, shapes):
        runs = shapes["body"].findall(".//a:p/a:r", NS)
        runs[1].find("a:rPr", NS).set("b", "0")
        runs[2].find("a:rPr/a:solidFill/a:srgbClr", NS).set("val", "17212D")
        del runs[2].find("a:rPr", NS).attrib["u"]

    edit_slide(path, restyle)
    observations = validate_pptx(path, data, project)[0]["observations"]
    assert sorted((item["run"], item["field"]) for item in observations) == [
        (2, "bold"),
        (3, "color"),
        (3, "underline"),
    ]
    # 区切りの数が違えば、区切りごとには照合しない。
    edit_slide(
        path, lambda root, shapes: shapes["body"].find(".//a:p", NS).remove(shapes["body"].findall(".//a:p/a:r", NS)[1])
    )
    issues = validate_pptx(path, data, project)
    assert {"font_mismatch", "text_mismatch"} == codes(issues)
    assert next(issue for issue in issues if issue["code"] == "font_mismatch")["observations"][0]["expected_runs"] == [
        "強調 ",
        "語",
        "リンク",
    ]


def test_math_content_uses_fragment_style(project, tmp_path):
    data = scene(
        styled(content=[{"kind": "text", "text": "大きい式 "}, math("a", style={"size_pt": 32, "bold": True})])
    )
    path = written(tmp_path, project, data)
    assert validate_pptx(path, data, project) == []
    data["slides"][0]["elements"][0]["content"][1]["style"] = {"size_pt": 20}
    assert codes(validate_pptx(path, data, project)) == {"math_content_mismatch"}


def layout_scene():
    aligns = ("left", "center", "right", "justify")
    elements = [
        styled(
            align, "equation", [math()], style={"align": align, "word_wrap": False}, area=(20, 20 + 70 * index, 300, 60)
        )
        for index, align in enumerate(aligns)
    ]
    elements += [
        styled("text", content=[math("y")], style={"align": "center"}, area=(380, 20, 300, 60)),
        styled("box", "box", [math("z")], shape="ellipse", area=(380, 100, 300, 60)),
        styled("mixed", content=[{"kind": "text", "text": "式 "}, math()], area=(380, 180, 300, 60)),
    ]
    return scene(*elements)


def test_validate_pptx_accepts_layout_modes(project, tmp_path):
    fonts = FakeFonts()
    data = layout_scene()
    path = written(tmp_path, project, data, fonts=fonts)
    assert validate_pptx(path, data, project, fonts=fonts) == []
    # fonts がなければ、中央揃えの補正の量は照合しない。
    assert validate_pptx(path, data, project) == []
    plain = written(tmp_path, project, data, name="plain.pptx")
    assert validate_pptx(plain, data, project) == []


def test_validate_pptx_detects_layout_mismatch(project, tmp_path):
    fonts = FakeFonts()
    data = layout_scene()

    def mismatch(change):
        path = written(tmp_path, project, data, fonts=fonts)
        edit_slide(path, change)
        return validate_pptx(path, data, project, fonts=fonts)

    def justification(root, shapes):
        shapes["right"].find(".//m:jc", NS).set(f"{{{NS['m']}}}val", "left")

    def zero_width(root, shapes):
        paragraph = shapes["text"].find(".//a:p", NS)
        paragraph.remove(paragraph.find("a:r", NS))

    def indent(root, shapes):
        del shapes["center"].find(".//a:pPr", NS).attrib["marL"]

    def end_font(root, shapes):
        shapes["left"].find(".//a:endParaRPr/a:latin", NS).set("typeface", "Meiryo")

    def inline(root, shapes):
        wrapper = shapes["justify"].find(".//a14:m", NS)
        paragraph = wrapper.find("m:oMathPara", NS)
        wrapper.replace(paragraph, paragraph.find("m:oMath", NS))

    for change, name in (
        (justification, "right"),
        (zero_width, "text"),
        (indent, "center"),
        (end_font, "left"),
        (inline, "justify"),
    ):
        issues = mismatch(change)
        assert located(issues, "math_layout_mismatch") == [name], change.__name__
    issue = next(item for item in mismatch(indent) if item["code"] == "math_layout_mismatch")
    assert issue["observations"] == [{"field": "a:pPr/@marL", "expected_pt": 10.0, "actual_pt": 0.0}]
    # 独立した数式の要素を、文中の数式の場面と照合しても見つける。
    path = written(tmp_path, project, data, fonts=fonts)
    data["slides"][0]["elements"][0]["kind"] = "text"
    assert located(validate_pptx(path, data, project, fonts=fonts), "math_layout_mismatch") == ["left"]


def test_fallback_expected_false_skips_missing_fallback(project, tmp_path):
    data = scene(styled(content=[math()]))
    path = written(tmp_path, project, data)
    assert "math_fallback_missing" not in codes(validate_pptx(path, data, project))
    project.config["math"]["fallback"] = "required"
    assert located(validate_pptx(path, data, project), "math_fallback_missing") == ["body"]
    # 下書き（互換画像を作らない）では互換画像の欠落を報告しない。
    assert validate_pptx(path, data, project, fallback_expected=False) == []
    project.config["math"]["fallback"] = "off"
    assert "math_fallback_missing" in codes(validate_pptx(path, data, project, fallback_expected=True))


def test_diagnostics_carry_slide_number(project, tmp_path):
    data = scene(styled())
    path = written(tmp_path, project, data)
    data["slides"][0]["elements"][0]["content"][0]["text"] = "changed"
    data["slides"][0]["number"] = 7
    project.config["metadata"]["title"] = "別の題名"
    issues = {issue["code"]: issue for issue in validate_pptx(path, data, project)}
    # 部分生成の場面でも、診断のページ番号は元のページ番号になる。資料全体の診断には付けない。
    assert issues["text_mismatch"]["slide_number"] == 7
    assert "slide_number" not in issues["metadata_mismatch"]
    del data["slides"][0]["number"]
    assert (
        next(issue for issue in validate_pptx(path, data, project) if issue["code"] == "text_mismatch")["slide_number"]
        == 1
    )


def test_normalize_text_removes_spaces_and_format_characters_like_pdf_layout():
    from slidekit.pdf_layout import _normalize

    samples = ["式​", "a⁠b﻿c", "全角　空白\tと改行\n", "ｶﾀｶﾅ ①ﬁ", "é", "​", ""]
    assert [_normalize_text(value) for value in samples] == ["式", "abc", "全角空白と改行", "カタカナ1fi", "é", "", ""]
    assert [_normalize_text(value) for value in samples] == [_normalize(value) for value in samples]


def group(id, *, parent=None, pptx_group=True):
    element = {
        "id": id,
        "kind": "group",
        "area": [0, 0, 0, 0],
        "role": None,
        "style": {},
        "content": [],
        "pptx_group": pptx_group,
    }
    if parent is not None:
        element["parent_id"] = parent
    return element


def boxed(id, area, *, parent=None, text="図形", **extra):
    element = styled(id, "box", [{"kind": "text", "text": text}], area=area, shape="rectangle", **extra)
    if parent is not None:
        element["parent_id"] = parent
    return element


def connector(id, start, end, **references):
    area = [min(start[0], end[0]), min(start[1], end[1]), abs(end[0] - start[0]), abs(end[1] - start[1])]
    return {
        "id": id,
        "kind": "line",
        "area": area,
        "role": None,
        "style": dict(STYLE),
        "content": [],
        "start": list(start),
        "end": list(end),
        **references,
    }


def group_scene(*, outer=True):
    return scene(
        group("diagram", pptx_group=outer),
        boxed("diagram/first", (40, 60, 100, 50), parent="diagram"),
        group("diagram/inner", parent="diagram"),
        boxed("diagram/inner/second", (200, 60, 100, 50), parent="diagram/inner"),
        boxed("diagram/inner/third", (360, 60, 100, 50), parent="diagram/inner"),
        boxed("outside", (40, 300, 100, 50)),
    )


def groups_by_name(root):
    return {node.find("p:nvGrpSpPr/p:cNvPr", NS).get("name"): node for node in root.iter(f"{{{NS['p']}}}grpSp")}


def test_pptx_shapes_inside_groups_are_matched(project, tmp_path):
    data = group_scene()
    path = written(tmp_path, project, data)
    assert validate_pptx(path, data, project) == []

    def change(root, shapes):
        shapes["diagram/inner/second"].find(".//a:t", NS).text = "別の文字"
        shapes["diagram/first"].find("p:spPr/a:xfrm/a:off", NS).set("x", str(45 * 12700))

    edit_slide(path, change)
    issues = validate_pptx(path, data, project)
    # グループの中の図形も名前で対応づけ、ほかの図形と同じく照合する。
    assert codes(issues) == {"text_mismatch", "font_mismatch", "geometry_mismatch"}
    assert located(issues, "text_mismatch") == ["diagram/inner/second"]
    assert located(issues, "geometry_mismatch") == ["diagram/first"]


def test_element_missing_inside_group(project, tmp_path):
    data = group_scene()
    path = written(tmp_path, project, data)

    def remove(root, shapes):
        shapes["diagram/inner/third"].getparent().remove(shapes["diagram/inner/third"])

    edit_slide(path, remove)
    issues = validate_pptx(path, data, project)
    assert codes(issues) == {"element_missing"} and located(issues, "element_missing") == ["diagram/inner/third"]


def test_group_mismatch_detects_missing_or_wrong_group(project, tmp_path):
    data = group_scene()

    def mismatch(change, scene_data=data):
        path = written(tmp_path, project, data)
        edit_slide(path, change)
        return validate_pptx(path, scene_data, project)

    def ungroup_inner(root, shapes):
        inner = groups_by_name(root)["diagram/inner"]
        for child in [child for child in inner if child.tag == f"{{{NS['p']}}}sp"]:
            inner.addprevious(child)
        inner.getparent().remove(inner)

    issues = mismatch(ungroup_inner)
    assert codes(issues) == {"group_mismatch"}
    assert located(issues, "group_mismatch") == ["diagram/inner", "diagram/inner/second", "diagram/inner/third"]
    observed = next(issue for issue in issues if issue["element_id"] == "diagram/inner/second")["observations"]
    assert observed == [{"expected_groups": ["diagram", "diagram/inner"], "actual_groups": ["diagram"]}]

    def move_out(root, shapes):
        root.find("p:cSld/p:spTree", NS).append(shapes["diagram/first"])

    issues = mismatch(move_out)
    assert codes(issues) == {"group_mismatch"} and located(issues, "group_mismatch") == ["diagram/first"]

    def shift_group(root, shapes):
        groups_by_name(root)["diagram"].find("p:grpSpPr/a:xfrm/a:off", NS).set("x", str(50 * 12700))

    issues = mismatch(shift_group)
    # 子の座標はページ上の座標のままなので、座標変換が恒等でないグループは表示の位置がずれる。
    assert codes(issues) == {"group_mismatch"} and located(issues, "group_mismatch") == ["diagram"]
    assert issues[0]["observations"][0]["field"] == "a:xfrm"
    # 書き出さないまとまりのグループがあれば不一致にする（場面だけ pptx_group を偽にする）。
    flat = group_scene(outer=False)
    issues = mismatch(lambda root, shapes: None, flat)
    assert codes(issues) == {"group_mismatch"}
    assert located(issues, "group_mismatch") == [
        "diagram",
        "diagram/first",
        "diagram/inner",
        "diagram/inner/second",
        "diagram/inner/third",
    ]
    path = written(tmp_path, project, flat, name="flat.pptx")
    assert validate_pptx(path, flat, project) == []


def wrap_in_group(node, name):
    """図形を恒等の座標変換の p:grpSp で包む。"""
    xfrm = node.find("p:spPr/a:xfrm", NS)
    off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
    a, p = NS["a"], NS["p"]
    group = etree.Element(f"{{{p}}}grpSp")
    properties = etree.SubElement(group, f"{{{p}}}nvGrpSpPr")
    etree.SubElement(properties, f"{{{p}}}cNvPr", id="900", name=name)
    etree.SubElement(properties, f"{{{p}}}cNvGrpSpPr")
    etree.SubElement(properties, f"{{{p}}}nvPr")
    transform = etree.SubElement(etree.SubElement(group, f"{{{p}}}grpSpPr"), f"{{{a}}}xfrm")
    for tag, source in (("off", off), ("ext", ext), ("chOff", off), ("chExt", ext)):
        etree.SubElement(transform, f"{{{a}}}{tag}", dict(source.attrib))
    node.addprevious(group)
    group.append(node)


def test_math_inside_group_is_math_grouped(project, tmp_path):
    data = scene(
        styled("formula", content=[{"kind": "text", "text": "式 "}, math()]), boxed("plain", (40, 300, 100, 50))
    )
    path = written(tmp_path, project, data)
    assert validate_pptx(path, data, project) == []
    edit_slide(path, lambda root, shapes: wrap_in_group(shapes["formula"], "wrapper"))
    issues = validate_pptx(path, data, project)
    # 数式を含む図形は、座標変換が恒等でもグループに入れない（互換画像はスライド直下の図形だけに付ける）。
    assert located(issues, "math_grouped") == ["formula"]
    assert codes(issues) == {"math_grouped", "group_mismatch"}
    # グループに入れない図形は math_grouped にしない。
    path = written(tmp_path, project, data, name="plain.pptx")
    edit_slide(path, lambda root, shapes: wrap_in_group(shapes["plain"], "wrapper"))
    assert codes(validate_pptx(path, data, project)) == {"group_mismatch"}


def connection_scene():
    return scene(
        boxed("target", (300, 40, 100, 50)),
        boxed("other", (300, 200, 100, 50)),
        connector("link", (100, 65), (300, 65), end_ref={"element_id": "target", "side": "left"}),
        connector("free", (100, 150), (200, 150)),
    )


def test_connection_mismatch_detects_wrong_site(project, tmp_path):
    data = connection_scene()

    def mismatch(change):
        path = written(tmp_path, project, data)
        assert validate_pptx(path, data, project) == []
        edit_slide(path, change)
        return validate_pptx(path, data, project)

    def connectors(root):
        return {
            node.find(".//p:cNvPr", NS).get("name"): node.find("p:nvCxnSpPr/p:cNvCxnSpPr", NS)
            for node in root.iter(f"{{{NS['p']}}}cxnSp")
        }

    def wrong_site(root, shapes):
        connectors(root)["link"].find("a:endCxn", NS).set("idx", "3")

    issues = mismatch(wrong_site)
    assert codes(issues) == {"connection_mismatch"} and located(issues, "connection_mismatch") == ["link"]
    observed = issues[0]["observations"][0]
    assert (observed["end"], observed["expected"]["element_id"], observed["expected"]["site"]) == ("end", "target", 1)
    assert (observed["actual"]["element_id"], observed["actual"]["site"]) == ("target", 3)

    def wrong_target(root, shapes):
        other = shapes["other"].find(".//p:cNvPr", NS).get("id")
        connectors(root)["link"].find("a:endCxn", NS).set("id", other)

    issues = mismatch(wrong_target)
    assert located(issues, "connection_mismatch") == ["link"]
    assert issues[0]["observations"][0]["actual"]["element_id"] == "other"

    def missing(root, shapes):
        properties = connectors(root)["link"]
        properties.remove(properties.find("a:endCxn", NS))

    assert located(mismatch(missing), "connection_mismatch") == ["link"]

    def unexpected(root, shapes):
        # anchor() で結んでいない端に接続がある。
        target = shapes["target"].find(".//p:cNvPr", NS).get("id")
        etree.SubElement(connectors(root)["free"], f"{{{NS['a']}}}stCxn", id=target, idx="0")

    issues = mismatch(unexpected)
    assert located(issues, "connection_mismatch") == ["free"]
    assert issues[0]["observations"][0]["expected"] is None


def test_connection_to_math_is_kept_with_fallback(project, tmp_path):
    # 数式を含む接続先は、互換画像を付ける資料でも接続する。書き出しと検査は同じ判定（writer.anchored_site）を使う。
    data = scene(
        styled("formula", "equation", [math("x")], area=(300, 40, 100, 50)),
        connector("link", (100, 65), (300, 65), end_ref={"element_id": "formula", "side": "left"}),
    )
    project.config["math"]["fallback"] = "required"
    connected = written(tmp_path, project, data, name="connected.pptx")
    assert codes(validate_pptx(connected, data, project, fallback_expected=False)) == set()


def vector_scene(tmp_path):
    from PIL import Image

    png, svg = tmp_path / "picture.png", tmp_path / "picture.svg"
    Image.new("RGB", (200, 100), "white").save(png)
    svg.write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100"><rect width="200" height="100"/></svg>'
    )
    image = {
        "id": "picture",
        "kind": "image",
        "area": [40, 40, 200, 100],
        "actual_area": [40, 40, 200, 100],
        "role": None,
        "path": str(png),
        "fit": "contain",
        "vector_path": str(svg),
        "style": {},
        "content": [],
    }
    return scene(image)


def test_image_vector_missing(project, tmp_path):
    data = vector_scene(tmp_path)
    path = written(tmp_path, project, data)
    assert validate_pptx(path, data, project) == []

    def drop_extension(root, shapes):
        blip = root.find(".//p:pic/p:blipFill/a:blip", NS)
        blip.remove(blip.find("a:extLst", NS))

    edit_slide(path, drop_extension)
    issues = validate_pptx(path, data, project)
    assert codes(issues) == {"image_vector_missing"} and located(issues, "image_vector_missing") == ["picture"]
    assert issues[0]["observations"] == [{"svg_blip": False, "target": None}]
    # SVG の部品がない。
    path = written(tmp_path, project, data, name="parts.pptx")
    with ZipFile(path) as source:
        contents = {name: source.read(name) for name in source.namelist() if not name.endswith(".svg")}
    with ZipFile(path, "w") as target:
        for name, content in contents.items():
            target.writestr(name, content)
    issues = validate_pptx(path, data, project)
    assert {"image_vector_missing", "package_reference_missing"} == codes(issues)
    # vector_path のない画像には SVG を求めない。
    del data["slides"][0]["elements"][0]["vector_path"]
    plain = written(tmp_path, project, data, name="plain.pptx")
    assert validate_pptx(plain, data, project) == []


def test_line_style_mismatch(project, tmp_path):
    data = scene(
        styled("dashed", "box", shape="rectangle", style={"line_color": "183153", "dashed": True}),
        styled("solid", "box", shape="rectangle", style={"line_color": "183153"}, area=(20, 100, 300, 60)),
        connector("line", (40, 200), (300, 200)),
        styled("unlined", "box", shape="rectangle", style={"dashed": True}, area=(20, 300, 300, 60)),
    )
    data["slides"][0]["elements"][2]["style"]["dashed"] = True
    path = written(tmp_path, project, data)
    assert validate_pptx(path, data, project) == []

    def undash(root, shapes):
        line = shapes["dashed"].find("p:spPr/a:ln", NS)
        line.remove(line.find("a:prstDash", NS))
        etree.SubElement(shapes["solid"].find("p:spPr/a:ln", NS), f"{{{NS['a']}}}prstDash", val="sysDot")
        connector_line = next(node for node in root.iter(f"{{{NS['p']}}}cxnSp")).find("p:spPr/a:ln/a:prstDash", NS)
        connector_line.set("val", "solid")

    edit_slide(path, undash)
    issues = validate_pptx(path, data, project)
    assert codes(issues) == {"line_style_mismatch"}
    assert located(issues, "line_style_mismatch") == ["dashed", "line", "solid"]
    observed = {issue["element_id"]: issue["observations"][0] for issue in issues}
    assert observed["dashed"] == {"expected": "dash", "actual": "solid"}
    assert observed["solid"] == {"expected": "solid", "actual": "sysDot"}


def test_line_break_size_is_checked(project, tmp_path):
    data = scene(
        styled(
            content=[
                {"kind": "text", "text": "一行目\n"},
                {"kind": "text", "text": "\n三行目", "style": {"size_pt": 28}},
            ]
        )
    )
    path = written(tmp_path, project, data)
    assert validate_pptx(path, data, project) == []

    def unsized(root, shapes):
        shapes["body"].findall(".//a:p/a:br", NS)[1].find("a:rPr", NS).attrib.pop("sz")

    edit_slide(path, unsized)
    issues = validate_pptx(path, data, project)
    # a:br の文字サイズがないと、PowerPoint は空の行を雛形の既定の大きさで置く。
    assert codes(issues) == {"font_size_mismatch"}
    assert issues[0]["observations"] == [{"break": 2, "expected_size_pt": 28, "actual_size_pt": None}]

    def removed(root, shapes):
        paragraph = shapes["body"].find(".//a:p", NS)
        paragraph.remove(paragraph.find("a:br", NS))

    path = written(tmp_path, project, data, name="removed.pptx")
    edit_slide(path, removed)
    issues = validate_pptx(path, data, project)
    assert codes(issues) == {"text_mismatch"}
    assert issues[0]["observations"] == [{"expected_breaks": 2, "actual_breaks": 1}]


def test_inline_math_without_no_break_box_is_detected(project, tmp_path):
    # 文中の数式は m:box（m:noBreak）で包む（PowerPoint が式の途中で改行しないように）。包みを外した PPTX は不一致にする。
    data = scene(styled(content=[{"kind": "text", "text": "式 "}, math("a+b")]))
    path = written(tmp_path, project, data)
    assert validate_pptx(path, data, project) == []

    def unwrap(root, shapes):
        box = shapes["body"].find(".//m:oMath/m:box", NS)
        equation = box.getparent()
        for child in list(box.find("m:e", NS)):
            equation.append(child)
        equation.remove(box)

    edit_slide(path, unwrap)
    assert codes(validate_pptx(path, data, project)) == {"math_content_mismatch"}
