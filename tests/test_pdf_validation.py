"""PDF の検査が、文字と数式の欠落・書体の置き換え・ページの外の文字・文字領域からのはみ出し・行数の違いを見つけ、
画像にした数式の文字を求めず、字形の割り当ての誤りをはみ出しにせず、予測と実際を並べた記録を作ることを確かめる。
PDF は試験用の補助（pdf_tools）で作る。"""

import io
from types import SimpleNamespace

from fake_fonts import FakeFonts
import pdf_tools
from PIL import Image
import pytest

from slidekit import pdf
from slidekit.measure import MEASURE_VERSION, layout_scene
from slidekit.pdf_validation import compare, validate_pdf


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


def element(id="body", *, area=(20, 20, 180, 40), text="Sample", font="Arial", kind="text", content=None, **extra):
    style = {
        "size_pt": 20,
        "min_size_pt": 18,
        "font_latin": font,
        "font_japanese": font,
        "font_math": "Cambria Math",
        "padding_pt": 0,
        "line_spacing": 1.0,
        **extra.pop("style", {}),
    }
    return {
        "id": id,
        "kind": kind,
        "area": list(area),
        "role": "body",
        "style": style,
        "content": [{"kind": "text", "text": text}] if content is None else content,
        **extra,
    }


def plain(id, area, text, **extra):
    """標準 14 書体の Helvetica で描く文字の要素。"""
    return element(id, area=area, text=text, font="Helvetica", **extra)


def formula(id, area, expression, **extra):
    """数式だけの要素。数式の字形は標準 14 書体の Times-Roman で描く。"""
    item = element(
        id, area=area, kind="equation", font="Helvetica", content=[{"kind": "math", "text": expression}], **extra
    )
    item["style"]["font_math"] = "Times-Roman"
    return item


def scene(*elements, notes="説明します。", number=None):
    page = {"id": "first", "kind": "content", "notes": notes, "elements": list(elements)}
    if number is not None:
        page["number"] = number
    return {
        "schema_version": 1,
        "size": [720, 540],
        "metadata": {"title": "Example", "author": "", "last_modified_by": ""},
        "slides": [page],
    }


def write_pdf(path, *pages, size=(720, 540)):
    """ページごとに、描画の関数を並べた順に描いた PDF を作る。"""
    with pdf_tools.Document() as document:
        for commands in pages:
            page = document.new_page(width=size[0], height=size[1])
            for command in commands:
                command(page)
        document.save(path)
    return path


def glyphs(x, y, value, *, font="helv", size=20, **options):
    return lambda page: page.insert_text((x, y), value, fontname=font, fontsize=size, **options)


def width(value, font="helv", size=20):
    return pdf_tools.text_length(value, fontname=font, fontsize=size)


def codes(issues):
    return {issue["code"] for issue in issues}


def found(issues, code, element_id=None):
    return [
        issue for issue in issues if issue["code"] == code and (element_id is None or issue["element_id"] == element_id)
    ]


def test_pdf_counts_fonts_and_detects_missing_text(project, tmp_path):
    data = scene(element(text="Sample"))
    data["slides"][0]["elements"][0]["style"].update(font_latin="Helvetica", font_japanese="Helvetica")
    path = write_pdf(tmp_path / "deck.pdf", [glyphs(20, 40, "Sample")])
    assert validate_pdf(path, data, project) == []
    data["slides"][0]["elements"][0]["content"][0]["text"] = "Sample missing"
    assert "pdf_text_missing" in codes(validate_pdf(path, data, project))
    data["slides"][0]["elements"][0]["style"]["font_latin"] = "Arial"
    data["slides"][0]["elements"][0]["style"]["font_japanese"] = "Arial"
    assert "pdf_font_substitution" in codes(validate_pdf(path, data, project))


def test_fallback_pdf_does_not_expect_rasterized_text(project, tmp_path):
    text = element()
    text["content"].append({"kind": "math", "text": "x"})
    data = scene(text)
    path = write_pdf(tmp_path / "fallback.pdf", [])
    assert "pdf_text_missing" in codes(validate_pdf(path, data, project, variant="standard"))
    assert "pdf_text_missing" not in codes(validate_pdf(path, data, project, variant="fallback"))


def test_pdf_line_count_mismatch_is_detected(project, tmp_path):
    # 計測では 1 行に収まる文を、PDF では 2 行に描く。字形のない要素は行数を比べない。
    data = scene(plain("lead", (40, 100, 600, 80), "First second"), plain("absent", (40, 300, 600, 40), "Gone"))
    path = write_pdf(tmp_path / "deck.pdf", [glyphs(40, 124, "First"), glyphs(40, 148, "second")])
    without_prediction = validate_pdf(path, data, project)
    assert "pdf_line_count_mismatch" not in codes(without_prediction)
    assert found(without_prediction, "pdf_text_missing")[0]["observations"][0]["missing_characters"] == {
        "G": 1,
        "e": 1,
        "n": 1,
        "o": 1,
    }
    record = {}
    issues = validate_pdf(path, data, project, fonts=FakeFonts(), layout_record=record)
    [mismatch] = found(issues, "pdf_line_count_mismatch")
    assert (mismatch["element_id"], mismatch["severity"], mismatch["scope"]) == (
        "lead",
        "warning",
        "pdf:powerpoint:standard",
    )
    assert mismatch["observations"] == [
        {
            "predicted_lines": 1,
            "actual_lines": 2,
            "predicted_line_widths_pt": [120.0],
            "actual_line_widths_pt": [round(width("First"), 2), round(width("second"), 2)],
            "available_width_pt": 600.0,
            "suggestion": "枠の幅を広げるか、改行を明示する",
        }
    ]
    # 行数が違えば、記録の行の幅の差は求めない。
    assert record["slides"][0]["elements"][0]["difference_pt"]["max_line_width"] is None
    # 計測の結果を渡しても同じになる。計測できない要素（未対応の数式）は予測なしとして扱う。
    data["slides"][0]["elements"].append(formula("broken", (40, 400, 200, 40), r"\unknown"))
    layouts = layout_scene(data, FakeFonts())
    assert found(validate_pdf(path, data, project, layouts=layouts), "pdf_line_count_mismatch") == [mismatch]


def test_pdf_ink_overflow_uses_visible_frame_tolerance(project, tmp_path):
    framed = {"fill": "FFFFFF"}
    data = scene(
        plain("hidden", (40, 100, 100, 40), "Sample", kind="box"),
        plain("framed", (40, 200, 100, 40), "Sample", kind="box", style=framed),
        plain("far", (40, 300, 100, 40), "Sample", kind="box"),
        plain("round", (300, 100, 200, 80), "Oval", kind="box", shape="ellipse"),
        plain("hidden_edge", (400, 300, 100, 40), "Sample", kind="box"),
        plain("framed_edge", (400, 400, 100, 40), "Sample", kind="box", style=framed),
    )
    path = write_pdf(
        tmp_path / "deck.pdf",
        [
            glyphs(
                141 - width("Sample"), 130, "Sample"
            ),  # 枠の右端を 1 pt 超える（塗りも線もない枠の許容 1.5 pt 以内）
            glyphs(141 - width("Sample"), 230, "Sample"),  # 塗りのある枠では、1 pt でも許容 0.5 pt を超える
            glyphs(142 - width("Sample"), 330, "Sample"),  # 2 pt は塗りも線もない枠でも超える
            glyphs(310, 150, "Oval"),  # 楕円の枠の中だが、文字領域（内接する矩形）の左端より外
            glyphs(501.4 - width("Sample"), 330, "Sample"),  # 許容幅の近く：塗りも線もない枠で 1.4 pt
            glyphs(500.4 - width("Sample"), 430, "Sample"),  # 塗りのある枠で 0.4 pt
        ],
    )
    for issues in (validate_pdf(path, data, project), validate_pdf(path, data, project, fonts=FakeFonts())):
        overflow = {issue["element_id"]: issue for issue in found(issues, "pdf_element_text_overflow")}
        assert set(overflow) == {"framed", "far", "round"}
        observation = overflow["framed"]["observations"][0]
        assert observation["text"] == "Sample" and observation["overflow_pt"] == {"side": "right", "amount": 1.0}
        assert observation["ink_pt"][2] == pytest.approx(141.0) and observation["text_area_pt"] == [
            40.0,
            200.0,
            100.0,
            40.0,
        ]
        assert overflow["far"]["observations"][0]["overflow_pt"] == {"side": "right", "amount": 2.0}
        assert overflow["round"]["observations"][0]["overflow_pt"]["side"] == "left"
        assert overflow["round"]["observations"][0]["text_area_pt"][0] == pytest.approx(
            300 + 200 * (1 - 0.5**0.5) / 2, abs=0.01
        )
        assert {issue["severity"] for issue in overflow.values()} == {"error"}
    # 予測があれば、予測したインクを併記する。
    # 上揃えの最初の基線は 200 + 18 pt で、インクは上 14 pt・下 2 pt（偽のフォント）。
    assert found(issues, "pdf_element_text_overflow", "framed")[0]["observations"][0]["predicted_ink_pt"] == [
        40.0,
        204.0,
        100.0,
        220.0,
    ]
    assert (
        found(validate_pdf(path, data, project), "pdf_element_text_overflow", "framed")[0]["observations"][0][
            "predicted_ink_pt"
        ]
        is None
    )


def test_pdf_layout_record_has_predicted_and_actual(project, tmp_path):
    data = scene(
        plain("lead", (40, 100, 400, 40), "Sample"),
        plain("turned", (40, 300, 300, 40), "Turned", rotation=90),
        formula("formula", (40, 200, 100, 40), "x"),
        number=3,
    )
    path = write_pdf(
        tmp_path / "deck.pdf",
        [glyphs(40, 118, "Sample"), glyphs(200, 360, "Turned", rotate=90), glyphs(44, 230, "x", font="tiro")],
    )
    record = {}
    assert validate_pdf(path, data, project, fonts=FakeFonts(), layout_record=record) == []
    assert {key: record[key] for key in ("schema_version", "renderer", "variant", "measure_version")} == {
        "schema_version": 1,
        "renderer": "powerpoint",
        "variant": "standard",
        "measure_version": MEASURE_VERSION,
    }
    [page] = record["slides"]
    assert (page["id"], page["number"], page["unassigned_glyphs"]) == ("first", 3, 0)
    # 回転した要素は読まないので、記録にも入れない。
    assert [item["id"] for item in page["elements"]] == ["lead", "formula"]
    lead = page["elements"][0]
    # 偽のフォントの予測：1 文字 0.5 em、最初の基線は行間 × ascent 0.9 em、インクは上 0.7 em・下 0.1 em。
    assert lead["predicted"] == {
        "lines": 1,
        "line_widths_pt": [60.0],
        "baselines_pt": [118.0],
        "ink_pt": [40.0, 104.0, 100.0, 120.0],
    }
    actual = lead["actual"]
    assert (actual["lines"], actual["baselines_pt"], actual["glyphs"], actual["drawings"], actual["sizes_pt"]) == (
        1,
        [118.0],
        6,
        0,
        [20.0],
    )
    assert actual["line_widths_pt"] == [round(width("Sample"), 2)] and actual["ink_pt"][0] == 40.0
    difference = lead["difference_pt"]
    assert difference["max_line_width"] == pytest.approx(width("Sample") - 60, abs=0.01)
    assert difference["ink"] == pytest.approx(
        [a - b for a, b in zip(actual["ink_pt"], lead["predicted"]["ink_pt"])], abs=0.01
    )
    assert page["elements"][1]["predicted"]["lines"] == 1 and page["elements"][1]["actual"]["glyphs"] == 1
    # 予測がなければ、実際だけを書く。
    record = {}
    validate_pdf(path, data, project, layout_record=record)
    assert record["slides"][0]["elements"][0] == {
        "id": "lead",
        "predicted": None,
        "actual": actual,
        "difference_pt": None,
    }


def test_compare_counts_only_lines_with_content(tmp_path):
    data = scene(plain("lead", (40, 100, 400, 120), "First\n\n \nlast"))
    path = write_pdf(tmp_path / "deck.pdf", [glyphs(40, 118, "First"), glyphs(40, 190, "last")])
    from slidekit.pdf_layout import read_pdf

    actual = read_pdf(path, data)[0].elements["lead"]
    predicted = layout_scene(data, FakeFonts())[("first", "lead")]
    assert len(predicted.text.lines) == 4
    result = compare(predicted, actual)
    # 空の行と空白だけの行は PDF に現れないので行数に含めない。行の幅の差は絶対値が最大のもの（符号付き）。
    assert result["predicted"]["lines"] == result["actual"]["lines"] == 2
    assert result["predicted"]["baselines_pt"] == [118.0, 190.0]
    differences = [width("First") - 50, width("last") - 40]
    assert result["difference_pt"]["max_line_width"] == pytest.approx(max(differences, key=abs), abs=0.01)


def test_pdf_warning_evidence_uses_page_evidence(project, tmp_path):
    # 2 つ目の資料では同じページの前に別のページを加え、ページ番号が 1 から 2 に変わる。
    page = {"id": "page", "number": 1, "kind": "content", "notes": "説明します。", "elements": [element(text="Sample")]}
    other = {
        "id": "other",
        "number": 1,
        "kind": "content",
        "notes": "説明します。",
        "elements": [plain("other", (20, 20, 180, 40), "Other")],
    }
    alone = {"schema_version": 1, "size": [720, 540], "slides": [page]}
    moved = {"schema_version": 1, "size": [720, 540], "slides": [other, {**page, "number": 2}]}
    one = write_pdf(tmp_path / "one.pdf", [glyphs(20, 40, "Sample")])
    two = write_pdf(tmp_path / "two.pdf", [glyphs(20, 40, "Other")], [glyphs(20, 40, "Sample")])

    def substitution(path, data, page_evidence=None):
        [issue] = found(validate_pdf(path, data, project, page_evidence=page_evidence), "pdf_font_substitution")
        return issue

    evidence = {"page": "visual-page", "other": "visual-other"}
    before, after = substitution(one, alone, evidence), substitution(two, moved, evidence)
    assert (before["slide_number"], after["slide_number"]) == (1, 2)
    assert before["diagnostic_id"] == after["diagnostic_id"] and before["evidence_hash"] == after["evidence_hash"]
    # 表示の証拠が変われば証拠も変わる。渡さなければページの全体（番号を含む）を証拠にする。
    assert substitution(two, moved, {**evidence, "page": "changed"})["evidence_hash"] != after["evidence_hash"]
    assert substitution(one, alone)["evidence_hash"] != substitution(two, moved)["evidence_hash"]


def test_pdf_page_count_size_and_text_outside_page(project, tmp_path):
    data = scene(plain("edge", (600, 20, 200, 40), "Sample"), number=7)
    inside = write_pdf(tmp_path / "inside.pdf", [glyphs(610, 40, "Sample")])
    assert validate_pdf(inside, data, project) == []
    # ページの外の文字は抽出されないので、欠落の候補にもなる。
    outside = validate_pdf(write_pdf(tmp_path / "outside.pdf", [glyphs(690, 40, "Sample")]), data, project)
    assert {(issue["code"], issue["severity"], issue["slide_id"], issue["slide_number"]) for issue in outside} == {
        ("pdf_text_overflow", "error", "first", 7),
        ("pdf_text_missing", "warning", "first", 7),
    }
    extra = write_pdf(tmp_path / "extra.pdf", [glyphs(610, 40, "Sample")], [])
    [count] = validate_pdf(extra, data, project)
    assert (count["code"], count["severity"], count["observations"]) == (
        "pdf_page_count",
        "error",
        [{"expected": 1, "actual": 2}],
    )
    assert count["slide_id"] is None and "slide_number" not in count
    wide = write_pdf(tmp_path / "wide.pdf", [glyphs(610, 40, "Sample")], size=(720, 405))
    [size] = validate_pdf(wide, data, project)
    assert (size["code"], size["severity"], size["slide_number"], size["observations"]) == (
        "pdf_page_size",
        "error",
        7,
        [{"expected": [720, 540], "actual": [720.0, 405.0]}],
    )
    broken = tmp_path / "broken.pdf"
    broken.write_bytes(b"not a pdf")
    assert [issue["code"] for issue in validate_pdf(broken, data, project)] == ["pdf_unreadable"]


def test_unassigned_glyphs_are_reported_as_info(project, tmp_path):
    data = scene(plain("title", (40, 20, 600, 40), "Overview"), number=4)
    path = write_pdf(tmp_path / "deck.pdf", [glyphs(600, 520, "Draft"), glyphs(44, 44, "Overview")])
    [issue] = validate_pdf(path, data, project)
    assert (issue["code"], issue["severity"], issue["slide_number"], issue["observations"]) == (
        "pdf_text_unassigned",
        "info",
        4,
        [{"count": 5}],
    )


def test_math_element_without_glyphs_is_missing_text(project, tmp_path):
    # 数式の字形は文字の照合に入らないので、字形のない数式の要素を欠落の候補として示す。行数とインクは比べない。
    data = scene(plain("intro", (40, 20, 300, 40), "Intro"), formula("formula", (40, 200, 200, 60), "x^2"))
    path = write_pdf(tmp_path / "deck.pdf", [glyphs(44, 44, "Intro")])
    record = {}
    issues = validate_pdf(path, data, project, fonts=FakeFonts(), layout_record=record)
    [issue] = issues
    assert (issue["code"], issue["severity"]) == ("pdf_text_missing", "warning")
    assert issue["observations"] == [
        {"missing_characters": {}, "missing_count": 0, "math_elements_without_glyphs": ["formula"]}
    ]
    assert record["slides"][0]["elements"][1]["actual"]["glyphs"] == 0
    # 数式を画像にした表示では、数式の字形を求めない。
    assert validate_pdf(path, data, project, variant="fallback") == []


def test_math_dropped_from_mixed_element_is_missing_text(project, tmp_path):
    # 文字と数式が混じる要素で数式だけが描かれない場合（PowerPoint は不正な数式を描かず、そのことを知らせない）も、欠落の候補として示す。
    # 文字の書体で描く \text{...} の字形も数式の字形に含める。空白の命令だけの数式は字形を描かないので、求めない。
    fonts = {"font_math": "Times-Roman"}
    data = scene(
        plain(
            "mixed",
            (40, 100, 600, 40),
            None,
            style=fonts,
            content=[{"kind": "text", "text": "Value "}, {"kind": "math", "text": "x^2"}],
        ),
        plain(
            "spaced",
            (40, 200, 600, 40),
            None,
            style=fonts,
            content=[{"kind": "text", "text": "Gap"}, {"kind": "math", "text": r"\quad"}],
        ),
        plain(
            "unit",
            (40, 300, 600, 40),
            None,
            style=fonts,
            content=[{"kind": "text", "text": "Speed 5"}, {"kind": "math", "text": r"\text{ km}"}],
        ),
    )
    drawn = write_pdf(
        tmp_path / "drawn.pdf",
        [
            glyphs(44, 130, "Value"),
            glyphs(110, 130, "x", font="tiro"),
            glyphs(120, 122, "2", font="tiro", size=14),
            glyphs(44, 230, "Gap"),
            glyphs(44, 330, "Speed 5"),
            glyphs(130, 330, "km"),
        ],
    )
    assert validate_pdf(drawn, data, project) == []
    dropped = write_pdf(
        tmp_path / "dropped.pdf", [glyphs(44, 130, "Value"), glyphs(44, 230, "Gap"), glyphs(44, 330, "Speed 5")]
    )
    [issue] = validate_pdf(dropped, data, project)
    assert (issue["code"], issue["observations"]) == (
        "pdf_text_missing",
        [
            {
                "missing_characters": {"k": 1, "m": 1},
                "missing_count": 2,
                "math_elements_without_glyphs": ["mixed", "unit"],
            }
        ],
    )


def test_spacing_only_math_draws_no_glyphs():
    # 空白の命令・\text の名前・波括弧・空白だけの数式は字形を描かない。記号の \{ や分数は描く。
    from slidekit.pdf_validation import _draws_glyphs

    expressions = [r"\quad", r"\, \; \qquad", r"\text{ }", "{}", "~", "x", r"\text{if}", r"\{", r"\frac{a}{b}"]
    assert [_draws_glyphs(expression) for expression in expressions] == [False] * 5 + [True] * 4


def test_far_ink_is_not_reported_when_glyphs_may_belong_to_another_element(project, tmp_path):
    # 同じ文字列の 2 要素のうち下だけが描かれると、字形の文字列だけの対応付けは、場面の順で上の要素に字形を割り当てる。
    # 欠落のあるページでは、枠を 1 em 広げた範囲と交わらないインクを割り当ての誤りとみなし、はみ出しと行数の違いにしない。
    # 枠のすぐ外（1 em 以内）に描かれた要素は、同じページでもはみ出しとして示す。
    data = scene(
        plain("near", (40, 150, 600, 20), "Near"),
        plain("top", (40, 20, 600, 40), "First second"),
        plain("bottom", (40, 300, 600, 80), "First second"),
    )
    path = write_pdf(
        tmp_path / "deck.pdf", [glyphs(40, 186, "Near"), glyphs(40, 324, "First"), glyphs(40, 348, "second")]
    )
    issues = validate_pdf(path, data, project, fonts=FakeFonts())
    assert sorted((issue["code"], issue["element_id"]) for issue in issues) == [
        ("pdf_element_text_overflow", "near"),
        ("pdf_text_missing", None),
    ]
    assert found(issues, "pdf_text_missing")[0]["observations"][0]["missing_count"] == len("Firstsecond")
    assert found(issues, "pdf_element_text_overflow")[0]["observations"][0]["overflow_pt"]["side"] == "bottom"
    # 数式の \text{...} の中身も文字として照らすので、数式が描かれず、後の要素の同じ文字列だけが描かれても欠落の候補になる。
    data = scene(formula("case", (40, 100, 300, 40), r"x \text{ if } y"), plain("word", (40, 200, 300, 40), "if"))
    path = write_pdf(tmp_path / "text.pdf", [glyphs(44, 230, "if")])
    [issue] = validate_pdf(path, data, project)
    assert (issue["code"], issue["observations"][0]["missing_characters"]) == ("pdf_text_missing", {"f": 1, "i": 1})
    # 雛形の同じ文字列が先に描かれたページも、割り当てられない字形があるので同じ扱いにする。
    data = scene(plain("title", (40, 100, 600, 40), "Overview"))
    path = write_pdf(tmp_path / "template.pdf", [glyphs(500, 520, "Overview"), glyphs(44, 124, "Overview")])
    [issue] = validate_pdf(path, data, project)
    assert (issue["code"], issue["observations"]) == ("pdf_text_unassigned", [{"count": 8}])
    # 欠落も割り当てられない字形もないページでは、読んだとおりに比べる（枠から離れたインクもはみ出しにする）。
    data = scene(plain("first", (40, 60, 400, 40), "Same words"), plain("second", (40, 140, 400, 40), "Same words"))
    path = write_pdf(tmp_path / "shifted.pdf", [glyphs(44, 176, "Same words"), glyphs(44, 164, "Same words")])
    assert [
        issue["element_id"]
        for issue in validate_pdf(path, data, project)
        if issue["code"] == "pdf_element_text_overflow"
    ] == ["first"]


def test_pdf_element_evidence_adds_element_to_page_evidence(project, tmp_path):
    # 要素の診断の証拠はページの表示の証拠に要素を加えたもの。ページ番号だけが変わっても変わらず、要素が変われば変わる。
    def page(number, color="17212D"):
        lead = plain("lead", (40, 100, 600, 80), "First second", style={"color": color})
        return {"id": "page", "number": number, "kind": "content", "notes": "説明します。", "elements": [lead]}

    other = {
        "id": "other",
        "number": 1,
        "kind": "content",
        "notes": "説明します。",
        "elements": [plain("other", (40, 100, 600, 40), "Other")],
    }
    lines = [glyphs(40, 124, "First"), glyphs(40, 148, "second")]
    one = write_pdf(tmp_path / "one.pdf", lines)
    two = write_pdf(tmp_path / "two.pdf", [glyphs(40, 124, "Other")], lines)
    evidence = {"page": "visual-page", "other": "visual-other"}

    def mismatch(path, *slides):
        data = {"schema_version": 1, "size": [720, 540], "slides": list(slides)}
        [issue] = found(
            validate_pdf(path, data, project, fonts=FakeFonts(), page_evidence=evidence), "pdf_line_count_mismatch"
        )
        return issue

    before, moved, restyled = mismatch(one, page(1)), mismatch(two, other, page(2)), mismatch(one, page(1, "336699"))
    assert (before["slide_number"], moved["slide_number"], restyled["slide_number"]) == (1, 2, 1)
    assert before["diagnostic_id"] == moved["diagnostic_id"] == restyled["diagnostic_id"]
    assert before["observations"] == restyled["observations"]
    assert before["evidence_hash"] == moved["evidence_hash"] != restyled["evidence_hash"]


def test_slide_number_defaults_to_position(project, tmp_path):
    # 場面のページに番号がなければ、診断と記録のページ番号は 1 始まりの位置にする。
    data = {
        "schema_version": 1,
        "size": [720, 540],
        "slides": [
            {"id": "first", "elements": [plain("alpha", (40, 100, 300, 40), "Alpha")]},
            {"id": "second", "elements": [plain("beta", (40, 100, 300, 40), "Beta")]},
        ],
    }
    path = write_pdf(
        tmp_path / "deck.pdf", [glyphs(44, 124, "Alpha")], [glyphs(600, 520, "Draft"), glyphs(44, 124, "Beta")]
    )
    record = {}
    [issue] = validate_pdf(path, data, project, layout_record=record)
    assert (issue["code"], issue["slide_id"], issue["slide_number"]) == ("pdf_text_unassigned", "second", 2)
    assert [page["number"] for page in record["slides"]] == [1, 2]


def test_pdf_text_is_read_without_images(project, tmp_path, monkeypatch):
    # validate_pdf は文字と描画だけを使うので、画像のあるページでも、ページを画像にしない。
    def render(page, scale):
        raise AssertionError("PDF の検査でページを画像にした")

    monkeypatch.setattr(pdf.Page, "render", render)
    picture = io.BytesIO()
    Image.new("RGB", (8, 8), (200, 40, 40)).save(picture, format="PNG")

    def image(page):
        page.insert_image((300, 300, 400, 400), stream=picture.getvalue())

    path = write_pdf(tmp_path / "deck.pdf", [image, glyphs(20, 40, "Sample")])
    assert validate_pdf(path, scene(plain("body", (20, 20, 180, 40), "Sample")), project) == []


def test_zero_width_text_is_not_checked_for_fonts(project, tmp_path, dejavu):
    # 数式だけの段落に加える幅 0 の空白は、要素の文字の書体で描かれるが、表示されないので書体を検査しない。
    data = scene(formula("formula", (40, 40, 200, 60), "x"))
    path = write_pdf(
        tmp_path / "deck.pdf",
        [glyphs(44, 80, "x", font="tiro"), glyphs(54, 80, " \u200b", font="DejaVuSans", fontfile=str(dejavu.path))],
    )
    assert validate_pdf(path, data, project) == []
    visible = write_pdf(
        tmp_path / "visible.pdf",
        [glyphs(44, 80, "x", font="tiro"), glyphs(54, 80, " a", font="DejaVuSans", fontfile=str(dejavu.path))],
    )
    assert "pdf_font_substitution" in codes(validate_pdf(visible, data, project))


def test_invisible_text_is_not_checked(project, tmp_path):
    # 見えない文字片（不透明度 0。PowerPoint が説明のある要素の数式に加える読み上げ用の文字など）は、書体の置き換えにも、
    # 欠けた文字の補いにも含めない。
    data = scene(plain("body", (20, 20, 180, 40), "Sample"))
    path = write_pdf(
        tmp_path / "deck.pdf", [glyphs(20, 40, "Sample"), glyphs(20, 90, "hidden", font="tiro", fill_opacity=0)]
    )
    assert validate_pdf(path, data, project) == []
    data["slides"][0]["elements"][0]["content"][0]["text"] = "Sample hidden"
    assert found(validate_pdf(path, data, project), "pdf_text_missing")


def test_text_in_math_expects_text_fonts(project, tmp_path):
    # \text{...} の中身は要素の文字の書体で描かれるので、数式だけの要素でも文字の書体を期待する。
    data = scene(formula("case", (40, 100, 300, 40), r"x \text{ if } y"))
    path = write_pdf(
        tmp_path / "deck.pdf",
        [glyphs(44, 130, "x", font="tiro"), glyphs(60, 130, "if"), glyphs(80, 130, "y", font="tiro")],
    )
    assert validate_pdf(path, data, project) == []
    data["slides"][0]["elements"][0]["content"][0]["text"] = "x + y"
    assert "pdf_font_substitution" in codes(validate_pdf(path, data, project))
