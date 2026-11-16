"""寸法の補助と接続位置を確かめる：内容から決める寸法（Rect の None）、slide.measure・element.measure、
積み上げ（stack）、anchor の間隔・位置・回転・反転・画像の表示範囲、画像ごとの倍率と埋め込みの指定。"""

import itertools
import math
import re

from PIL import Image
import pytest

from fake_fonts import FakeFonts
from pdf_tools import Document
from slidekit import Deck, Math, Rect, Text
from slidekit.assets import PDF_VECTOR_MESSAGE
from slidekit.config import load_project
from slidekit.errors import ArgumentNameError, InputError
from slidekit.measure import MeasureError, TextLayout, fit_findings, layout_element
from slidekit.validators import validate_scene

LONG = "長い文を枠の幅で折り返して高さを内容から決める例です。" * 2


@pytest.fixture
def make_deck(tmp_path, write_deck):
    """資料を書いて読み込み、計測を偽のフォント（1 文字 0.5 em）にした Deck を返す関数。"""
    numbers = itertools.count()

    def make(theme=None):
        project = load_project(write_deck(tmp_path / f"deck{next(numbers)}", theme=theme))
        project.fonts = FakeFonts()
        return Deck(project)

    return make


def by_id(page):
    return {element["id"]: element for element in page["elements"]}


def png(root, name="figure.png", size=(200, 100)):
    Image.new("RGB", size, "white").save(root / name)
    return name


def pdf(root, name="diagram.pdf", size=(200, 100)):
    with Document() as document:
        document.new_page(width=size[0], height=size[1])
        document.save(root / name)
    return name


def svg(root, name, attributes):
    (root / name).write_text(
        f'<svg xmlns="http://www.w3.org/2000/svg" {attributes}><rect width="10" height="10"/></svg>', encoding="utf-8"
    )
    return name


def tenth(value):
    """0.1 pt 単位の値か。"""
    return abs(value * 10 - round(value * 10)) < 1e-9


# ---- 内容から決める寸法 ----


def test_auto_height_matches_measure(make_deck):
    deck = make_deck()
    slide = deck.slide(id="auto")
    lead = slide.text(id="lead", text=LONG, area=Rect(48, 100, 305, None))
    expected = slide.measure(LONG, width=305)
    assert isinstance(expected, TextLayout) and len(expected.lines) > 1
    assert lead.area == Rect(48, 100, 305, expected.height) and tenth(lead.area.h)
    note = slide.box(id="note", text=LONG, area=Rect(48, 300, 300, None), role="note", padding_pt=8)
    assert note.area.h == slide.measure(LONG, kind="box", width=300, role="note", padding_pt=8).height
    # 場面の領域は数値で、要素は配置の検査の計測でもはみ出さない。
    elements = by_id(deck.to_dict()["slides"][0])
    assert elements["lead"]["area"] == [48, 100, 305, expected.height]
    for name in ("lead", "note"):
        findings = fit_findings(layout_element(elements[name], deck.project.fonts))
        assert None not in elements[name]["area"] and not [item for item in findings if item["severity"] != "info"]


def test_auto_width_without_wrap(make_deck):
    slide = make_deck().slide(id="auto")
    # 偽のフォントは 1 文字 0.5 em なので、20 pt の 4 文字は 40 pt。
    label = slide.text(id="label", text="ABCD", area=Rect(48, 100, None, 40), word_wrap=False)
    assert label.area == Rect(48, 100, 40, 40)
    # 左右の距離を加える（上下の余白は box の既定 6 pt を引き継ぐ）。
    tag = slide.box(
        id="tag", text="ABCD", area=Rect(48, 160, None, 40), padding_pt={"left": 6, "right": 10}, word_wrap=False
    )
    assert tag.area.w == 56
    # 折り返す要素は、計測の誤差で PowerPoint が折り返さないよう、誤差の帯の分だけ広げる。
    wrapped = slide.text(id="wrapped", text="ABCD", area=Rect(48, 220, None, 40))
    assert wrapped.area.w == slide.measure("ABCD").width and 40 < wrapped.area.w <= 41
    assert len(layout_element(by_id(slide.to_dict(1))["wrapped"], slide.project.fonts).text.lines) == 1
    # 幅と高さの両方を内容から決めることができる。明示改行の最も長い行が幅になる。
    both = slide.text(id="both", text="AB\nCDEF", area=(48, 300, None, None), word_wrap=False)
    measured = slide.measure("AB\nCDEF", word_wrap=False)
    assert both.area == Rect(48, 300, 40, measured.height) and len(measured.lines) == 2


def test_auto_height_for_shapes(make_deck):
    deck = make_deck()
    slide = deck.slide(id="shapes")
    fonts = deck.project.fonts
    for index, shape in enumerate(("rectangle", "rounded_rectangle", "ellipse", "diamond")):
        radius = {"corner_radius_pt": 24} if shape == "rounded_rectangle" else {}
        tall = slide.box(
            id=f"{shape}_tall", text=LONG, shape=shape, area=Rect(40, 40 + index * 100, 240, None), **radius
        )
        wide = slide.box(
            id=f"{shape}_wide", text="ABCD", shape=shape, area=Rect(300, 40 + index * 100, None, None), **radius
        )
        assert tall.area.h == slide.measure(LONG, kind="box", shape=shape, width=240, **radius).height
        assert (wide.area.w, wide.area.h) == (
            slide.measure("ABCD", kind="box", shape=shape, **radius).width,
            slide.measure("ABCD", kind="box", shape=shape, **radius).height,
        )
    elements = by_id(deck.to_dict()["slides"][0])
    for element_id in [element.id for element in slide.elements]:
        layout = layout_element(elements[element_id], fonts)
        assert not [finding for finding in fit_findings(layout) if finding["code"] == "text_overflow"], element_id
    # 図形の文字領域は枠より狭いので、同じ幅の長方形より高い枠になる。
    assert elements["ellipse_tall"]["area"][3] > elements["rectangle_tall"]["area"][3]
    assert elements["diamond_tall"]["area"][3] > elements["rectangle_tall"]["area"][3]


def test_image_height_from_aspect_ratio(make_deck):
    deck = make_deck()
    root = deck.project.root
    slide = deck.slide(id="images")
    sources = [
        png(root),
        pdf(root),
        svg(root, "sized.svg", 'width="4cm" height="2cm"'),
        svg(root, "box.svg", 'width="100%" viewBox="0 0 300 150"'),
    ]
    for index, path in enumerate(sources):
        page = {"page": 1} if path.endswith(".pdf") else {}
        tall = slide.image(id=f"width_{index}", path=path, area=Rect(10, 10, 300, None), **page)
        wide = slide.image(id=f"height_{index}", path=path, area=Rect(10, 200, None, 50), fit="cover", **page)
        assert (tall.area.h, wide.area.w) == (150, 100), path
    elements = by_id(deck.to_dict()["slides"][0])
    assert elements["width_0"]["area"] == [10, 10, 300, 150]
    # 縦横比の同じ PNG の表示範囲は、領域と同じになる。
    assert elements["width_0"]["actual_area"] == [10, 10, 300, 150]
    # 両方を内容から決めることはできない。寸法の分からない SVG では、内容から決めることができない。
    unknown = svg(root, "unknown.svg", "")
    for path, area, words in (
        (sources[0], Rect(0, 0, None, None), "画像の幅か高さの一方を指定してください"),
        (unknown, Rect(0, 0, 100, None), "画像の縦横比が分からないため、高さを内容から決められません"),
        (unknown, Rect(0, 0, None, 100), "画像の縦横比が分からないため、幅を内容から決められません"),
    ):
        with pytest.raises(InputError, match=words) as caught:
            slide.image(id="unsized", path=path, area=area)
        assert (caught.value.code, caught.value.element_id) == ("invalid_geometry", "unsized")
    # 寸法の分からない SVG も、領域を指定すれば置くことができる。
    slide.image(id="placed", path=unknown, area=Rect(0, 0, 100, 100))


def test_unknown_dimensions_are_rejected_where_content_cannot_decide(make_deck):
    slide = make_deck().slide(id="kinds")
    for operation in (
        lambda: slide.group(id="group", area=Rect(0, 0, 100, None)),
        lambda: slide.text(id="title", text="題名", area=Rect(0, 0, 100, -1)),
    ):
        with pytest.raises(InputError) as caught:
            operation()
        assert caught.value.code == "invalid_geometry"
    with pytest.raises(InputError, match="text・box・equation・image と table で使えます"):
        slide.group(id="group", area=Rect(0, 0, None, 100))
    # 題名の領域は内容から決めない。
    with pytest.raises(InputError, match="領域の幅と高さを数値で指定してください") as caught:
        slide.deck.slide(id="titled", title="題名", title_area=Rect(40, 20, 400, None))
    assert caught.value.element_id == "_theme/title"
    # 内容のない文字は幅を決めることができない（高さは 1 行分になる）。
    with pytest.raises(InputError, match="内容のない要素は幅を内容から決められません") as caught:
        slide.text(id="empty", text="", area=Rect(0, 0, None, 40))
    assert (caught.value.code, caught.value.element_id) == ("invalid_geometry", "empty")
    assert slide.text(id="blank", text="", area=Rect(0, 0, 100, None)).area.h == slide.measure("", width=100).height > 0
    # 作成後の領域は数値で、幅を変えても高さは計算し直さない。
    lead = slide.text(id="lead", text=LONG, area=Rect(0, 0, 300, None))
    height = lead.area.h
    lead.area = Rect(0, 0, 150, height)
    assert lead.area.h == height
    with pytest.raises(InputError, match="数値で指定してください"):
        lead.area = Rect(0, 0, 150, None)


def test_group_children_use_unknown_dimensions_in_group_coordinates(make_deck):
    slide = make_deck().slide(id="group")
    group = slide.group(id="panel", area=Rect(100, 50, 300, 200))
    child = group.text(id="label", text="ABCD", area=(10, 20, None, 40), word_wrap=False)
    assert (child.id, child.area) == ("panel/label", Rect(110, 70, 40, 40))


def test_measurement_errors_name_the_element_and_leave_nothing(make_deck):
    deck = make_deck()
    deck.project.fonts = FakeFonts(missing={"★"})
    slide = deck.slide(id="errors")
    with pytest.raises(MeasureError) as caught:
        slide.text(id="star", text="★印", area=Rect(0, 0, 200, None))
    error = caught.value
    assert (error.code, error.slide_id, error.element_id) == ("font_glyph_missing", "errors", "star")
    assert str(error).startswith("ページ errors の要素 star: ")
    with pytest.raises(MeasureError, match=re.escape(r"\unknown")) as caught:
        slide.equation(id="formula", expression=r"\unknown", area=Rect(0, 0, None, 40))
    assert caught.value.element_id == "formula"
    # 寸法を指定した要素は、作成では計測しない（配置の検査が報告する）。
    slide.text(id="fixed", text="★", area=Rect(0, 0, 100, 40))
    assert [element.id for element in slide.elements] == ["fixed"]


# ---- 計測の関数 ----


def test_measure_uses_same_style_as_element(make_deck):
    deck = make_deck(
        '[roles.callout]\ninherits = "note"\nsize_pt = 16\npadding_pt = 4\n'
        '[roles.small]\ninherits = "note"\nrestricted = true\n'
    )
    fonts = deck.project.fonts
    slide = deck.slide(id="measure")
    # 役割の文字サイズ（16 pt の 4 文字は 32 pt）と余白、種類の既定（box の余白 6 pt は役割の 4 pt より下）を使う。
    callout = slide.measure("ABCD", role="callout", word_wrap=False)
    assert (callout.content_width, callout.width, callout.insets.left) == (32, 40, 4)
    assert slide.measure("ABCD", kind="box", word_wrap=False).insets.left == 6
    assert slide.measure("ABCD", word_wrap=False).insets.left == 0
    measured = slide.measure("ABCD", kind="box", role="callout")
    box = slide.box(id="callout", text="ABCD", role="callout", area=Rect(0, 0, None, None))
    assert box.area == Rect(0, 0, measured.width, measured.height)
    # style= の辞書よりキーワード引数が優先する（要素と同じ）。
    assert slide.measure("ABCD", style={"size_pt": 30}, size_pt=10, word_wrap=False).content_width == 20
    # 文字の断片も受け付け、幅を指定すると折り返す。
    assert len(slide.measure([Text("AB "), Math("x"), Text(" CD")], width=30).lines) > 1
    # equation で数式 1 つだけなら独立した数式として計測し、文中の数式は文中の数式として計測する。
    for content, kind, display in (("x^2", "equation", True), ([Text("値 "), Math("x^2")], "text", False)):
        fonts.calls.clear()
        slide.measure(content, kind=kind)
        assert fonts.calls and {call["display"] for call in fonts.calls} == {display}, kind
    # 使う場所を限った役割も、計測には宣言が要らない。計測は要素を作らない。
    slide.measure("注", role="small")
    assert [element.id for element in slide.elements] == ["callout"]


def test_measure_rejects_invalid_arguments(make_deck):
    slide = make_deck().slide(id="measure")
    for operation, code, words in (
        (lambda: slide.measure("A", kind="line"), "input_error", "kind は text、box、equation のいずれか"),
        (lambda: slide.measure("A", width=0), "invalid_geometry", "幅は正の有限値にしてください"),
        (lambda: slide.measure("A", width=float("inf")), "invalid_geometry", "幅は正の有限値"),
        (lambda: slide.measure("A", shape="ellipse"), "input_error", "shape は kind が box のときだけ"),
        (lambda: slide.measure("A", kind="box", shape="star"), "input_error", "未対応の図形です"),
        (lambda: slide.measure("", kind="equation"), "invalid_fragment", "数式を空でない文字列で"),
        (lambda: slide.measure("A", kind="box", corner_radius_pt=4), "invalid_style", "rounded_rectangle の box だけ"),
        (lambda: slide.measure("A", colour="accent"), "unknown_argument", "colour（候補: color）"),
        (lambda: slide.measure("A", rotation=90), "unsupported_argument", "rotation は text では使えません"),
    ):
        with pytest.raises(InputError, match=re.escape(words)) as caught:
            operation()
        assert (caught.value.code, caught.value.slide_id) == (code, "measure"), words
        assert str(caught.value).startswith("ページ measure: ")
    with pytest.raises(ArgumentNameError):
        slide.measure("A", size=20)


def test_element_measure_matches_validation(make_deck):
    deck = make_deck()
    fonts = deck.project.fonts
    slide = deck.slide(id="elements")
    card = slide.box(id="card", text=LONG, area=Rect(40, 40, 240, 200), shape="rounded_rectangle", padding_pt=8)
    formula = slide.equation(id="formula", expression="x^2 + y^2", area=Rect(40, 300, 240, 40))
    data = by_id(slide.to_dict(1))
    assert card.measure() == layout_element(data["card"], fonts).text
    assert formula.measure() == layout_element(data["formula"], fonts).text
    # 文字を持たない要素は計測できない。まとまりの measure() も同じ。
    rule = slide.line(id="rule", start=(0, 10), end=(100, 10))
    group = slide.group(id="group", area=Rect(0, 0, 100, 100))
    for element in (rule, group):
        with pytest.raises(InputError, match="文字を持たない要素は計測できません") as caught:
            element.measure()
        assert caught.value.element_id == element.id
    # まとまりの measure(text, ...) は slide.measure と同じ。
    assert group.measure("ABCD", width=100) == slide.measure("ABCD", width=100)


# ---- 積み上げ ----


def test_stack_places_elements_with_gap_and_reports_bounds(make_deck):
    deck = make_deck()
    root, fonts = deck.project.root, deck.project.fonts
    png(root)
    slide = deck.slide(id="stack")
    area = Rect(48, 100, 400, 340)
    stack = slide.stack(area, gap=12)
    assert (stack.area, stack.gap, stack.cursor, stack.elements) == (area, 12, 100, ())
    lead = stack.text(id="lead", text=LONG)
    height = slide.measure(LONG, width=400).height
    assert lead.area == Rect(48, 100, 400, height) and stack.cursor == 100 + height + 12
    # 幅が領域より狭いときの横位置は place、その要素の後の間隔は gap。
    top = stack.cursor
    note = stack.box(id="note", text="ABCD", width=200, place="center", gap=20, padding_pt=8)
    assert note.area == Rect(148, top, 200, slide.measure("ABCD", kind="box", width=200, padding_pt=8).height)
    assert stack.cursor == note.area.bottom + 20
    stack.space(8)
    top = stack.cursor
    left, right = stack.take(60).columns(2, gap=24)
    assert (left, right.x) == (Rect(48, top, 188, 60), 260) and stack.cursor == top + 72
    formula = stack.equation(id="formula", expression="x", indent=16, align="center")
    assert (formula.area.x, formula.area.w) == (64, 384) and formula.data["style"]["align"] == "center"
    figure = stack.image(id="figure", path="figure.png", width=100, place="right", gap=0)
    assert (figure.area.x, figure.area.w, figure.area.h) == (348, 100, 50)
    assert stack.elements == (lead, note, formula, figure)
    assert stack.remaining == Rect(48, figure.area.bottom, 400, 440 - figure.area.bottom)
    for element in stack.elements:
        assert element.data["layout_bounds"] == [48, 100, 400, 340]
    # 領域を超えた要素は、配置の検査が layout_bounds_overflow にする。
    stack.text(id="overflow", text=LONG, height=200)
    scene = deck.to_dict()
    # 領域に収まる take は、場面の issues に何も加えない。
    assert "issues" not in scene
    issues = [
        issue for issue in validate_scene(scene, deck.project, fonts) if issue["code"] == "layout_bounds_overflow"
    ]
    assert [issue["element_id"] for issue in issues] == ["overflow"]
    # 高さを内容から決める領域では、領域の検査を記録しない。
    open_stack = slide.stack(Rect(500, 100, 200, None))
    tail = open_stack.text(id="tail", text="ABCD")
    assert "layout_bounds" not in tail.data and open_stack.remaining.h is None
    assert open_stack.take(500) == Rect(500, tail.area.bottom + 12, 200, 500)


def test_stack_avoids_page_number_area(make_deck):
    """avoid="footer" は、幅を指定しない要素がページ番号の領域に重なるとき、領域の左に収まる幅で作り直す。"""
    deck = make_deck()
    slide = deck.slide(id="avoid")
    footer = slide.footer_area
    assert footer is not None
    width, height = deck.project.size
    area = Rect(48, footer.y - 40, footer.right - 48, height - (footer.y - 40))
    stack = slide.stack(area, gap=12, avoid="footer")
    stack.space(30)
    line = stack.text(id="line", text="ABCD EFGH")
    # 領域の左へ 8 pt（gap が 8 より大きいときは 8）空けて収める。重ならない要素はそのまま。
    assert line.area.right == pytest.approx(footer.x - 8) and line.area.x == 48
    assert not (line.area.right > footer.x and line.area.bottom > footer.y)
    plain = slide.stack(Rect(48, 100, 400, None), avoid="footer").text(id="plain", text="ABCD")
    assert plain.area.w == 400
    # 幅を指定した要素は動かさない。
    fixed = slide.stack(area, avoid="footer")
    fixed.space(30)
    kept = fixed.text(id="fixed", text="ABCD", width=area.w)
    assert kept.area.w == area.w
    # 全幅を覆う領域は、要素をその下へ送る。
    band = Rect(40, 200, 900, 30)
    below = slide.stack(Rect(48, 190, 400, None), avoid=[band]).text(id="below", text="ABCD")
    assert below.area.y == band.bottom + 12
    with pytest.raises(InputError, match="avoid は Rect"):
        slide.stack(area, avoid=["header"])
    # ページ番号を置かないページでは footer を避ける領域にしない。
    cover = deck.slide(id="cover", kind="cover")
    assert cover.footer_area is None and cover.stack(area, avoid="footer").text(id="c", text="AB").area.w == area.w


def test_stack_rejects_invalid_arguments_and_leaves_nothing(make_deck):
    deck = make_deck()
    png(deck.project.root, size=(100, 100))
    slide = deck.slide(id="stack")
    with pytest.raises(InputError, match="積み上げる間隔は 0 以上にしてください") as caught:
        slide.stack(Rect(0, 0, 300, 200), gap=-1)
    assert caught.value.code == "invalid_geometry"
    with pytest.raises(InputError, match="積み上げる領域の幅は正の数値で"):
        slide.stack(Rect(0, 0, None, 200))
    stack = slide.stack(Rect(0, 0, 300, 200))
    for operation, code, words in (
        (lambda: stack.text(id="a", text="A", gap=-2), "invalid_geometry", "積み上げる間隔は 0 以上"),
        (lambda: stack.text(id="a", text="A", indent=-1), "invalid_geometry", "indent は 0 以上"),
        (
            lambda: stack.text(id="a", text="A", indent=300),
            "invalid_geometry",
            "indent 300 pt が積み上げる領域の幅 300 pt 以上",
        ),
        (lambda: stack.text(id="a", text="A", place="middle"), "input_error", "place は left、center、right"),
        (lambda: stack.text(id="a", text="A", width=301), "invalid_geometry", "幅 301 pt が積み上げる領域の幅"),
        (lambda: stack.text(id="a", text="A", height=0), "invalid_geometry", "height は正の有限値"),
        (
            lambda: stack.text(id="a", text="A", area=Rect(0, 0, 10, 10)),
            "unsupported_argument",
            "area は stack の text では使えません",
        ),
        # 縦横比から決めた幅が領域を超える画像は、作った要素を残さない。
        (
            lambda: stack.image(id="a", path="figure.png", height=400),
            "invalid_geometry",
            "画像の縦横比から決めた幅 400 pt",
        ),
        (lambda: stack.text(id="a", text="A", colour="accent"), "unknown_argument", "colour"),
    ):
        with pytest.raises(InputError, match=re.escape(words)) as caught:
            operation()
        assert (caught.value.code, caught.value.slide_id) == (code, "stack"), words
        assert caught.value.element_id == "a", words
    assert not slide.elements and stack.cursor == 0 and stack.elements == ()
    for operation, words in (
        (lambda: stack.space(-1), "積み上げに足す間隔は 0 以上"),
        (lambda: stack.take(0), "height は正の有限値"),
        (lambda: stack.take(float("nan")), "height は正の有限値"),
    ):
        with pytest.raises(InputError, match=re.escape(words)):
            operation()
    assert stack.cursor == 0


def test_stack_take_beyond_area_is_reported_as_diagnostic(make_deck):
    deck = make_deck()
    slide = deck.slide(id="stack")
    stack = slide.stack(Rect(48, 100, 400, 60), gap=12)
    # 前に置いた要素が領域を超えても、続く take は例外にせず、領域を返して進む。
    lead = stack.text(id="lead", text=LONG * 2)
    assert lead.area.bottom > 160 and stack.remaining.h == 0
    top = stack.cursor
    assert stack.take(20) == Rect(48, top, 400, 20) and stack.cursor == top + 32
    # 下端までちょうど受け取る take と、高さを内容から決める領域の take は記録しない。
    assert slide.stack(Rect(500, 100, 200, 100)).take(100) == Rect(500, 100, 200, 100)
    slide.stack(Rect(500, 300, 200, None)).take(500)
    # まとまりの stack の超過は場面にするときのまとまりの位置で、ページの座標にする。
    group = slide.group(id="panel", area=Rect(0, 300, 200, 100))
    inner = group.stack(Rect(10, 10, 100, 50), gap=0)
    inner.take(30)
    inner.take(30)
    group.move(dy=10)
    scene = deck.to_dict()
    issues = [issue for issue in scene["issues"] if issue["code"] == "layout_bounds_overflow"]
    assert len(issues) == 1
    issue = issues[0]
    assert (issue["severity"], issue["scope"], issue["slide_id"], issue["element_id"], issue["slide_number"]) == (
        "error",
        "layout",
        "stack",
        None,
        1,
    )
    assert issue["message"] == "stack.take で受け取った領域が積み上げる領域を超えている"
    assert sorted(issue["observations"], key=lambda item: item["area"][0]) == [
        {"area": [10, 350, 100, 30], "bounds": [10, 320, 100, 50], "overflow_pt": {"side": "bottom", "amount": 10}},
        {
            "area": [48, top, 400, 20],
            "bounds": [48, 100, 400, 60],
            "overflow_pt": {"side": "bottom", "amount": round(top + 20 - 160, 2)},
        },
    ]
    # 置いた要素の超過は、これまでどおり配置の検査が報告する。
    found = validate_scene(scene, deck.project, deck.project.fonts)
    assert [item["element_id"] for item in found if item["code"] == "layout_bounds_overflow"] == ["lead"]


def test_stack_places_tables_as_one_element(make_deck):
    slide = make_deck().slide(id="stack")
    stack = slide.stack(Rect(48, 100, 400, 300), gap=10)
    table = {"headers": ["項目", "値"], "rows": [["A", "1"]], "row_keys": ["a"], "column_keys": ["item", "value"]}
    facts = stack.table(id="facts", height=60, width=200, place="center", **table)
    # 表はまとまりとして 1 つの要素とみなし、位置を変えるとセルも動く。積み上げの領域はまとまりに記録する。
    assert facts.area == Rect(148, 100, 200, 60) and stack.elements == (facts,) and stack.cursor == 170
    cells = {element.id: element for element in slide.elements}
    assert cells["facts/a/value"].area == Rect(248, 130, 100, 30)
    assert facts.data["layout_bounds"] == [48, 100, 400, 300] and "layout_bounds" not in cells["facts/a/value"].data
    # 失敗した表は、まとまりもセルも残さない。
    count = len(slide.elements)
    with pytest.raises(ValueError):
        stack.table(id="broken", height=60, **{**table, "rows": [["A"]]})
    assert len(slide.elements) == count and stack.cursor == 170


def test_group_stack_uses_group_coordinates(make_deck):
    slide = make_deck().slide(id="group")
    group = slide.group(id="panel", area=Rect(100, 50, 300, 400))
    stack = group.stack(Rect(10, 20, 280, 300), gap=10)
    first = stack.text(id="first", text="ABCD")
    assert (first.id, first.data["parent_id"]) == ("panel/first", "panel")
    assert first.area == Rect(110, 70, 280, slide.measure("ABCD", width=280).height)
    # cursor・take・remaining はまとまりの座標で、そのまま group の関数に渡すことができる。
    assert stack.cursor == 20 + first.area.h + 10
    taken = stack.take(30)
    assert taken == Rect(10, 30 + first.area.h, 280, 30)
    holder = group.box(id="holder", area=taken)
    assert holder.area == taken.translated(100, 50)
    second = stack.text(id="second", text="ABCD", indent=20, width=100, place="right")
    assert (second.area.x, second.area.y) == (290, 50 + taken.bottom + 10)
    assert first.data["layout_bounds"] == second.data["layout_bounds"] == [110, 70, 280, 300]
    # まとまりを動かすと、子と積み上げの領域も動く。
    group.move(5, 5)
    assert first.area.x == 115 and first.data["layout_bounds"] == [115, 75, 280, 300]
    assert by_id(slide.to_dict(1))["panel/second"]["layout_bounds"] == [115, 75, 280, 300]
    # まとまりの中のまとまりでも積み上げることができる。
    inner = group.group(id="inner", area=Rect(0, 0, 100, 100))
    nested = inner.stack(Rect(0, 0, 100, None)).text(id="deep", text="AB")
    assert (nested.id, nested.area.x, nested.area.y) == ("panel/inner/deep", 105, 55)


# ---- anchor ----


def test_anchor_gap_and_position_along_side(make_deck):
    slide = make_deck().slide(id="flow")
    first = slide.box(id="first", text="", area=Rect(100, 100, 200, 80))
    assert first.anchor().point() == (200, 140)
    assert first.anchor("left", gap=6, at=0.25).point() == (94, 120)
    assert first.anchor("right", gap=4).point() == (304, 140)
    assert first.anchor("top", at=0).point() == (100, 100)
    assert first.anchor("bottom", gap=2, at=1).point() == (300, 182)
    # 領域の anchor も同じ規則で、既定値では 0.1 と同じ点を返す。
    assert Rect(10, 20, 30, 40).anchor("left") == (10, 40) and Rect(10, 20, 30, 40).anchor("bottom") == (25, 60)
    assert Rect(0, 0, 100, 50).anchor("top", gap=5, at=0.2) == (20, -5)
    second = slide.box(id="second", text="", area=Rect(400, 100, 100, 80))
    slide.arrow(id="link", start=first.anchor("right", gap=6), end=second.anchor("left", gap=6, at=0.25))
    # 要素を動かすと端点も追従する。
    first.move(10, 20)
    link = by_id(slide.to_dict(1))["link"]
    assert (link["start"], link["end"]) == ([316, 160], [394, 120])
    assert link["area"] == [316, 120, 78, 40]


def test_anchor_follows_rotation_and_flip(make_deck):
    slide = make_deck().slide(id="turned")
    # 中心 (200, 150) の 200 × 100 の図形を時計回りに 90 度回すと、右の辺は下に、上の辺は右に来る。
    turned = slide.box(id="turned", text="", area=Rect(100, 100, 200, 100), rotation=90)
    assert turned.anchor("right").point() == (200, 250)
    assert turned.anchor("top").point() == (250, 150)
    assert turned.anchor("left", gap=10, at=0).point() == (250, 40)
    # 反転は中心について行い、その後に回す（PowerPoint の図形の変換の順）。
    flipped = slide.box(id="flipped", text="", area=Rect(100, 300, 200, 100), flip_h=True)
    assert flipped.anchor("left", at=0.25).point() == (300, 325)
    upside = slide.box(id="upside", text="", area=Rect(100, 300, 200, 100), flip_v=True)
    assert upside.anchor("top", at=0.25).point() == (150, 400)
    both = slide.box(id="both", text="", area=Rect(100, 300, 200, 100), flip_h=True, rotation=90)
    assert both.anchor("left", at=0.25).point() == (225, 450)
    slanted = slide.box(id="slanted", text="", area=Rect(0, 0, 200, 100), rotation=45)
    x, y = slanted.anchor("right").point()
    assert math.isclose(x, 100 + 100 / math.sqrt(2)) and math.isclose(y, 50 + 100 / math.sqrt(2))
    # 回転した要素に結んだ線の端点は、回転後の辺の点になる。
    slide.line(id="rule", start=turned.anchor("bottom"), end=(500, 500))
    assert by_id(slide.to_dict(1))["rule"]["start"] == [150, 150]


def test_anchor_uses_visible_image_area(make_deck):
    deck = make_deck()
    root = deck.project.root
    slide = deck.slide(id="images")
    sources = [
        png(root),
        pdf(root),
        svg(root, "sized.svg", 'width="200" height="100"'),
        svg(root, "box.svg", 'viewBox="0 0 400 200"'),
    ]
    for index, path in enumerate(sources):
        page = {"page": 1} if path.endswith(".pdf") else {}
        # 横長の画像を正方形の領域に contain で置くと、表示される範囲は Rect(100, 150, 200, 100)。
        image = slide.image(id=f"image_{index}", path=path, area=Rect(100, 100, 200, 200), **page)
        assert image.anchor("top").point() == (200, 150), path
        assert image.anchor("bottom", gap=4).point() == (200, 254), path
        assert image.anchor("left", at=0).point() == (100, 150), path
    elements = by_id(slide.to_dict(1))
    assert elements["image_0"]["actual_area"] == [100, 150, 200, 100]
    # PDF・SVG の表示範囲は、素材の準備が変換後の画素数で決める（場面には書かない）。
    assert "actual_area" not in elements["image_1"] and "image_size" not in elements["image_1"]
    # cover と、寸法の分からない SVG は、配置領域を基準にする。
    cover = slide.image(id="cover", path=sources[0], area=Rect(100, 100, 200, 200), fit="cover")
    unknown = slide.image(id="unknown", path=svg(root, "unknown.svg", ""), area=Rect(100, 100, 200, 200))
    assert cover.anchor("top").point() == unknown.anchor("top").point() == (200, 100)


def test_anchor_rejects_invalid_gap_and_position(make_deck):
    slide = make_deck().slide(id="flow")
    first = slide.box(id="first", text="", area=Rect(100, 100, 200, 80))
    holder = slide.group(id="holder")
    for operation, words in (
        (lambda: first.anchor("left", gap=-2), "gap は 0 以上の数値で指定してください: -2"),
        (lambda: first.anchor("left", gap=True), "gap は 0 以上の数値で指定してください: True"),
        (lambda: first.anchor("top", at=1.5), "at は 0 以上 1 以下で指定してください: 1.5"),
        (lambda: first.anchor("top", at=float("nan")), "at は 0 以上 1 以下で指定してください: nan"),
        (lambda: first.anchor("center", gap=4), "center には gap・at を指定できません"),
        (lambda: first.anchor("center", at=0), "center には gap・at を指定できません"),
        (lambda: first.anchor("middle"), "未対応の接続位置です: 'middle'（候補: left、right、top、bottom、center）"),
        (lambda: holder.anchor("left"), "領域のないまとまりには接続できません"),
    ):
        with pytest.raises(InputError, match=re.escape(words)) as caught:
            operation()
        error = caught.value
        assert (error.code, error.slide_id) == ("invalid_anchor", "flow"), words
        assert error.element_id in ("first", "holder") and str(error).startswith(
            f"ページ flow の要素 {error.element_id}: "
        )
    with pytest.raises(InputError) as caught:
        Rect(0, 0, 10, 10).anchor("right", gap=-1)
    assert caught.value.code == "invalid_anchor"
    # 領域のあるまとまりには接続できる。
    framed = slide.group(id="framed", area=Rect(0, 0, 100, 50))
    assert framed.anchor("right").point() == (100, 25)


def test_anchor_reference_records_only_changed_values(make_deck):
    slide = make_deck().slide(id="flow")
    first = slide.box(id="first", text="", area=Rect(100, 100, 200, 80))
    second = slide.box(id="second", text="", area=Rect(400, 100, 100, 80))
    slide.arrow(id="plain", start=first.anchor("right"), end=second.anchor("left"))
    slide.arrow(id="spaced", start=first.anchor("right", gap=6, at=0.5), end=second.anchor("left", gap=0, at=0.25))
    elements = by_id(slide.to_dict(1))
    assert elements["plain"]["start_ref"] == {"element_id": "first", "side": "right"}
    assert elements["plain"]["end_ref"] == {"element_id": "second", "side": "left"}
    assert elements["spaced"]["start_ref"] == {"element_id": "first", "side": "right", "gap": 6.0}
    assert elements["spaced"]["end_ref"] == {"element_id": "second", "side": "left", "at": 0.25}


# ---- 画像の倍率と埋め込み（作成 API） ----


def test_raster_options_are_rejected_for_png_and_out_of_range(make_deck):
    deck = make_deck()
    root = deck.project.root
    png(root)
    for name in ("photo.jpg", "photo.jpeg", "PHOTO.JPG"):
        Image.new("RGB", (200, 100), "white").save(root / name, format="JPEG")
    pdf(root)
    svg(root, "figure.svg", 'width="200" height="100"')
    slide = deck.slide(id="images")
    area = Rect(0, 0, 200, 100)
    for operation, words in (
        (
            lambda: slide.image(id="photo", path="figure.png", area=area, raster_scale=3),
            "raster_scale は PDF・SVG の素材だけに指定できます",
        ),
        (
            lambda: slide.image(id="photo", path="figure.png", area=area, vector=False),
            "vector は PDF・SVG の素材だけに",
        ),
        # JPEG も画素数のまま使う素材で、拡張子の大文字・小文字によらない。
        (
            lambda: slide.image(id="photo", path="photo.jpg", area=area, raster_scale=2),
            "raster_scale は PDF・SVG の素材だけに指定できます（.jpg の素材",
        ),
        (
            lambda: slide.image(id="photo", path="photo.jpeg", area=area, vector=True),
            "vector は PDF・SVG の素材だけに指定できます（.jpeg の素材",
        ),
        (
            lambda: slide.image(id="photo", path="PHOTO.JPG", area=area, raster_scale=2, vector=False),
            "raster_scale・vector は PDF・SVG の素材だけに",
        ),
        (
            lambda: slide.image(id="photo", path="diagram.pdf", page=1, area=area, raster_scale=0.5),
            "raster_scale は 1 以上 16 以下の数",
        ),
        (
            lambda: slide.image(id="photo", path="diagram.pdf", page=1, area=area, raster_scale=17),
            "raster_scale は 1 以上 16 以下",
        ),
        (
            lambda: slide.image(id="photo", path="diagram.pdf", page=1, area=area, raster_scale=True),
            "raster_scale は 1 以上 16 以下",
        ),
        (
            lambda: slide.image(id="photo", path="figure.svg", area=area, raster_scale="2"),
            "raster_scale は 1 以上 16 以下",
        ),
        (
            lambda: slide.image(id="photo", path="figure.svg", area=area, vector="yes"),
            "vector は真偽値で指定してください",
        ),
        # PDF の素材は PNG だけにする（vector は SVG の素材だけ）。
        (
            lambda: slide.image(id="photo", path="diagram.pdf", page=1, area=area, raster_scale=4, vector=True),
            PDF_VECTOR_MESSAGE,
        ),
    ):
        with pytest.raises(InputError, match=re.escape(words)) as caught:
            operation()
        assert (caught.value.code, caught.value.element_id) == ("invalid_image_option", "photo"), words
    assert not slide.elements
    # 指定した項目だけを場面に書く（None は deck.toml の値を使う。JPEG にも None は渡すことができる）。
    vector = slide.image(id="vector", path="figure.svg", area=area, raster_scale=4, vector=True)
    scaled = slide.image(id="scaled", path="figure.svg", area=area, raster_scale=1.5, vector=None)
    raster = slide.image(id="raster", path="diagram.pdf", page=1, area=area, raster_scale=4, vector=False)
    plain = slide.image(id="plain", path="diagram.pdf", page=1, area=area)
    slide.image(id="jpeg", path="photo.jpg", area=area, raster_scale=None, vector=None)
    elements = by_id(slide.to_dict(1))
    assert (elements["vector"]["raster_scale"], elements["vector"]["vector"]) == (4.0, True)
    assert elements["scaled"]["raster_scale"] == 1.5 and "vector" not in elements["scaled"]
    assert (elements["raster"]["raster_scale"], elements["raster"]["vector"]) == (4.0, False)
    assert not {"raster_scale", "vector"} & (set(elements["plain"]) | set(elements["jpeg"]))
    assert [vector.id, scaled.id, raster.id, plain.id] == ["vector", "scaled", "raster", "plain"]


def test_image_input_errors_name_the_element(make_deck):
    deck = make_deck()
    root = deck.project.root
    (root / "broken.png").write_bytes(b"not an image")
    pdf(root)
    slide = deck.slide(id="images")
    area = Rect(0, 0, 200, 100)
    for operation, code, words in (
        (
            lambda: slide.image(id="photo", path="missing.png", area=area),
            "asset_missing",
            "入力ファイルが見つかりません",
        ),
        (lambda: slide.image(id="photo", path="../outside.png", area=area), "input_error", "参照範囲外"),
        (lambda: slide.image(id="photo", path="broken.png", area=area), "input_error", "画像を読み込めません"),
        (
            lambda: slide.image(id="photo", path="diagram.pdf", page=3, area=area),
            "input_error",
            "PDF にページ 3 がありません（全 1 ページ）",
        ),
    ):
        with pytest.raises(InputError, match=re.escape(words)) as caught:
            operation()
        error = caught.value
        assert (error.code, error.slide_id, error.element_id) == (code, "images", "photo"), words
        assert str(error).startswith("ページ images の要素 photo: ")
    assert not slide.elements


def test_phrase_wrap_inserts_breaks_at_phrase_boundaries(make_deck):
    """phrase_wrap は文節の境で折り返し、場面の断片にその位置の明示改行を入れる（PowerPoint の折り返しを固定する）。"""
    deck = make_deck()
    slide = deck.slide(id="phrase")
    text = "届いたため、末尾を指す処理を書く"
    plain = slide.text(id="plain", text=text, area=Rect(48, 100, 70, None), size_pt=20, role="body")
    phrased = slide.text(id="phrased", text=text, area=Rect(48, 300, 70, None), size_pt=20, phrase_wrap=True)
    assert [line.wrap for line in plain.measure().lines].count("mid_phrase") >= 1
    lines = phrased.measure().lines
    assert "mid_phrase" not in [line.wrap for line in lines]
    elements = {element["id"]: element for element in deck.to_dict()["slides"][0]["elements"]}
    written = "".join(fragment["text"] for fragment in elements["phrased"]["content"])
    assert written.split("\n") == [line.text.rstrip() for line in lines]
    # 指定のない要素の断片は変えない。
    assert "".join(fragment["text"] for fragment in elements["plain"]["content"]) == text


def test_stack_places_callouts_by_body_and_rechecks_the_tail(make_deck):
    deck = make_deck()
    deck.project.collect_errors = True
    slide = deck.slide(id="stack", kind="blank")
    stack = slide.stack(Rect(48, 100, 400, None), gap=12)
    first = stack.box(id="first", text="ABCD", tail_side="bottom", height=40)
    second = stack.box(id="second", text="ABCD", tail_side="left", height=40)
    # 積み上げは本体（area）で並べる。下の辺のしっぽの長さは間隔に加えない。
    assert second.area.y == first.area.bottom + 12
    found = by_id(deck.to_dict()["slides"][0])
    assert found["first"]["actual_area"] == [48, 100, 400, 48]
    assert found["second"]["tail"]["side"] == "left"
    # 作った後に辺が短くなり、付け根が収まらなくなったときは、場面を作るときに報告してしっぽを付けない。
    second.area = Rect(second.area.x, second.area.y, 400, 9)
    scene = deck.to_dict()
    assert "tail" not in by_id(scene["slides"][0])["second"]
    assert [(issue["element_id"], issue["code"]) for issue in scene["issues"]] == [("second", "invalid_style")]
