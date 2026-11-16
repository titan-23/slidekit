"""作成 API の中核を確かめる：ページの題名、数式の役割、断片の書式とリンク、識別子の補助、続行できる誤りの集約、
折り返しと辺ごとの余白、内容から決める寸法の Rect、使う場所を限った役割の宣言、警告の許容の宣言、構造の誤りの識別子。"""

import importlib.util
import itertools
import re

import pytest

from fake_fonts import FakeFonts
from slidekit import Deck, Emphasis, Math, Rect, Text, key_id
from slidekit.config import load_project
from slidekit.errors import InputError
from slidekit.markup import parse
from slidekit.math.syntax import MathSyntaxError

AREA = Rect(40, 80, 400, 60)
SMALL = '[roles.small]\ninherits = "note"\nrestricted = true\n'


@pytest.fixture
def make_deck(tmp_path, write_deck):
    """資料を書いて読み込み、Deck を返す関数。theme は資料の theme.toml、collect は collect_errors。"""
    numbers = itertools.count()

    def make(theme=None, *, collect=False):
        project = load_project(write_deck(tmp_path / f"deck{next(numbers)}", theme=theme))
        project.collect_errors = collect
        return Deck(project)

    return make


def by_id(page):
    return {element["id"]: element for element in page["elements"]}


# ---- ページの題名 ----


def test_title_accepts_fragments_and_scene_title_is_plain(make_deck):
    deck = make_deck()
    slide = deck.slide(id="sum", title=[Text("区間 "), Math("[l, r)"), Text(" の和")])
    deck.slide(id="parsed", title=parse("**重要** な $x$"))
    deck.slide(id="single", title=Emphasis("強調"))
    deck.slide(id="none", title=None, notes=None)
    pages = {page["id"]: page for page in deck.to_dict()["slides"]}
    assert slide.title == pages["sum"]["title"] == "区間 $[l, r)$ の和"
    assert by_id(pages["sum"])["_theme/title"]["content"] == [
        {"kind": "text", "text": "区間 "},
        {"kind": "math", "text": "[l, r)"},
        {"kind": "text", "text": " の和"},
    ]
    assert pages["parsed"]["title"] == "重要 な $x$"
    assert by_id(pages["single"])["_theme/title"]["content"] == [{"kind": "emphasis", "text": "強調"}]
    # 題名なしは空の文字列で、題名の要素を作らない。ノートの None は空のノート。
    assert (pages["none"]["title"], pages["none"]["notes"]) == ("", "")
    assert "_theme/title" not in by_id(pages["none"])


def test_title_role_area_and_style_override_per_slide(make_deck):
    deck = make_deck('[page_kinds.content]\ntitle_style = {align = "center", size_pt = 30}\n')
    deck.slide(id="default", title="既定")
    deck.slide(
        id="custom",
        title="上書き",
        title_role="heading",
        title_area=Rect(48, 20, 600, 50),
        title_style={"size_pt": 28, "color": "accent"},
    )
    deck.slide(id="blank", kind="blank", title="白紙")
    deck.slide(id="placed", kind="blank", title="白紙に置く", title_area=(40, 30, 500, 40))
    pages = {page["id"]: by_id(page) for page in deck.to_dict()["slides"]}
    default = pages["default"]["_theme/title"]
    assert (default["role"], default["style"]["align"], default["style"]["size_pt"]) == ("slide_title", "center", 30)
    # 役割 < ページ種別の title_style < ページの title_style の順に重ねる。
    custom = pages["custom"]["_theme/title"]
    assert (custom["role"], custom["area"]) == ("heading", [48, 20, 600, 50])
    assert (custom["style"]["size_pt"], custom["style"]["align"], custom["style"]["color"]) == (28, "center", "C43D3D")
    assert custom["style"]["bold"] is True
    # blank は題名の領域を指定したときだけ題名を置き、役割は slide_title にする。
    assert "_theme/title" not in pages["blank"]
    placed = pages["placed"]["_theme/title"]
    assert (placed["role"], placed["area"]) == ("slide_title", [40, 30, 500, 40])


def test_title_role_restrictions_and_style_names_are_checked(make_deck):
    deck = make_deck('[roles.caption]\ninherits = "note"\nallowed_slide_ids = ["summary"]\n')
    with pytest.raises(InputError) as caught:
        deck.slide(id="intro", title="題名", title_role="caption")
    assert (caught.value.code, caught.value.slide_id, caught.value.element_id) == (
        "role_usage",
        "intro",
        "_theme/title",
    )
    deck.slide(id="summary", title="まとめ", title_role="caption")
    with pytest.raises(TypeError) as caught:
        deck.slide(id="styled", title="題名", title_style={"colour": "accent"})
    assert "colour（候補: color）" in str(caught.value) and caught.value.element_id == "_theme/title"
    # 題名の役割の名前は、題名の領域がないページ（blank）でも確かめる。
    for kind in ("content", "blank"):
        with pytest.raises(InputError, match="未定義の文字の役割です: 'missing'") as caught:
            deck.slide(id=f"unknown_{kind}", kind=kind, title="題名", title_role="missing")
        assert (caught.value.code, caught.value.element_id) == ("unknown_role", "_theme/title")
    assert [slide.id for slide in deck.slides] == ["summary"]


# ---- 数式の役割 ----


def test_equation_defaults_to_equation_role_with_body_values(make_deck):
    slide = make_deck().slide(id="formula")
    equation = slide.equation(id="sum", expression="a + b", area=AREA)
    text = slide.text(id="body", text="本文", area=AREA)
    assert equation.data["role"] == "equation"
    equation_style, body_style = dict(equation.data["style"]), dict(text.data["style"])
    # 独立した数式は既定で折り返さない（同梱テーマの [defaults.equation]）。ほかの値は body と同じ。
    assert equation_style.pop("word_wrap") is False and body_style.pop("word_wrap") is True
    assert equation_style == body_style
    explicit = slide.equation(id="explicit", expression="x", area=AREA, role="body")
    assert explicit.data["role"] == "body" and explicit.data["style"]["word_wrap"] is False


def test_deck_theme_can_style_equation_role(make_deck):
    slide = make_deck(
        '[roles.equation]\nsize_pt = 24\nmin_size_pt = 20\ncolor = "primary"\nalign = "center"\nvalign = "middle"\n'
    ).slide(id="formula")
    styled = slide.equation(id="styled", expression="x^2", area=AREA)
    plain = slide.equation(id="plain", expression="x", area=AREA, role="body")
    values = ("size_pt", "min_size_pt", "color", "align", "valign")
    assert styled.data["role"] == "equation"
    assert tuple(styled.data["style"][name] for name in values) == (24, 20, "183153", "center", "middle")
    assert tuple(plain.data["style"][name] for name in values) == (20, 18, "20252B", "left", "top")


def test_equation_requires_expression(make_deck):
    slide = make_deck().slide(id="formula")
    # 使われていなかった text= は受け付けない。
    with pytest.raises(TypeError, match="expression で指定") as caught:
        slide.equation(id="old", text="x", area=AREA)
    assert caught.value.code == "unsupported_argument"
    for value in (None, "", 3):
        with pytest.raises(InputError, match="expression") as caught:
            slide.equation(id="empty", expression=value, area=AREA)
        assert caught.value.element_id == "empty"
    assert not slide.elements


# ---- 断片の書式とリンク ----


def test_fragment_styles_are_resolved_and_written_only_when_given(make_deck):
    slide = make_deck().slide(id="contact")
    element = slide.text(
        id="contact",
        area=Rect(48, 420, 864, 40),
        text=[
            Text("詳細は "),
            Text("利用ガイド", color="primary", underline=True, link="https://example.com/guide"),
            Text(" を参照。"),
            Text("（2026 年版）", size_pt=14, color="#606b78"),
            Emphasis("注", bold=False),
            Text("1", superscript=True),
        ],
    )
    assert element.data["content"] == [
        {"kind": "text", "text": "詳細は "},
        {
            "kind": "text",
            "text": "利用ガイド",
            "style": {"color": "183153", "underline": True},
            "link": "https://example.com/guide",
        },
        {"kind": "text", "text": " を参照。"},
        {"kind": "text", "text": "（2026 年版）", "style": {"size_pt": 14, "color": "606B78"}},
        {"kind": "emphasis", "text": "注", "style": {"bold": False}},
        {"kind": "text", "text": "1", "style": {"superscript": True}},
    ]
    with pytest.raises(InputError, match="Text の superscript は真偽値で指定してください"):
        Text("1", superscript="yes")
    # 要素の書式は断片の指定で変わらない。
    assert element.data["style"]["color"] == "20252B"


def test_fragment_colors_and_links_are_validated(make_deck):
    slide = make_deck().slide(id="links")
    colors = "（テーマの色: accent、background、line、muted、primary、surface、text）"
    cases = [
        (Text("a", color="tela"), "undefined_color", f"断片の color に未定義の色を指定しています: 'tela'{colors}"),
        (
            Emphasis("a", color="none"),
            "invalid_style",
            f"断片の color に none は指定できません。none は fill・line_color だけに使えます{colors}",
        ),
        (
            Math("x", color="#12345"),
            "invalid_style",
            f"断片の color は色の名前か 6 桁の 16 進数で指定してください: '#12345'{colors}",
        ),
        (Text("a", link="ftp://example.com/"), "invalid_link", "https://"),
        (Text("a", link="https://example.com/a b"), "invalid_link", "空白を含まない"),
        (Emphasis("a", link="example.com"), "invalid_link", "mailto:"),
    ]
    for number, (fragment, code, words) in enumerate(cases):
        with pytest.raises(InputError, match=re.escape(words)) as caught:
            slide.text(id=f"t{number}", text=[fragment], area=AREA)
        assert (caught.value.code, caught.value.element_id) == (code, f"t{number}")
    assert not slide.elements
    element = slide.text(
        id="links",
        area=AREA,
        text=[Text("連絡", link="mailto:info@example.com"), Text("旧", link="http://example.com/")],
    )
    assert [fragment["link"] for fragment in element.data["content"]] == [
        "mailto:info@example.com",
        "http://example.com/",
    ]


def test_math_emphasis_uses_element_emphasis_color(make_deck):
    slide = make_deck().slide(id="math")
    element = slide.text(
        id="formula",
        area=AREA,
        emphasis_color="primary",
        text=parse("**最大 $k$ 件**") + [Math("y", emphasis=True, bold=False), Math("z", emphasis=True, color="muted")],
    )
    content = element.data["content"]
    assert content[1] == {"kind": "math", "text": "k", "style": {"bold": True, "color": "183153"}}
    assert content[3]["style"] == {"bold": False, "color": "183153"}
    assert content[4]["style"] == {"bold": True, "color": "606B78"}
    assert not any("emphasis" in fragment for fragment in content)


def test_parse_result_can_be_used_in_text_table_and_title(make_deck):
    deck = make_deck()
    slide = deck.slide(id="parsed", title=parse("件数 $n$ の**上限**"))
    slide.text(id="premise", text=parse("対象の件数を $n$ とし、**最大 $k$ 件**まで扱う"), area=Rect(48, 100, 864, 60))
    slide.table(
        id="limits",
        headers=[parse("**項目**"), "値"],
        rows=[[parse("上限 $k$"), "10"]],
        row_keys=["upper"],
        column_keys=["item", "value"],
        area=Rect(48, 200, 600, 120),
    )
    page = deck.to_dict()["slides"][0]
    elements = by_id(page)
    assert page["title"] == "件数 $n$ の上限"
    assert [fragment["kind"] for fragment in elements["_theme/title"]["content"]] == [
        "text",
        "math",
        "text",
        "emphasis",
    ]
    assert [fragment["kind"] for fragment in elements["premise"]["content"]] == [
        "text",
        "math",
        "text",
        "emphasis",
        "math",
        "emphasis",
        "text",
    ]
    assert elements["premise"]["content"][4]["style"] == {"bold": True, "color": "C43D3D"}
    assert elements["limits/_header/item"]["content"] == [{"kind": "emphasis", "text": "項目"}]
    assert elements["limits/upper/item"]["content"] == [
        {"kind": "text", "text": "上限 "},
        {"kind": "math", "text": "k"},
    ]


# ---- 識別子の補助 ----


def test_key_id_joins_and_validates_keys(make_deck):
    assert key_id("step", "prepare") == "step_prepare"
    assert key_id("row", 3, -1) == "row_3_-1"
    for keys in ((), ("",), (" ",), (True,), (1.5,), (None,), ("a/b",), ("a\tb",), ("_theme", "x")):
        with pytest.raises(InputError) as caught:
            key_id(*keys)
        assert caught.value.code == "invalid_element_id"
    slide = make_deck().slide(id="steps")
    for index, (key, label) in enumerate((("prepare", "準備"), ("review", "確認"))):
        slide.box(id=key_id("step", key), text=label, area=Rect(48 + 300 * index, 150, 260, 60))
    assert [element.id for element in slide.elements] == ["step_prepare", "step_review"]


def test_group_without_area_prefixes_ids_without_moving(make_deck):
    slide = make_deck().slide(id="layout")
    steps = slide.group(id="steps")
    title = steps.text(id="title", text="手順", area=Rect(48, 90, 300, 36))
    steps.line(id="rule", start=(48, 130), end=(348, 130))
    assert (title.id, title.area) == ("steps/title", Rect(48, 90, 300, 36))
    elements = by_id(slide.to_dict(1))
    assert elements["steps/title"]["parent_id"] == "steps"
    assert elements["steps/rule"]["start"] == [48, 130] and elements["steps/rule"]["end"] == [348, 130]


# ---- 続行できる誤りの集約 ----


def test_errors_raise_immediately_by_default(make_deck):
    # collect_errors を持たない簡易な資料でも、既定では最初の誤りで例外にする。
    assert Deck(object()).collect_errors is False
    deck = make_deck()
    assert deck.collect_errors is False
    slide = deck.slide(id="plan")
    with pytest.raises(InputError) as caught:
        slide.box(id="first", text="最初", area=AREA, fill="tela", colour="accent")
    error = caught.value
    assert (error.code, error.slide_id, error.element_id) == ("unknown_argument", "plan", "first")
    assert not slide.elements and not deck.issues


def test_collected_errors_are_reported_with_fallbacks(make_deck):
    deck = make_deck(SMALL, collect=True)
    slide = deck.slide(id="plan", title="計画", title_style={"colour": "accent"})
    first = slide.box(id="first", text="最初", area=AREA, fill="tela", colour="accent", size_pt=24)
    second = slide.text(id="second", text="二番目", area=AREA, role="small", min_size_pt=1)
    third = slide.text(
        id="third", text=[Text("参照", color="nothing", link="example.com")], area=AREA, semantic="title"
    )
    formula = slide.equation(id="formula", expression="x", area=AREA, role="missing")
    # 文字列でない semantic と、style= の文字列でないキーも、続行できる誤りとして集める。
    fourth = slide.text(id="fourth", text="四番目", area=AREA, semantic=["background"], style={1: 2, "bold": True})
    scene = deck.to_dict()
    assert {(issue["element_id"], issue["code"]) for issue in scene["issues"]} == {
        ("_theme/title", "unknown_argument"),
        ("first", "unknown_argument"),
        ("first", "undefined_color"),
        ("second", "role_usage"),
        ("second", "minimum_override"),
        ("third", "undefined_color"),
        ("third", "invalid_link"),
        ("third", "input_error"),
        ("formula", "unknown_role"),
        ("fourth", "input_error"),
        ("fourth", "unknown_argument"),
    }
    assert all(
        (issue["severity"], issue["scope"], issue["slide_id"], issue["slide_number"]) == ("error", "input", "plan", 1)
        for issue in scene["issues"]
    )
    assert any("colour（候補: color）" in issue["message"] for issue in scene["issues"])
    assert any(
        issue["message"] == "ページ plan の要素 fourth: text に未知の書式があります: 1" for issue in scene["issues"]
    )
    # 間違った指定は使わず、下の層の値とページの既定の役割で作る。
    assert (first.data["style"]["fill"], first.data["style"]["size_pt"]) == ("F3F6FA", 24)
    assert (second.data["role"], second.data["style"]["min_size_pt"]) == ("body", 18)
    assert third.data["content"] == [{"kind": "text", "text": "参照"}] and "semantic" not in third.data
    assert formula.data["role"] == "body"
    assert "semantic" not in fourth.data and fourth.data["style"]["bold"] is True and 1 not in fourth.data["style"]
    # 誤りのない資料の場面には issues を作らない。
    assert "issues" not in make_deck(collect=True).to_dict()


def test_unrecoverable_errors_raise_even_when_collecting(make_deck):
    deck = make_deck(collect=True)
    slide = deck.slide(id="plan")
    slide.text(id="body", text="本文", area=AREA)
    for operation, code in (
        (lambda: slide.text(id="body", text="重複", area=AREA), "invalid_element_id"),
        (lambda: slide.text(id="flat", text="平ら", area=Rect(40, 120, 200, 0)), "invalid_geometry"),
        (lambda: slide.text(id="tab", text="a\tb", area=AREA), "invalid_fragment"),
        (
            lambda: slide.table(
                id="table",
                headers=["項目"],
                rows=[["a\tb"]],
                row_keys=["row"],
                column_keys=["item"],
                area=Rect(40, 160, 400, 120),
            ),
            "invalid_fragment",
        ),
        (lambda: deck.slide(id="plan"), "invalid_slide_id"),
        # 警告の許容の宣言の誤りは、許容の範囲が変わらないように集めない。
        (lambda: slide.allow_warning("element_overlap", reason="理由"), "input_error"),
    ):
        with pytest.raises(InputError) as caught:
            operation()
        assert caught.value.code == code
    assert not deck.issues and [element.id for element in slide.elements] == ["body"]


def test_failed_table_leaves_no_elements_and_names_the_cell(make_deck):
    class SyntaxFonts(FakeFonts):
        def math_box(self, expr, size, bold=False, **options):
            raise MathSyntaxError("未対応の数式命令", expression=expr, position=0, token=expr)

    deck = make_deck()
    deck.project.fonts = FakeFonts()
    slide = deck.slide(id="facts")
    slide.text(id="before", text="前", area=AREA)
    table = {
        "headers": ["項目", "値"],
        "row_keys": ["count"],
        "column_keys": ["item", "value"],
        "area": Rect(40, 160, 400, 120),
    }
    # セルの文章の誤りは自動寸法の有無によらずセルの識別子で報告する。
    for widths in (None, "auto"):
        with pytest.raises(InputError, match=re.escape("ページ facts の要素 table/count/item: 文章にタブ")) as caught:
            slide.table(id="table", rows=[["件\t数", "3"]], column_widths=widths, **table)
        assert (caught.value.code, caught.value.slide_id, caught.value.element_id) == (
            "invalid_fragment",
            "facts",
            "table/count/item",
        )
    # 自動寸法の計測の誤り（計測できない数式、数式の記法の誤り）も同じ。
    with pytest.raises(InputError, match="ページ facts の要素 table/count/value: ") as caught:
        slide.table(id="table", rows=[["件数", Math(r"\unknown")]], row_heights="auto", **table)
    assert (caught.value.code, caught.value.element_id) == ("text_measurement_failed", "table/count/value")
    deck.project.fonts = SyntaxFonts()
    with pytest.raises(
        MathSyntaxError, match=re.escape(r"ページ facts の要素 table/count/value: 未対応の数式命令: \frac")
    ) as caught:
        slide.table(id="table", rows=[["件数", Math(r"\frac")]], column_widths="auto", **table)
    assert (caught.value.slide_id, caught.value.element_id) == ("facts", "table/count/value")
    # 見出しのセルを作った後の誤り（本文のセルの役割）でも、まとまりと作ったセルを残さない。
    with pytest.raises(InputError) as caught:
        slide.table(id="table", rows=[["件数", "3"]], cell_roles={"count/value": "missing"}, **table)
    assert (caught.value.code, caught.value.element_id) == ("unknown_role", "table/count/value")
    assert [element.id for element in slide.elements] == ["before"]
    slide.table(id="table", rows=[["件数", "3"]], **table)
    assert [element.id for element in slide.elements] == [
        "before",
        "table",
        "table/_header/item",
        "table/_header/value",
        "table/count/item",
        "table/count/value",
    ]


# ---- 折り返し・辺ごとの余白・内容から決める寸法 ----


def test_control_characters_are_rejected(make_deck):
    slide = make_deck().slide(id="text")
    for number, value in enumerate(("列1\t列2", "行1\r\n行2", ["前", Emphasis("\x0b")], "a b")):
        with pytest.raises(InputError, match="制御文字は使えません（改行は \\\\n だけを使えます）") as caught:
            slide.text(id=f"t{number}", text=value, area=AREA)
        assert (caught.value.code, caught.value.element_id) == ("invalid_fragment", f"t{number}")
    assert caught.value.details["character"] == "U+2028"
    assert not slide.elements
    slide.text(id="lines", text="1 行目\n2 行目", area=AREA)
    with pytest.raises(InputError, match="U\\+0009"):
        make_deck().slide(id="title", title="題\t名")


def test_word_wrap_style_key(make_deck):
    library = make_deck().slide(id="library")
    assert library.text(id="body", text="本文", area=AREA).data["style"]["word_wrap"] is True
    assert library.text(id="off", text="1 行", area=AREA, word_wrap=False).data["style"]["word_wrap"] is False
    slide = make_deck('[defaults]\nword_wrap = false\n[roles.wrapped]\ninherits = "body"\nword_wrap = true\n').slide(
        id="theme"
    )
    # ライブラリ < テーマ < 役割 < 要素の順。None は下の層の値を使う。
    assert slide.text(id="theme", text="本文", area=AREA).data["style"]["word_wrap"] is False
    assert slide.text(id="role", text="本文", area=AREA, role="wrapped").data["style"]["word_wrap"] is True
    assert (
        slide.text(id="none", text="本文", area=AREA, role="wrapped", word_wrap=None).data["style"]["word_wrap"] is True
    )
    assert (
        slide.text(id="element", text="本文", area=AREA, role="wrapped", style={"word_wrap": False}).data["style"][
            "word_wrap"
        ]
        is False
    )
    with pytest.raises(InputError, match="word_wrap は真偽値で指定してください") as caught:
        slide.text(id="bad", text="本文", area=AREA, word_wrap="no")
    assert caught.value.code == "invalid_style"


def test_partial_padding_inherits_kind_default(make_deck):
    slide = make_deck().slide(id="padding")
    assert slide.box(id="default", area=AREA, text="既定").data["style"]["padding_pt"] == 6
    left = slide.box(id="left", area=AREA, text="左", padding_pt={"left": 12})
    assert left.data["style"]["padding_pt"] == {"left": 12, "top": 6, "right": 6, "bottom": 6}
    # 4 辺が同じなら数にそろえる。
    assert (
        slide.box(id="same", area=AREA, text="同じ", padding_pt={"left": 6, "right": 6}).data["style"]["padding_pt"]
        == 6
    )
    text = slide.text(id="text", area=AREA, text="文字", style={"padding_pt": {"top": 4}}, padding_pt={"bottom": 2})
    assert text.data["style"]["padding_pt"] == {"left": 0, "top": 4, "right": 0, "bottom": 2}
    for value, words in (
        ([6, 6], "辞書で指定"),
        ({"lft": 3}, "lft（候補: left）"),
        (-1, "0 以上"),
        ({"left": -2}, "0 以上"),
    ):
        with pytest.raises(InputError, match=words) as caught:
            slide.box(id="bad", area=AREA, text="誤り", padding_pt=value)
        assert caught.value.code == "invalid_style"


def test_rect_with_unknown_height_rejects_bottom():
    area = Rect(48, 100, 864, None)
    assert list(area) == [48, 100, 864, None] and area.translated(10, 5) == Rect(58, 105, 864, None)
    with pytest.raises(
        InputError,
        match=re.escape(
            "高さを内容から決める領域（h=None）では bottom を求められません。"
            "要素を作成した後の element.area を使ってください"
        ),
    ) as caught:
        area.bottom
    assert caught.value.code == "invalid_geometry"
    with pytest.raises(InputError, match="h=None）では rows"):
        area.rows(2)
    left, right = area.columns(2, gap=24)
    assert left.h is None and right == Rect(492, 100, 420, None)
    for operation in (
        lambda: area.center,
        lambda: area.anchor("top"),
        lambda: area.contain(10, 10),
        lambda: area.inset(top=4),
    ):
        with pytest.raises(InputError, match="element.area"):
            operation()
    assert area.inset(left=8, right=8) == Rect(56, 100, 848, None)
    width_unknown = Rect(48, 100, None, 40)
    with pytest.raises(InputError, match="幅を内容から決める領域（w=None）では right"):
        width_unknown.right
    with pytest.raises(InputError, match="columns"):
        width_unknown.columns(2)
    assert width_unknown.rows(2, gap=0)[1] == Rect(48, 120, None, 20)
    for values in ((0, 0, -1, 1), (0, 0, float("nan"), 1), (0, None, 1, 1)):
        with pytest.raises(InputError) as caught:
            Rect(*values)
        assert caught.value.code == "invalid_geometry"


# ---- 使う場所を限った役割の宣言 ----


def test_restricted_role_requires_declaration(make_deck):
    deck = make_deck(SMALL + '[roles.listed]\ninherits = "note"\nallowed_slide_ids = ["listed"]\n')
    with pytest.raises(InputError) as caught:
        deck.slide(id="plain").text(id="note", text="注", area=AREA, role="small")
    error = caught.value
    assert (error.code, error.slide_id, error.element_id) == ("role_usage", "plain", "note")
    assert (
        'deck.slide(..., allow_roles=["small"])' in str(error)
        and error.details["setting"] == "theme.toml の roles.small"
    )
    # ページ、要素、親のまとまり、表の宣言のどれかがあれば使うことができる。
    deck.slide(id="page", allow_roles=["small"]).text(id="note", text="注", area=AREA, role="small")
    deck.slide(id="element").text(id="note", text="注", area=AREA, role="small", allow_roles=("small",))
    deck.slide(id="group").group(id="notes", allow_roles=["small"]).text(id="note", text="注", area=AREA, role="small")
    deck.slide(id="table").table(
        id="facts",
        headers=["項目", "値"],
        rows=[["件数", "3"]],
        row_keys=["count"],
        column_keys=["item", "value"],
        area=Rect(40, 40, 400, 120),
        role="small",
        header_role="small",
        allow_roles=["small"],
    )
    # allowed_slide_ids に挙げたページでは、宣言なしで使うことができる。ほかのページでは使うことができない。
    deck.slide(id="listed").text(id="note", text="注", area=AREA, role="listed")
    with pytest.raises(InputError, match="使えません"):
        deck.slide(id="other").text(id="note", text="注", area=AREA, role="listed")


def test_allow_roles_are_recorded_and_inherited(make_deck):
    deck = make_deck(SMALL)
    slide = deck.slide(id="facts", allow_roles=["small", "small"])
    slide.text(id="plain", text="本文", area=AREA)
    slide.table(
        id="table",
        headers=["項目", "値"],
        rows=[["件数", "3"]],
        row_keys=["count"],
        column_keys=["item", "value"],
        area=Rect(40, 160, 400, 120),
        allow_roles=["small"],
        role="small",
    )
    deck.slide(id="other").text(id="plain", text="本文", area=AREA)
    first, other = deck.to_dict()["slides"]
    elements = by_id(first)
    assert first["allow_roles"] == ["small"]
    assert elements["table"]["allow_roles"] == ["small"] and elements["table/count/value"]["role"] == "small"
    assert "allow_roles" not in elements["plain"] and "allow_roles" not in elements["table/count/value"]
    assert all("allow_roles" not in element["style"] for element in first["elements"])
    assert "allow_roles" not in other and "allow_roles" not in by_id(other)["plain"]


def test_allow_roles_reject_unknown_roles_and_strings(make_deck):
    deck = make_deck(SMALL)
    with pytest.raises(InputError, match=re.escape('["small"]')) as caught:
        deck.slide(id="string", allow_roles="small")
    assert caught.value.code == "input_error"
    with pytest.raises(InputError) as caught:
        deck.slide(id="unknown", allow_roles=["smal"])
    assert caught.value.code == "unknown_role" and "（候補: small）" in str(caught.value)
    slide = deck.slide(id="elements")
    with pytest.raises(InputError, match="list か tuple"):
        slide.text(id="note", text="注", area=AREA, allow_roles="small")
    with pytest.raises(InputError) as caught:
        slide.text(id="note", text="注", area=AREA, allow_roles=["missing"])
    assert (caught.value.code, caught.value.element_id) == ("unknown_role", "note")
    assert [slide.id for slide in deck.slides] == ["elements"] and not slide.elements


def test_theme_title_role_is_declared_but_page_override_is_not(make_deck):
    deck = make_deck("[roles.slide_title]\nrestricted = true\n" + SMALL)
    slide = deck.slide(id="intro", title="題名")
    title = by_id(slide.to_dict(1))["_theme/title"]
    assert (title["role"], title["allow_roles"]) == ("slide_title", ["slide_title"])
    with pytest.raises(InputError) as caught:
        deck.slide(id="override", title="題名", title_role="small")
    assert (caught.value.code, caught.value.element_id) == ("role_usage", "_theme/title")
    declared = by_id(deck.slide(id="declared", title="題名", title_role="small", allow_roles=["small"]).to_dict(2))
    assert declared["_theme/title"]["role"] == "small" and "allow_roles" not in declared["_theme/title"]
    # 題名の役割を本文で使うには、宣言が要る。
    with pytest.raises(InputError, match="使えません"):
        slide.text(id="heading", text="見出し", area=AREA, role="slide_title")


def test_text_only_role_rejects_math(make_deck):
    slide = make_deck('[roles.prose]\ninherits = "body"\nallowed_content = "text"\n').slide(id="prose")
    slide.text(id="words", text="文章だけ", area=AREA, role="prose")
    with pytest.raises(InputError, match="役割 prose に数式は使えません") as caught:
        slide.text(id="mixed", text=[Text("値 "), Math("x")], area=AREA, role="prose")
    assert caught.value.code == "role_usage"
    with pytest.raises(InputError, match="数式は使えません"):
        slide.equation(id="formula", expression="x", area=AREA, role="prose")


# ---- 警告の許容の宣言 ----


def test_slide_allow_warning_validates_arguments(make_deck):
    deck = make_deck()
    slide = deck.slide(id="figure", title="図", notes="図を説明します。")
    frame = slide.box(id="frame", area=Rect(40, 80, 400, 200), text="")
    label = slide.text(id="label", text="ラベル", area=Rect(60, 100, 200, 40))
    slide.allow_warning("element_overlap", element=label, related=frame, reason="ラベルを枠に重ねて示す")
    slide.allow_warning("math_not_marked", element="label", reason="記号ではない")
    slide.allow_warning("notes_style", reason="引用した文の常体を残す")
    slide.allow_warning("content_area_overflow", element="_theme/title", reason="題名の位置はテーマで決める")
    failures = [
        ("element_overla", {}, "許容できる警告の種類ではない: 'element_overla'（候補: element_overlap）"),
        ("text_overflow", {}, "エラーは許容できない: text_overflow"),
        ("allow_warning_unused", {}, "許容できる警告の種類ではない"),
        # 組の警告はページ全体には宣言できない。
        ("element_overlap", {}, "element に対象の要素を指定してください"),
        ("math_not_marked", {"element": "label", "related": "frame"}, "related は組の警告"),
        ("element_overlap", {"element": "label", "related": "label"}, "同じ要素"),
        ("element_overlap", {"element": frame, "related": label}, "すでに宣言されています"),
        ("notes_style", {}, "すでに宣言されています"),
        ("math_not_marked", {"element": 3}, "要素か要素の識別子"),
    ]
    for code, arguments, words in failures:
        with pytest.raises(InputError, match=re.escape(words)) as caught:
            slide.allow_warning(code, reason="理由", **arguments)
        assert (caught.value.code, caught.value.slide_id) == ("input_error", "figure")
    with pytest.raises(InputError, match="reason"):
        slide.allow_warning("line_break", reason=" ")
    with pytest.raises(InputError, match="別のページ"):
        deck.slide(id="other").allow_warning("math_not_marked", element=label, reason="理由")
    page = deck.to_dict()["slides"][0]
    declarations = page["allowed_warnings"]
    assert [
        (item["code"], item["element_id"], item["related_element_id"], item["reason"]) for item in declarations
    ] == [
        ("element_overlap", "label", "frame", "ラベルを枠に重ねて示す"),
        ("math_not_marked", "label", None, "記号ではない"),
        ("notes_style", None, None, "引用した文の常体を残す"),
        ("content_area_overflow", "_theme/title", None, "題名の位置はテーマで決める"),
    ]
    assert all(re.fullmatch(r".+test_model_core\.py:\d+", item["location"]) for item in declarations)
    assert "allowed_warnings" not in deck.to_dict()["slides"][1]
    # 要素の存在は場面を作るときに確かめる（宣言の後に作る要素も指すことができる）。
    slide.allow_warning("element_overlap", element="later", reason="後で作る")
    with pytest.raises(InputError, match="later がページにありません"):
        deck.to_dict()
    slide.text(id="later", text="後", area=Rect(60, 300, 200, 40))
    assert by_id(deck.to_dict()["slides"][0])["later"]["id"] == "later"


def test_allow_warning_location_is_relative_to_the_deck(make_deck):
    deck = make_deck()
    helper = deck.project.root / "helpers.py"
    helper.write_text(
        'def declare(slide):\n    slide.allow_warning("notes_style", reason="引用を残す")\n', encoding="utf-8"
    )
    spec = importlib.util.spec_from_file_location("slidekit_test_location_helper", helper)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    slide = deck.slide(id="quote", notes="引用である。")
    module.declare(slide)
    assert slide.to_dict(1)["allowed_warnings"][0]["location"] == "helpers.py:2"


# ---- 構造の誤り ----


def test_structural_errors_have_diagnostic_codes(make_deck):
    deck = make_deck()
    slide = deck.slide(id="shapes")
    other = deck.slide(id="other")
    first = slide.box(id="first", text="最初", area=AREA)
    holder = other.group(id="holder", area=Rect(10, 10, 400, 300))
    cases = [
        (lambda: deck.slide(id="shapes"), "invalid_slide_id", None),
        (lambda: deck.slide(id="a/b"), "invalid_slide_id", None),
        (lambda: slide.text(id="first", text="重複", area=AREA), "invalid_element_id", "first"),
        (lambda: slide.text(id="", text="空", area=AREA), "invalid_element_id", None),
        (lambda: slide.text(id="flat", text="平ら", area=Rect(40, 40, 200, 0)), "invalid_geometry", "flat"),
        (lambda: slide.text(id="short", text="短い", area=(40, 40, 200)), "invalid_geometry", "short"),
        (lambda: slide.line(id="dot", start=(10, 10), end=(10, 10)), "invalid_geometry", "dot"),
        (lambda: slide.line(id="bad", start=(10, "x"), end=(20, 20)), "invalid_geometry", "bad"),
        (lambda: slide.text(id="role", text="役割", area=AREA, role="missing"), "unknown_role", "role"),
        (lambda: slide.text(id="minimum", text="最小", area=AREA, min_size_pt=10), "minimum_override", "minimum"),
        (lambda: other.arrow(id="link", start=first.anchor("right"), end=(500, 100)), "connection_missing", "link"),
        (lambda: slide.image(id="picture", path="figure.png", area=AREA, fit="stretch"), "image_fit", "picture"),
        (lambda: slide.text(id="turned", text=[Math("x")], area=AREA, rotation=90), "math_transform", "turned"),
        (lambda: slide.text(id="child", text="子", area=AREA, _parent_id="missing"), "parent_missing", "child"),
        (lambda: slide._add("circle", id="circle", area=AREA), "unknown_element", None),
        # まとまりの代理で要素を作るときと、作成後に領域を変更・移動するときの誤りにも、要素の識別子を付ける。
        (lambda: holder.text(id="child", text="子", area=(1, 2, 3)), "invalid_geometry", "holder/child"),
        (lambda: holder.box(id="wide", text="広い", area=(0, 0, float("inf"), 40)), "invalid_geometry", "holder/wide"),
        (lambda: setattr(first, "area", (1, 2, 3)), "invalid_geometry", "first"),
        (lambda: first.move(dx="右"), "invalid_geometry", "first"),
        (lambda: holder.move(dy=float("nan")), "invalid_geometry", "holder"),
    ]
    for operation, code, element_id in cases:
        with pytest.raises(InputError) as caught:
            operation()
        error = caught.value
        assert (error.code, error.element_id) == (code, element_id), code
        if element_id is not None:
            assert str(error).startswith(f"ページ {error.slide_id} の要素 {element_id}: "), str(error)
    assert [element.id for element in slide.elements] == ["first"]
    assert [element.id for element in other.elements] == ["holder"] and holder.area == Rect(10, 10, 400, 300)
    first.area = Rect(40, 40, 200, 60)
    slide.elements[0].data["area"] = [40, 40, 0, 60]
    with pytest.raises(InputError) as caught:
        deck.to_dict()
    assert caught.value.code == "invalid_geometry"
    deck.project.config["size"] = [0, 540]
    with pytest.raises(InputError) as caught:
        deck.slide(id="size")
    assert caught.value.code == "invalid_page_size"


# ---- 吹き出し ----

CALLOUT = (
    '[roles.callout]\ninherits = "note"\nfill = "none"\nline_color = "text"\nline_width_pt = 1\n'
    'padding_pt = 6\ntail_side = "left"\ntail_at = 0.2\n'
)
TAIL_NAMES = {"tail_side", "tail_at", "tail_width_pt", "tail_length_pt"}


def setting(issue):
    """続行できる誤りの診断の観測値の setting（なければ None）。同じ要素と識別子の誤りは 1 件にまとまる。"""
    [observations] = issue["observations"] or [{}]
    return observations.get("setting")


def issues_of(scene):
    return {
        (issue["element_id"], issue["code"], (observations or {}).get("setting"))
        for issue in scene.get("issues", [])
        for observations in issue["observations"] or [None]
    }


def callout_deck(make_deck, theme=None, *, collect=False):
    deck = make_deck(theme, collect=collect)
    deck.project.fonts = FakeFonts()
    return deck


def test_callout_scene_has_tail_actual_area_and_outside_parts(make_deck):
    deck = callout_deck(make_deck)
    slide = deck.slide(id="page", kind="blank")
    slide.box(
        id="note", area=(420, 150, 240, 70), text="補足", fill="none", line_color="text", tail_side="left", tail_at=0.2
    )
    slide.box(
        id="hint",
        area=(100, 80, 200, 50),
        shape="rounded_rectangle",
        text="角丸",
        corner_radius_pt=6,
        tail_side="bottom",
    )
    slide.box(id="plain", area=(100, 300, 200, 50), text="しっぽなし")
    found = by_id(deck.to_dict()["slides"][0])
    note = found["note"]
    assert note["actual_area"] == [412, 150, 248, 70]
    assert note["outside_parts"] == [{"part": "tail", "area": [412, 159, 8, 10]}]
    assert note["tail"] == {
        "side": "left",
        "at": 0.2,
        "width_pt": 10,
        "length_pt": 8,
        "points": [[420, 159], [412, 164], [420, 169]],
    }
    assert note["style"]["tail_side"] == "left" and note["style"]["tail_at"] == 0.2
    assert found["hint"]["tail"]["points"] == [[195, 130], [200, 138], [205, 130]]
    assert found["hint"]["actual_area"] == [100, 80, 200, 58]
    # しっぽのない要素の場面は今までと同じ（項目を加えない）。
    assert not {"tail", "actual_area", "outside_parts"} & set(found["plain"])
    assert not TAIL_NAMES & set(found["plain"]["style"])


def test_callout_role_layers_and_cancellation(make_deck):
    deck = callout_deck(make_deck, CALLOUT)
    slide = deck.slide(id="page", kind="blank")
    slide.box(id="role", area=(420, 100, 240, 60), role="callout", text="役割")
    slide.box(id="moved", area=(420, 200, 240, 60), role="callout", text="上書き", tail_side="right", tail_at=0.8)
    slide.box(id="kept", area=(420, 300, 240, 60), role="callout", text="None", tail_side=None)
    slide.box(id="cancelled", area=(420, 400, 240, 60), role="callout", text="取り消し", tail_side="none")
    slide.box(
        id="styled",
        area=(100, 400, 240, 60),
        role="callout",
        text="style",
        style={"tail_side": "top"},
        tail_side="bottom",
    )
    found = by_id(deck.to_dict()["slides"][0])
    assert (found["role"]["tail"]["side"], found["role"]["tail"]["at"]) == ("left", 0.2)
    assert (found["moved"]["tail"]["side"], found["moved"]["tail"]["at"]) == ("right", 0.8)
    assert found["kept"]["tail"]["side"] == "left"
    # none は役割のしっぽを取り消し、書式からしっぽの項目を除く。
    assert "tail" not in found["cancelled"] and not TAIL_NAMES & set(found["cancelled"]["style"])
    # キーワード引数が style= より優先する。
    assert found["styled"]["tail"]["side"] == "bottom"


def test_tail_items_do_not_reach_other_elements(make_deck):
    theme = CALLOUT + "[defaults]\ntail_length_pt = 12\n[defaults.box]\ntail_width_pt = 14\n"
    deck = callout_deck(make_deck, theme)
    slide = deck.slide(id="page", kind="blank")
    slide.text(id="text", area=(40, 40, 200, 40), role="callout", text="文字")
    slide.box(id="oval", area=(40, 100, 200, 40), shape="ellipse", role="callout", text="楕円")
    slide.box(id="plain", area=(40, 160, 200, 40), text="図形")
    slide.line(id="line", start=(40, 220), end=(240, 220))
    slide.box(id="callout", area=(300, 160, 200, 40), text="吹き出し", tail_side="bottom")
    slide.table(id="table", rows=[["a"]], row_keys=["r"], column_keys=["c"], area=Rect(300, 300, 200, None))
    scene = deck.to_dict()
    found = by_id(scene["slides"][0])
    # 役割のしっぽが text・ellipse の box に当たってもエラーにせず、しっぽも付けない。
    assert "issues" not in scene
    for element_id, element in found.items():
        if element_id != "callout":
            assert not TAIL_NAMES & set(element.get("style") or {}), element_id
            assert "tail" not in element, element_id
    # [defaults]・[defaults.box] の寸法は、しっぽの付く box ではライブラリの既定より優先する。
    assert (found["callout"]["tail"]["width_pt"], found["callout"]["tail"]["length_pt"]) == (14, 12)
    assert found["callout"]["style"]["tail_width_pt"] == 14


def test_callout_auto_size_matches_box_without_tail(make_deck):
    deck = callout_deck(make_deck)
    slide = deck.slide(id="page", kind="blank")
    plain = slide.box(id="plain", area=(40, 40, None, None), text="$O(\\log M)$", line_color="text")
    callout = slide.box(
        id="callout",
        area=(40, 140, None, None),
        text="$O(\\log M)$",
        line_color="text",
        tail_side="right",
        tail_length_pt=14,
        tail_width_pt=12,
    )
    assert (callout.area.w, callout.area.h) == (plain.area.w, plain.area.h)
    rounded = slide.box(
        id="rounded", area=(40, 240, 160, None), shape="rounded_rectangle", text="角丸の文", tail_side="top"
    )
    other = slide.box(id="other", area=(40, 340, 160, None), shape="rounded_rectangle", text="角丸の文")
    assert rounded.area.h == other.area.h
    tail = by_id(deck.to_dict()["slides"][0])["callout"]["tail"]
    assert tail["points"][1][0] == pytest.approx(callout.area.x + callout.area.w + 14)


def test_callout_shape_follows_the_last_area(make_deck):
    deck = callout_deck(make_deck, collect=True)
    slide = deck.slide(id="page", kind="blank")
    box = slide.box(id="note", area=(420, 150, 240, 70), text="補足", tail_side="left", tail_at=0.2)
    box.move(dx=-100, dy=10)
    narrow = slide.box(id="narrow", area=(40, 300, 240, 40), text="狭くなる", tail_side="left")
    narrow.area = Rect(40, 300, 240, 8)
    scene = deck.to_dict()
    found = by_id(scene["slides"][0])
    assert found["note"]["tail"]["points"] == [[320, 169], [312, 174], [320, 179]]
    # 場面を作るときに確かめ直し、付け根が辺に収まらなければ報告してしっぽを付けない。
    assert "tail" not in found["narrow"] and not TAIL_NAMES & set(found["narrow"]["style"])
    [issue] = [issue for issue in scene["issues"] if issue["element_id"] == "narrow"]
    assert issue["code"] == "invalid_style" and setting(issue) == "tail_width_pt"
    assert issue["message"] == (
        "ページ page の要素 narrow: しっぽの付け根の幅 tail_width_pt（10 pt）が左の辺の直線部分に収まりません"
        "（辺の長さ 8 pt、角丸の半径 0 pt、指定できる最大の幅 8 pt）"
    )
    # 場面を作り直しても、同じ誤りを重ねて集めない。
    assert len([issue for issue in deck.to_dict()["issues"] if issue["element_id"] == "narrow"]) == 1


def test_callout_errors_continue_with_codes_and_settings(make_deck):
    deck = callout_deck(make_deck, collect=True)
    slide = deck.slide(id="page", kind="blank")
    slide.box(id="side", area=(40, 40, 200, 40), text="辺", tail_side="center")
    slide.box(id="at", area=(40, 100, 200, 40), text="位置", tail_side="left", tail_at=1.5, tail_width_pt=0)
    slide.box(
        id="oval", area=(40, 160, 200, 40), shape="ellipse", text="楕円", style={"tail_side": "left"}, tail_at=0.3
    )
    slide.box(id="missing", area=(40, 220, 200, 40), text="辺なし", tail_at=0.3, tail_length_pt=4)
    slide.box(
        id="wide", area=(40, 280, 200, 20), shape="rounded_rectangle", text="幅", corner_radius_pt=6, tail_side="left"
    )
    scene = deck.to_dict()
    assert issues_of(scene) == {
        ("side", "invalid_style", "tail_side"),
        ("at", "invalid_style", "tail_at"),
        ("at", "invalid_style", "tail_width_pt"),
        ("oval", "invalid_style", "tail_at"),
        ("oval", "invalid_style", "tail_side"),
        ("missing", "invalid_style", "tail_at"),
        ("missing", "invalid_style", "tail_length_pt"),
        ("wide", "invalid_style", "tail_width_pt"),
    }
    messages = {error.element_id + error.details["setting"]: str(error) for error in deck.issues}
    assert messages["missingtail_at"] == (
        "ページ page の要素 missing: tail_at はしっぽを付けるときだけ使えます。"
        "tail_side に left・right・top・bottom のどれかを指定してください"
    )
    assert messages["ovaltail_side"].endswith(
        "tail_side は shape が rectangle か rounded_rectangle の box だけに指定できます（shape: ellipse）"
    )
    assert "辺の長さ 20 pt、角丸の半径 6 pt、指定できる最大の幅 8 pt" in messages["widetail_width_pt"]
    found = by_id(scene["slides"][0])
    # 誤った値を除いた結果、at は既定の位置・幅でしっぽを付ける。ほかはしっぽを付けない。
    assert found["at"]["tail"]["at"] == 0.5 and found["at"]["tail"]["width_pt"] == 10
    for element_id in ("side", "oval", "missing", "wide"):
        assert "tail" not in found[element_id] and not TAIL_NAMES & set(found[element_id]["style"]), element_id


def test_callout_rejects_rotation_and_flip(make_deck):
    deck = callout_deck(make_deck, CALLOUT)
    slide = deck.slide(id="page", kind="blank")
    for arguments in ({"rotation": 90}, {"flip_h": True, "flip_v": True}):
        with pytest.raises(InputError) as caught:
            slide.box(id="turned", area=(40, 40, 200, 40), text="回転", role="callout", **arguments)
        assert caught.value.code == "unsupported_argument"
        assert caught.value.details["names"] == list(arguments)
        assert str(caught.value).startswith("ページ page の要素 turned: しっぽのある box は回転・反転できません")
    # しっぽを取り消した box と、しっぽのない box は回転できる。
    slide.box(id="free", area=(40, 40, 200, 40), text="回転", role="callout", tail_side="none", rotation=90)
    slide.box(id="plain", area=(40, 100, 200, 40), text="回転", rotation=90)


def test_unknown_callout_shape_names_point_to_tail_side(make_deck):
    deck = callout_deck(make_deck)
    slide = deck.slide(id="page", kind="blank")
    for operation in (
        lambda: slide.box(id="bubble", area=(40, 40, 200, 40), shape="callout"),
        lambda: slide.measure("吹き出し", kind="box", shape="wedgeRectCallout"),
    ):
        with pytest.raises(InputError, match="吹き出しは tail_side で指定します"):
            operation()
    with pytest.raises(InputError) as caught:
        slide.box(id="star", area=(40, 40, 200, 40), shape="star")
    assert "tail_side" not in str(caught.value)
    with pytest.raises(InputError, match="tail_side は text では使えません"):
        slide.text(id="text", area=(40, 40, 200, 40), text="文字", tail_side="left")


def test_measure_reports_tail_items_on_ellipse(make_deck):
    deck = callout_deck(make_deck, collect=True)
    slide = deck.slide(id="page", kind="blank")
    plain = slide.measure("長さ", kind="box", width=120)
    assert slide.measure("長さ", kind="box", width=120, tail_side="left", tail_length_pt=20).height == plain.height
    slide.measure("長さ", kind="box", shape="ellipse", width=120, tail_side="left")
    assert [(error.code, error.details["setting"]) for error in deck.issues] == [("invalid_style", "tail_side")]


def test_tail_toward_chooses_side_and_records_target(make_deck):
    deck = callout_deck(make_deck, CALLOUT)
    slide = deck.slide(id="page", kind="blank")
    body = slide.text(id="body", area=(40, 120, 340, 200), text="本文")
    slide.box(
        id="auto", area=(410, 130, 220, 60), text="補足", line_color="text", tail_toward=body.anchor("right", at=0.2)
    )
    slide.box(id="role", area=(410, 260, 220, 60), role="callout", text="役割", tail_toward=(300, 100))
    slide.box(id="fixed", area=(410, 400, 220, 60), text="固定", tail_side="left", tail_toward=(300, 100))
    slide.box(id="below", area=(100, 400, 200, 40), text="下", tail_toward=body.anchor("bottom", gap=4))
    found = by_id(deck.to_dict()["slides"][0])
    # 横の距離 30 pt、縦の距離 0 で左の辺。付け根は指す点を辺へ垂直に下ろした位置（160 - 130 = 30 / 60）。
    assert found["auto"]["tail"]["side"] == "left"
    assert found["auto"]["tail"]["at"] == pytest.approx(0.5)
    assert found["auto"]["tail"]["target"] == {"element_id": "body", "side": "right", "at": 0.2}
    # 役割の tail_side は tail_toward が上書きし（縦の距離が大きいので上の辺）、引数の tail_side は辺を固定する。
    assert found["role"]["tail"]["side"] == "top" and found["role"]["tail"]["target"] == {"point": [300, 100]}
    assert found["fixed"]["tail"]["side"] == "left"
    assert found["below"]["tail"]["side"] == "top"
    assert found["below"]["tail"]["target"] == {"element_id": "body", "side": "bottom", "gap": 4}
    # 長さは tail_length_pt のまま（先端は指す点まで伸ばさない）。
    assert found["auto"]["tail"]["points"][1] == [402, 160]
    # 点が付け根の正面にあれば、先端はずらさない（場面に tip_shift_pt を入れない）。
    assert "tip_shift_pt" not in found["auto"]["tail"]
    # 斜めの点では、先端を付け根の中心から点へ向かう直線の上へずらす。
    (x0, y0), (tx, ty), (x2, y2) = found["role"]["tail"]["points"]
    bx, by = (x0 + x2) / 2, (y0 + y2) / 2
    assert found["role"]["tail"]["tip_shift_pt"] < 0
    assert (ty - by) * (300 - bx) == pytest.approx((100 - by) * (tx - bx))
    # ずらしは ±tail_length_pt（45°）まで。
    assert found["fixed"]["tail"]["tip_shift_pt"] == pytest.approx(-8)
    # 書き出しは場面の tail から同じ輪郭を求める。
    from slidekit.writer import callout_geometry

    for key in ("role", "fixed"):
        outline = callout_geometry(found[key], found[key]["style"])
        assert [value for point in outline.tail for value in point] == pytest.approx(
            [value for point in found[key]["tail"]["points"] for value in point]
        )
        assert list(outline.actual_area) == pytest.approx(found[key]["actual_area"])


def test_tail_toward_errors(make_deck):
    deck = callout_deck(make_deck, collect=True)
    slide = deck.slide(id="page", kind="blank")
    other = deck.slide(id="other", kind="blank")
    target = other.text(id="target", area=(40, 40, 100, 40), text="別")
    cases = [
        (lambda: slide.box(id="kind", area=(40, 40, 200, 40), tail_toward="右"), "input_error"),
        (lambda: slide.box(id="far", area=(40, 40, 200, 40), tail_toward=target.anchor("left")), "input_error"),
        (
            lambda: slide.box(id="oval", area=(40, 40, 200, 40), shape="ellipse", tail_toward=(0, 0)),
            "unsupported_argument",
        ),
    ]
    for operation, code in cases:
        with pytest.raises(InputError) as caught:
            operation()
        assert caught.value.code == code
    slide.box(id="inside", area=(40, 100, 200, 40), text="内側", tail_toward=(100, 120))
    slide.box(id="wrong", area=(40, 200, 200, 40), text="固定", tail_side="right", tail_toward=(0, 220))
    slide.box(id="both", area=(40, 300, 200, 40), text="両方", tail_at=0.3, tail_toward=(0, 320))
    slide.box(id="cancel", area=(40, 400, 200, 40), text="none", tail_side="none", tail_toward=(0, 420))
    scene = deck.to_dict()
    assert issues_of(scene) == {
        ("inside", "invalid_geometry", None),
        ("wrong", "invalid_geometry", None),
        ("both", "invalid_style", "tail_at"),
        ("cancel", "invalid_style", "tail_side"),
    }
    found = by_id(scene["slides"][0])
    assert "tail" not in found["inside"] and "tail" not in found["wrong"] and "tail" not in found["cancel"]
    assert found["both"]["tail"]["side"] == "left" and "tail_at" not in found["both"]["style"]
    message = next(issue["message"] for issue in scene["issues"] if issue["element_id"] == "wrong")
    assert message == (
        "ページ page の要素 wrong: tail_toward の点 (0, 220) が tail_side で指定した右の辺の外側にありません。"
        "しっぽを向ける点は本体の外に置いてください"
    )


def test_tail_toward_points_in_groups_move_with_the_group(make_deck):
    deck = callout_deck(make_deck)
    slide = deck.slide(id="page", kind="blank")
    holder = slide.group(id="holder", area=Rect(100, 100, 300, 200))
    holder.box(id="note", area=(100, 20, 120, 40), text="補足", tail_toward=(20, 40))
    holder.move(dx=50)
    found = by_id(deck.to_dict()["slides"][0])
    # まとまりの子の座標はまとまりの左上が原点で、まとまりを動かすと指す点も動く。
    assert found["holder/note"]["tail"]["target"] == {"point": [170, 140]}
    assert found["holder/note"]["tail"]["side"] == "left"
