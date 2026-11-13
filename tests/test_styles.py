"""書式名の一覧、候補と理由、書式を重ねる順、None と "none"、断片の実効書式を確かめる。

前半は styles の関数の単位で、後半は Deck を通して（作成 API の書式の解決順、None と "none"、引数名の検査）確かめる。
"""

import itertools
import re

from PIL import Image
import pytest

from slidekit import Deck, Rect, styles
from slidekit.config import load_project
from slidekit.errors import ArgumentNameError, InputError
from slidekit.styles import (
    ALIASES,
    ALL_STYLE_KEYS,
    ATTRIBUTE_KEYS,
    DEFAULT_CHAINS,
    DEFAULT_KINDS,
    GROUP_ATTRIBUTE_KEYS,
    LIBRARY_DEFAULTS,
    ROLE_KEYS,
    ROLE_ONLY_KEYS,
    STYLE_KEYS,
    UNSUPPORTED,
    attribute_keys,
    check_names,
    check_style_value,
    element_layer,
    fragment_style,
    library_layer,
    role_layer,
    theme_layers,
    unknown_names,
)


def raised(kind, names, **arguments):
    with pytest.raises(ArgumentNameError) as caught:
        check_names(kind, names, **arguments)
    return caught.value


def test_style_key_tables_follow_element_kinds():
    assert set(STYLE_KEYS) == {
        "text",
        "equation",
        "box",
        "table_cell",
        "table_header",
        "line",
        "arrow",
        "image",
        "group",
    }
    assert "word_wrap" in STYLE_KEYS["text"] and "word_wrap" in STYLE_KEYS["equation"]
    assert STYLE_KEYS["line"] == STYLE_KEYS["arrow"] == styles.STROKE_KEYS
    assert STYLE_KEYS["box"] - STYLE_KEYS["text"] == {"corner_radius_pt"}
    assert STYLE_KEYS["image"] == STYLE_KEYS["group"] == frozenset()
    assert DEFAULT_KINDS == {"text", "equation", "box", "table_cell", "table_header", "line", "arrow"}
    assert DEFAULT_CHAINS["table_header"] == ("box", "table_cell", "table_header")
    assert DEFAULT_CHAINS["arrow"] == ("line", "arrow")
    assert {"restricted", "inherits", "allowed_slide_ids", "allowed_content", "skip_checks"} == ROLE_ONLY_KEYS
    assert ROLE_KEYS == ALL_STYLE_KEYS | ROLE_ONLY_KEYS
    assert styles.PAINT_KEYS <= styles.FRAME_KEYS and styles.PAINT_KEYS <= styles.COLOR_KEYS
    assert ATTRIBUTE_KEYS == {"description", "semantic", "rotation", "flip_h", "flip_v", "allow_roles"}
    assert not ATTRIBUTE_KEYS & ALL_STYLE_KEYS
    # まとまりは回転・反転できず、semantic も持たない。
    assert attribute_keys("group") == GROUP_ATTRIBUTE_KEYS == {"description", "allow_roles"}
    assert attribute_keys("text") == attribute_keys("line") == ATTRIBUTE_KEYS
    assert styles.SEMANTIC_VALUES == {"background", "decoration"}


def test_aliases_and_unsupported_names_do_not_shadow_real_names():
    names = set(ALIASES) | set(UNSUPPORTED)
    assert not names & (ALL_STYLE_KEYS | ATTRIBUTE_KEYS | ROLE_ONLY_KEYS)
    assert not set(ALIASES) & set(UNSUPPORTED)
    assert set(ALIASES.values()) <= ALL_STYLE_KEYS


def test_suggest_returns_close_names_only_for_strings():
    assert styles.suggest("colr", ["color", "line_color", "bold"]) == ("color",)
    assert styles.suggest("zzz", ["color"]) == ()
    assert styles.suggest(["color"], ["color"]) == ()
    assert styles.format_candidates(("a", "b")) == "（候補: a、b）" and styles.format_candidates(()) == ""


def test_unknown_names_are_reported_together_with_candidates():
    error = raised(
        "text",
        ["size", "colour", "valing", "font_size", "align"],
        context="ページ intro の要素 premise",
        slide_id="intro",
        element_id="premise",
    )
    assert isinstance(error, TypeError) and isinstance(error, ValueError) and isinstance(error, InputError)
    assert str(error) == (
        "ページ intro の要素 premise: text に未知の書式があります: size（候補: size_pt）、colour（候補: color）、"
        "valing（候補: valign、align）、font_size（候補: size_pt）"
    )
    assert error.code == "unknown_argument"
    assert (error.slide_id, error.element_id) == ("intro", "premise")
    assert error.details["names"] == ["size", "colour", "valing", "font_size"]
    assert error.details["problems"][0] == {"name": "size", "suggestions": ["size_pt"], "reason": None}
    # 文字列でない名前（style= の辞書のキー）は未知の名前とし、文言では repr で示す。
    error = raised("text", [1, ("size",), "colour"])
    assert str(error) == "text に未知の書式があります: 1、('size',)、colour（候補: color）"
    assert (error.code, error.details["names"]) == ("unknown_argument", [1, ("size",), "colour"])


def test_aliases_are_suggested_only_when_the_kind_accepts_them():
    assert unknown_names("box", ["background"]) == [("background", ("fill",), None)]
    assert unknown_names("text", ["wrap"]) == [("wrap", ("word_wrap",), None)]
    # 線は fill を受け付けないので、別名の行き先ではなく近い名前を探す。
    [(name, found, reason)] = unknown_names("line", ["background"])
    assert name == "background" and "fill" not in found and reason is None
    assert unknown_names("line", ["colour"]) == [("colour", ("line_color",), None)]


def test_check_names_explains_unsupported_names():
    error = raised("text", ["z_index", "italic"], context="ページ plan の要素 first")
    assert error.code == "unsupported_argument"
    assert str(error) == (
        "ページ plan の要素 first: z_index " + UNSUPPORTED["z_index"] + "。italic " + UNSUPPORTED["italic"]
    )
    assert "描画順は作成順" in str(error) and "未知" not in str(error)
    assert unknown_names("box", ["underline", "link"]) == [
        ("underline", (), UNSUPPORTED["underline"]),
        ("link", (), UNSUPPORTED["link"]),
    ]


def test_names_of_other_kinds_and_attributes_are_explained():
    assert unknown_names("line", ["fill"]) == [("fill", (), "は line では使えません")]
    assert unknown_names("image", ["size_pt"]) == [("size_pt", (), "は image では使えません")]
    assert unknown_names("group", ["color"]) == [("color", (), "は group では使えません")]
    # style= の辞書に属性を書いたとき（回転・反転は要素の引数だけで受け付ける）。
    assert unknown_names("text", ["rotation", "flip_h"]) == [
        ("rotation", (), "は書式に書けません。要素の引数で指定してください"),
        ("flip_h", (), "は書式に書けません。要素の引数で指定してください"),
    ]
    error = raised("line", ["fill", "colour"])
    assert error.code == "unknown_argument"
    assert str(error) == "line に未知の書式があります: colour（候補: line_color）。fill は line では使えません"
    assert raised("group", ["color"]).code == "unsupported_argument"


def test_attributes_given_as_arguments_are_explained_per_kind():
    # 要素を作る関数のキーワード引数（arguments=True）では、その種類が受け付けない属性を「使えない」とする。
    assert unknown_names("group", ["semantic", "rotation"], extra=GROUP_ATTRIBUTE_KEYS, arguments=True) == [
        ("semantic", (), "は group では使えません"),
        ("rotation", (), "は group では使えません"),
    ]
    assert unknown_names("text", ["semantic"], extra=ATTRIBUTE_KEYS, arguments=True) == []
    error = raised("group", ["semantic"], extra=GROUP_ATTRIBUTE_KEYS, arguments=True)
    assert (str(error), error.code) == ("semantic は group では使えません", "unsupported_argument")


def test_extra_names_are_accepted_and_suggested():
    check_names("box", ["shape", "fill", "rotation"], extra={"shape", "rotation"})
    assert unknown_names("box", ["shap", "shape", "shape"], extra={"shape"}) == [("shap", ("shape",), None)]
    assert unknown_names(None, ["corner_radius_pt", "dashed", "word_wrap"]) == []
    assert unknown_names(None, ["inherits"], extra=ROLE_ONLY_KEYS) == []
    assert unknown_names(None, ["inherit"], extra=ROLE_ONLY_KEYS) == [("inherit", ("inherits",), None)]
    with pytest.raises(ValueError, match="未対応の要素の種類"):
        unknown_names("circle", ["fill"])


@pytest.mark.parametrize(
    ("name", "value"),
    [
        ("size_pt", 20),
        ("size_pt", 10.5),
        ("min_size_pt", 12),
        ("line_spacing", 1.2),
        ("line_spacing_pt", 24),
        ("line_width_pt", 0),
        ("corner_radius_pt", 6),
        ("padding_pt", 0),
        ("padding_pt", {"left": 12}),
        ("padding_pt", {"left": 1, "top": 2, "right": 3, "bottom": 4}),
        ("bold", False),
        ("word_wrap", True),
        ("dashed", True),
        ("align", "justify"),
        ("valign", "middle"),
        ("font_latin", "Arial"),
        ("color", "primary"),
        ("color", "#1A2B3C"),
        ("emphasis_color", "C43D3D"),
        ("fill", "none"),
        ("line_color", "surface"),
        ("size_pt", None),
    ],
)
def test_check_style_value_accepts_valid_values(name, value):
    assert check_style_value(name, value) is None


@pytest.mark.parametrize(
    ("name", "value", "message"),
    [
        ("size_pt", 0, "は正の有限の数で指定してください"),
        ("size_pt", True, "は正の有限の数で指定してください"),
        ("line_spacing", float("nan"), "は正の有限の数で指定してください"),
        ("line_width_pt", -1, "は 0 以上の有限の数で指定してください"),
        ("bold", "yes", "は真偽値で指定してください"),
        ("word_wrap", 1, "は真偽値で指定してください"),
        ("align", "middle", "は left、center、right、justify のどれかで指定してください"),
        ("valign", "center", "は top、middle、bottom のどれかで指定してください"),
        ("font_math", " ", "は空でない文字列で指定してください"),
        ("color", "#12345", "は色の名前か 6 桁の 16 進数で指定してください"),
        ("fill", 123, "は色の名前か 6 桁の 16 進数で指定してください"),
        ("padding_pt", [6, 6], "は 0 以上の数値か、left・top・right・bottom を持つ辞書で指定してください"),
        ("padding_pt", -2, "の値は 0 以上の有限値にしてください"),
        ("padding_pt", {"left": -1}, "の値は 0 以上の有限値にしてください"),
        ("padding_pt", {"lft": 3}, "に未知の辺があります: lft（候補: left）"),
    ],
)
def test_check_style_value_rejects_invalid_values(name, value, message):
    assert check_style_value(name, value) == message


def test_every_style_key_has_a_value_rule():
    assert all(check_style_value(name, object()) for name in ALL_STYLE_KEYS)
    assert check_style_value("unknown_name", object()) is None
    # 色の名前の有無と "none" の使い方は、テーマの色と合わせて確かめる（schema・作成 API）。
    assert check_style_value("color", "none") is None


def test_theme_layers_follow_default_chains():
    theme = {
        "defaults": {
            "color": "text",
            "padding_pt": {"left": 4},
            "box": {"fill": "surface"},
            "table_cell": {"fill": "background"},
            "line": {"line_width_pt": 2},
            "arrow": {"line_color": "accent"},
        }
    }
    first = {"color": "text", "padding_pt": {"left": 4}}
    assert theme_layers(theme, "table_header") == [first, {"fill": "surface"}, {"fill": "background"}, {}]
    assert theme_layers(theme, "arrow") == [first, {"line_width_pt": 2}, {"line_color": "accent"}]
    assert theme_layers(theme, "text") == [first, {}]
    assert theme_layers(theme, None) == theme_layers(theme, "image") == theme_layers(theme, "group") == [first]
    layers = theme_layers(theme, "box")
    layers[0]["padding_pt"]["left"] = 99
    layers[1]["fill"] = "changed"
    assert theme["defaults"]["padding_pt"] == {"left": 4} and theme["defaults"]["box"] == {"fill": "surface"}
    assert theme_layers({}, "box") == [{}, {}]
    with pytest.raises(ValueError):
        theme_layers(theme, "circle")


def test_library_and_role_layers():
    assert LIBRARY_DEFAULTS["word_wrap"] is True
    assert LIBRARY_DEFAULTS["fill"] is None and LIBRARY_DEFAULTS["line_color"] is None
    assert not {"size_pt", "min_size_pt", "font_latin"} & set(LIBRARY_DEFAULTS)
    with pytest.raises(TypeError):
        LIBRARY_DEFAULTS["bold"] = True
    layer = library_layer({"fonts": {"japanese": "和文", "latin": "Latin", "math": "Math"}})
    assert layer == {"font_japanese": "和文", "font_latin": "Latin", "font_math": "Math", **LIBRARY_DEFAULTS}
    role = {
        "size_pt": 18,
        "min_size_pt": 16,
        "color": "muted",
        "inherits": "body",
        "allowed_slide_ids": ["intro"],
        "allowed_content": "math",
        "restricted": True,
    }
    assert role_layer(role) == {"size_pt": 18, "min_size_pt": 16, "color": "muted"}


def test_element_layer_none_and_none_string():
    overrides = {"fill": None, "line_color": None, "size_pt": None, "color": None, "bold": False, "align": "center"}
    assert element_layer(overrides, "box") == {"fill": "none", "line_color": "none", "bold": False, "align": "center"}
    for kind in styles.FRAME_KINDS:
        assert element_layer({"fill": None}, kind) == {"fill": "none"}
    assert element_layer({"fill": "none"}, "text") == {"fill": "none"}
    # 線の line_color の None は指定しない（文字色へ戻る）。
    assert element_layer({"line_color": None, "dashed": True}, "line") == {"dashed": True}
    assert element_layer(None, "text") == {}
    # style= の辞書とキーワード引数は別々に通して重ねる。キーワード引数の None は辞書の値を消さない。
    merged = {
        **element_layer({"size_pt": 30, "fill": "accent"}, "box"),
        **element_layer({"size_pt": None, "fill": None}, "box"),
    }
    assert merged == {"size_pt": 30, "fill": "none"}


def test_fragment_style_resolves_fragment_over_element():
    element = {"size_pt": 20, "bold": False, "color": "20252B", "emphasis_color": "C43D3D"}
    assert fragment_style(element, {"kind": "text", "text": "a"}) == {
        "size_pt": 20,
        "bold": False,
        "color": "20252B",
        "underline": False,
        "superscript": False,
        "link": None,
    }
    assert fragment_style(element, {"kind": "emphasis", "text": "a"})["bold"] is True
    assert fragment_style(element, {"kind": "emphasis", "text": "a"})["color"] == "C43D3D"
    given = {
        "kind": "text",
        "text": "詳細",
        "style": {"size_pt": 14, "color": "183153", "underline": True, "bold": True, "superscript": True},
        "link": "https://example.com/",
    }
    assert fragment_style(element, given) == {
        "size_pt": 14,
        "bold": True,
        "color": "183153",
        "underline": True,
        "superscript": True,
        "link": "https://example.com/",
    }
    assert (
        fragment_style(element, {"kind": "emphasis", "text": "a", "style": {"bold": False, "color": "606B78"}})["bold"]
        is False
    )
    assert fragment_style({**element, "bold": True}, {"kind": "math", "text": "x"})["bold"] is True
    math = fragment_style(element, {"kind": "math", "text": "x"})
    assert (math["bold"], math["color"]) == (False, "20252B")
    emphasized = fragment_style(element, {"kind": "math", "text": "x", "emphasis": True})
    assert (emphasized["bold"], emphasized["color"]) == (True, "C43D3D")
    assert fragment_style(element, {"kind": "math", "text": "x", "style": {"size_pt": 28}})["size_pt"] == 28


# ---- Deck を通す書式の試験（作成 API） ----

AREA = Rect(40, 80, 300, 60)


@pytest.fixture
def make_deck(tmp_path, write_deck):
    """資料を書いて読み込み、Deck を返す関数。theme は資料の theme.toml。"""
    numbers = itertools.count()

    def make(theme=None):
        return Deck(load_project(write_deck(tmp_path / f"deck{next(numbers)}", theme=theme)))

    return make


def style_of(element):
    return element.data["style"]


def test_resolution_order_library_theme_kind_role_element(make_deck):
    slide = make_deck(
        "[defaults]\nline_spacing = 1.3\n[defaults.box]\nline_spacing = 1.4\n"
        '[roles.card]\ninherits = "body"\nline_spacing = 1.5\n'
    ).slide(id="order")

    def spacing(element):
        return style_of(element)["line_spacing"]

    # テーマの [defaults] < [defaults.box] < 役割 < style= < キーワード引数。
    assert spacing(slide.text(id="theme", text="a", area=AREA)) == 1.3
    assert spacing(slide.box(id="kind", text="a", area=AREA)) == 1.4
    assert spacing(slide.box(id="role", text="a", area=AREA, role="card")) == 1.5
    assert spacing(slide.box(id="style", text="a", area=AREA, role="card", style={"line_spacing": 1.6})) == 1.6
    assert (
        spacing(
            slide.box(id="keyword", text="a", area=AREA, role="card", style={"line_spacing": 1.6}, line_spacing=1.7)
        )
        == 1.7
    )
    # テーマにない値はライブラリの既定（styles.LIBRARY_DEFAULTS）になる。
    library = style_of(make_deck().slide(id="library").text(id="plain", text="a", area=AREA))
    assert (library["word_wrap"], library["emphasis_color"], library["bold"]) == (True, "C43D3D", False)


def test_box_defaults_come_from_bundled_theme(make_deck):
    slide = make_deck("[defaults]\npadding_pt = 3\n").slide(id="boxes")
    box = style_of(slide.box(id="box", area=AREA, text="図形"))
    text = style_of(slide.text(id="text", area=AREA, text="文字"))
    # 同梱テーマの [defaults.box] がテーマ共通の [defaults] より優先する（0.1 の box() と同じ値）。
    assert (box["fill"], box["line_color"], box["padding_pt"]) == ("F3F6FA", "BBC5D0", 6)
    assert (text["fill"], text["line_color"], text["padding_pt"]) == (None, None, 3)


def test_role_fill_and_padding_apply_to_box(make_deck):
    slide = make_deck('[roles.callout]\ninherits = "body"\nfill = "FFF4D6"\npadding_pt = 10\n').slide(id="callout")
    box = style_of(slide.box(id="callout", area=AREA, text="注意", role="callout"))
    assert (box["fill"], box["padding_pt"], box["line_color"]) == ("FFF4D6", 10, "BBC5D0")
    assert style_of(slide.box(id="given", area=AREA, text="指定", role="callout", fill="surface"))["fill"] == "F3F6FA"


def test_kind_chains_for_table_cells_and_arrows(make_deck):
    slide = make_deck(
        '[defaults.box]\nline_width_pt = 2\n[defaults.line]\nline_color = "accent"\n'
        "[defaults.arrow]\nline_width_pt = 3\n"
    ).slide(id="chains")
    slide.table(
        id="table",
        headers=["項目", "値"],
        rows=[["件数", "3"]],
        row_keys=["count"],
        column_keys=["item", "value"],
        area=Rect(40, 40, 400, 120),
    )
    line = style_of(slide.line(id="line", start=(40, 200), end=(400, 200)))
    arrow = style_of(slide.arrow(id="arrow", start=(40, 240), end=(400, 240)))
    cells = {element.id: style_of(element) for element in slide.elements if element.id.startswith("table/")}
    # セルは box の後に table_cell（見出しはさらに table_header）を重ねる。矢印は line の後に arrow を重ねる。
    assert cells["table/_header/item"]["fill"] == "F3F6FA" and cells["table/count/item"]["fill"] == "FFFFFF"
    assert all(
        (style["line_width_pt"], style["line_color"], style["padding_pt"]) == (2, "BBC5D0", 6)
        for style in cells.values()
    )
    assert (line["line_color"], line["line_width_pt"]) == ("C43D3D", 1)
    assert (arrow["line_color"], arrow["line_width_pt"]) == ("C43D3D", 3)


def test_theme_background_ignores_kind_defaults(make_deck):
    slide = make_deck(
        '[defaults.box]\nfill = "accent"\nline_color = "primary"\npadding_pt = 9\n'
        "[defaults.text]\nline_spacing = 1.4\n"
    ).slide(id="theme", title="題名")
    elements = {element["id"]: element["style"] for element in slide.to_dict(1)["elements"]}
    background = elements["_theme/background"]
    assert (background["fill"], background["line_color"], background["padding_pt"], background["line_spacing"]) == (
        "FFFFFF",
        None,
        0,
        1.15,
    )
    # 題名とページ番号は text の既定を使う。
    assert elements["_theme/title"]["line_spacing"] == elements["_theme/footer"]["line_spacing"] == 1.4
    assert elements["_theme/footer"]["align"] == "right"


def test_unknown_defaults_names_are_rejected(tmp_path, write_deck):
    for number, (theme, words) in enumerate(
        (
            ('[defaults.boxes]\nfill = "surface"\n', "defaults.boxes は不明な設定です（候補: box）"),
            ('[defaults.box]\ncolour = "text"\n', "defaults.box.colour は不明な設定です（候補: color）"),
        )
    ):
        with pytest.raises(ValueError, match=re.escape(words)):
            load_project(write_deck(tmp_path / f"deck{number}", theme=theme))


def test_none_means_unspecified_for_text_style(make_deck):
    slide = make_deck(
        '[roles.lead]\ninherits = "body"\nsize_pt = 24\ncolor = "primary"\nbold = true\nalign = "center"\n'
        'valign = "middle"\npadding_pt = 4\n'
    ).slide(id="none")
    plain = slide.text(id="plain", text="既定", area=AREA, role="lead")
    given = slide.text(
        id="given",
        text="未指定",
        area=AREA,
        role="lead",
        size_pt=None,
        color=None,
        bold=None,
        align=None,
        valign=None,
        padding_pt=None,
        line_spacing=None,
        word_wrap=None,
    )
    assert style_of(given) == style_of(plain)
    assert (style_of(given)["size_pt"], style_of(given)["color"], style_of(given)["padding_pt"]) == (24, "183153", 4)
    # キーワード引数の None は style= の辞書の値を消さない。
    assert style_of(slide.text(id="kept", text="辞書", area=AREA, style={"size_pt": 30}, size_pt=None))["size_pt"] == 30


def test_none_paint_means_transparent_and_matches_none_string(make_deck):
    slide = make_deck().slide(id="paint")
    omitted = style_of(slide.box(id="omitted", area=AREA, text="既定"))
    assert (omitted["fill"], omitted["line_color"]) == ("F3F6FA", "BBC5D0")
    for number, arguments in enumerate(
        (
            {"fill": None, "line_color": None},
            {"fill": "none", "line_color": "none"},
            {"style": {"fill": None, "line_color": "none"}},
        )
    ):
        box = style_of(slide.box(id=f"clear{number}", area=AREA, text="透明", **arguments))
        assert (box["fill"], box["line_color"]) == (None, None)
    assert style_of(slide.text(id="filled", area=AREA, text="文字", fill="surface"))["fill"] == "F3F6FA"
    assert style_of(slide.text(id="clear", area=AREA, text="文字", fill=None))["fill"] is None


def test_line_color_none_on_line_uses_default(make_deck):
    line = (
        make_deck('[defaults.line]\nline_color = "accent"\n')
        .slide(id="lines")
        .line(id="rule", start=(40, 100), end=(400, 100), line_color=None)
    )
    assert style_of(line)["line_color"] == "C43D3D"
    # 線の色を決める層がなければ None のままで、書き出しは文字色を使う（0.1 と同じ）。
    plain = make_deck().slide(id="plain").line(id="rule", start=(40, 100), end=(400, 100), line_color=None)
    assert (style_of(plain)["line_color"], style_of(plain)["color"]) == (None, "20252B")


def test_invalid_color_reports_page_element_and_names(make_deck):
    slide = make_deck().slide(id="plan")
    colors = "（テーマの色: accent、background、line、muted、primary、surface、text）"
    with pytest.raises(InputError) as caught:
        slide.box(id="first", area=AREA, text="最初", fill="tela")
    error = caught.value
    assert str(error) == "ページ plan の要素 first: fill に未定義の色を指定しています: 'tela'" + colors
    assert (error.code, error.slide_id, error.element_id) == ("undefined_color", "plan", "first")
    for name in ("color", "emphasis_color"):
        with pytest.raises(InputError) as caught:
            slide.text(id="second", area=AREA, text="二番目", **{name: "none"})
        assert (
            str(caught.value)
            == f"ページ plan の要素 second: {name} に none は指定できません。none は fill・line_color だけに使えます{colors}"
        )
        assert caught.value.code == "invalid_style"
    # 形の間違った値（6 桁でない 16 進数、文字列でない値）にも、テーマの色名の一覧を添える。
    for arguments, setting, shown in (
        ({"color": "#12345"}, "color", "'#12345'"),
        ({"fill": 3}, "fill", "3"),
        ({"style": {"line_color": "1a2b3"}}, "line_color", "'1a2b3'"),
    ):
        with pytest.raises(InputError) as caught:
            slide.box(id="third", area=AREA, text="三番目", **arguments)
        assert (
            str(caught.value)
            == f"ページ plan の要素 third: {setting} は色の名前か 6 桁の 16 進数で指定してください: {shown}{colors}"
        )
        assert (caught.value.code, caught.value.element_id, caught.value.details["setting"]) == (
            "invalid_style",
            "third",
            setting,
        )
    assert style_of(slide.text(id="hex", area=AREA, text="16 進数", color="#1a2b3c"))["color"] == "1A2B3C"
    assert [element.id for element in slide.elements] == ["hex"]


def test_optional_attributes_accept_none(make_deck):
    deck = make_deck()
    Image.new("RGB", (40, 20), "white").save(deck.project.root / "figure.png")
    slide = deck.slide(id="optional")
    text = slide.text(
        id="text",
        text="文字",
        area=AREA,
        rotation=None,
        flip_h=None,
        flip_v=None,
        description=None,
        semantic=None,
        allow_roles=None,
    )
    assert not {"rotation", "flip_h", "flip_v", "description", "semantic", "allow_roles"} & set(text.data)
    box = slide.box(id="box", text="図形", area=AREA, shape=None, role=None)
    assert (box.data["shape"], box.data["role"]) == ("rectangle", "body")
    image = slide.image(id="image", path="figure.png", area=AREA, fit=None, page=None)
    assert image.data["fit"] == "contain" and "page" not in image.data


def test_unknown_style_names_are_rejected_with_candidates(make_deck):
    slide = make_deck().slide(id="intro")
    with pytest.raises(TypeError) as caught:
        slide.text(id="premise", text="前提", area=AREA, size=20, colour="accent", valing="top", font_size=20)
    error = caught.value
    assert str(error) == (
        "ページ intro の要素 premise: text に未知の書式があります: size（候補: size_pt）、colour（候補: color）、"
        "valing（候補: valign、align）、font_size（候補: size_pt）"
    )
    assert isinstance(error, ArgumentNameError) and isinstance(error, ValueError)
    assert (error.code, error.slide_id, error.element_id) == ("unknown_argument", "intro", "premise")
    assert not slide.elements


def test_style_names_are_checked_per_kind(make_deck):
    deck = make_deck()
    Image.new("RGB", (40, 20), "white").save(deck.project.root / "figure.png")
    slide = deck.slide(id="kinds")
    for operation, words in (
        (lambda: slide.line(id="rule", start=(0, 10), end=(100, 10), fill="surface"), "fill は line では使えません"),
        (lambda: slide.image(id="figure", path="figure.png", area=AREA, size_pt=20), "size_pt は image では使えません"),
        (lambda: slide.group(id="group", color="text"), "color は group では使えません"),
        (
            lambda: slide.text(id="text", text="文字", area=AREA, corner_radius_pt=4),
            "corner_radius_pt は text では使えません",
        ),
    ):
        with pytest.raises(TypeError, match=words) as caught:
            operation()
        assert caught.value.code == "unsupported_argument"
    # 角丸の大きさは角丸長方形だけに指定できる。
    with pytest.raises(InputError, match="rounded_rectangle の box だけ") as caught:
        slide.box(id="square", area=AREA, text="四角", corner_radius_pt=4)
    assert caught.value.code == "invalid_style"
    rounded = slide.box(id="rounded", area=AREA, text="角丸", shape="rounded_rectangle", corner_radius_pt=4)
    assert style_of(rounded)["corner_radius_pt"] == 4
    assert [element.id for element in slide.elements] == ["rounded"]


def test_unsupported_names_explain_reason(make_deck):
    slide = make_deck().slide(id="plan")
    with pytest.raises(TypeError) as caught:
        slide.text(id="first", text="最初", area=AREA, z_index=2)
    assert (
        str(caught.value)
        == "ページ plan の要素 first: z_index には対応していません。描画順は作成順で、後に作った要素が前面になります"
    )
    assert caught.value.code == "unsupported_argument"
    with pytest.raises(TypeError) as caught:
        slide.box(id="second", text="二番目", area=AREA, italic=True, underline=True)
    assert "italic には対応していません" in str(caught.value) and "Text(..., underline=True)" in str(caught.value)


def test_group_proxy_and_style_dict_are_checked(make_deck):
    slide = make_deck().slide(id="proxy")
    group = slide.group(id="group", area=Rect(40, 40, 400, 200))
    with pytest.raises(TypeError, match="colour（候補: color）") as caught:
        group.box(id="child", text="子", area=Rect(0, 0, 100, 40), colour="accent")
    assert caught.value.element_id == "group/child"
    with pytest.raises(TypeError, match="colour（候補: color）"):
        slide.box(id="styled", text="辞書", area=AREA, style={"colour": "accent"})
    # style= の辞書の文字列でないキーも、未知の書式名として扱う。
    with pytest.raises(TypeError) as caught:
        group.text(id="keys", text="キー", area=Rect(0, 0, 100, 40), style={1: "accent", None: 2, "bold": True})
    assert str(caught.value) == "ページ proxy の要素 group/keys: text に未知の書式があります: 1、None"
    assert (caught.value.code, caught.value.details["names"]) == ("unknown_argument", [1, None])
    # 回転と反転は要素の引数だけで受け付け、style= には書くことができない（場面の検査の math_transform に届かない）。
    for name, value in (("rotation", 20), ("flip_h", True), ("flip_v", True)):
        with pytest.raises(TypeError, match=f"{name} は書式に書けません") as caught:
            group.text(id="label", text="ラベル", area=Rect(0, 50, 100, 40), style={name: value})
        assert caught.value.code == "unsupported_argument"
    with pytest.raises(InputError, match="書式の辞書") as caught:
        slide.text(id="list", text="一覧", area=AREA, style=["bold"])
    assert caught.value.code == "invalid_style"
    turned = group.text(id="label", text="回転", area=Rect(0, 50, 100, 40), rotation=20)
    assert turned.data["rotation"] == 20 and "rotation" not in style_of(turned)
    assert [element.id for element in slide.elements] == ["group", "group/label"]


def test_internal_parent_argument_is_not_public(make_deck):
    slide = make_deck().slide(id="parent")
    slide.group(id="group")
    with pytest.raises(TypeError, match="まとまりの子は group の戻り値から作成してください") as caught:
        slide.text(id="child", text="子", area=AREA, parent_id="group")
    assert caught.value.code == "unsupported_argument"
    with pytest.raises(TypeError, match="未知の書式"):
        slide.text(id="child", text="子", area=AREA, _id="group/child")


def test_semantic_values_are_restricted(make_deck):
    slide = make_deck().slide(id="semantic")
    assert slide.box(id="frame", area=AREA, text="", semantic="decoration").data["semantic"] == "decoration"
    assert slide.box(id="back", area=AREA, text="", semantic="background").data["semantic"] == "background"
    # 文字列でない値（list など）も、同じ識別子の誤りにする。
    for value in ("title", ["background"], {"decoration"}):
        with pytest.raises(InputError, match="semantic は background、decoration") as caught:
            slide.text(id="title", text="題名", area=AREA, semantic=value)
        assert (caught.value.code, caught.value.element_id) == ("input_error", "title")
    with pytest.raises(TypeError, match="semantic は group では使えません"):
        slide.group(id="group", semantic="decoration")
    assert [element.id for element in slide.elements] == ["frame", "back"]
