"""表の作成（tables）を確かめる：見出しと行の塗り、書式と役割の重ね方、寸法、戻り値の Table、入力の誤り。

自動の寸法は project.fonts を FakeFonts に差し替えて確かめる。行の高さの値は計測のモデルで変わるので、
同じ書式で measure.layout_text を呼んだ値と比べる（幅は FakeFonts の送り幅で決まるので値で確かめる）。
"""

import itertools
import re

import pytest

from fake_fonts import FakeFonts
from slidekit import Deck, Emphasis, Math, Rect, Text
from slidekit.config import load_project
from slidekit.errors import InputError
from slidekit.measure import layout_text
from slidekit.tables import Table

AREA = Rect(40, 100, 600, 240)
# 同梱テーマの色
SURFACE, BACKGROUND, PRIMARY, ACCENT, MUTED, TEXT = "F3F6FA", "FFFFFF", "183153", "C43D3D", "606B78", "20252B"
ROLES = """[roles.shaded]
inherits = "body"
fill = "FFF4D6"

[roles.label]
inherits = "body"
bold = true

[roles.formula]
inherits = "body"
size_pt = 22
"""


@pytest.fixture
def make_deck(tmp_path, write_deck):
    """資料を書いて読み込み、Deck を返す関数。theme は資料の theme.toml、collect は collect_errors。"""
    numbers = itertools.count()

    def make(theme=None, *, collect=False, fonts=None):
        project = load_project(write_deck(tmp_path / f"deck{next(numbers)}", theme=theme))
        project.collect_errors = collect
        if fonts is not None:
            project.fonts = fonts
        return Deck(project)

    return make


def by_id(slide):
    return {element["id"]: element for element in slide.to_dict(1)["elements"]}


def simple(slide, **arguments):
    """2 列・2 行と見出しの表。"""
    values = {
        "id": "table",
        "headers": ["項目", "値"],
        "rows": [["件数", "3"], ["上限", "10"]],
        "row_keys": ["a", "b"],
        "column_keys": ["item", "value"],
        "area": AREA,
    }
    values.update(arguments)
    return slide.table(**values)


def cell_style(table, row, column):
    return table.cell(row, column).data["style"]


# ---- 書式と役割 ----


def test_header_and_body_fills_come_from_theme_and_role(make_deck):
    slide = make_deck(ROLES).slide(id="fills")
    table = simple(slide)
    # 同梱テーマの既定：見出しは surface、本文は background。
    assert cell_style(table, "_header", "item")["fill"] == SURFACE
    assert cell_style(table, "a", "item")["fill"] == BACKGROUND
    # 役割の fill は種類の既定より優先し、header_style は役割より優先する。
    shaded = simple(slide, id="shaded", role="shaded", header_role="shaded", header_style={"fill": "primary"})
    assert cell_style(shaded, "a", "value")["fill"] == "FFF4D6"
    assert cell_style(shaded, "_header", "value")["fill"] == PRIMARY
    assert cell_style(shaded, "a", "value")["line_color"] == "BBC5D0"  # 枠線は [defaults.box] のまま


def test_band_fills_repeat_over_body_rows(make_deck):
    slide = make_deck().slide(id="bands")
    table = slide.table(
        id="table",
        headers=["番号"],
        rows=[["1"], ["2"], ["3"], ["4"]],
        row_keys=["r1", "r2", "r3", "r4"],
        column_keys=["number"],
        area=AREA,
        band_fills=["accent", None, "muted"],
    )
    fills = [cell_style(table, key, "number")["fill"] for key in ("r1", "r2", "r3", "r4")]
    # None は塗らない。3 色を繰り返し、見出し行には適用されない。
    assert fills == [ACCENT, None, MUTED, ACCENT]
    assert cell_style(table, "_header", "number")["fill"] == SURFACE
    # 行の書式は交互の塗りより優先する。
    table = slide.table(
        id="rows",
        rows=[["1"], ["2"]],
        row_keys=["r1", "r2"],
        column_keys=["number"],
        area=AREA,
        band_fills=("surface",),
        row_styles={"r2": {"fill": "accent"}},
    )
    assert [cell_style(table, key, "number")["fill"] for key in ("r1", "r2")] == [SURFACE, ACCENT]


def test_column_row_and_cell_styles_override_in_order(make_deck):
    slide = make_deck().slide(id="order")
    table = simple(
        slide,
        style={"color": "muted"},
        valign="middle",
        padding_pt=10,
        column_styles={"value": {"color": "primary", "align": "right"}, "item": {"padding_pt": {"left": 20}}},
        header_style={"align": "center"},
        row_styles={"b": {"color": "accent", "bold": True}},
        cell_styles={
            "b/value": {"color": "FF0000"},
            "_header/item": {"align": "left"},
            "b/item": {"padding_pt": {"top": 2}},
        },
    )
    styles = {
        f"{row}/{column}": cell_style(table, row, column)
        for row in ("_header", "a", "b")
        for column in ("item", "value")
    }
    # 本文：セル > 行 > 列 > 表全体。
    assert (styles["a/item"]["color"], styles["a/item"]["align"]) == (MUTED, "left")
    assert (styles["a/value"]["color"], styles["a/value"]["align"]) == (PRIMARY, "right")
    assert (styles["b/item"]["color"], styles["b/item"]["bold"]) == (ACCENT, True)
    assert (styles["b/value"]["color"], styles["b/value"]["bold"], styles["b/value"]["align"]) == (
        "FF0000",
        True,
        "right",
    )
    # 見出し：セル > 見出し > 列 > 表全体 > 役割（heading の色より表全体の色が優先する）。
    assert (styles["_header/item"]["align"], styles["_header/item"]["color"]) == ("left", MUTED)
    assert (styles["_header/value"]["align"], styles["_header/value"]["color"]) == ("center", PRIMARY)
    assert styles["_header/value"]["bold"] is True
    assert all(style["valign"] == "middle" for style in styles.values())
    # 余白の一部の辺は前の層の値を引き継ぐ。
    assert styles["a/value"]["padding_pt"] == 10
    assert styles["a/item"]["padding_pt"] == {"left": 20, "top": 10, "right": 10, "bottom": 10}
    assert styles["b/item"]["padding_pt"] == {"left": 20, "top": 2, "right": 10, "bottom": 10}
    # 列と行が同じ項目を指定し、セルの指定がなければ行の指定が優先する。列の塗りは交互の塗りより優先する。
    table = simple(
        slide,
        id="overlap",
        band_fills=["surface"],
        column_styles={"value": {"color": "primary", "fill": "accent"}},
        row_styles={"a": {"color": "accent"}},
    )
    assert cell_style(table, "a", "value")["color"] == ACCENT  # 行 > 列
    assert cell_style(table, "b", "value")["color"] == PRIMARY  # 行の指定がなければ列
    assert [cell_style(table, row, "value")["fill"] for row in ("a", "b")] == [ACCENT, ACCENT]  # 列 > 交互の塗り
    assert [cell_style(table, row, "item")["fill"] for row in ("a", "b")] == [SURFACE, SURFACE]


def test_column_roles_and_math_role_select_roles(make_deck):
    slide = make_deck(ROLES).slide(id="roles")
    rows = ("first", "second", "third", "fourth")
    table = slide.table(
        id="table",
        headers=["名前", "値"],
        column_keys=["name", "value"],
        row_keys=list(rows),
        rows=[["名前", Math("x")], [Math("y"), "説明"], [[Text("値 "), Math("z")], "補足"], ["名前", Math("w")]],
        area=AREA,
        column_roles={"name": "label"},
        math_role="formula",
        cell_roles={"second/value": "note", "_header/value": "body", "fourth/name": "note", "fourth/value": "note"},
    )
    roles = {
        f"{row}/{column}": table.cell(row, column).data["role"]
        for row in ("_header", *rows)
        for column in ("name", "value")
    }
    assert roles == {
        "_header/name": "heading",
        "_header/value": "body",  # 見出しは header_role。cell_roles が最優先
        "first/name": "label",
        "first/value": "formula",  # 数式だけのセルは math_role
        "second/name": "formula",
        "second/value": "note",  # math_role は column_roles より優先
        "third/name": "label",
        "third/value": "body",  # 数式と文字が混じるセルは column_roles、ほかは表の役割
        "fourth/name": "note",
        "fourth/value": "note",  # 本文でも cell_roles は column_roles・math_role より優先
    }
    assert cell_style(table, "first", "value")["size_pt"] == 22
    assert cell_style(table, "first", "name")["bold"] is True


def test_group_style_excludes_cell_formats(make_deck):
    # 表のまとまりの書式はセル用の書式を含めずに解決する。セルには表全体の書式が付く。
    slide = make_deck().slide(id="group")
    table = simple(
        slide, valign="middle", padding_pt=16, style={"fill": "accent", "color": "muted"}, description="手順の表"
    )
    plain = slide.group(id="plain", area=AREA)
    group_style = table.data["style"]
    assert group_style == plain.data["style"]
    # 説明はまとまりに付き、セルには付かない。
    assert table.data["description"] == "手順の表" and "description" not in table.cell("a", "item").data
    assert (group_style["valign"], group_style["padding_pt"], group_style["fill"]) == ("top", 0, None)
    cell = cell_style(table, "a", "item")
    assert (cell["valign"], cell["padding_pt"], cell["fill"], cell["color"]) == ("middle", 16, ACCENT, MUTED)


# ---- 見出し行と戻り値 ----


def test_table_without_header_needs_no_headers(make_deck):
    slide = make_deck().slide(id="plain")
    table = slide.table(id="table", rows=[["1", "2"]], row_keys=["only"], column_keys=["left", "right"], area=AREA)
    assert isinstance(table, Table)
    assert [element.id for element in slide.elements] == ["table", "table/only/left", "table/only/right"]
    assert table.row_heights == (AREA.h,)
    with pytest.raises(InputError, match="表に行キー '_header' がありません") as caught:
        table.cell("_header", "left")
    assert caught.value.code == "invalid_table_key"
    # header=True だけでは見出し行の内容がない。headers があっても header=False なら見出し行を作らない。
    with pytest.raises(InputError, match="headers") as caught:
        slide.table(id="missing", rows=[["1"]], row_keys=["only"], column_keys=["value"], area=AREA, header=True)
    assert (caught.value.code, caught.value.element_id) == ("input_error", "missing")
    hidden = slide.table(
        id="hidden", headers=["値"], header=False, rows=[["1"]], row_keys=["only"], column_keys=["value"], area=AREA
    )
    assert [element.id for element in slide.elements if element.id.startswith("hidden/")] == ["hidden/only/value"]
    assert hidden.row_heights == (AREA.h,)


def test_scalar_sizes_and_header_height(make_deck):
    slide = make_deck().slide(id="sizes")
    # 数値の列幅・行高は全列・全行に使い、header_height は見出し行だけを替える。指定領域の寸法は使わない。
    table = simple(slide, column_widths=200, row_heights=40, header_height=60)
    assert (table.column_widths, table.row_heights) == ((200, 200), (60, 40, 40))
    assert table.area == Rect(AREA.x, AREA.y, 400, 140)
    assert table.cell("b", "value").area == Rect(AREA.x + 200, AREA.y + 100, 200, 40)
    # row_heights を省くと、見出し行を除いた残りの高さを本文の行で等分する。
    table = simple(slide, id="rest", area=Rect(40, 100, 600, 200), header_height=50)
    assert table.row_heights == (50, 75, 75) and table.area == Rect(40, 100, 600, 200)
    # 行高の列は見出し行を含む。
    assert simple(slide, id="list", row_heights=[30, 40, 50]).row_heights == (30, 40, 50)
    for arguments, pattern in (
        ({"column_widths": 0}, "column_widths は正の有限の数"),
        ({"row_heights": [30, 40]}, "row_heights の数が表と一致しません"),
        ({"column_widths": [100, float("nan")]}, "column_widths は正の有限の数"),
        ({"area": Rect(40, 100, 600, 50), "header_height": 50}, "見出し行の高さ"),
        ({"area": Rect(40, 100, 0, 240)}, "列幅を等分する領域の寸法が 0 以下です: 0$"),
    ):
        with pytest.raises(InputError, match=pattern) as caught:
            simple(slide, id="bad", **arguments)
        assert (caught.value.code, caught.value.element_id) == ("invalid_geometry", "bad")
    with pytest.raises(InputError, match="header_height") as caught:
        slide.table(id="bad", rows=[["1"]], row_keys=["only"], column_keys=["value"], area=AREA, header_height=40)
    assert caught.value.code == "input_error"
    assert not any(element.id.startswith("bad") for element in slide.elements)


def test_auto_size_updates_returned_area(make_deck):
    fonts = FakeFonts()
    slide = make_deck(fonts=fonts).slide(id="auto")
    table = simple(slide, column_widths="auto", row_heights="auto", word_wrap=False)
    # FakeFonts の送り幅は 0.5 em。見出し（24 pt）の「項目」と本文（20 pt）の「10」が各列で最も広く、左右の余白は 6 pt ずつ。
    assert table.column_widths == (2 * 12 + 12, 2 * 10 + 12)
    assert table.area == Rect(AREA.x, AREA.y, sum(table.column_widths), sum(table.row_heights))
    assert table.area != AREA
    group = by_id(slide)["table"]
    assert group["area"] == list(table.area) and group["required_size"] == [table.area.w, table.area.h]
    # 領域の幅・高さが None なら、内容から決める。
    unknown = simple(slide, id="unknown", area=Rect(40, 360, None, None), word_wrap=False)
    assert unknown.column_widths == table.column_widths and unknown.row_heights == table.row_heights
    assert by_id(slide)["unknown"]["area"] == [40, 360, table.area.w, table.area.h]
    # 高さだけが None の領域では、幅は等分し、行高の和を高さにする。
    stacked = simple(slide, id="stacked", area=Rect(40, 360, 500, None), row_heights=[30, 40, 50])
    assert stacked.area == Rect(40, 360, 500, 120) and stacked.column_widths == (250, 250)


def test_automatic_table_dimensions_use_header_and_explicit_breaks(make_deck):
    fonts = FakeFonts(bold_advance=1.0)
    slide = make_deck(fonts=fonts).slide(id="comparison")
    table = slide.table(
        id="table",
        headers=["H"],
        rows=[["AB\nC"]],
        row_keys=["value"],
        column_keys=["value"],
        area=Rect(20, 100, 200, 150),
        column_widths="auto",
        row_heights="auto",
    )
    header, body = table.cell("_header", "value").data, table.cell("value", "value").data
    width = layout_text(header["content"], header["style"], fonts).width
    # 列幅は太字の見出し（24 pt × 1.0 em と左右の余白、折り返しの誤差の帯）で決まる。
    assert table.column_widths == (width,)
    assert width > layout_text(body["content"], body["style"], fonts).width
    assert width == pytest.approx(24 + 12 + 0.004 * 24 + 0.25, abs=0.1)
    # 行高は列幅で計測した各行の高さ。本文は明示改行の 2 行。
    heights = tuple(layout_text(cell["content"], cell["style"], fonts, width=width).height for cell in (header, body))
    assert table.row_heights == heights
    one_line = layout_text([{"kind": "text", "text": "AB"}], body["style"], fonts, width=width).height
    assert heights[1] > one_line


def test_table_auto_rows_wrap_with_fixed_columns(make_deck):
    fonts = FakeFonts()
    slide = make_deck(fonts=fonts).slide(id="wrap")
    text = "abcdefghij klmnopqrst uvwxyz"
    table = slide.table(
        id="table",
        rows=[[text, "短い"]],
        row_keys=["row"],
        column_keys=["long", "short"],
        area=Rect(40, 100, 300, 200),
        column_widths=[100, 200],
        row_heights="auto",
    )
    style = table.cell("row", "long").data["style"]
    wrapped = layout_text([{"kind": "text", "text": text}], style, fonts, width=100)
    assert len(wrapped.lines) > 1
    assert table.row_heights == (wrapped.height,)
    single = layout_text([{"kind": "text", "text": text}], style, fonts, width=None)
    assert wrapped.height > single.height
    # 折り返さないセルでは、列幅が決まっていても 1 行の高さになる。
    table = slide.table(
        id="nowrap",
        rows=[[text, "短い"]],
        row_keys=["row"],
        column_keys=["long", "short"],
        area=Rect(40, 100, 300, 200),
        column_widths=[100, 200],
        row_heights="auto",
        word_wrap=False,
    )
    nowrap_style = table.cell("row", "long").data["style"]
    assert table.row_heights == (layout_text([{"kind": "text", "text": text}], nowrap_style, fonts, width=100).height,)
    assert table.row_heights[0] < wrapped.height


def test_auto_width_counts_trailing_space_before_math(make_deck):
    # 明示改行ごとの行全体（文字と数式の断片）を計測するので、断片の末尾の空白も幅に入る。
    slide = make_deck(fonts=FakeFonts()).slide(id="space")
    table = slide.table(
        id="table",
        rows=[[[Text("Total "), Math("x")]], [[Emphasis("ab"), Text("\nc")]]],
        row_keys=["sum", "other"],
        column_keys=["value"],
        area=AREA,
        column_widths="auto",
        word_wrap=False,
    )
    # 「Total 」6 文字と数式 1 文字（0.5 em × 20 pt）と左右の余白。
    assert table.column_widths == (7 * 10 + 12,)


def test_table_keys_survive_row_reordering(make_deck):
    deck = make_deck()

    def build(keys):
        slide = Deck(deck.project).slide(id="table")
        slide.table(
            id="comparison",
            headers=["方式", "値"],
            rows=[[key, Math("x")] for key in keys],
            row_keys=keys,
            column_keys=["method", "value"],
            area=Rect(20, 100, 500, 200),
            header_role="body",
        )
        return by_id(slide)

    first, second = build(["a", "b"]), build(["b", "a"])
    assert first["comparison/a/value"]["content"] == second["comparison/a/value"]["content"]
    assert first["comparison/a/value"]["area"] != second["comparison/a/value"]["area"]
    assert first["comparison/_header/method"]["style"]["fill"] == SURFACE
    assert first["comparison/a/method"]["style"]["fill"] == BACKGROUND


# ---- cell() と入力の誤り ----


def test_cell_accessor_and_unknown_keys(make_deck):
    slide = make_deck().slide(id="keys")
    table = simple(slide)
    assert table.cell("a", "value").id == "table/a/value" and table.cell("_header", "item").id == "table/_header/item"
    assert (table.row_keys, table.column_keys) == (("a", "b"), ("item", "value"))
    with pytest.raises(InputError, match=re.escape("表に列キー 'vlue' がありません（候補: value）")) as caught:
        table.cell("a", "vlue")
    assert (caught.value.code, caught.value.slide_id, caught.value.element_id) == ("invalid_table_key", "keys", "table")
    for arguments, pattern in (
        (
            {"column_styles": {"vlue": {"bold": True}}},
            "column_styles に存在しないキーがあります: 'vlue'（候補: value）",
        ),
        ({"row_styles": {"c": {"bold": True}}}, "row_styles に存在しないキーがあります: 'c'"),
        (
            {"row_styles": {"_header": {"bold": True}}},
            "row_styles に _header は使えません。見出し行の書式は header_style",
        ),
        (
            {"cell_styles": {"a/vlue": {"bold": True}}},
            "cell_styles に存在しないキーがあります: 'a/vlue'（候補: a/value",
        ),
        (
            {"cell_roles": {"_header/itm": "body"}},
            "cell_roles に存在しないキーがあります: '_header/itm'（候補: _header/item",
        ),
        ({"column_roles": {"name": "body"}}, "column_roles に存在しないキーがあります: 'name'"),
    ):
        with pytest.raises(InputError, match=re.escape(pattern)) as caught:
            simple(slide, id="bad", **arguments)
        assert (caught.value.code, caught.value.element_id) == ("invalid_table_key", "bad")
    assert not any(element.id.startswith("bad") for element in slide.elements)
    # 表に要素は追加できない。
    with pytest.raises(AttributeError, match="cell\\(\\) でセルを参照してください"):
        table.text(id="extra", text="追加", area=AREA)


def test_table_inside_group_uses_group_coordinates(make_deck):
    slide = make_deck().slide(id="nested")
    group = slide.group(id="panel", area=Rect(100, 50, 600, 400))
    table = group.table(
        id="facts",
        headers=["項目", "値"],
        rows=[["件数", "3"]],
        row_keys=["count"],
        column_keys=["item", "value"],
        area=Rect(10, 20, 400, 100),
    )
    assert isinstance(table, Table) and table.id == "panel/facts" and table.area == Rect(110, 70, 400, 100)
    cell = table.cell("count", "value")
    assert (cell.id, cell.data["parent_id"], cell.area) == (
        "panel/facts/count/value",
        "panel/facts",
        Rect(310, 120, 200, 50),
    )
    assert by_id(slide)["panel/facts"]["parent_id"] == "panel"
    # 表を動かすとセルも動く。
    table.move(5, 5)
    assert table.cell("count", "value").area == Rect(315, 125, 200, 50)


def test_table_errors_are_collected_with_fallbacks(make_deck):
    deck = make_deck(collect=True)
    slide = deck.slide(id="collect")
    table = simple(
        slide,
        cell_roles={"a/vlue": "note"},
        band_fills=["accent", "tela"],
        colour="text",
        column_styles={"value": {"align": "middel"}},
    )
    # 間違った指定を無視して表を作り、誤りを識別子付きで集める。
    assert cell_style(table, "b", "item")["fill"] == BACKGROUND and cell_style(table, "a", "item")["fill"] == ACCENT
    assert cell_style(table, "a", "value")["align"] == "left" and table.cell("a", "value").data["role"] == "body"
    issues = deck.to_dict()["issues"]
    assert sorted((issue["code"], issue["element_id"]) for issue in issues) == [
        ("invalid_style", "table"),
        ("invalid_table_key", "table"),
        ("undefined_color", "table"),
        ("unknown_argument", "table"),
    ]
    assert any("band_fills の 2 番目" in issue["message"] for issue in issues)


def test_table_rejects_rotation_and_unknown_style(make_deck):
    slide = make_deck().slide(id="reject")
    with pytest.raises(InputError, match="回転・反転") as caught:
        simple(slide, rotation=90)
    assert (caught.value.code, caught.value.element_id) == ("unsupported_argument", "table")
    for arguments, candidate in (
        ({"colour": "text"}, "colour（候補: color）"),
        ({"style": {"size": 20}}, "size（候補: size_pt）"),
        ({"header_style": {"wrap": False}}, "wrap（候補: word_wrap）"),
        ({"column_styles": {"value": {"halign": "right"}}}, "halign（候補: align）"),
        ({"row_height": 40}, "row_height（候補: row_heights"),
    ):
        with pytest.raises(TypeError, match=re.escape(candidate)) as caught:
            simple(slide, **arguments)
        assert caught.value.code == "unknown_argument" and caught.value.element_id == "table"
    for arguments, pattern in (
        ({"table_format": "native"}, "table_format は shapes"),
        ({"pptx_group": "yes"}, "pptx_group は真偽値か None"),
        ({"header": 1}, "header は真偽値か None"),
    ):
        with pytest.raises(InputError, match=pattern) as caught:
            simple(slide, **arguments)
        assert (caught.value.code, caught.value.element_id) == ("input_error", "table")
    with pytest.raises(InputError, match="重複") as caught:
        simple(slide, row_keys=["a", "a"])
    assert caught.value.code == "invalid_element_id"
    assert not slide.elements
