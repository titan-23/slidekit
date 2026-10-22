"""表の作成。セルを図形の要素として展開し、列幅・行高を決める。

セルの役割は前ほど優先して次の順に選ぶ：`cell_roles["行キー/列キー"]`、見出し行なら `header_role`、
内容が数式の断片だけなら `math_role`、`column_roles[列キー]`、`role`、ページの既定の役割。

セルの書式は役割（とテーマの既定）の上に次の順に重ねる（後ろほど優先）。
見出しのセルは種類 `table_header`、本文のセルは `table_cell` の既定を使う。

- 見出しのセル：表全体（`style=` < キーワード引数） < 列（`column_styles`） < 見出し（`header_style`） < セル（`cell_styles`）
- 本文のセル：表全体 < 交互の塗り（`band_fills`） < 列 < 行（`row_styles`） < セル

表のまとまりの書式はセル用の書式を含めずに役割とテーマだけで解決する（まとまり自身は描かない）。
`"auto"` の寸法はセルごとに解決した書式で `measure.layout_text` を使って計測する。列幅が決まっていれば、行高は折り返しを反映する。
"""

from __future__ import annotations

from collections.abc import Mapping

from . import styles
from .errors import InputError
from .fragments import content_fragments
from .layout import Rect, as_rect
from .measure import MeasureError, layout_text
from .model import Group, _identifier, _shown

HEADER = "_header"  # 見出し行の行キー

# 表のまとまりに渡す引数（セルの書式でないもの）。回転・反転はまとまりが拒否する。
_GROUP_ARGUMENTS = styles.GROUP_ATTRIBUTE_KEYS | {
    "sources",
    "rotation",
    "flip_h",
    "flip_v",
    "_parent_id",
    "_element_id",
}
# 受け付けない名前の候補に使う、表の引数。
_PARAMETERS = (
    frozenset(
        {
            "id",
            "rows",
            "row_keys",
            "column_keys",
            "area",
            "headers",
            "header",
            "column_widths",
            "row_heights",
            "header_height",
            "role",
            "header_role",
            "column_roles",
            "math_role",
            "cell_roles",
            "header_style",
            "band_fills",
            "column_styles",
            "row_styles",
            "cell_styles",
            "style",
            "pptx_group",
            "table_format",
        }
    )
    | _GROUP_ARGUMENTS
)
_AUTO = "auto"
_REST = object()  # 表の高さから見出し行を除いた残りを、本文の行で等分する行


class Table(Group):
    """`slide.table()` の戻り値。表のまとまりで、セルは `cell()` で参照する。

    `area` は実際の表の範囲（列幅の和と行高の和）。表に要素は追加できない。
    """

    def __getattr__(self, name):
        if name in Group._CREATORS:
            raise AttributeError(f"表に要素は追加できません。cell() でセルを参照してください（{name}）")
        raise AttributeError(name)

    @property
    def row_keys(self) -> tuple[str, ...]:
        """本文の行キー（見出し行の _header は含めない）。"""
        return tuple(self.data["row_keys"])

    @property
    def column_keys(self) -> tuple[str, ...]:
        return tuple(self.data["column_keys"])

    @property
    def column_widths(self) -> tuple[float, ...]:
        return tuple(self.data["column_widths"])

    @property
    def row_heights(self) -> tuple[float, ...]:
        """見出し行を含む行の高さ。"""
        return tuple(self.data["row_heights"])

    def cell(self, row_key, column_key):
        """セルの要素を返す。見出し行のセルは row_key に "_header" を指定する。"""
        cells = {element.id for element in self.slide.elements if element.data.get("parent_id") == self.id}
        header = f"{self.id}/{HEADER}/{self.data['column_keys'][0]}" in cells
        rows = ([HEADER] if header else []) + list(self.row_keys)
        for name, key, keys in (("行キー", row_key, rows), ("列キー", column_key, list(self.column_keys))):
            if not isinstance(key, str) or key not in keys:
                raise InputError(
                    f"{self.slide._context(self.id)}: 表に{name} {_shown(key)} がありません"
                    f"{styles.format_candidates(styles.suggest(key, keys))}",
                    code="invalid_table_key",
                    slide_id=self.slide.id,
                    element_id=self.id,
                    key=_shown(key),
                )
        element_id = f"{self.id}/{row_key}/{column_key}"
        return next(element for element in self.slide.elements if element.id == element_id)


def add_table(
    slide,
    *,
    id,
    rows,
    row_keys,
    column_keys,
    area,
    headers=None,
    header=None,
    column_widths=None,
    row_heights=None,
    header_height=None,
    role=None,
    header_role="heading",
    column_roles=None,
    math_role=None,
    cell_roles=None,
    header_style=None,
    band_fills=None,
    column_styles=None,
    row_styles=None,
    cell_styles=None,
    style=None,
    pptx_group=None,
    table_format="shapes",
    **kwargs,
) -> Table:
    """表を作り、`Table` を返す。

    - headers：見出し行の内容（None なら見出し行を作らない）。header は互換用で、False は見出し行を作らず、
      True は headers を必須にする。
    - column_widths：None（領域の幅を等分）、"auto"（内容から）、数値（全列）、数値の列。
      row_heights：None（領域の高さを等分）、"auto"、数値（全行）、見出し行を含む数値の列。
      header_height：見出し行だけの高さ（数値か "auto"）。row_heights より優先し、row_heights が None なら残りを本文の行で等分する。
      領域の幅・高さが None のときは、None の列幅・行高を "auto" として扱う。
    - column_roles・column_styles のキーは列キー、row_styles のキーは本文の行キー、cell_roles・cell_styles のキーは
      "行キー/列キー"（見出しは "_header/列キー"）。band_fills は本文の行の塗りをこの順に繰り返す。
    - sources・description・allow_roles は表のまとまりに付ける。pptx_group はまとまりへ渡す。
    - そのほかのキーワード引数は、すべてのセルの書式（表全体の層）である。
    """
    try:
        _identifier(id)
    except InputError as error:
        raise slide._locate(error, None) from None
    table_id = kwargs.get("_element_id") or id

    def failure(message, code="input_error", **details):
        return InputError(
            f"{slide._context(table_id)}: {message}", code=code, slide_id=slide.id, element_id=table_id, **details
        )

    if table_format != "shapes":
        raise failure(f"table_format は shapes だけに対応しています: {_shown(table_format)}")
    if header is None:
        header = headers is not None
    elif not isinstance(header, bool):
        raise failure(f"header は真偽値か None で指定してください: {_shown(header)}")
    elif header and headers is None:
        raise failure("header=True の表には headers（見出し行の内容）を指定してください")
    if pptx_group is not None and not isinstance(pptx_group, bool):
        raise failure(f"pptx_group は真偽値か None で指定してください: {_shown(pptx_group)}")
    for name, value in (("rows", rows), ("row_keys", row_keys), ("column_keys", column_keys)) + (
        (("headers", headers),) if header else ()
    ):
        if not isinstance(value, (list, tuple)):
            raise failure(f"{name} は list か tuple で指定してください: {_shown(value)}")
    if not column_keys:
        raise failure("表には列キーを 1 つ以上指定してください")
    if (header and len(headers) != len(column_keys)) or len(rows) != len(row_keys):
        raise failure("表の見出し・行列とキーの数が一致しません")
    for keys in (row_keys, column_keys):
        for key in keys:
            try:
                _identifier(key, "表のキー")
            except InputError as error:
                raise slide._locate(error, table_id) from None
        if len(set(keys)) != len(keys):
            raise failure(
                f"表のキーが重複しています: {_shown(sorted({key for key in keys if keys.count(key) > 1}))}",
                code="invalid_element_id",
            )
    if HEADER in row_keys:
        raise failure("表の行キー _header は見出し用に予約されています", code="invalid_element_id")
    if any(not isinstance(row, (list, tuple)) or len(row) != len(column_keys) for row in rows):
        raise failure("表の列数が一致しません")
    content_rows = ([list(headers)] if header else []) + [list(row) for row in rows]
    keys = ([HEADER] if header else []) + list(row_keys)
    if not content_rows:
        raise failure("空の表は作成できません")
    if header_height is not None and not header:
        raise failure("見出し行のない表には header_height を指定できません")
    try:
        area = as_rect(area)
    except InputError as error:
        raise slide._locate(error, table_id) from None

    group_arguments = {name: value for name, value in kwargs.items() if name in _GROUP_ARGUMENTS}
    cell_style = {name: value for name, value in kwargs.items() if name not in _GROUP_ARGUMENTS}
    if pptx_group is not None:
        group_arguments["pptx_group"] = pptx_group
    start = len(slide.elements)
    # 内容から決める寸法（None）は 0 で作り、寸法を決めた後に実際の範囲へ直す。
    group = slide.group(id=id, area=Rect(area.x, area.y, area.w or 0.0, area.h or 0.0), role=role, **group_arguments)
    table = Table(slide, group.data)
    slide.elements[slide.elements.index(group)] = table
    try:
        cells = _cells(
            slide,
            table,
            content_rows,
            keys,
            column_keys,
            header=header,
            role=role,
            header_role=header_role,
            column_roles=column_roles,
            math_role=math_role,
            cell_roles=cell_roles,
            style=style,
            cell_style=cell_style,
            header_style=header_style,
            band_fills=band_fills,
            column_styles=column_styles,
            row_styles=row_styles,
            cell_styles=cell_styles,
        )
        widths, heights = _sizes(slide, table, cells, area, column_widths, row_heights, header_height, header=header)
        table.data.update(
            area=[area.x, area.y, sum(widths), sum(heights)],
            table_format="shapes",
            row_keys=list(row_keys),
            column_keys=list(column_keys),
            column_widths=widths,
            row_heights=heights,
            required_size=[sum(widths), sum(heights)],
        )
        y = area.y
        for row, height in zip(cells, heights):
            x = area.x
            for cell, width in zip(row, widths):
                slide.box(
                    id=id,
                    _element_id=cell.id,
                    _parent_id=table.id,
                    _style_kind=cell.kind,
                    area=Rect(x, y, width, height),
                    text=cell.value,
                    role=cell.role,
                    style=cell.style,
                )
                x += width
            y += height
    except BaseException:
        # 失敗した表は、まとまりも作ったセルも残さない（ほかの要素と同じく、失敗した呼び出しは要素を残さない）。
        del slide.elements[start:]
        raise
    return table


class _Cell:
    """セルを作る前に決めた値（識別子、書式の種類、役割、重ねた書式、断片）。"""

    __slots__ = ("id", "kind", "value", "fragments", "role", "style", "resolved")

    def __init__(self, id, kind, value, fragments, role, style):
        self.id, self.kind, self.value, self.fragments, self.role, self.style = id, kind, value, fragments, role, style
        self.resolved = None  # 計測に使う解決済みの書式（自動の寸法のときだけ求める）


def _mapping(slide, table, name, value, *, code) -> dict:
    """辞書の引数。None は空の辞書、辞書でなければ誤り（集めずに送出する）。"""
    if value is None:
        return {}
    if not isinstance(value, Mapping):
        raise InputError(
            f"{slide._context(table.id)}: {name} は辞書で指定してください: {_shown(value)}",
            code=code,
            slide_id=slide.id,
            element_id=table.id,
        )
    return dict(value)


def _known_keys(slide, table, name, values: dict, valid, *, hint=None) -> dict:
    """存在しないキーを候補付きの誤り（invalid_table_key）にする。集めるときは、その指定を除いて続ける。"""
    result = {}
    for key, value in values.items():
        if key in valid:
            result[key] = value
            continue
        message = (hint(key) if hint is not None else None) or (
            f"{name} に存在しないキーがあります: {_shown(key)}{styles.format_candidates(styles.suggest(key, valid))}"
        )
        slide._problem(
            InputError(
                f"{slide._context(table.id)}: {message}",
                code="invalid_table_key",
                slide_id=slide.id,
                element_id=table.id,
                argument=name,
                key=_shown(key),
            )
        )
    return result


def _style_layer(slide, table, name, value, kind, definition, element_id=None) -> dict:
    """書式の辞書を、名前と値を確かめた層にする。集めるときは、間違った項目を除く。"""
    element_id = table.id if element_id is None else element_id
    if value is None:
        return {}
    if not isinstance(value, Mapping):
        raise InputError(
            f"{slide._context(element_id)}: {name} は書式の辞書で指定してください: {_shown(value)}",
            code="invalid_style",
            slide_id=slide.id,
            element_id=element_id,
        )
    return slide._clean(slide._known_names(kind, value, element_id), element_id, definition)


def _band_layers(slide, table, band_fills) -> list[dict]:
    """交互の塗りの層。間違った色は、集めるときはその行の塗りを指定しなかったものとして続ける。"""
    if band_fills is None:
        return []
    if isinstance(band_fills, (str, bytes)) or not isinstance(band_fills, (list, tuple)) or not band_fills:
        error = InputError(
            f"{slide._context(table.id)}: band_fills は 1 つ以上の色の list か tuple で指定してください: "
            f"{_shown(band_fills)}",
            code="invalid_style",
            slide_id=slide.id,
            element_id=table.id,
            setting="band_fills",
            value=_shown(band_fills),
        )
        slide._problem(error)
        return []
    layers = []
    for number, value in enumerate(band_fills, 1):
        label = f"band_fills の {number} 番目"
        problem = styles.check_style_value("fill", value)
        if problem:
            error = slide._value_error("fill", problem, value, table.id, label=label)
        else:
            error = None if value is None else slide._color_problem("fill", value, table.id, label=label)
        if error is not None:
            slide._problem(error)
            layers.append({})
        else:
            layers.append({"fill": value})
    return layers


def _merge(layers, kind) -> dict:
    """層を 1 つの書式の辞書に重ねる（後ろほど優先）。None は未指定で、セルの fill・line_color の None は "none"。

    padding_pt の一部の辺の辞書は、前の層の値を引き継ぐ（前の層がなければ役割とテーマの値を引き継ぐ）。
    """
    result = {}
    for layer in layers:
        for name, value in styles.element_layer(layer, kind).items():
            if name == "padding_pt" and isinstance(value, Mapping) and "padding_pt" in result:
                previous = result["padding_pt"]
                sides = (
                    dict(previous) if isinstance(previous, Mapping) else dict.fromkeys(styles.PADDING_SIDES, previous)
                )
                value = {**sides, **value}
            result[name] = value
    return result


def _cells(
    slide,
    table,
    content_rows,
    keys,
    column_keys,
    *,
    header,
    role,
    header_role,
    column_roles,
    math_role,
    cell_roles,
    style,
    cell_style,
    header_style,
    band_fills,
    column_styles,
    row_styles,
    cell_styles,
):
    """セルごとに、断片・役割・重ねた書式を決める。書式の名前と値は、指定ごとに 1 回だけ確かめる。"""
    definition = slide.project.theme["roles"][table.data["role"]]
    body_keys = keys[1:] if header else keys
    cell_keys = [f"{row_key}/{column_key}" for row_key in keys for column_key in column_keys]
    # 表全体の書式はセルの書式（table_cell）として名前と値を確かめる。
    table_layers = [
        _style_layer(slide, table, "style", style, "table_cell", definition),
        slide._clean(
            slide._known_names("table_cell", cell_style, table.id, extra=_PARAMETERS, arguments=True),
            table.id,
            definition,
        ),
    ]
    header_layer = _style_layer(slide, table, "header_style", header_style, "table_header", definition)
    bands = _band_layers(slide, table, band_fills)
    columns = {
        key: _style_layer(slide, table, f"column_styles の {key}", value, "table_cell", definition)
        for key, value in _known_keys(
            slide,
            table,
            "column_styles",
            _mapping(slide, table, "column_styles", column_styles, code="invalid_style"),
            column_keys,
        ).items()
    }
    header_hint = lambda key: (
        "row_styles に _header は使えません。見出し行の書式は header_style で指定してください"
        if key == HEADER
        else None
    )
    body_rows = {
        key: _style_layer(slide, table, f"row_styles の {key}", value, "table_cell", definition)
        for key, value in _known_keys(
            slide,
            table,
            "row_styles",
            _mapping(slide, table, "row_styles", row_styles, code="invalid_style"),
            body_keys,
            hint=header_hint,
        ).items()
    }
    own_styles = _known_keys(
        slide, table, "cell_styles", _mapping(slide, table, "cell_styles", cell_styles, code="invalid_style"), cell_keys
    )
    roles_by_column = _known_keys(
        slide,
        table,
        "column_roles",
        _mapping(slide, table, "column_roles", column_roles, code="input_error"),
        column_keys,
    )
    roles_by_cell = _known_keys(
        slide, table, "cell_roles", _mapping(slide, table, "cell_roles", cell_roles, code="input_error"), cell_keys
    )
    declared = slide._declared((), table)
    result = []
    for index, (row_key, values) in enumerate(zip(keys, content_rows)):
        is_header = header and index == 0
        kind = "table_header" if is_header else "table_cell"
        row = []
        for column_key, value in zip(column_keys, values):
            cell_id = f"{table.id}/{row_key}/{column_key}"
            name = f"{row_key}/{column_key}"
            try:
                fragments = content_fragments(value)
            except InputError as error:
                raise slide._locate(error, cell_id) from None
            selected = _select_role(
                name,
                column_key,
                fragments,
                is_header,
                role=role,
                header_role=header_role,
                math_role=math_role,
                column_roles=roles_by_column,
                cell_roles=roles_by_cell,
            )
            selected = slide._resolve_role(selected, content=fragments, element_id=cell_id, declared=declared)
            # セルの書式の値はセルの識別子で確かめる（最小文字サイズはセルの役割で確かめる）。
            own = _style_layer(
                slide,
                table,
                f"cell_styles の {name}",
                own_styles.get(name),
                kind,
                slide.project.theme["roles"][selected],
                element_id=cell_id,
            )
            if is_header:
                layers = [*table_layers, columns.get(column_key, {}), header_layer, own]
            else:
                band = bands[(index - (1 if header else 0)) % len(bands)] if bands else {}
                layers = [*table_layers, band, columns.get(column_key, {}), body_rows.get(row_key, {}), own]
            row.append(_Cell(cell_id, kind, value, fragments, selected, _merge(layers, kind)))
        result.append(row)
    return result


def _select_role(name, column_key, fragments, is_header, *, role, header_role, math_role, column_roles, cell_roles):
    """セルの役割の名前（None はページの既定の役割）。"""
    if cell_roles.get(name) is not None:
        return cell_roles[name]
    if is_header and header_role is not None:
        return header_role
    if math_role is not None and fragments and all(fragment["kind"] == "math" for fragment in fragments):
        return math_role
    if column_roles.get(column_key) is not None:
        return column_roles[column_key]
    return role


def _size(slide, table, value, name) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not value > 0 or value == float("inf"):
        raise InputError(
            f"{slide._context(table.id)}: {name} は正の有限の数で指定してください: {_shown(value)}",
            code="invalid_geometry",
            slide_id=slide.id,
            element_id=table.id,
            value=_shown(value),
        )
    return float(value)


def _given(slide, table, value, count, name) -> list:
    """列幅・行高の指定を、数値か自動（None）の列にする。"""
    if value == _AUTO:
        return [None] * count
    if isinstance(value, (list, tuple)):
        if len(value) != count:
            raise InputError(
                f"{slide._context(table.id)}: {name} の数が表と一致しません（{len(value)} 個、表は {count} 個）",
                code="invalid_geometry",
                slide_id=slide.id,
                element_id=table.id,
            )
        return [_size(slide, table, item, name) for item in value]
    return [_size(slide, table, value, name)] * count


def _equal(slide, table, available, count, name) -> list:
    """領域の寸法の等分。"""
    if available <= 0:
        raise InputError(
            f"{slide._context(table.id)}: {name}を等分する領域の寸法が 0 以下です: {available:g}",
            code="invalid_geometry",
            slide_id=slide.id,
            element_id=table.id,
        )
    return [available / count] * count


def _sizes(slide, table, cells, area, column_widths, row_heights, header_height, *, header):
    """列幅と行高を決める。自動の寸法はセルの解決済みの書式で計測する。"""
    columns, count = len(cells[0]), len(cells)
    if column_widths is None and area.w is None:
        column_widths = _AUTO
    widths = (
        _equal(slide, table, area.w, columns, "列幅")
        if column_widths is None
        else _given(slide, table, column_widths, columns, "column_widths")
    )
    if row_heights is None and area.h is None:
        row_heights = _AUTO
    if row_heights is None:
        if header and header_height is not None:
            heights = [None] + [_REST] * (count - 1)
        else:
            heights = _equal(slide, table, area.h, count, "行高")
    else:
        heights = _given(slide, table, row_heights, count, "row_heights")
    if header and header_height is not None:
        heights[0] = None if header_height == _AUTO else _size(slide, table, header_height, "header_height")
    if any(width is None for width in widths):
        measured = [[_measure(slide, cell, None).width for cell in row] for row in cells]
        widths = [
            width if width is not None else max(row[column] for row in measured) for column, width in enumerate(widths)
        ]
    for index, height in enumerate(heights):
        if height is None:
            heights[index] = max(_measure(slide, cell, width).height for cell, width in zip(cells[index], widths))
    rest = [index for index, height in enumerate(heights) if height is _REST]
    if rest:
        remaining = area.h - sum(height for height in heights if height is not _REST)
        if remaining <= 0:
            raise InputError(
                f"{slide._context(table.id)}: 見出し行の高さ（{heights[0]:g}）が表の領域の高さ（{area.h:g}）以上です",
                code="invalid_geometry",
                slide_id=slide.id,
                element_id=table.id,
            )
        for index in rest:
            heights[index] = remaining / len(rest)
    return [float(width) for width in widths], [float(height) for height in heights]


def _measure(slide, cell, width):
    """セルの必要な枠の寸法（measure.TextLayout）。width が None なら折り返さない内容の幅から枠の幅を決める。"""
    try:
        if cell.resolved is None:
            cell.resolved = slide._style(
                cell.kind,
                cell.role,
                [(cell.style, False)],
                content=cell.fragments,
                element_id=cell.id,
                declared={cell.role},
            )[1]
        return layout_text(cell.fragments, cell.resolved, slide.project.fonts, width=width)
    except ValueError as error:
        # 入力の誤り（計測、数式の記法）に、セルを作るときと同じページと要素の識別子を付ける。
        # 数式の記法の誤りは、計測が包んだ元の誤り（math_syntax_error）で報告する。
        cause = error.__cause__
        if isinstance(error, MeasureError) and getattr(cause, "code", None) == "math_syntax_error":
            raise slide._locate(cause, cell.id) from None
        slide._locate(error, cell.id)
        raise
