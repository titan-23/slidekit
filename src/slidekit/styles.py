"""書式名の一覧と、書式を重ねる順序。

要素の書式は「ライブラリの既定 < テーマの `[defaults]` < テーマの `[defaults.<種類>]` < 役割 < 要素（`style=` < キーワード引数）」
の順に重ね、後ろほど優先する。`min_size_pt` は常に役割の値にする。色の名前は、すべてを重ねた後にテーマの `[colors]` で 16 進数に直す。

値の `None` は「指定しない」を表す。ただし文字枠を持つ種類（`FRAME_KINDS`）の `fill`・`line_color` の `None` は
「塗らない」「枠線を引かない」を表し、`"none"` と同じである。`"none"` は入力の表記で、解決後の場面の書式では
塗らない・線なしを `None` で表す。

作成 API・設定の検査（`schema`）・書き出し・検査・文書の試験がこの一覧を共有する。
このモジュールは `layout.finite_number` と `errors` の例外だけを使う。
"""

from __future__ import annotations

from collections.abc import Collection, Iterable, Mapping
import copy
import difflib
import re
from types import MappingProxyType
from typing import Any

from .errors import ArgumentNameError
from .layout import finite_number

TEXT_KEYS = frozenset(
    {
        "size_pt",
        "min_size_pt",
        "color",
        "bold",
        "align",
        "valign",
        "line_spacing",
        "line_spacing_pt",
        "padding_pt",
        "font_japanese",
        "font_latin",
        "font_math",
        "emphasis_color",
        "word_wrap",
        "phrase_wrap",
    }
)
FRAME_KEYS = frozenset({"fill", "line_color", "line_width_pt", "dashed"})
STROKE_KEYS = frozenset({"line_color", "line_width_pt", "dashed"})
# 吹き出しのしっぽ（box の rectangle・rounded_rectangle だけ）。既定値は outline の定数で、LIBRARY_DEFAULTS には置かない。
TAIL_KEYS = frozenset({"tail_side", "tail_at", "tail_width_pt", "tail_length_pt"})
TAIL_SIDE_VALUES = ("left", "right", "top", "bottom", "none")

# 要素・セルの引数で受け付ける書式。
STYLE_KEYS = MappingProxyType(
    {
        "text": TEXT_KEYS | FRAME_KEYS,
        "equation": TEXT_KEYS | FRAME_KEYS,
        "box": TEXT_KEYS | FRAME_KEYS | {"corner_radius_pt"} | TAIL_KEYS,
        "table_cell": TEXT_KEYS | FRAME_KEYS,
        "table_header": TEXT_KEYS | FRAME_KEYS,
        "line": STROKE_KEYS,
        "arrow": STROKE_KEYS,
        "image": frozenset(),
        "group": frozenset(),
    }
)
ALL_STYLE_KEYS = frozenset().union(*STYLE_KEYS.values())
# 文字枠を持つ種類。fill・line_color の None を「塗らない」「枠線を引かない」として扱う。
FRAME_KINDS = frozenset({"text", "box", "equation", "table_cell", "table_header"})

# [defaults.<名前>] を重ねる順（後ろほど優先）。表のセルは box の後に、矢印は line の後に重ねる。
DEFAULT_CHAINS = MappingProxyType(
    {
        "text": ("text",),
        "equation": ("equation",),
        "box": ("box",),
        "table_cell": ("box", "table_cell"),
        "table_header": ("box", "table_cell", "table_header"),
        "line": ("line",),
        "arrow": ("line", "arrow"),
    }
)
DEFAULT_KINDS = frozenset(DEFAULT_CHAINS)  # [defaults] の下で受け付ける表の名前

# 役割だけが持ち、書式へ写さない項目。
ROLE_ONLY_KEYS = frozenset({"inherits", "allowed_slide_ids", "allowed_content", "restricted", "skip_checks"})
ROLE_KEYS = TEXT_KEYS | FRAME_KEYS | {"corner_radius_pt"} | TAIL_KEYS | ROLE_ONLY_KEYS
PAINT_KEYS = frozenset({"fill", "line_color"})  # "none" を受け付ける色の項目
COLOR_KEYS = PAINT_KEYS | {"color", "emphasis_color"}

# 要素の属性（書式でない引数）。
ATTRIBUTE_KEYS = frozenset({"description", "semantic", "rotation", "flip_h", "flip_v", "allow_roles"})
# まとまり（表を含む）が受け付ける属性。まとまりは回転・反転できず、semantic も持たない。
GROUP_ATTRIBUTE_KEYS = frozenset({"description", "allow_roles"})
SEMANTIC_VALUES = frozenset({"background", "decoration"})  # 利用者が semantic に指定できる値

SHAPES = ("rectangle", "rounded_rectangle", "ellipse", "diamond")
ALIGN_VALUES = ("left", "center", "right", "justify")
VALIGN_VALUES = ("top", "middle", "bottom")
PADDING_SIDES = ("left", "top", "right", "bottom")
HEX_COLOR = re.compile(r"#?[0-9A-Fa-f]{6}")
COLOR_NAME = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")

# ライブラリの既定。書体はテーマの fonts、文字サイズと最小文字サイズは役割が決める（library_layer）。
# 色はテーマの [colors] の名前で書き、解決後の場面では 16 進数になる（塗らない・線なしは None）。
# math.office_math の関数の引数の既定値は公開 API の既定値として残し、この表には含めない。
LIBRARY_DEFAULTS = MappingProxyType(
    {
        "color": "text",
        "emphasis_color": "accent",
        "bold": False,
        "align": "left",
        "valign": "top",
        "line_spacing": 1.15,
        "padding_pt": 0.0,
        "fill": None,
        "line_color": None,
        "line_width_pt": 1.0,
        "word_wrap": True,
        "phrase_wrap": False,
    }
)

# 間違えやすい名前と、その正しい書式名。候補の提示だけに使い、読み替えには使わない。
ALIASES = MappingProxyType(
    {
        "size": "size_pt",
        "font_size": "size_pt",
        "fontsize": "size_pt",
        "font_size_pt": "size_pt",
        "text_size": "size_pt",
        "min_size": "min_size_pt",
        "colour": "color",
        "font_color": "color",
        "text_color": "color",
        "emphasis_colour": "emphasis_color",
        "padding": "padding_pt",
        "radius": "corner_radius_pt",
        "corner_radius": "corner_radius_pt",
        "line": "line_color",
        "border_color": "line_color",
        "border": "line_color",
        "stroke": "line_color",
        "line_colour": "line_color",
        "line_width": "line_width_pt",
        "width": "line_width_pt",
        "border_width": "line_width_pt",
        "stroke_width": "line_width_pt",
        "dash": "dashed",
        "wrap": "word_wrap",
        "wordwrap": "word_wrap",
        "background": "fill",
        "fill_color": "fill",
        "background_color": "fill",
        "fill_colour": "fill",
        "halign": "align",
        "text_align": "align",
        "vertical_align": "valign",
        "line_height": "line_spacing",
        "line_height_pt": "line_spacing_pt",
        "tail": "tail_side",
        "callout": "tail_side",
        "pointer": "tail_side",
        "pointer_side": "tail_side",
        "tail_position": "tail_at",
        "tail_width": "tail_width_pt",
        "tail_length": "tail_length_pt",
    }
)

# 受け付けない名前と、その理由（名前に続く文）。
UNSUPPORTED = MappingProxyType(
    {
        "z_index": "には対応していません。描画順は作成順で、後に作った要素が前面になります",
        "italic": "には対応していません。強調には Emphasis の断片か bold を使ってください",
        "shadow": "には対応していません",
        "autofit": "には対応していません。文字は自動で縮小しないので、枠か文を調整してください",
        "shrink_to_fit": "には対応していません。文字は自動で縮小しないので、枠か文を調整してください",
        "underline": "は要素の書式に使えません。断片の Text(..., underline=True) で指定してください",
        "link": "は要素の書式に使えません。断片の Text(..., link=...) で指定してください",
        "parent_id": "は指定できません。まとまりの子は group の戻り値から作成してください",
        "font": "は使えません。書体は font_japanese・font_latin・font_math で指定してください",
        "font_family": "は使えません。書体は font_japanese・font_latin・font_math で指定してください",
    }
)


def suggest(name: Any, candidates: Iterable[str]) -> tuple[str, ...]:
    """name に近い名前を最大 3 つ返す（difflib の類似度 0.6 以上）。name が文字列でなければ空のタプルを返す。"""
    if not isinstance(name, str):
        return ()
    return tuple(difflib.get_close_matches(name, sorted(set(map(str, candidates))), n=3, cutoff=0.6))


def format_candidates(names: Collection[str]) -> str:
    """候補の注記「（候補: a、b）」。候補がなければ空文字列。"""
    return f"（候補: {'、'.join(names)}）" if names else ""


def _label(name: Any) -> str:
    """誤りの文言に入れる名前。文字列でない名前（style= の辞書のキーなど）は repr で示す。"""
    return name if isinstance(name, str) else repr(name)


def _accepted(kind: str | None) -> frozenset[str]:
    if kind is None:
        return ALL_STYLE_KEYS
    if kind not in STYLE_KEYS:
        raise ValueError(f"未対応の要素の種類です: {kind}")
    return STYLE_KEYS[kind]


def attribute_keys(kind: str) -> frozenset[str]:
    """要素の種類が引数で受け付ける属性。まとまりは `GROUP_ATTRIBUTE_KEYS`、ほかは `ATTRIBUTE_KEYS`。"""
    return GROUP_ATTRIBUTE_KEYS if kind == "group" else ATTRIBUTE_KEYS


def unknown_names(
    kind: str | None, names: Iterable[str], *, extra: Iterable[str] = (), arguments: bool = False
) -> list[tuple[str, tuple[str, ...], str | None]]:
    """受け付けない名前を、(名前, 候補, 理由) の列で返す。

    受け付ける名前は `STYLE_KEYS[kind]` と extra の和で、kind が None なら全種類の書式の和である。
    理由は名前に続く文で、`UNSUPPORTED` の名前、ほかの種類の書式、要素の属性に付く。未知の名前では理由が None で、
    候補は `ALIASES` の行き先を受け付けるならそれを、なければ受け付ける名前のうち近いものを返す。
    arguments が真なら、names は要素を作る関数のキーワード引数で、extra にない属性は「その種類では使えない」とする
    （偽なら書式の辞書なので、属性は「書式に書けない」とする）。
    """
    accepted = _accepted(kind) | frozenset(extra)
    result, seen = [], set()
    for name in names:
        if name in accepted or name in seen:
            continue
        seen.add(name)
        if name in UNSUPPORTED:
            result.append((name, (), UNSUPPORTED[name]))
        elif name in ALL_STYLE_KEYS or (arguments and name in ATTRIBUTE_KEYS):
            result.append((name, (), f"は {kind} では使えません"))
        elif name in ATTRIBUTE_KEYS:
            result.append((name, (), "は書式に書けません。要素の引数で指定してください"))
        else:
            alias = ALIASES.get(name)
            result.append((name, (alias,) if alias in accepted else suggest(name, accepted), None))
    return result


def check_names(
    kind: str,
    names: Iterable[str],
    *,
    extra: Iterable[str] = (),
    context: str = "",
    slide_id: str | None = None,
    element_id: str | None = None,
    arguments: bool = False,
) -> None:
    """受け付けない名前をすべて集め、1 回の `ArgumentNameError` で報告する。

    context は文言の先頭（例：「ページ intro の要素 premise」）。未知の名前が 1 つでもあれば code は
    `unknown_argument`、理由の付く名前（ほかの種類の書式、要素の属性、`UNSUPPORTED`）だけなら `unsupported_argument` にする。
    文字列でない名前（style= の辞書のキー）は未知の名前として扱う。arguments は `unknown_names` と同じ。
    """
    problems = unknown_names(kind, names, extra=extra, arguments=arguments)
    if not problems:
        return
    unknown = [(name, suggestions) for name, suggestions, reason in problems if reason is None]
    parts = (
        [
            f"{kind} に未知の書式があります: "
            + "、".join(_label(name) + format_candidates(found) for name, found in unknown)
        ]
        if unknown
        else []
    )
    parts += [f"{_label(name)} {reason}" for name, _, reason in problems if reason is not None]
    message = "。".join(parts)
    raise ArgumentNameError(
        f"{context}: {message}" if context else message,
        code="unknown_argument" if unknown else "unsupported_argument",
        slide_id=slide_id,
        element_id=element_id,
        kind=kind,
        names=[name for name, _, _ in problems],
        problems=[{"name": name, "suggestions": list(found), "reason": reason} for name, found, reason in problems],
    )


def _number(value: Any) -> float | None:
    try:
        return finite_number(value, "値")
    except ValueError:
        return None


def _padding(value: Any) -> str | None:
    if isinstance(value, Mapping):
        unknown = [str(side) for side in value if side not in PADDING_SIDES]
        if unknown:
            return "に未知の辺があります: " + "、".join(
                side + format_candidates(suggest(side, PADDING_SIDES)) for side in unknown
            )
        if any(_number(amount) is None or amount < 0 for amount in value.values()):
            return "の値は 0 以上の有限値にしてください"
        return None
    number = _number(value)
    if number is None:
        return "は 0 以上の数値か、left・top・right・bottom を持つ辞書で指定してください"
    return "の値は 0 以上の有限値にしてください" if number < 0 else None


def _positive(value: Any) -> str | None:
    number = _number(value)
    return None if number is not None and number > 0 else "は正の有限の数で指定してください"


def _non_negative(value: Any) -> str | None:
    number = _number(value)
    return None if number is not None and number >= 0 else "は 0 以上の有限の数で指定してください"


def _boolean(value: Any) -> str | None:
    return None if isinstance(value, bool) else "は真偽値で指定してください"


def _fraction(value: Any) -> str | None:
    number = _number(value)
    return None if number is not None and 0 <= number <= 1 else "は 0 以上 1 以下の数で指定してください"


def _choice(values: tuple[str, ...]):
    def check(value: Any) -> str | None:
        return (
            None if isinstance(value, str) and value in values else f"は {'、'.join(values)} のどれかで指定してください"
        )

    return check


def _font(value: Any) -> str | None:
    return None if isinstance(value, str) and value.strip() else "は空でない文字列で指定してください"


def _color(value: Any) -> str | None:
    # 名前の有無と "none" の使い方は、テーマの色と合わせて確かめる（schema.check_theme_references、作成 API）。
    if isinstance(value, str) and (HEX_COLOR.fullmatch(value) or COLOR_NAME.fullmatch(value)):
        return None
    return "は色の名前か 6 桁の 16 進数で指定してください"


_VALUE_RULES = {
    **dict.fromkeys(
        ("size_pt", "min_size_pt", "line_spacing", "line_spacing_pt", "tail_width_pt", "tail_length_pt"), _positive
    ),
    **dict.fromkeys(("line_width_pt", "corner_radius_pt"), _non_negative),
    **dict.fromkeys(("bold", "word_wrap", "phrase_wrap", "dashed"), _boolean),
    **dict.fromkeys(("font_japanese", "font_latin", "font_math"), _font),
    **dict.fromkeys(COLOR_KEYS, _color),
    "align": _choice(ALIGN_VALUES),
    "valign": _choice(VALIGN_VALUES),
    "padding_pt": _padding,
    "tail_side": _choice(TAIL_SIDE_VALUES),
    "tail_at": _fraction,
}


def check_style_value(name: str, value: Any) -> str | None:
    """書式の値の問題を、名前に続く文（例：「は真偽値で指定してください」）で返す。問題がなければ None。

    None は未指定なので問題にしない。色は形だけを確かめ、テーマにある名前かどうかと、
    `"none"` を受け付ける項目かどうかは、テーマの色と合わせて確かめる。未知の名前は unknown_names が扱う。
    """
    rule = _VALUE_RULES.get(name)
    return None if value is None or rule is None else rule(value)


def library_layer(theme: Mapping) -> dict:
    """1 層目：テーマの書体とライブラリの既定。"""
    fonts = theme.get("fonts", {})
    return {
        "font_japanese": fonts.get("japanese"),
        "font_latin": fonts.get("latin"),
        "font_math": fonts.get("math"),
        **LIBRARY_DEFAULTS,
    }


def theme_layers(theme: Mapping, kind: str | None) -> list[dict]:
    """テーマの既定の層を、重ねる順に返す。

    `[defaults]` のうち種類の表でない値を 1 層目とし、`DEFAULT_CHAINS[kind]` の順に `[defaults.<名前>]` を重ねる
    （ない表は空の層）。kind が None（テーマの背景）と、種類別の既定のない種類（image・group）は 1 層目だけを返す。
    """
    _accepted(kind)  # 未対応の種類はここで ValueError にする
    defaults = theme.get("defaults") or {}
    layers = [{key: copy.deepcopy(value) for key, value in defaults.items() if key not in DEFAULT_KINDS}]
    for name in DEFAULT_CHAINS.get(kind, ()):
        table = defaults.get(name)
        layers.append(copy.deepcopy(dict(table)) if isinstance(table, Mapping) else {})
    return layers


def role_layer(role: Mapping) -> dict:
    """役割の層。役割だけが持つ項目（`ROLE_ONLY_KEYS`）を除く。"""
    return {key: copy.deepcopy(value) for key, value in role.items() if key not in ROLE_ONLY_KEYS}


def element_layer(overrides: Mapping | None, kind: str) -> dict:
    """要素の指定の層。値が None の項目を除く。

    ただし文字枠を持つ種類の fill・line_color の None は、"none"（塗らない・枠線なし）に置き換えて残す。
    `style=` の辞書とキーワード引数は別々に通し、この順に重ねる（キーワード引数の None で辞書の値を消さないため）。
    """
    layer = {}
    for key, value in (overrides or {}).items():
        if value is None:
            if key in PAINT_KEYS and kind in FRAME_KINDS:
                layer[key] = "none"
            continue
        layer[key] = value
    return layer


def fragment_style(style: Mapping, fragment: Mapping) -> dict:
    """断片の実効書式 {size_pt, bold, color, underline, superscript, link} を返す。

    style は要素の解決済みの書式、fragment は場面の断片（kind、text、任意の style・link）。断片の指定が要素より優先する。
    強調（emphasis の断片と、emphasis が真の数式）は、指定がなければ太字と要素の強調色を使う。
    書き出し・検査・計測は、断片の書式をこの関数だけで求める。
    """
    own = fragment.get("style") or {}
    kind = fragment.get("kind", "text")
    emphasized = kind == "emphasis" or (kind == "math" and bool(fragment.get("emphasis", False)))
    bold = own.get("bold")
    if bold is None:
        bold = bool(style.get("bold", LIBRARY_DEFAULTS["bold"])) or emphasized
    color = own.get("color")
    if color is None:
        color = style.get("emphasis_color") if emphasized else style.get("color")
    size = own.get("size_pt")
    return {
        "size_pt": style.get("size_pt") if size is None else size,
        "bold": bool(bold),
        "color": color,
        "underline": bool(own.get("underline")),
        "superscript": bool(own.get("superscript")),
        "link": fragment.get("link"),
    }
