"""文章の断片（Text・Emphasis・Math）と、場面データの断片の列への変換。

断片の書式（`size_pt`・`bold`・`color`・`underline`）とリンクは、指定した項目だけが要素の書式より優先される。
値の型は断片を作るときに確かめる。色の名前（テーマが要る）とリンクの形式は、要素を作るときに確かめる。
"""

from __future__ import annotations

from dataclasses import KW_ONLY, dataclass, fields
from math import isfinite
import re
import unicodedata

from .errors import InputError

# 場面データの断片の "style" に書く項目（指定したものだけを書く）。
STYLE_KEYS = ("size_pt", "bold", "color", "underline")

# リンク先の形式。slidekit はリンク先へ接続しない。
LINK = re.compile(r"(?:https?://|mailto:)\S+")

# 文章で受け付けない文字の分類（制御文字、サロゲート、行・段落の区切り）。改行（\n）だけは受け付ける。
CONTROL_CATEGORIES = frozenset({"Cc", "Cs", "Zl", "Zp"})


def control_character(text: str) -> str | None:
    """文章で受け付けない文字（\\n 以外の制御文字など）のうち、最初の文字を返す。なければ None。"""
    return next((char for char in text if char != "\n" and unicodedata.category(char) in CONTROL_CATEGORIES), None)


def valid_link(value) -> bool:
    """`https://`・`http://`・`mailto:` で始まり、空白を含まない文字列か。"""
    return isinstance(value, str) and LINK.fullmatch(value) is not None


def _shown(value) -> str:
    """誤りの文言に入れる値（repr を 40 文字まで）。"""
    text = repr(value)
    return text if len(text) <= 40 else text[:39] + "…"


def _invalid(fragment, name: str, expected: str, value) -> InputError:
    return InputError(
        f"{type(fragment).__name__} の {name} は{expected}で指定してください: {_shown(value)}",
        code="invalid_fragment",
        argument=name,
        value=_shown(value),
    )


def _repr(fragment) -> str:
    """既定と違う項目だけを示す表記（例：Text('詳細', underline=True)）。"""
    options = "".join(
        f", {item.name}={getattr(fragment, item.name)!r}"
        for item in fields(fragment)
        if item.name != "text" and getattr(fragment, item.name) != item.default
    )
    return f"{type(fragment).__name__}({fragment.text!r}{options})"


def _validate(fragment) -> None:
    """断片の値の型を確かめ、size_pt を float にそろえる。"""
    if not isinstance(fragment.text, str):
        raise _invalid(fragment, "text", "文字列", fragment.text)
    for name in ("bold", "underline"):
        value = getattr(fragment, name, None)
        if value is not None and not isinstance(value, bool):
            raise _invalid(fragment, name, "真偽値", value)
    for name in ("color", "link"):
        value = getattr(fragment, name, None)
        if value is not None and not isinstance(value, str):
            raise _invalid(fragment, name, "文字列", value)
    size = fragment.size_pt
    if size is not None:
        if isinstance(size, bool) or not isinstance(size, (int, float)) or not isfinite(size) or size <= 0:
            raise _invalid(fragment, "size_pt", "正の有限の数", size)
        object.__setattr__(fragment, "size_pt", float(size))


@dataclass(frozen=True, repr=False)
class Text:
    """通常の文字。`link` は外部のリンク先（`https://`・`http://`・`mailto:` で始まる文字列）。"""

    text: str
    _: KW_ONLY
    bold: bool | None = None
    color: str | None = None
    size_pt: float | None = None
    underline: bool | None = None
    link: str | None = None

    def __post_init__(self):
        _validate(self)

    __repr__ = _repr


@dataclass(frozen=True, repr=False)
class Emphasis:
    """強調する文字。既定は太字と要素の強調色で、Text と同じ引数で上書きできる。"""

    text: str
    _: KW_ONLY
    bold: bool | None = None
    color: str | None = None
    size_pt: float | None = None
    underline: bool | None = None
    link: str | None = None

    def __post_init__(self):
        _validate(self)

    __repr__ = _repr


@dataclass(frozen=True, repr=False)
class Math:
    """数式。`emphasis=True` は要素の強調色と太字にする（`bold`・`color` の指定が優先）。"""

    text: str
    _: KW_ONLY
    bold: bool | None = None
    color: str | None = None
    size_pt: float | None = None
    emphasis: bool = False

    def __post_init__(self):
        _validate(self)
        if not isinstance(self.emphasis, bool):
            raise _invalid(self, "emphasis", "真偽値", self.emphasis)

    __repr__ = _repr


KINDS = ((Text, "text"), (Emphasis, "emphasis"), (Math, "math"))


def fragment_data(fragment: Text | Emphasis | Math) -> dict:
    """1 つの断片を場面データの辞書（kind と text と、指定した書式とリンク）にする。

    色の名前と数式の強調（emphasis）は、要素の書式で解決する前の値のまま入れる
    （要素を作る側が 16 進数の色と太字に直す）。指定がなければ、0.1 と同じ kind と text だけの辞書になる。
    """
    kind = next(name for cls, name in KINDS if isinstance(fragment, cls))
    data = {"kind": kind, "text": fragment.text}
    style = {key: getattr(fragment, key) for key in STYLE_KEYS if getattr(fragment, key, None) is not None}
    if style:
        data["style"] = style
    if getattr(fragment, "link", None) is not None:
        data["link"] = fragment.link
    if getattr(fragment, "emphasis", False):
        data["emphasis"] = True
    return data


def content_fragments(value):
    """文章を場面データの断片の列にする。文字列は Text として扱い、記法（$…$ など）は解釈しない。

    通常の文字と強調の断片に、改行（\\n）以外の制御文字があれば誤りにする（数式の中身は数式の解析が確かめる）。
    """
    if value is None:
        return []
    if isinstance(value, (str, Text, Emphasis, Math)):
        value = [value]
    if not isinstance(value, (list, tuple)):
        raise InputError(
            f"文章は文字列または Text、Emphasis、Math の断片で指定してください: {_shown(value)}",
            code="invalid_fragment",
            value=_shown(value),
        )
    result = []
    for fragment in value:
        if isinstance(fragment, str):
            fragment = Text(fragment)
        if not isinstance(fragment, (Text, Emphasis, Math)) or not isinstance(fragment.text, str):
            raise InputError(
                f"不正な文章の断片です: {_shown(fragment)}", code="invalid_fragment", value=_shown(fragment)
            )
        if not isinstance(fragment, Math) and (char := control_character(fragment.text)) is not None:
            raise InputError(
                f"文章にタブなどの制御文字は使えません（改行は \\n だけを使えます）: U+{ord(char):04X}",
                code="invalid_fragment",
                character=f"U+{ord(char):04X}",
            )
        result.append(fragment_data(fragment))
    return result
