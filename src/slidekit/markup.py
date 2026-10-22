"""`$…$` を数式、`**…**` を強調として、文字列を文章の断片の列にする。

要素を作る関数は文字列の記法を自動では解釈しない。記法を使うときは `parse()` を明示して呼ぶ。

    from slidekit.markup import parse
    slide.text(id="premise", text=parse("対象の件数を $n$ とし、**最大 $k$ 件**まで扱う"), area=...)
"""

from __future__ import annotations

from .errors import InputError
from .fragments import Emphasis, Math, Text

# 記法の外で、次の文字の前の \ はエスケープとして扱う。それ以外の \ はそのまま残す（例：C:\path）。
_ESCAPED = frozenset("$*\\")
_CONTEXT = 10


def _context(source: str, index: int) -> str:
    """誤りの位置の前後の文字列（repr）。長い文字列は前後 10 文字に縮める。"""
    start, end = max(0, index - _CONTEXT), min(len(source), index + _CONTEXT)
    return ("…" if start else "") + repr(source[start:end]) + ("…" if end < len(source) else "")


def _error(source: str, index: int, problem: str) -> InputError:
    return InputError(
        f"{index + 1} 文字目の{problem}: {_context(source, index)}", code="invalid_markup", position=index + 1
    )


def parse(source: str, *, math: bool = True, emphasis: bool = True) -> list[Text | Emphasis | Math]:
    """`$...$` を Math、`**...**` を Emphasis にし、それ以外を Text にする。

    - `$...$` の中身はそのまま Math に渡す（最初の `$` で閉じる）。空の数式と、数式の中の改行は誤りにする。
    - `**...**` は強調で、中で改行してよい。中の `$...$` は `Math(..., emphasis=True)` にする。
      `**` は現れた順に開く・閉じるを繰り返す（入れ子はない）。空の強調（`****`）は誤りにする。
    - `\\$`・`\\*`・`\\\\` は `$`・`*`・`\\` にする。1 つの `*` は通常の文字である。
    - `math=False`・`emphasis=False` で、その記号を通常の文字として扱う（エスケープは常に処理する）。
    - 隣り合う同じ種類の文字の断片はつなげる（数式はつなげない）。空の文字列は空の列を返す。

    閉じていない記号などの誤りは、位置（1 始まりの文字数）と前後の文字列を添えた `InputError`
    （`ValueError` の派生、`code` は `invalid_markup`）にする。
    """
    if not isinstance(source, str):
        raise InputError(f"記法の文章は文字列で指定してください: {source!r:.40}", code="invalid_markup")
    for name, value in (("math", math), ("emphasis", emphasis)):
        if not isinstance(value, bool):
            raise InputError(f"parse() の {name} は真偽値で指定してください: {value!r:.40}", code="invalid_markup")
    result: list[Text | Emphasis | Math] = []
    buffer: list[str] = []
    opened = None  # 開いている強調の ** の位置
    start = 0  # 強調を開いたときの result の長さ

    def add(fragment):
        previous = result[-1] if result else None
        if type(fragment) is not Math and type(previous) is type(fragment):
            result[-1] = type(fragment)(previous.text + fragment.text)
        else:
            result.append(fragment)

    def flush():
        if buffer:
            add((Text if opened is None else Emphasis)("".join(buffer)))
            buffer.clear()

    index, length = 0, len(source)
    while index < length:
        character = source[index]
        if character == "\\" and index + 1 < length and source[index + 1] in _ESCAPED:
            buffer.append(source[index + 1])
            index += 2
        elif math and character == "$":
            end = source.find("$", index + 1)
            if end < 0:
                raise _error(source, index, " $ が閉じていません")
            expression = source[index + 1 : end]
            if not expression.strip():
                raise _error(source, index, "数式が空です")
            if "\n" in expression or "\r" in expression:
                raise _error(source, index, "数式の中で改行しています")
            flush()
            add(Math(expression, emphasis=opened is not None))
            index = end + 1
        elif emphasis and source.startswith("**", index):
            if opened is None:
                flush()
                opened, start = index, len(result)
            else:
                if not buffer and len(result) == start:
                    raise _error(source, opened, "強調が空です")
                flush()
                opened = None
            index += 2
        else:
            buffer.append(character)
            index += 1
    if opened is not None:
        raise _error(source, opened, " ** が閉じていません")
    flush()
    return result
