"""ページの指定（`--slides`）の解釈。部分生成と確認記録が同じ書式を使う。

書式はカンマ区切りの項目の列である。前後の空白は無視する。

- `N`・`N-M`：ページ番号（1 始まり、両端を含む）。番号は資料の元のページ番号である。
- `id:識別子`：ページの識別子。数字だけの識別子や `2024-2025` のような識別子は、この形で書く。
- それ以外の項目：ページの識別子。

カンマを含む識別子は指定できない。
指定の誤り（書式、範囲外の番号、ない識別子）は `SelectionError` にする。利用者の入力の誤りとして、
トレースバックを出さずに文言だけを示す（診断 `invalid_selection`）。
"""

from __future__ import annotations

import difflib
import re
from typing import Iterable, Sequence

from .errors import InputError

_NUMBER = re.compile(r"[0-9]+")
_RANGE = re.compile(r"([0-9]+)\s*-\s*([0-9]+)")
# 範囲の書きかけ（"3-"、"-3"、"-"）。識別子として扱わずに誤りにする。
_PARTIAL_RANGE = re.compile(r"[0-9]*\s*-\s*[0-9]*")


class SelectionError(InputError):
    """ページの指定（--slides）の誤り。`InputError`（`ValueError` の派生）なので、`ValueError` としても捕まえることができる。"""

    code = "invalid_selection"


def _unreadable(item: str, hint: str = "") -> SelectionError:
    return SelectionError(f"--slides の指定を解釈できません: '{item}'{hint}")


def parse(text: str) -> list[tuple]:
    """指定を項目の列にする。

    "intro, 3, 5-7, id:2024" は [("id", "intro"), ("range", 3, 3), ("range", 5, 7), ("id", "2024")] になる。
    空の項目は無視する。項目が 1 つもなければ誤りにする。
    """
    if not isinstance(text, str):
        raise TypeError("ページの指定は文字列にしてください")
    items: list[tuple] = []
    for part in text.split(","):
        item = part.strip()
        if not item:
            continue
        if item.startswith("id:"):
            identifier = item[3:].strip()
            if not identifier:
                raise _unreadable(item)
            items.append(("id", identifier))
        elif _NUMBER.fullmatch(item):
            items.append(("range", int(item), int(item)))
        elif match := _RANGE.fullmatch(item):
            start, end = int(match.group(1)), int(match.group(2))
            if start > end:
                raise _unreadable(item, f"（範囲は {end}-{start} のように小さい番号から書く）")
            items.append(("range", start, end))
        elif _PARTIAL_RANGE.fullmatch(item):
            raise _unreadable(item)
        else:
            items.append(("id", item))
    if not items:
        raise SelectionError("ページを 1 つ以上指定してください")
    return items


def _numbers(slides: Sequence[dict]) -> list[int]:
    numbers = []
    for position, slide in enumerate(slides, 1):
        number = slide.get("number")
        numbers.append(number if isinstance(number, int) and not isinstance(number, bool) else position)
    return numbers


def _bounds(numbers: Sequence[int]) -> str:
    return f"{min(numbers)}〜{max(numbers)}" if numbers else "ページがありません"


def select(slides: Sequence[dict], text: str, *, known: Iterable[str] = ()) -> list[dict]:
    """指定に当たるページを、資料の順に重複なく返す。

    slides の各要素は "id" と "number"（元のページ番号。なければ位置）を持つ。戻り値は slides の要素そのものである。
    範囲外の番号と、ない識別子は誤りにする（候補は近い識別子）。known は候補に加える識別子
    （資料のコードが選択の外として作らなかったページの識別子など）で、選択の対象にはしない。
    """
    items = parse(text)
    numbers = _numbers(slides)
    ids = [slide.get("id") for slide in slides]
    names = list(
        dict.fromkeys([str(identifier) for identifier in ids if identifier is not None] + sorted(map(str, known)))
    )
    chosen: set[int] = set()
    for item in items:
        if item[0] == "id":
            matches = [index for index, identifier in enumerate(ids) if identifier == item[1]]
            if not matches:
                candidates = difflib.get_close_matches(item[1], names)
                hint = f"（候補: {'、'.join(candidates)}）" if candidates else ""
                raise SelectionError(f"ページの識別子がありません: {item[1]}{hint}")
            chosen.update(matches)
            continue
        start, end = item[1], item[2]
        for value in (start, end):
            if not numbers or not min(numbers) <= value <= max(numbers):
                raise SelectionError(f"ページ番号が範囲外です: {value}（{_bounds(numbers)}）")
        matches = [index for index, number in enumerate(numbers) if start <= number <= end]
        if not matches:
            # 部分の資料など、番号が飛んでいる一覧で、指定の番号のページがない場合。
            raise SelectionError(
                f"指定した番号のページがありません: {format_numbers(range(start, end + 1))}（ある番号: {format_numbers(numbers)}）"
            )
        chosen.update(matches)
    return [slide for index, slide in enumerate(slides) if index in chosen]


def describe(selected: Sequence[dict], text: str, total: int) -> dict:
    """生成記録と結果に書く、部分生成の指定の記録。"""
    return {
        "spec": text,
        "slide_ids": [slide.get("id") for slide in selected],
        "numbers": [slide.get("number") for slide in selected],
        "total_slides": total,
    }


def format_numbers(numbers: Iterable[int]) -> str:
    """ページ番号の列を、指定の書式の範囲表記にする。[3, 4, 5, 8] は "3-5,8" になる。"""
    ordered = sorted(set(numbers))
    parts = []
    index = 0
    while index < len(ordered):
        end = index
        while end + 1 < len(ordered) and ordered[end + 1] == ordered[end] + 1:
            end += 1
        parts.append(str(ordered[index]) if end == index else f"{ordered[index]}-{ordered[end]}")
        index = end + 1
    return ",".join(parts)
