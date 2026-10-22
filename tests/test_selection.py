"""ページの指定（--slides）の解釈と選択を確かめる。"""

import pytest

from slidekit.errors import InputError
from slidekit.selection import SelectionError, describe, format_numbers, parse, select


def slides(*ids):
    return [{"id": identifier, "number": number} for number, identifier in enumerate(ids, 1)]


def test_parse_ids_numbers_and_ranges():
    assert parse("intro, 3, 5-7, id:2024") == [("id", "intro"), ("range", 3, 3), ("range", 5, 7), ("id", "2024")]
    # 前後と範囲の記号のまわりの空白は無視する。空の項目も無視する。
    assert parse(" 5 - 7 ,, id: method ,") == [("range", 5, 7), ("id", "method")]
    # 数字と記号を含む識別子は、範囲の形でなければそのまま識別子として読む。
    assert parse("step-2, 3a, 方法") == [("id", "step-2"), ("id", "3a"), ("id", "方法")]


def test_numeric_ids_need_prefix():
    deck = slides("2024", "summary", "2025-2026")
    # 数字だけの項目と範囲の形は常にページ番号として読む。
    assert [slide["id"] for slide in select(deck, "1")] == ["2024"]
    with pytest.raises(ValueError, match=r"^ページ番号が範囲外です: 2024（1〜3）$"):
        select(deck, "2024")
    # 範囲は、範囲外の最初の端を示す。
    with pytest.raises(ValueError, match=r"^ページ番号が範囲外です: 2025（1〜3）$"):
        select(deck, "2025-2026")
    with pytest.raises(ValueError, match=r"^ページ番号が範囲外です: 7（1〜3）$"):
        select(deck, "2-7")
    assert [slide["id"] for slide in select(deck, "id:2024, id:2025-2026")] == ["2024", "2025-2026"]


def test_errors_name_candidates_and_range():
    deck = slides(*(f"page{number}" for number in range(1, 45)), "intro")
    for text in ("3-", "-3", "-", "id:"):
        with pytest.raises(ValueError, match=rf"^--slides の指定を解釈できません: '{text}'$"):
            parse(text)
    with pytest.raises(
        ValueError, match=r"^--slides の指定を解釈できません: '7-5'（範囲は 5-7 のように小さい番号から書く）$"
    ):
        parse("7-5")
    for text in ("", " , ,"):
        with pytest.raises(ValueError, match=r"^ページを 1 つ以上指定してください$"):
            select(deck, text)
    with pytest.raises(ValueError, match=r"^ページ番号が範囲外です: 50（1〜45）$"):
        select(deck, "50")
    with pytest.raises(ValueError, match=r"^ページ番号が範囲外です: 0（1〜45）$"):
        select(deck, "0-3")
    with pytest.raises(ValueError, match=r"^ページの識別子がありません: intro2（候補: intro）$"):
        select(deck, "intro2")
    # 近い識別子がなければ候補を示さない。
    with pytest.raises(ValueError, match=r"^ページの識別子がありません: まとめ$"):
        select(deck, "まとめ")
    # 資料のコードが作らなかったページの識別子（known）も候補にする。ページが 1 つもなくても示す。選択の対象にはしない。
    with pytest.raises(ValueError, match=r"^ページの識別子がありません: summery（候補: summary）$"):
        select([], "summery", known={"summary", "intro"})
    with pytest.raises(ValueError, match=r"^ページの識別子がありません: summary（候補: summary）$"):
        select(deck, "summary", known=["summary"])
    with pytest.raises(ValueError, match=r"^ページの識別子がありません: intro2（候補: intro）$"):
        select(deck, "intro2", known=["intro"])
    # 部分の資料のように番号が飛んでいる一覧では、ない番号を誤りにする。
    with pytest.raises(ValueError, match=r"^指定した番号のページがありません: 4（ある番号: 3,5）$"):
        select([{"id": "c", "number": 3}, {"id": "e", "number": 5}], "4")
    with pytest.raises(TypeError):
        parse(None)


def test_errors_are_input_errors_with_selection_code():
    # 指定の誤りは利用者の入力の誤り（InputError、ValueError の派生）で、診断の識別子は invalid_selection である。
    deck = slides("intro", "summary")
    for call in (
        lambda: parse("3-"),
        lambda: parse(","),
        lambda: select(deck, "9"),
        lambda: select(deck, "intr"),
        lambda: select([{"id": "c", "number": 3}], "2-2"),
    ):
        with pytest.raises(SelectionError) as raised:
            call()
        assert isinstance(raised.value, InputError) and isinstance(raised.value, ValueError)
        assert raised.value.code == "invalid_selection"


def test_order_follows_deck_and_removes_duplicates():
    deck = slides("cover", "intro", "method", "result", "summary")
    selected = select(deck, "summary, 2-3, id:intro, 3, 1")
    assert [slide["id"] for slide in selected] == ["cover", "intro", "method", "summary"]
    # 戻り値は元のページの辞書そのものである。
    assert all(any(slide is original for original in deck) for slide in selected)
    # 番号は元のページ番号を使う。
    partial = [{"id": "method", "number": 3}, {"id": "summary", "number": 5}]
    assert select(partial, "3-5") == partial
    assert describe(selected, "summary, 2-3", 5) == {
        "spec": "summary, 2-3",
        "slide_ids": ["cover", "intro", "method", "summary"],
        "numbers": [1, 2, 3, 5],
        "total_slides": 5,
    }
    assert format_numbers([8, 3, 4, 5, 4]) == "3-5,8"
    assert format_numbers([]) == ""
    assert parse(format_numbers([1, 3, 4, 5, 9])) == [("range", 1, 1), ("range", 3, 5), ("range", 9, 9)]
