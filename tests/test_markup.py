import math

import pytest

from slidekit.errors import InputError
from slidekit.fragments import Emphasis, Math, Text, content_fragments
from slidekit.markup import parse


def test_parse_math_emphasis_and_escapes():
    assert parse("対象の件数を $n$ とし、**最大 $k$ 件**まで扱う") == [
        Text("対象の件数を "),
        Math("n"),
        Text(" とし、"),
        Emphasis("最大 "),
        Math("k", emphasis=True),
        Emphasis(" 件"),
        Text("まで扱う"),
    ]
    # \$ \* \\ だけをエスケープとして扱い、ほかの \ は残す。1 つの * は通常の文字。
    assert parse(r"価格は \$100、a\*b、\\ と C:\path、2 * 3") == [Text("価格は $100、a*b、\\ と C:\\path、2 * 3")]
    # 数式の中身はそのまま渡し、最初の $ で閉じる。
    assert parse(r"$\frac{a}{b}$ と $x_1$") == [Math(r"\frac{a}{b}"), Text(" と "), Math("x_1")]
    # 隣り合う同じ種類の文字の断片はつなげる。数式はつなげない（命令名が続く文字とつながらないように）。
    assert parse("**a****b**") == [Emphasis("ab")]
    assert parse(r"$\alpha$$b$") == [Math(r"\alpha"), Math("b")]
    assert parse("***a**") == [Emphasis("*a")]
    assert parse("") == []


def test_parse_emphasis_can_contain_math():
    assert parse("**重要な $x$ の値**") == [Emphasis("重要な "), Math("x", emphasis=True), Emphasis(" の値")]
    assert parse("**$x$**") == [Math("x", emphasis=True)]
    # 強調の中では改行してよい。
    assert parse("前置き\n**次の行\nまで強調**") == [Text("前置き\n"), Emphasis("次の行\nまで強調")]


@pytest.mark.parametrize(
    ("source", "message"),
    [
        ("a $x", "3 文字目の $ が閉じていません: 'a $x'"),
        ("**a", "1 文字目の ** が閉じていません: '**a'"),
        ("$$", "1 文字目の数式が空です: '$$'"),
        ("$ $", "1 文字目の数式が空です: '$ $'"),
        ("****", "1 文字目の強調が空です: '****'"),
        ("$a\nb$", "1 文字目の数式の中で改行しています: '$a\\nb$'"),
        ("**a** **b", "7 文字目の ** が閉じていません: '**a** **b'"),  # 3 つ目の ** が閉じていない
        ("前置きの長い文字列のあとで閉じない $x", "19 文字目の $ が閉じていません: …'列のあとで閉じない $x'"),
    ],
)
def test_parse_rejects_unclosed_or_empty_markers(source, message):
    with pytest.raises(ValueError) as caught:
        parse(source)
    assert str(caught.value) == message
    assert isinstance(caught.value, InputError) and caught.value.code == "invalid_markup"
    assert caught.value.details["position"] == int(message.split(" ", 1)[0])


def test_parse_options_disable_markers():
    assert parse("$x$ と **y**", math=False) == [Text("$x$ と "), Emphasis("y")]
    assert parse("$x$ と **y**", emphasis=False) == [Math("x"), Text(" と **y**")]
    # 記号を通常の文字にしても、エスケープは常に処理する。
    assert parse(r"\$ \* \\ $x$ **y**", math=False, emphasis=False) == [Text(r"$ * \ $x$ **y**")]
    with pytest.raises(InputError, match="文字列"):
        parse(["a"])
    with pytest.raises(InputError, match="math は真偽値"):
        parse("a", math="no")


def test_fragment_values_are_validated():
    given = Text("詳細", bold=True, color="primary", size_pt=14, underline=True, link="https://example.com/")
    assert given.size_pt == 14.0 and given.link == "https://example.com/"
    assert Math("x", size_pt=12.5, emphasis=True).size_pt == 12.5
    invalid = [
        (Text, {"text": 1}, "text は文字列"),
        (Text, {"text": "a", "bold": "yes"}, "bold は真偽値"),
        (Emphasis, {"text": "a", "underline": 1}, "underline は真偽値"),
        (Text, {"text": "a", "color": 0x183153}, "color は文字列"),
        (Emphasis, {"text": "a", "link": b"https://example.com/"}, "link は文字列"),
        (Math, {"text": "x", "emphasis": None}, "emphasis は真偽値"),
    ]
    invalid += [
        (kind, {"text": "a", "size_pt": size}, "size_pt は正の有限の数")
        for kind in (Text, Emphasis, Math)
        for size in (0, -1, True, "12", math.nan, math.inf)
    ]
    for kind, arguments, message in invalid:
        with pytest.raises(ValueError, match=message) as caught:
            kind(**arguments)
        assert isinstance(caught.value, InputError) and caught.value.code == "invalid_fragment"
    # 数式に下線とリンクはない。位置引数は text だけ。
    with pytest.raises(TypeError):
        Math("x", underline=True)
    with pytest.raises(TypeError):
        Text("a", True)


def test_content_fragments_keep_plain_form_and_add_given_items():
    # 書式とリンクを指定しない断片は、 0.1 と同じ kind と text だけの辞書になる。
    assert content_fragments(["a", Text("b"), Emphasis("c"), Math("x")]) == [
        {"kind": "text", "text": "a"},
        {"kind": "text", "text": "b"},
        {"kind": "emphasis", "text": "c"},
        {"kind": "math", "text": "x"},
    ]
    assert content_fragments(Text("a $x$")) == [{"kind": "text", "text": "a $x$"}]  # 記法は解釈しない
    assert content_fragments(None) == []
    given = content_fragments(
        [Text("詳細", size_pt=14, underline=True, link="https://example.com/guide"), Emphasis("注", bold=False)]
    )
    assert given == [
        {
            "kind": "text",
            "text": "詳細",
            "style": {"size_pt": 14.0, "underline": True},
            "link": "https://example.com/guide",
        },
        {"kind": "emphasis", "text": "注", "style": {"bold": False}},
    ]
    for value in ({"text": "a"}, 3, [Text("a"), 3]):
        with pytest.raises(InputError) as caught:
            content_fragments(value)
        assert caught.value.code == "invalid_fragment"
