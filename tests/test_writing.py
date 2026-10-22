"""記述規則の判定（数式の候補、明示改行の禁則、ノートの文末、禁止文字）を確かめる。"""

import ast
from pathlib import Path
import random
import re
import sys

import pytest

from slidekit import writing
from slidekit.writing import (
    candidate_text,
    forbidden_characters,
    line_break_problems,
    math_candidates,
    math_check_applies,
    notes_style_candidates,
)

COMMAND_NAMES = writing._command_names


@pytest.fixture(autouse=True)
def command_names(monkeypatch):
    # 数式の命令の一覧は数式の領域が持つ。ここでは固定の一覧に差し替える。
    monkeypatch.setattr(writing, "_command_names", lambda: frozenset({"alpha", "beta", "frac", "le", "sum", "to"}))


def text(value, **extra):
    return {"kind": "text", "text": value, **extra}


def math(value):
    return {"kind": "math", "text": value}


def previous_candidates(value, config):
    """0.1 の validators の判定（単独の英字、数字、指定した名前）。"""
    for pattern in config.get("exclude_patterns", []):
        value = re.sub(pattern, " ", value)
    found = set(re.findall(r"(?<![A-Za-z0-9_])[A-Za-z](?![A-Za-z0-9_])", value))
    if config.get("numbers", True):
        found.update(re.findall(r"(?<![A-Za-z0-9_])\d+(?:\.\d+)?(?![A-Za-z0-9_])", value))
    for name in config.get("variables", []) + config.get("operations", []):
        if name and re.search(r"(?<![A-Za-z0-9_])" + re.escape(name) + r"(?![A-Za-z0-9_])", value):
            found.add(name)
    return found


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("配列の長さ n0 を求める", ["n0"]),
        ("添字 x_1 と x_{ij} を使う", ["x_1", "x_{ij}"]),
        ("円周率 π の値", ["π"]),
        ("n² 回の比較", ["n²"]),
        ("要素 x₁ を選ぶ", ["x₁"]),
        ("ｎ 個の要素", ["ｎ"]),
        (r"\alpha の値", [r"\alpha"]),
    ],
)
def test_math_candidates_include_letter_digit_subscript_and_greek(value, expected):
    assert math_candidates(value, {}) == expected


def test_math_candidates_exclude_words_identifiers_and_paths():
    assert math_candidates("λόγος と size_pt を使う", {}) == []
    # ドライブ名は 0.1 と同じく単独の英字の候補である。\Users・\tools は数式の命令ではない（\to の続きに英字がある）。
    assert math_candidates(r"保存先は C:\Users\tools です", {}) == ["C"]


def test_math_candidates_keep_previous_detections():
    assert math_candidates("n 個の要素を O(n) 時間で 3 回調べる", {}) == ["3", "O", "n"]
    assert math_candidates("Access と insert を使う", {"variables": ["Access"], "operations": ["insert"]}) == [
        "Access",
        "insert",
    ]
    assert math_candidates("２n 個", {}) == ["n"]
    assert math_candidates("xα の値", {}) == ["x"]
    # 0.1 の判定で候補のある文字列は、新しい判定でも候補がある。
    alphabet = [
        *"abnxyzABOT0123456789_^{}.,()（）「」、。 　\n",
        *"ｎｘＡ０１２",
        *"αβπμ²₁",
        *"の値を個年",
        r"\alpha",
        "Access",
    ]
    config = {"numbers": True, "variables": ["Access"], "operations": []}
    generator = random.Random(1)
    for _ in range(3000):
        value = "".join(generator.choice(alphabet) for _ in range(generator.randint(1, 8)))
        assert math_candidates(value, config) or not previous_candidates(value, config), value


def test_math_candidates_skip_urls_doi_and_arxiv():
    assert math_candidates("https://example.com/x の n", {}) == ["n"]
    assert math_candidates("（https://example.com/a/b）の n", {}) == ["n"]
    assert math_candidates("www.example.org/p の m", {}) == ["m"]
    assert math_candidates("doi 10.1145/1234.5678 と x", {}) == ["x"]
    assert math_candidates("arXiv:2101.00001 の y", {}) == ["y"]
    # 年は既定では除外しない。
    assert math_candidates("2026 年", {}) == ["2026"]


def test_math_candidates_numbers_option_and_exclude_patterns():
    assert math_candidates("3 個の n", {"numbers": False}) == ["n"]
    assert math_candidates("2026 年の n 件", {"exclude_patterns": [r"\d{4}\s*年"]}) == ["n"]
    # 除外した部分は同じ長さの空白になり、後ろの候補の位置と表記が変わらない。
    assert math_candidates("Fig. 3 の ｘ", {"exclude_patterns": [r"Fig\. \d+"]}) == ["ｘ"]
    assert math_candidates(
        "2026年に n 個を Access する。", {"exclude_patterns": [r"2026年"], "operations": ["Access"]}
    ) == ["Access", "n"]


def test_candidate_text_skips_math_and_link_fragments():
    content = [
        text("長さ "),
        math("n"),
        text(" の列。詳細は "),
        text("Example Guide", link="https://example.com/"),
        {"kind": "emphasis", "text": "x"},
    ]
    assert candidate_text(content) == "長さ  の列。詳細は x"
    assert candidate_text("文字列の内容 n") == "文字列の内容 n"


def test_math_check_applies_to_roles_page_kinds_and_semantics():
    body = {"id": "body", "kind": "text", "role": "body"}
    content_page, references_page = {"id": "intro", "kind": "content"}, {"id": "refs", "kind": "references"}
    config = {"enabled": True, "exclude_roles": ["note"], "exclude_page_kinds": ["references"]}
    assert math_check_applies(body, content_page, config)
    assert not math_check_applies(body, content_page, {**config, "enabled": False})
    assert not math_check_applies(body, content_page, {})
    assert not math_check_applies({**body, "semantic": "footer"}, content_page, config)
    assert not math_check_applies({**body, "semantic": "decoration"}, content_page, config)
    assert not math_check_applies({**body, "role": "note"}, content_page, config)
    assert not math_check_applies(body, references_page, config)
    title = {"id": "_theme/title", "kind": "text", "role": "slide_title", "semantic": "title"}
    assert math_check_applies(title, content_page, config)


def test_line_break_treats_math_as_one_character():
    assert line_break_problems([text("前半\n"), math("x"), text("、後半")]) == []
    assert line_break_problems([text("前半\n"), math("a\nb"), text("、後半")]) == []
    assert line_break_problems([text("前半\n、後半")]) == [{"line": 2, "side": "start", "character": "、"}]
    # 空の数式は 0.1 と同じく除く。
    assert line_break_problems([text("前半\n"), math(""), text("、後半")]) == [
        {"line": 2, "side": "start", "character": "、"}
    ]


def test_line_break_detects_line_start_and_end_rules():
    assert line_break_problems("「引用\n」の後") == [{"line": 2, "side": "start", "character": "」"}]
    assert line_break_problems("本文「\n引用」") == [{"line": 1, "side": "end", "character": "「"}]
    assert line_break_problems([text("ブロ\nック")]) == [{"line": 2, "side": "start", "character": "ッ"}]
    assert line_break_problems([text("100\n%")]) == [{"line": 2, "side": "start", "character": "%"}]
    assert line_break_problems([text("一行目\n　、二行目「 \n三行目")]) == [
        {"line": 2, "side": "start", "character": "、"},
        {"line": 2, "side": "end", "character": "「"},
    ]


def test_line_break_allows_bullets_and_plain_breaks():
    assert line_break_problems([text("項目\n・箇条")]) == []
    assert line_break_problems([text("項目\n･箇条")]) == []
    assert line_break_problems([text("一行目\n二行目")]) == []
    assert line_break_problems([text("一行目「引用」\n")]) == []


def test_line_break_tables_are_the_kinsoku_base_with_explicit_differences():
    # 基準の表は 1 か所にあり、明示改行の検査は基準との差を定数で持つ。
    assert writing.LINE_START_ALLOWED <= writing.KINSOKU_LINE_START
    assert not writing.LINE_START_EXTRA & writing.KINSOKU_LINE_START
    assert not writing.LINE_END_EXTRA & writing.KINSOKU_LINE_END
    assert (
        writing.LINE_START_FORBIDDEN
        == (writing.KINSOKU_LINE_START - writing.LINE_START_ALLOWED) | writing.LINE_START_EXTRA
    )
    assert writing.LINE_END_FORBIDDEN == writing.KINSOKU_LINE_END | writing.LINE_END_EXTRA
    assert not writing.LINE_START_FORBIDDEN & writing.LINE_END_FORBIDDEN
    # 0.1 の判定の 6 文字を含む。
    assert {"、", "。", ",", ".", ")", "）"} <= writing.LINE_START_FORBIDDEN


def test_notes_style_detects_plain_endings():
    notes = "結果を求める。これは定義である。簡単ではない。以上が結果だ！\n進むのだ\n補足である（注）。結果を求める。"
    assert notes_style_candidates(notes) == [
        "結果を求める。",
        "これは定義である。",
        "簡単ではない。",
        "以上が結果だ！",
        "進むのだ",
        "補足である（注）。",
    ]


@pytest.mark.parametrize(
    "notes",
    [
        "定義を示します。",
        "できますか？",
        "そうですね。",
        "ご覧ください。",
        "出典：文献 第 7 章。",
        "作れます（出典：教科書）。",
        "「定義だ。」と述べます。",
        "O(n) です。",
        "手順を確かめること。",
        "鈴木ら。",
        "Sun ら。",
    ],
)
def test_notes_style_accepts_polite_noun_and_quoted_endings(notes):
    assert notes_style_candidates(notes) == []


def test_forbidden_characters_are_sorted_without_duplicates():
    assert forbidden_characters("記号「。」と「、」", ["、", "。", "！", "、"]) == ["、", "。"]
    assert forbidden_characters("文字", None) == []


def test_command_names_come_from_math_commands():
    commands = pytest.importorskip("slidekit.math.commands", reason="数式の命令の一覧（W1-5）がまだない")
    names = COMMAND_NAMES()
    assert names == {name for name in commands.command_names() if name.isascii() and name.isalpha()}
    assert {"alpha", "frac", "sqrt"} <= names


def test_writing_imports_only_standard_library_and_measure_tables():
    # 数式の命令の一覧は関数の中で読み込む。モジュールの最上位では標準ライブラリと measure の禁則の表だけを読み込む。
    tree = ast.parse(Path(writing.__file__).read_text(encoding="utf-8"))
    imports = [node for node in tree.body if isinstance(node, (ast.Import, ast.ImportFrom))]
    relative = [node for node in imports if isinstance(node, ast.ImportFrom) and node.level]
    assert [(node.level, node.module, sorted(alias.name for alias in node.names)) for node in relative] == [
        (1, "measure", ["NO_LINE_END", "NO_LINE_START"])
    ]
    modules = {alias.name.split(".")[0] for node in imports if isinstance(node, ast.Import) for alias in node.names}
    modules |= {node.module.split(".")[0] for node in imports if isinstance(node, ast.ImportFrom) and not node.level}
    assert modules - {"__future__"} <= sys.stdlib_module_names


def test_kinsoku_base_tables_are_shared_with_measure():
    # 禁則の基準の表は 1 か所（measure）に置き、明示改行の検査は基準との差だけを持つ。
    from slidekit import measure

    assert writing.KINSOKU_LINE_START is measure.NO_LINE_START
    assert writing.KINSOKU_LINE_END is measure.NO_LINE_END
