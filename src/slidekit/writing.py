"""記述規則の判定。数式として書くべき文字、明示改行の禁則、ノートの文末、禁止文字を調べる。

どの関数も入力を変更しない。依存は標準ライブラリと、禁則の基準の表を持つ measure だけで、
数式の命令の一覧は最初に使うときに読み込む。
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping, Sequence
import functools
import re
from typing import Any

# 日本語の禁則処理の基準の表（Office の日本語の「標準」の禁則文字）は、自動折り返しの計測と同じ measure の表を使う。
# PowerPoint の実際の一覧と違えば measure の表だけを直し、下の差の定数は変えない。
from .measure import NO_LINE_END as KINSOKU_LINE_END, NO_LINE_START as KINSOKU_LINE_START

# 明示改行の検査と基準の表の差。
# 中点は箇条書きの行頭に使うので、行頭に置いてよい。
LINE_START_ALLOWED = frozenset("・･")
# 基準の表にない閉じ括弧・句読点・繰り返し記号を加える（小書きの仮名と長音記号は基準の表にある）。
LINE_START_EXTRA = frozenset("〻〙〗〟｠»‼⁇⁈⁉")
# 基準の表にない開き括弧を加える。
LINE_END_EXTRA = frozenset("〘〖〝｟«")

LINE_START_FORBIDDEN = (KINSOKU_LINE_START - LINE_START_ALLOWED) | LINE_START_EXTRA
LINE_END_FORBIDDEN = KINSOKU_LINE_END | LINE_END_EXTRA

# 明示改行の検査で、数式の各行の代わりに置く 1 文字。
_MATH_CHARACTER = "￼"


def _fragments(content: str | Sequence[Mapping[str, Any]]) -> Sequence[Mapping[str, Any]]:
    if isinstance(content, str):
        return [{"kind": "text", "text": content}]
    return content


def forbidden_characters(text: str, characters: Iterable[str] | None) -> list[str]:
    """text に含まれる禁止文字を、重複なく並べ替えて返す。"""
    return sorted({character for character in characters or () if character in text})


def candidate_text(content: str | Sequence[Mapping[str, Any]]) -> str:
    """数式の候補を探す文字。数式の断片と、リンクを持つ断片（URL や題名）を除いてつなぐ。"""
    return "".join(
        str(fragment.get("text", ""))
        for fragment in _fragments(content)
        if fragment.get("kind") != "math" and not fragment.get("link")
    )


def math_check_applies(element: Mapping[str, Any], slide: Mapping[str, Any], config: Mapping[str, Any]) -> bool:
    """要素を数式の候補の検査の対象にするか。config は記述規則の math_candidates である。

    ページ番号（footer）とテーマの飾り（decoration）は検査しない。題名は検査する。
    """
    if not config.get("enabled", False):
        return False
    if element.get("semantic") in {"footer", "decoration"}:
        return False
    if (element.get("role") or "body") in (config.get("exclude_roles") or ()):
        return False
    return slide.get("kind") not in (config.get("exclude_page_kinds") or ())


# 数式の候補の正規表現。英字の候補は全角の英数字を半角に読み替えた写しに当てる。
_FULLWIDTH = str.maketrans(
    {chr(code): chr(code - 0xFEE0) for code in (*range(0xFF10, 0xFF1A), *range(0xFF21, 0xFF3B), *range(0xFF41, 0xFF5B))}
)
_GREEK = "Α-ΡΣ-Ωα-ωϑϕϖϰϱϵ"
_SCRIPTS = "²³¹⁰-₟"  # 上付き・下付きの数字と文字
SUFFIX = rf"(?:\d+|[{_SCRIPTS}]+|[_^](?:\{{[^{{}}\n]*\}}|[A-Za-z0-9{_GREEK}]+))?"
_LETTER = re.compile(rf"(?<![A-Za-z0-9_])[A-Za-z]{SUFFIX}(?![A-Za-z0-9_])")
_GREEK_LETTER = re.compile(rf"(?<![Ͱ-ϿA-Za-z])[{_GREEK}]{SUFFIX}(?![Ͱ-ϿA-Za-z])")
_SINGLE_LETTER = re.compile(r"(?<![A-Za-z0-9_])[A-Za-z](?![A-Za-z0-9_])")  # 0.1 の判定。元の文字列にも当てる
_NUMBER = re.compile(r"(?<![A-Za-z0-9_])\d+(?:\.\d+)?(?![A-Za-z0-9_])")
_BUILTIN_EXCLUSIONS = (
    re.compile(r"(?:https?://|www\.)[A-Za-z0-9\-._~:/?#\[\]@!$&'()*+,;=%]+"),  # URL。日本語の文字で止まる
    re.compile(r"(?<![A-Za-z0-9.])10\.\d{4,9}/[A-Za-z0-9\-._;()/:]+"),  # DOI
    re.compile(r"(?i:arXiv):\s?[A-Za-z0-9.\-/]+"),  # arXiv の番号
)


@functools.cache
def _command_names() -> frozenset[str]:
    """数式の変換が受け付ける英字の命令名。"""
    from .math.commands import command_names

    return frozenset(name for name in command_names() if name.isascii() and name.isalpha())


@functools.lru_cache(maxsize=8)
def _command_pattern(names: frozenset[str]) -> re.Pattern[str] | None:
    if not names:
        return None
    alternatives = "|".join(re.escape(name) for name in sorted(names, key=lambda name: (-len(name), name)))
    return re.compile(rf"\\(?:{alternatives})(?![A-Za-z])")


def _blank(text: str, spans: Iterable[tuple[int, int]]) -> str:
    characters = list(text)
    for start, end in spans:
        characters[start:end] = " " * (end - start)
    return "".join(characters)


def math_candidates(text: str, config: Mapping[str, Any]) -> list[str]:
    """数式として書くべき候補を、元の表記のまま重複なく並べ替えて返す。config は記述規則の math_candidates である。

    候補は、英字（数字・上付き・下付き・`_` や `^` の添字が続くものを含む）、ギリシャ文字、数式の命令、
    数字（numbers が真のとき）、variables・operations の名前である。全角の英数字も候補にする。
    exclude_patterns に一致した部分と、URL・DOI・arXiv の番号の中は候補にしない。ほかの候補に含まれる候補は返さない。
    enabled は参照しない（math_check_applies で判定する）。
    """
    work = text
    for pattern in config.get("exclude_patterns") or ():
        # 同じ長さの空白に置き換えて、位置を保つ。
        work = re.sub(pattern, lambda match: " " * len(match.group()), work)
    folded = work.translate(_FULLWIDTH)
    excluded = [match.span() for pattern in _BUILTIN_EXCLUSIONS for match in pattern.finditer(folded)]
    work, folded = _blank(work, excluded), _blank(folded, excluded)
    patterns = [(_LETTER, (folded,)), (_GREEK_LETTER, (folded,)), (_SINGLE_LETTER, (work,))]
    commands = _command_pattern(_command_names())
    if commands is not None:
        patterns.append((commands, (folded,)))
    if config.get("numbers", True):
        patterns.append((_NUMBER, (folded, work)))
    for name in [*(config.get("variables") or ()), *(config.get("operations") or ())]:
        if name:
            patterns.append((re.compile(r"(?<![A-Za-z0-9_])" + re.escape(name) + r"(?![A-Za-z0-9_])"), (folded, work)))
    spans = {match.span() for pattern, sources in patterns for source in sources for match in pattern.finditer(source)}
    kept = [
        span
        for span in spans
        if not any(other != span and other[0] <= span[0] and span[1] <= other[1] for other in spans)
    ]
    return sorted({text[start:end] for start, end in kept})


def line_break_problems(content: str | Sequence[Mapping[str, Any]]) -> list[dict]:
    """明示改行の直後の行頭と直前の行末にある禁則の文字を返す。

    数式の断片は各行を 1 文字とみなす（空の数式は除く）。戻り値は
    `{"line": 行番号, "side": "start" | "end", "character": 文字}` の列で、行番号は 1 から始まり、
    同じ行では start を先にする。
    自動折り返しは PowerPoint が禁則処理をするので、ここでは扱わない。
    """
    parts = []
    for fragment in _fragments(content):
        text = str(fragment.get("text", ""))
        if fragment.get("kind") != "math":
            parts.append(text)
        elif text.strip():
            parts.append("\n".join(_MATH_CHARACTER if line.strip() else "" for line in text.split("\n")))
    lines = "".join(parts).split("\n")
    problems = []
    for number, line in enumerate(lines, 1):
        start, end = line.lstrip(), line.rstrip()
        if number > 1 and start and start[0] in LINE_START_FORBIDDEN:
            problems.append({"line": number, "side": "start", "character": start[0]})
        if number < len(lines) and end and end[-1] in LINE_END_FORBIDDEN:
            problems.append({"line": number, "side": "end", "character": end[-1]})
    return problems


# ノートの文末の判定。
_POLITE_ENDING = re.compile(
    r"(?:です|ます|でした|ました|でしょう|ません|ましょう|ください|下さい|ませ)"
    r"(?:か|ね|よ|よね|が|けど|けれど|けれども|から|ので|し|かね)?$"
)
_HIRAGANA_NOUNS = ("こと", "もの", "とき", "ところ", "など", "ため", "まで", "だけ", "ほど", "ごと")
_SENTENCE_ENDS = frozenset("。！？!?")
_TRAILING_MARKS = frozenset("。．.！!？?…‥")
_BRACKETS = {
    "「": "」",
    "『": "』",
    "（": "）",
    "(": ")",
    "【": "】",
    "〔": "〕",
    "〈": "〉",
    "《": "》",
    "［": "］",
    "[": "]",
    "｛": "｝",
    "{": "}",
    "“": "”",
    "‘": "’",
    "〘": "〙",
    "〖": "〗",
    "〝": "〟",
    "｟": "｠",
    "«": "»",
}
_OPENING = {closing: opening for opening, closing in _BRACKETS.items()}


def _is_hiragana(character: str) -> bool:
    return "ぁ" <= character <= "ゟ"


def _sentences(notes: str) -> list[str]:
    """括弧の外の文末の記号と改行で文に分ける。括弧の対応は改行で打ち切る。"""
    sentences: list[str] = []
    current: list[str] = []
    closers: list[str] = []
    for character in notes:
        if character == "\n":
            sentences.append("".join(current))
            current, closers = [], []
            continue
        current.append(character)
        if character in _BRACKETS:
            closers.append(_BRACKETS[character])
        elif closers and character == closers[-1]:
            closers.pop()
        elif not closers and character in _SENTENCE_ENDS:
            sentences.append("".join(current))
            current = []
    sentences.append("".join(current))
    return [sentence.strip() for sentence in sentences if sentence.strip()]


def _strip_end(text: str) -> str:
    end = len(text)
    while end and (text[end - 1].isspace() or text[end - 1] in _TRAILING_MARKS):
        end -= 1
    return text[:end]


def _without_last_bracket(text: str) -> str:
    """末尾の閉じ括弧に対応する括弧書きを除く。対応する開き括弧がなければ閉じ括弧だけを除く。"""
    closing = text[-1]
    opening = _OPENING[closing]
    depth = 0
    for index in range(len(text) - 1, -1, -1):
        if text[index] == closing:
            depth += 1
        elif text[index] == opening:
            depth -= 1
            if depth == 0:
                return text[:index]
    return text[:-1]


def _sentence_end(sentence: str) -> str:
    """末尾の空白・句読点と、末尾の括弧書きを除いた文末。"""
    end = _strip_end(sentence)
    while end and end[-1] in _OPENING:
        end = _strip_end(_without_last_bracket(end))
    return end


def _plain_ending(end: str) -> bool:
    if not end or not _is_hiragana(end[-1]):
        return False
    if _POLITE_ENDING.search(end) or end.endswith(_HIRAGANA_NOUNS):
        return False
    # 「鈴木ら」「Sun ら」のように、ひらがな以外に続く「ら」は名詞として扱う。
    return not (end[-1] == "ら" and len(end) > 1 and not _is_hiragana(end[-2]))


def notes_style_candidates(notes: str) -> list[str]:
    """です・ます調でない文を、出現順に重複なく返す。

    文末（末尾の句読点と括弧書きを除いた部分）がひらがなで終わり、です・ます調の文末、
    ひらがなで書く名詞（こと・ものなど）、名前に続く「ら」のどれにも当たらない文を候補にする。
    ひらがなで終わらない文（名詞止めや出典の行）と、括弧の中の引用は候補にしない。
    """
    return list(
        dict.fromkeys(sentence for sentence in _sentences(str(notes)) if _plain_ending(_sentence_end(sentence)))
    )
