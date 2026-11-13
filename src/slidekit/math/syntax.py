r"""数式の記法の正規化・字句解析・構文解析。

出力（`office_math`）と寸法の計算は、どちらも `parse()` の構文木を使う。受け付ける記法は必ず一致する。
命令の表は `commands` にある。このパッケージは slidekit のほかのモジュール（`errors` など）に依存しない。

字句は ``\\[A-Za-z]+|\\.|[A-Za-z]+|\d+(?:\.\d+)?|\s+|.`` で分ける。主な規則は次のとおりである。

- 波括弧は見えないまとまり。波括弧のない添字は次の 1 文字（``x_12`` は ``x_{1}2``）。
- ``/`` は直前と直後のまとまりの横書き分数。プライムは上付きに加える。
- 普通の括弧は通常の文字で、対応だけを確かめる。``(``・``[`` は ``)``・``]`` のどちらでも閉じることができる。
  ``\left``・``\right`` の括弧だけが伸びる。縦線は組にしない。
- 大型演算子の引数は次のまとまりから、同じ段の次の二項演算・関係・句読点の前まで（なければ群の終わりまで）。
- ``\`` のない語は、関数名と ``sqrt`` だけを特別に扱い、ほかは斜体の文字の並びにする。
- 東アジアの文字は ``\text{...}`` の中にだけ書くことができる。
"""

from __future__ import annotations

from dataclasses import dataclass
import difflib
from functools import lru_cache
import re
from typing import Iterable, Iterator, Union
import unicodedata

from . import commands


class MathSyntaxError(ValueError):
    """数式の記法の誤り。

    `position` は元の式の 0 始まりの文字位置（決まらなければ None）、`token` は該当する字句、
    `reason` は理由、`hint` は助言（近い命令名など）である。文言は「{理由}: {字句}（{n} 文字目）。{助言}」。
    """

    code = "math_syntax_error"

    def __init__(
        self,
        reason: str,
        *,
        expression: str = "",
        position: int | None = None,
        token: str | None = None,
        hint: str | None = None,
    ) -> None:
        self.reason = reason
        self.expression = expression
        self.position = position
        self.token = token
        self.hint = hint
        message = reason
        if token is not None:
            message += ": " + token
        if position is not None:
            message += f"（{position + 1} 文字目）"
        if hint:
            message += "。" + hint
        super().__init__(message)


@dataclass(frozen=True)
class Run:
    """数式の文字。`style` は italic・upright・bold・bold-italic、`variant` は Office Math の m:scr の値。"""

    text: str
    style: str = "italic"
    variant: str | None = None
    kind: str = "ord"


@dataclass(frozen=True)
class Text:
    """\\text の文章。要素の文字の書体で書き出す。"""

    text: str


@dataclass(frozen=True)
class Space:
    """空白。`char` が空なら何も出力しない。"""

    char: str
    em: float


@dataclass(frozen=True)
class Fenced:
    """括弧で囲んだまとまり。`stretchy` が真なら伸びる括弧（\\left・\\right）、偽なら通常の文字。

    `parts` は \\middle で分けた中身の列（普通の括弧では 1 つ）、`separator` は \\middle の記号。
    """

    open: str
    close: str
    parts: tuple[tuple[Node, ...], ...]
    stretchy: bool = False
    separator: str | None = None


@dataclass(frozen=True)
class Fraction:
    numerator: tuple[Node, ...]
    denominator: tuple[Node, ...]
    kind: str = "bar"  # bar・lin（横書き）・noBar（線なし）


@dataclass(frozen=True)
class Radical:
    degree: tuple[Node, ...]
    body: tuple[Node, ...]


@dataclass(frozen=True)
class Scripts:
    base: tuple[Node, ...]
    sub: tuple[Node, ...] | None = None
    sup: tuple[Node, ...] | None = None


@dataclass(frozen=True)
class Limits:
    """上下に置くもの（\\overset・\\underset、\\limits を付けた関数名の添字）。"""

    base: tuple[Node, ...]
    over: tuple[Node, ...] | None = None
    under: tuple[Node, ...] | None = None


@dataclass(frozen=True)
class Nary:
    """大型演算子。`limits` は subSup（添字を右）か undOvr（添字を上下）。"""

    char: str
    sub: tuple[Node, ...] | None
    sup: tuple[Node, ...] | None
    body: tuple[Node, ...]
    limits: str = "subSup"


@dataclass(frozen=True)
class Function:
    """関数名。`name` は名前の文字（添字を含む）、`body` は引数（なければ空）。"""

    name: tuple[Node, ...]
    body: tuple[Node, ...]


@dataclass(frozen=True)
class Accent:
    char: str
    body: tuple[Node, ...]


@dataclass(frozen=True)
class Bar:
    body: tuple[Node, ...]
    position: str = "top"  # top・bot


Node = Union[Run, Text, Space, Fenced, Fraction, Radical, Scripts, Limits, Nary, Function, Accent, Bar]

# \text の中の - を書き出すハイフン（U+2010）。PowerPoint は数式の中の - を文章でもマイナス記号で描く。
TEXT_HYPHEN = "\u2010"
SUBS = dict(zip("₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ", "0123456789+-=()aehijklmnoprstuvx"))
SUPS = dict(zip("⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱ", "0123456789+-=()ni"))
TOKEN = re.compile(r"\\[A-Za-z]+|\\.|[A-Za-z]+|\d+(?:\.\d+)?|\s+|.")
# 正規化で \text{...} の中身と命令の字句を飛ばすための走査
_SCAN = re.compile(r"\\text(?![A-Za-z])\s*\{|\\[A-Za-z]+|\\.", re.DOTALL)

_PRIMES = {"'": "′", "′": "′", "″": "″", "‴": "‴", "⁗": "⁗"}
_SCRIPT_TOKENS = frozenset({"_", "^", *_PRIMES})
_CLOSERS = frozenset().union(*commands.BRACKETS.values()) | {"}", "\\right", "_", "^"}
# 関数名の後にこの字句が続くときは、引数を取らない
_FUNCTION_STOPS = frozenset({"", "=", "+", "-", ",", ";", "\\middle"}) | (_CLOSERS - {"_", "^"})
_OPERATORS = {name: symbol.char for name, symbol in commands.SYMBOLS.items() if symbol.group == commands.OPERATOR}
_CHARACTER_KINDS: dict[str, str] = {}
for _symbol in (*commands.CHARACTERS.values(), *commands.SYMBOLS.values()):
    _CHARACTER_KINDS.setdefault(_symbol.char, _symbol.kind)
# 大型演算子の引数を終える記号の種類
_OPERAND_STOPS = frozenset({"bin", "rel", "punct"})


def normalise(expression: str) -> tuple[str, tuple[int, ...]]:
    """Unicode の添字を _{...}・^{...} に、− を - にそろえる。2 つ目の値は正規化後の各文字に対応する元の式の位置。

    前後の空白は除く。\\text{...} の中身と命令の字句は変えない。
    """
    if not isinstance(expression, str):
        raise TypeError("数式は文字列で指定する")
    index = len(expression) - len(expression.lstrip())
    end = len(expression.rstrip())
    characters: list[str] = []
    positions: list[int] = []
    while index < end:
        match = _SCAN.match(expression, index, end)
        if match:
            stop = match.end()
            if match.group().endswith("{") and match.group().startswith("\\text"):
                close = expression.find("}", stop, end)
                stop = end if close < 0 else close + 1
            characters.extend(expression[index:stop])
            positions.extend(range(index, stop))
            index = stop
            continue
        character = expression[index]
        mapping, marker = (SUBS, "_") if character in SUBS else (SUPS, "^") if character in SUPS else (None, "")
        if mapping is None:
            characters.append("-" if character == "−" else character)
            positions.append(index)
            index += 1
            continue
        stop = index
        while stop < end and expression[stop] in mapping:
            stop += 1
        characters += [marker, "{", *(mapping[c] for c in expression[index:stop]), "}"]
        positions += [index, index, *range(index, stop), stop - 1]
        index = stop
    return "".join(characters), tuple(positions)


def function_names(functions: Iterable[str] = ()) -> frozenset[str]:
    """資料で追加した関数名を検証し、既定の関数名と合わせた不変な集合として返す。"""
    if isinstance(functions, str):
        raise ValueError("追加の関数名は文字列の列で指定する")
    values = tuple(functions)
    if any(not isinstance(name, str) or not re.fullmatch(r"[A-Za-z]+", name) for name in values):
        raise ValueError("追加の関数名は英字だけで指定する")
    return commands.FUNCTIONS.union(values)


def parse(expression: str, functions: Iterable[str] = ()) -> tuple[Node, ...]:
    """式を構文木にする。出力と計測は必ずこの結果を使う。(式, 関数名の集合) ごとに結果を保持する（最大 4096 件）。"""
    if not isinstance(expression, str):
        raise TypeError("数式は文字列で指定する")
    return _parse(expression, function_names(functions))


@lru_cache(maxsize=4096)
def _parse(expression: str, functions: frozenset[str]) -> tuple[Node, ...]:
    return _Parser(expression, functions).parse()


def children(node: Node) -> tuple[tuple[Node, ...], ...]:
    """構文木の節の子の列を、式の中の順に返す。"""
    if isinstance(node, Fenced):
        return node.parts
    if isinstance(node, Fraction):
        return node.numerator, node.denominator
    if isinstance(node, Radical):
        return node.degree, node.body
    if isinstance(node, Scripts):
        return node.base, node.sub or (), node.sup or ()
    if isinstance(node, Limits):
        return node.over or (), node.base, node.under or ()
    if isinstance(node, Nary):
        return node.sub or (), node.sup or (), node.body
    if isinstance(node, Function):
        return node.name, node.body
    if isinstance(node, (Accent, Bar)):
        return (node.body,)
    return ()


def walk(nodes: Iterable[Node]) -> Iterator[Node]:
    """構文木のすべての節を、式の中の順に返す。"""
    for node in nodes:
        yield node
        for group in children(node):
            yield from walk(group)


def text_contents(nodes: Iterable[Node]) -> list[str]:
    """\\text の中身を出現順に返す（字形の検査とフォントの検査に使う）。"""
    return [node.text for node in walk(nodes) if isinstance(node, Text)]


def _east_asian(text: str) -> bool:
    return any(unicodedata.east_asian_width(character) in ("W", "F", "H") for character in text)


def _letter_style(text: str) -> str:
    """直接書いた文字の既定の書式。英字と小文字のギリシャ文字は斜体、ほかは立体。"""
    for character in text:
        name = unicodedata.name(character, "")
        category = unicodedata.category(character)
        latin = category.startswith("L") and name.startswith("LATIN ")
        greek = category == "Ll" and name.startswith("GREEK ")
        if not (latin or greek):
            return "upright"
    return "italic"


def _flatten(items: list[tuple[Node, ...]]) -> tuple[Node, ...]:
    return tuple(node for item in items for node in item)


def _plain_text(nodes: Iterable[Node]) -> str:
    parts = []
    for node in walk(nodes):
        if isinstance(node, (Run, Text)):
            parts.append(node.text)
        elif isinstance(node, Space):
            parts.append(node.char)
    return "".join(parts)


class _Parser:
    def __init__(self, expression: str, functions: frozenset[str]) -> None:
        self.expression = expression
        self.functions = functions
        for position, character in enumerate(expression):
            if character in "$\n\r":
                self._raise("数式を囲む $ と改行は指定しない", position, "$" if character == "$" else None)
            if unicodedata.category(character) == "Cc" and character != "\t":
                self._raise("数式に制御文字は使えない", position, f"U+{ord(character):04X}")
        text, self.positions = normalise(expression)
        self.tokens = [(match.group(), match.start()) for match in TOKEN.finditer(text)]
        self.index = 0
        # 書体の命令による書式（None は文字ごとの既定）、m:scr の値、\boldsymbol の太字
        self.style: str | None = None
        self.variant: str | None = None
        self.bold = False

    def _raise(self, reason: str, position: int | None, token: str | None = None, hint: str | None = None):
        raise MathSyntaxError(reason, expression=self.expression, position=position, token=token, hint=hint)

    def fail(self, reason: str, index: int | None = None, hint: str | None = None):
        """字句の番号 index の位置で誤りにする。index がなければ、位置を示さない。"""
        if index is not None and 0 <= index < len(self.tokens):
            token, start = self.tokens[index]
            self._raise(reason, self.positions[start], token, hint)
        self._raise(reason, None, None, hint)

    def peek(self) -> str:
        return self.tokens[self.index][0] if self.index < len(self.tokens) else ""

    def pop(self) -> str:
        value = self.peek()
        self.index += bool(value)
        return value

    def skip_spaces(self) -> None:
        while self.peek() and self.peek().isspace():
            self.pop()

    def run(self, text: str, default: str, kind: str) -> Run:
        style = self.style or default
        if self.bold:
            style = {"italic": "bold-italic", "upright": "bold"}.get(style, style)
        return Run(text, style, self.variant, kind)

    def symbol(self, symbol: commands.Symbol) -> Run:
        return self.run(symbol.char, "italic" if symbol.italic else "upright", symbol.kind)

    def parse(self) -> tuple[Node, ...]:
        nodes, _ = self.sequence(frozenset())
        return nodes

    def sequence(self, stops: frozenset[str], opener: int | None = None) -> tuple[tuple[Node, ...], str | None]:
        """stops のどれかの字句まで読む。stops が空なら式の終わりまで読む。戻り値は中身と閉じた字句。"""
        items: list[tuple[Node, ...]] = []
        while True:
            token = self.peek()
            if not token:
                if stops:
                    self.fail("閉じ括弧がない", opener)
                return _flatten(items), None
            if token in stops:
                self.pop()
                return _flatten(items), token
            if token.isspace():
                self.pop()
                continue
            if token == "/":
                slash = self.index
                self.pop()
                if not items:
                    self.fail("分子のない分数", slash)
                numerator = items.pop()
                items.append((Fraction(numerator, self.scripted_atom(slash), "lin"),))
                continue
            items.append(self.scripted_atom())

    def group(self) -> tuple[Node, ...]:
        opener = self.index
        self.pop()
        nodes, _ = self.sequence(frozenset({"}"}), opener)
        return nodes

    def scripts(self) -> tuple[tuple[Node, ...] | None, tuple[Node, ...] | None]:
        sub = sup = None
        self.skip_spaces()
        while self.peek() in _SCRIPT_TOKENS:
            index = self.index
            kind = self.pop()
            if kind in _PRIMES:
                # プライムは上付きに置き、後続の下付きと一つの構造へまとめる
                sup = (*(sup or ()), self.run(_PRIMES[kind], "upright", "ord"))
                self.skip_spaces()
                continue
            value = self.script_argument(index)
            if kind == "_":
                if sub is not None:
                    self.fail("重複した下付き", index)
                sub = value
            else:
                if sup is not None:
                    self.fail("重複した上付き", index)
                sup = value
            self.skip_spaces()
        return sub, sup

    def script_argument(self, owner: int) -> tuple[Node, ...]:
        self.skip_spaces()
        if self.peek() == "{":
            return self.group()
        # 添字を波括弧で囲まない場合は LaTeX と同じく次の 1 文字とする
        token = self.peek()
        if token and re.fullmatch(r"[A-Za-z]+|\d+", token) and len(token) > 1:
            _, start = self.tokens[self.index]
            self.tokens[self.index : self.index + 1] = [(token[0], start), (token[1:], start + 1)]
        return self.atom(owner)

    def scripted_atom(self, owner: int | None = None) -> tuple[Node, ...]:
        self.skip_spaces()
        base = self.atom(owner)
        sub, sup = self.scripts()
        if sub is None and sup is None:
            return base
        return (Scripts(base, sub, sup),)

    def atom(self, owner: int | None = None) -> tuple[Node, ...]:
        self.skip_spaces()
        token = self.peek()
        index = self.index
        if not token:
            self.fail("引数のない数式", owner if owner is not None else index - 1)
        if token == "{":
            return self.group()
        if token in commands.BRACKETS:
            return self.bracket()
        if token in _CLOSERS:
            self.fail("対応する左側のない記号", index)
        self.pop()
        if token.startswith("\\") and len(token) > 1:
            return self.command(token, index)
        if token == "\\":
            self.fail("未対応の数式命令", index)
        if _east_asian(token):
            self.fail("日本語は \\text{...} の中に書く", index)
        if token in commands.SPACES:
            return (Space(*commands.SPACES[token]),)
        if token in ("sqrt", "√"):
            return self.radical(index, bare=True)
        if token in ("∑", "∏"):
            return self.nary(token, index)
        if token in self.functions:
            return self.function((self.run(token, "upright", "op"),))
        if token in commands.CHARACTERS:
            return (self.symbol(commands.CHARACTERS[token]),)
        if token[0].isdigit():
            return (self.run(token, "upright", "ord"),)
        return (self.run(token, _letter_style(token), _CHARACTER_KINDS.get(token, "ord")),)

    def command(self, token: str, index: int) -> tuple[Node, ...]:
        name = token[1:]
        if token in commands.SPACES:
            return (Space(*commands.SPACES[token]),)
        if not re.fullmatch(r"[A-Za-z]+", name):
            if name in commands.SYMBOLS:
                return (self.symbol(commands.SYMBOLS[name]),)
            self.fail("未対応の数式命令", index)
        if name == "left":
            return self.stretchy(index)
        if name == "middle":
            self.fail("\\middle は \\left と \\right の間で使う", index)
        if name in ("limits", "nolimits"):
            self.fail("\\limits は大型演算子か関数名の直後に置く", index)
        structure = _STRUCTURE_METHODS.get(name)
        if structure is not None:
            return structure(self, index)
        if name in commands.FONTS:
            return self.styled(index, *commands.FONTS[name])
        if name in commands.ACCENTS:
            return (Accent(commands.ACCENTS[name], self.script_argument(index)),)
        if name in _OPERATORS:
            return self.nary(_OPERATORS[name], index)
        if name in self.functions:
            return self.function((self.run(name, "upright", "op"),))
        if name in commands.SYMBOLS:
            return (self.symbol(commands.SYMBOLS[name]),)
        hint = commands.UNSUPPORTED.get(name)
        if hint is None:
            known = sorted(commands.command_names() | self.functions)
            candidates = difflib.get_close_matches(name, known, n=3, cutoff=0.6)
            hint = "近い命令: " + "、".join("\\" + candidate for candidate in candidates) if candidates else None
        self.fail("未対応の数式命令", index, hint)

    def bracket(self) -> tuple[Node, ...]:
        """普通の括弧。中身を読み、対応する閉じ括弧を確かめる。括弧は通常の文字として出力する。"""
        opener = self.index
        token = self.pop()
        nodes, closing = self.sequence(commands.BRACKETS[token], opener)
        return (Fenced(commands.DELIMITERS[token], commands.DELIMITERS[closing], (nodes,)),)

    def delimiter(self, owner: int, reason: str, *, empty: bool = True) -> str:
        self.skip_spaces()
        index = self.index
        token = self.pop()
        if token not in commands.DELIMITERS or (not empty and token == "."):
            self.fail(reason, index if token else owner)
        return commands.DELIMITERS[token]

    def stretchy(self, index: int) -> tuple[Node, ...]:
        """\\left … \\middle … \\right。"""
        opening = self.delimiter(index, "伸縮括弧の始まりが不正")
        parts: list[tuple[Node, ...]] = []
        separator = None
        while True:
            nodes, closing = self.sequence(frozenset({"\\right", "\\middle"}), index)
            parts.append(nodes)
            closing_index = self.index - 1
            if closing == "\\right":
                break
            character = self.delimiter(closing_index, "伸縮括弧の区切りが不正", empty=False)
            if separator is not None and character != separator:
                self.fail("\\middle の記号は 1 種類にする", closing_index)
            separator = character
        closing_character = self.delimiter(closing_index, "伸縮括弧の終わりが不正")
        return (Fenced(opening, closing_character, tuple(parts), stretchy=True, separator=separator),)

    def nary(self, character: str, index: int) -> tuple[Node, ...]:
        self.skip_spaces()
        limits = "subSup"
        if self.peek() in ("\\limits", "\\nolimits"):
            limits = "undOvr" if self.pop() == "\\limits" else "subSup"
        sub, sup = self.scripts()
        return (Nary(character, sub, sup, self.operand(index), limits),)

    def operand(self, owner: int) -> tuple[Node, ...]:
        """大型演算子の引数。次のまとまりから、同じ段の次の二項演算・関係・句読点の前まで（なければ群の終わりまで）。

        PowerPoint は引数（m:e）と後ろの原子の間に細い空きを入れる（\\int f(x)\\,dx が「f (x)」になる）ので、
        続く原子も引数に入れる。二項演算・関係・句読点の前で終えても、その空きは変わらない。
        """
        items = [self.scripted_atom(owner)]
        while True:
            self.skip_spaces()
            token = self.peek()
            if not token or token in _CLOSERS or token == "\\middle" or self.token_kind(self.index) in _OPERAND_STOPS:
                return _flatten(items)
            if token == "/":
                slash = self.index
                self.pop()
                numerator = items.pop()
                items.append((Fraction(numerator, self.scripted_atom(slash), "lin"),))
                continue
            items.append(self.scripted_atom())

    def token_kind(self, position: int) -> str | None:
        """字句の番号 position が記号なら、その間隔の種類。\\not は打ち消す記号の種類。記号でなければ None。"""
        token = self.tokens[position][0] if position < len(self.tokens) else ""
        if token == "\\not":
            position += 1
            while position < len(self.tokens) and self.tokens[position][0].isspace():
                position += 1
            token = self.tokens[position][0] if position < len(self.tokens) else ""
        if token in commands.CHARACTERS:
            return commands.CHARACTERS[token].kind
        if token.startswith("\\"):
            symbol = commands.SYMBOLS.get(token[1:])
            return None if symbol is None or token[1:] in self.functions else symbol.kind
        return _CHARACTER_KINDS.get(token)

    def function(self, name: tuple[Node, ...], limits: bool = False) -> tuple[Node, ...]:
        self.skip_spaces()
        if self.peek() in ("\\limits", "\\nolimits"):
            limits = self.pop() == "\\limits"
        sub, sup = self.scripts()
        if sub is not None or sup is not None:
            name = (Limits(name, over=sup, under=sub),) if limits else (Scripts(name, sub, sup),)
        self.skip_spaces()
        if self.peek() in _FUNCTION_STOPS:
            return (Function(name, ()),)
        return (Function(name, self.scripted_atom()),)

    def radical(self, index: int, bare: bool = False) -> tuple[Node, ...]:
        self.skip_spaces()
        degree: tuple[Node, ...] = ()
        if self.peek() == "[" and not bare:
            opener = self.index
            self.pop()
            degree, _ = self.sequence(frozenset({"]"}), opener)
            self.skip_spaces()
        if self.peek() == "{":
            body = self.group()
        elif self.peek() == "(":
            # sqrt(x) の丸括弧は根号の中身を表し、出力しない
            opener = self.index
            self.pop()
            body, _ = self.sequence(frozenset({")"}), opener)
        else:
            body = self.scripted_atom(index)
        return (Radical(degree, body),)

    def styled(self, index: int, style: str | None, variant: str | None, *, bold: bool = False) -> tuple[Node, ...]:
        """書体の命令の引数を、その書体で読む。数式用の書体命令は \\mathcal U のように次の 1 文字も引数に取る。"""
        saved = self.style, self.variant, self.bold
        if style is not None:
            self.style, self.variant = style, variant
        self.bold = self.bold or bold
        try:
            return self.script_argument(index)
        finally:
            self.style, self.variant, self.bold = saved

    # 構造の命令（_STRUCTURE_METHODS から呼ぶ）

    def _sqrt(self, index: int) -> tuple[Node, ...]:
        return self.radical(index)

    def _frac(self, index: int, kind: str = "bar") -> tuple[Node, ...]:
        numerator = self.script_argument(index)
        return (Fraction(numerator, self.script_argument(index), kind),)

    def _binom(self, index: int) -> tuple[Node, ...]:
        fraction = self._frac(index, "noBar")
        return (Fenced("(", ")", (fraction,), stretchy=True),)

    def _text(self, index: int) -> tuple[Node, ...]:
        # 文章中の空白と記号を、通常の数式の演算子として解釈しない
        self.skip_spaces()
        if self.peek() != "{":
            self.fail("数式内の文章は波括弧で囲む", index)
        opener = self.index
        self.pop()
        parts = []
        while self.peek() != "}":
            if not self.peek():
                self.fail("閉じ括弧がない", opener)
            if self.peek() == "{" or self.peek().startswith("\\"):
                self.fail("数式内の文章に入れ子や命令は使えない", self.index)
            parts.append(self.pop())
        self.pop()
        # PowerPoint は数式の中の - を文章（m:nor）でもマイナス記号で描くので、ハイフン（U+2010）にする。
        return (Text("".join(parts).replace("-", TEXT_HYPHEN)),)

    def _boldsymbol(self, index: int) -> tuple[Node, ...]:
        return self.styled(index, None, None, bold=True)

    def _operatorname(self, index: int) -> tuple[Node, ...]:
        self.skip_spaces()
        limits = self.peek() == "*"
        if limits:
            self.pop()
        name = _plain_text(self.styled(index, "upright", None))
        return self.function((self.run(name, "upright", "op"),), limits)

    def _bar(self, index: int, position: str = "top") -> tuple[Node, ...]:
        return (Bar(self.script_argument(index), position),)

    def _underline(self, index: int) -> tuple[Node, ...]:
        return self._bar(index, "bot")

    def _overset(self, index: int) -> tuple[Node, ...]:
        over = self.script_argument(index)
        return (Limits(self.script_argument(index), over=over),)

    def _underset(self, index: int) -> tuple[Node, ...]:
        under = self.script_argument(index)
        return (Limits(self.script_argument(index), under=under),)

    def _not(self, index: int) -> tuple[Node, ...]:
        self.skip_spaces()
        target = self.index
        token = self.pop()
        if token not in commands.NEGATIONS:
            self.fail("\\not の後に打ち消せない記号がある", target if token else index)
        symbol = commands.CHARACTERS[token] if token in commands.CHARACTERS else commands.SYMBOLS[token[1:]]
        return (self.run(commands.NEGATIONS[token], "upright", symbol.kind),)

    def _bmod(self, index: int) -> tuple[Node, ...]:
        # 前後に \: の空白（U+205F）を置く。空白がないと PowerPoint は「amodb」と詰めて描く。
        # 空白は出力と計測の両方に入るので、mod は前後の空きを加えない通常の文字にする（二項演算の空きと同じ幅になる）
        space = Space(*commands.SPACES["\\:"])
        return space, self.run("mod", "upright", "ord"), space

    def _pmod(self, index: int) -> tuple[Node, ...]:
        argument = self.script_argument(index)
        content = (self.run("mod", "upright", "ord"), Space("\u2004", 1 / 3), *argument)
        return Space("\u2003", 1.0), Fenced("(", ")", (content,))


# 構造の命令の処理。書体（commands.FONTS）と記号付き（commands.ACCENTS）の命令は表で処理する。
_STRUCTURE_METHODS = {
    "sqrt": _Parser._sqrt,
    "frac": _Parser._frac,
    "dfrac": _Parser._frac,
    "tfrac": _Parser._frac,
    "binom": _Parser._binom,
    "dbinom": _Parser._binom,
    "tbinom": _Parser._binom,
    "text": _Parser._text,
    "boldsymbol": _Parser._boldsymbol,
    "bm": _Parser._boldsymbol,
    "operatorname": _Parser._operatorname,
    "bar": _Parser._bar,
    "overline": _Parser._bar,
    "underline": _Parser._underline,
    "overset": _Parser._overset,
    "stackrel": _Parser._overset,
    "underset": _Parser._underset,
    "not": _Parser._not,
    "bmod": _Parser._bmod,
    "pmod": _Parser._pmod,
}
