r"""数式の命令の表と、対応する命令の一覧の生成。

構文解析（`syntax`）、Office Math の出力（`office_math`）、記述規則の数式の候補（`writing`）、
文書 `docs/math.md` の一覧が、この表を使う。表は命令ごとに 1 行で書き、一覧はこの表の順に並べる。

開発用のコマンド（資料の作成には使わない）::

    python -m slidekit.math.commands                       # 一覧を標準出力へ出す
    python -m slidekit.math.commands --write docs/math.md  # 文書の印の間を書き換える
    python -m slidekit.math.commands --check docs/math.md  # 文書の印の間が一覧と一致するかを確かめる

このモジュールは slidekit のほかのモジュールを読み込まない。
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
import re
import sys


@dataclass(frozen=True)
class Symbol:
    """記号を出す命令、または命令でない文字の出力。"""

    char: str  # 出力する文字
    kind: str  # 間隔の種類（ord・op・bin・rel・open・close・punct・inner）。計測（metrics）の原子の間の空きはこの値で決める
    group: str  # 一覧の分類
    italic: bool = False  # 既定で斜体にするか
    note: str = ""  # 一覧の備考


@dataclass(frozen=True)
class Structure:
    """構造を作る命令。解析は `syntax` が行う。"""

    syntax: str  # 一覧に載せる書き方（そのまま解析できる例）
    output: str  # 一覧に載せる出力の説明
    group: str
    note: str = ""


GREEK = "ギリシャ文字"
RELATION = "関係"
ARROW = "矢印"
BINARY = "二項演算"
OPERATOR = "大型演算子"
OTHER = "そのほかの記号"
BAR = "縦線"

# 間隔の種類の一覧での名前
KIND_NAMES = {
    "ord": "通常",
    "op": "大型演算子",
    "bin": "二項演算",
    "rel": "関係",
    "open": "開き",
    "close": "閉じ",
    "punct": "句読点",
    "inner": "内部",
}

# キーは \ を除いた命令名。大型演算子（分類 OPERATOR）は m:nary になり、後に引数が要る。
SYMBOLS: dict[str, Symbol] = {
    "alpha": Symbol("α", "ord", GREEK, italic=True),
    "beta": Symbol("β", "ord", GREEK, italic=True),
    "gamma": Symbol("γ", "ord", GREEK, italic=True),
    "delta": Symbol("δ", "ord", GREEK, italic=True),
    "epsilon": Symbol("ϵ", "ord", GREEK, italic=True),
    "varepsilon": Symbol("ε", "ord", GREEK, italic=True),
    "zeta": Symbol("ζ", "ord", GREEK, italic=True),
    "eta": Symbol("η", "ord", GREEK, italic=True),
    "theta": Symbol("θ", "ord", GREEK, italic=True),
    "vartheta": Symbol("ϑ", "ord", GREEK, italic=True),
    "iota": Symbol("ι", "ord", GREEK, italic=True),
    "kappa": Symbol("κ", "ord", GREEK, italic=True),
    "varkappa": Symbol("ϰ", "ord", GREEK, italic=True),
    "lambda": Symbol("λ", "ord", GREEK, italic=True),
    "mu": Symbol("μ", "ord", GREEK, italic=True),
    "nu": Symbol("ν", "ord", GREEK, italic=True),
    "xi": Symbol("ξ", "ord", GREEK, italic=True),
    "pi": Symbol("π", "ord", GREEK, italic=True),
    "varpi": Symbol("ϖ", "ord", GREEK, italic=True),
    "rho": Symbol("ρ", "ord", GREEK, italic=True),
    "varrho": Symbol("ϱ", "ord", GREEK, italic=True),
    "sigma": Symbol("σ", "ord", GREEK, italic=True),
    "varsigma": Symbol("ς", "ord", GREEK, italic=True),
    "tau": Symbol("τ", "ord", GREEK, italic=True),
    "upsilon": Symbol("υ", "ord", GREEK, italic=True),
    "phi": Symbol("ϕ", "ord", GREEK, italic=True),
    "varphi": Symbol("φ", "ord", GREEK, italic=True),
    "chi": Symbol("χ", "ord", GREEK, italic=True),
    "psi": Symbol("ψ", "ord", GREEK, italic=True),
    "omega": Symbol("ω", "ord", GREEK, italic=True),
    "Gamma": Symbol("Γ", "ord", GREEK),
    "Delta": Symbol("Δ", "ord", GREEK),
    "Theta": Symbol("Θ", "ord", GREEK),
    "Lambda": Symbol("Λ", "ord", GREEK),
    "Xi": Symbol("Ξ", "ord", GREEK),
    "Pi": Symbol("Π", "ord", GREEK),
    "Sigma": Symbol("Σ", "ord", GREEK),
    "Upsilon": Symbol("Υ", "ord", GREEK),
    "Phi": Symbol("Φ", "ord", GREEK),
    "Psi": Symbol("Ψ", "ord", GREEK),
    "Omega": Symbol("Ω", "ord", GREEK),
    "varGamma": Symbol("Γ", "ord", GREEK, italic=True),
    "varDelta": Symbol("Δ", "ord", GREEK, italic=True),
    "varTheta": Symbol("Θ", "ord", GREEK, italic=True),
    "varLambda": Symbol("Λ", "ord", GREEK, italic=True),
    "varXi": Symbol("Ξ", "ord", GREEK, italic=True),
    "varPi": Symbol("Π", "ord", GREEK, italic=True),
    "varSigma": Symbol("Σ", "ord", GREEK, italic=True),
    "varUpsilon": Symbol("Υ", "ord", GREEK, italic=True),
    "varPhi": Symbol("Φ", "ord", GREEK, italic=True),
    "varPsi": Symbol("Ψ", "ord", GREEK, italic=True),
    "varOmega": Symbol("Ω", "ord", GREEK, italic=True),
    "le": Symbol("≤", "rel", RELATION),
    "leq": Symbol("≤", "rel", RELATION),
    "ge": Symbol("≥", "rel", RELATION),
    "geq": Symbol("≥", "rel", RELATION),
    "ne": Symbol("≠", "rel", RELATION),
    "neq": Symbol("≠", "rel", RELATION),
    "lt": Symbol("<", "rel", RELATION),
    "gt": Symbol(">", "rel", RELATION),
    "leqq": Symbol("≦", "rel", RELATION),
    "geqq": Symbol("≧", "rel", RELATION),
    "leqslant": Symbol("⩽", "rel", RELATION),
    "geqslant": Symbol("⩾", "rel", RELATION),
    "ll": Symbol("≪", "rel", RELATION),
    "gg": Symbol("≫", "rel", RELATION),
    "prec": Symbol("≺", "rel", RELATION),
    "succ": Symbol("≻", "rel", RELATION),
    "preceq": Symbol("⪯", "rel", RELATION),
    "succeq": Symbol("⪰", "rel", RELATION),
    "equiv": Symbol("≡", "rel", RELATION),
    "approx": Symbol("≈", "rel", RELATION),
    "sim": Symbol("∼", "rel", RELATION),
    "simeq": Symbol("≃", "rel", RELATION),
    "cong": Symbol("≅", "rel", RELATION),
    "propto": Symbol("∝", "rel", RELATION),
    "asymp": Symbol("≍", "rel", RELATION),
    "doteq": Symbol("≐", "rel", RELATION),
    "coloneqq": Symbol("≔", "rel", RELATION),
    "triangleq": Symbol("≜", "rel", RELATION),
    "in": Symbol("∈", "rel", RELATION),
    "notin": Symbol("∉", "rel", RELATION),
    "ni": Symbol("∋", "rel", RELATION),
    "subset": Symbol("⊂", "rel", RELATION),
    "supset": Symbol("⊃", "rel", RELATION),
    "subseteq": Symbol("⊆", "rel", RELATION),
    "supseteq": Symbol("⊇", "rel", RELATION),
    "subsetneq": Symbol("⊊", "rel", RELATION),
    "supsetneq": Symbol("⊋", "rel", RELATION),
    "sqsubseteq": Symbol("⊑", "rel", RELATION),
    "sqsupseteq": Symbol("⊒", "rel", RELATION),
    "mid": Symbol("∣", "rel", RELATION, note="関係の縦線（前後に空きが入る）"),
    "nmid": Symbol("∤", "rel", RELATION),
    "parallel": Symbol("∥", "rel", RELATION),
    "nparallel": Symbol("∦", "rel", RELATION),
    "perp": Symbol("⊥", "rel", RELATION),
    "vdash": Symbol("⊢", "rel", RELATION),
    "dashv": Symbol("⊣", "rel", RELATION),
    "models": Symbol("⊨", "rel", RELATION),
    "to": Symbol("→", "rel", ARROW),
    "rightarrow": Symbol("→", "rel", ARROW),
    "leftarrow": Symbol("←", "rel", ARROW),
    "gets": Symbol("←", "rel", ARROW),
    "leftrightarrow": Symbol("↔", "rel", ARROW),
    "Rightarrow": Symbol("⇒", "rel", ARROW),
    "Leftarrow": Symbol("⇐", "rel", ARROW),
    "Leftrightarrow": Symbol("⇔", "rel", ARROW),
    "longrightarrow": Symbol("⟶", "rel", ARROW),
    "longleftarrow": Symbol("⟵", "rel", ARROW),
    "longleftrightarrow": Symbol("⟷", "rel", ARROW),
    "Longrightarrow": Symbol("⟹", "rel", ARROW),
    "Longleftarrow": Symbol("⟸", "rel", ARROW),
    "Longleftrightarrow": Symbol("⟺", "rel", ARROW),
    "implies": Symbol("⟹", "rel", ARROW),
    "impliedby": Symbol("⟸", "rel", ARROW),
    "iff": Symbol("⟺", "rel", ARROW),
    "mapsto": Symbol("↦", "rel", ARROW),
    "longmapsto": Symbol("⟼", "rel", ARROW),
    "hookrightarrow": Symbol("↪", "rel", ARROW),
    "uparrow": Symbol("↑", "rel", ARROW),
    "downarrow": Symbol("↓", "rel", ARROW),
    "Uparrow": Symbol("⇑", "rel", ARROW),
    "Downarrow": Symbol("⇓", "rel", ARROW),
    "nearrow": Symbol("↗", "rel", ARROW),
    "searrow": Symbol("↘", "rel", ARROW),
    "pm": Symbol("±", "bin", BINARY),
    "mp": Symbol("∓", "bin", BINARY),
    "times": Symbol("×", "bin", BINARY),
    "div": Symbol("÷", "bin", BINARY),
    "cdot": Symbol("⋅", "bin", BINARY),
    "ast": Symbol("∗", "bin", BINARY),
    "star": Symbol("⋆", "bin", BINARY),
    "circ": Symbol("∘", "bin", BINARY),
    "bullet": Symbol("∙", "bin", BINARY),
    "diamond": Symbol("⋄", "bin", BINARY),
    "setminus": Symbol("∖", "bin", BINARY),
    "cup": Symbol("∪", "bin", BINARY),
    "cap": Symbol("∩", "bin", BINARY),
    "sqcup": Symbol("⊔", "bin", BINARY),
    "sqcap": Symbol("⊓", "bin", BINARY),
    "uplus": Symbol("⊎", "bin", BINARY),
    "wedge": Symbol("∧", "bin", BINARY),
    "land": Symbol("∧", "bin", BINARY),
    "vee": Symbol("∨", "bin", BINARY),
    "lor": Symbol("∨", "bin", BINARY),
    "oplus": Symbol("⊕", "bin", BINARY),
    "ominus": Symbol("⊖", "bin", BINARY),
    "otimes": Symbol("⊗", "bin", BINARY),
    "odot": Symbol("⊙", "bin", BINARY),
    "sum": Symbol("∑", "op", OPERATOR),
    "prod": Symbol("∏", "op", OPERATOR),
    "coprod": Symbol("∐", "op", OPERATOR),
    "int": Symbol("∫", "op", OPERATOR),
    "iint": Symbol("∬", "op", OPERATOR),
    "iiint": Symbol("∭", "op", OPERATOR),
    "oint": Symbol("∮", "op", OPERATOR),
    "bigcup": Symbol("⋃", "op", OPERATOR),
    "bigcap": Symbol("⋂", "op", OPERATOR),
    "bigvee": Symbol("⋁", "op", OPERATOR),
    "bigwedge": Symbol("⋀", "op", OPERATOR),
    "bigoplus": Symbol("⨁", "op", OPERATOR),
    "bigotimes": Symbol("⨂", "op", OPERATOR),
    "bigodot": Symbol("⨀", "op", OPERATOR),
    "bigsqcup": Symbol("⨆", "op", OPERATOR),
    "infty": Symbol("∞", "ord", OTHER),
    "partial": Symbol("∂", "ord", OTHER),
    "nabla": Symbol("∇", "ord", OTHER),
    "emptyset": Symbol("∅", "ord", OTHER),
    "varnothing": Symbol("∅", "ord", OTHER),
    "forall": Symbol("∀", "ord", OTHER),
    "exists": Symbol("∃", "ord", OTHER),
    "nexists": Symbol("∄", "ord", OTHER),
    "neg": Symbol("¬", "ord", OTHER),
    "lnot": Symbol("¬", "ord", OTHER),
    "complement": Symbol("∁", "ord", OTHER),
    "ell": Symbol("ℓ", "ord", OTHER),
    "hbar": Symbol("ℏ", "ord", OTHER),
    "aleph": Symbol("ℵ", "ord", OTHER),
    "Re": Symbol("ℜ", "ord", OTHER),
    "Im": Symbol("ℑ", "ord", OTHER),
    "wp": Symbol("℘", "ord", OTHER),
    "top": Symbol("⊤", "ord", OTHER),
    "bot": Symbol("⊥", "rel", OTHER, note="前後に関係の空きが入る（PowerPoint と同じ）"),
    "angle": Symbol("∠", "ord", OTHER),
    "triangle": Symbol("△", "ord", OTHER),
    "square": Symbol("□", "ord", OTHER),
    "prime": Symbol("′", "ord", OTHER),
    "dagger": Symbol("†", "bin", OTHER),
    "ddagger": Symbol("‡", "bin", OTHER),
    "ldots": Symbol("…", "inner", OTHER, note="前後に細い空きが入る（添字の中を除く）"),
    "dots": Symbol("…", "inner", OTHER, note="前後に細い空きが入る（添字の中を除く）"),
    "cdots": Symbol("⋯", "inner", OTHER, note="前後に細い空きが入る（添字の中を除く）"),
    "vdots": Symbol("⋮", "ord", OTHER),
    "ddots": Symbol("⋱", "ord", OTHER),
    "backslash": Symbol("\\", "ord", OTHER),
    "slash": Symbol("/", "ord", OTHER, note="記号の斜線。/ だけを書くと横書きの分数になる"),
    "colon": Symbol(":", "punct", OTHER),
    "_": Symbol("_", "ord", OTHER),
    "%": Symbol("%", "ord", OTHER),
    "#": Symbol("#", "ord", OTHER),
    "&": Symbol("&", "ord", OTHER),
    "vert": Symbol("|", "ord", BAR),
    "lvert": Symbol("|", "open", BAR),
    "rvert": Symbol("|", "close", BAR),
    "|": Symbol("‖", "ord", BAR),
    "Vert": Symbol("‖", "ord", BAR),
    "lVert": Symbol("‖", "open", BAR),
    "rVert": Symbol("‖", "close", BAR),
}

# 命令でない文字のうち、出力の文字か間隔の種類が決まっているもの。ほかの文字はそのまま通常の文字として出す。
CHARACTERS: dict[str, Symbol] = {
    "+": Symbol("+", "bin", "文字"),
    "-": Symbol("−", "bin", "文字", note="負号にする（U+2212 の − も同じ）"),
    "*": Symbol("∗", "bin", "文字"),
    "=": Symbol("=", "rel", "文字"),
    "<": Symbol("<", "rel", "文字"),
    ">": Symbol(">", "rel", "文字"),
    ":": Symbol(":", "punct", "文字", note="後ろだけが空く（PowerPoint と同じ）"),
    ",": Symbol(",", "punct", "文字", note="数字の間（1,2）では数の一部（太字の数字の間を除く。PowerPoint と同じ）"),
    ";": Symbol(";", "punct", "文字"),
    ".": Symbol(".", "punct", "文字", note="数の中（3.14）では数の一部"),
    "!": Symbol("!", "close", "文字"),
    "?": Symbol("?", "close", "文字"),
    "'": Symbol("′", "ord", "文字", note="プライム。添字の位置では上付きになる"),
}

# 立体で書く関数名。\ を付けても付けなくてもよい。
_FUNCTION_ORDER = (
    "log",
    "ln",
    "lg",
    "exp",
    "sin",
    "cos",
    "tan",
    "cot",
    "sec",
    "csc",
    "arcsin",
    "arccos",
    "arctan",
    "sinh",
    "cosh",
    "tanh",
    "coth",
    "min",
    "max",
    "inf",
    "sup",
    "lim",
    "det",
    "gcd",
    "arg",
    "deg",
    "dim",
    "ker",
    "hom",
    "Pr",
)
FUNCTIONS: frozenset[str] = frozenset(_FUNCTION_ORDER)

# 括弧として扱う字句と出力の文字。\left・\right・\middle の後にはこの字句を置く（. は見えない括弧）。
DELIMITERS: dict[str, str] = {
    "(": "(",
    ")": ")",
    "[": "[",
    "]": "]",
    "\\lbrack": "[",
    "\\rbrack": "]",
    "\\{": "{",
    "\\}": "}",
    "\\lbrace": "{",
    "\\rbrace": "}",
    "\\langle": "⟨",
    "\\rangle": "⟩",
    "⟨": "⟨",
    "⟩": "⟩",
    "\\lceil": "⌈",
    "\\rceil": "⌉",
    "⌈": "⌈",
    "⌉": "⌉",
    "\\lfloor": "⌊",
    "\\rfloor": "⌋",
    "⌊": "⌊",
    "⌋": "⌋",
    "|": "|",
    "\\vert": "|",
    "\\lvert": "|",
    "\\rvert": "|",
    "\\|": "‖",
    "‖": "‖",
    "\\Vert": "‖",
    "\\lVert": "‖",
    "\\rVert": "‖",
    ".": "",
}

# 普通の括弧（\left・\right を付けない括弧）の組。開き括弧の字句と、それを閉じる字句。縦線は組にしない。
_ROUND_OR_SQUARE = frozenset({")", "]", "\\rbrack"})
_BRACE = frozenset({"\\}", "\\rbrace"})
_ANGLE = frozenset({"\\rangle", "⟩"})
_CEIL = frozenset({"\\rceil", "⌉"})
_FLOOR = frozenset({"\\rfloor", "⌋"})
BRACKETS: dict[str, frozenset[str]] = {
    "(": _ROUND_OR_SQUARE,
    "[": _ROUND_OR_SQUARE,
    "\\lbrack": _ROUND_OR_SQUARE,
    "\\{": _BRACE,
    "\\lbrace": _BRACE,
    "\\langle": _ANGLE,
    "⟨": _ANGLE,
    "\\lceil": _CEIL,
    "⌈": _CEIL,
    "\\lfloor": _FLOOR,
    "⌊": _FLOOR,
}

# \not の後の字句と、打ち消し付きの文字。
NEGATIONS: dict[str, str] = {
    "=": "≠",
    "<": "≮",
    ">": "≯",
    "\\in": "∉",
    "\\ni": "∌",
    "\\subset": "⊄",
    "\\supset": "⊅",
    "\\subseteq": "⊈",
    "\\supseteq": "⊉",
    "\\equiv": "≢",
    "\\sim": "≁",
    "\\simeq": "≄",
    "\\approx": "≉",
    "\\cong": "≇",
    "\\le": "≰",
    "\\leq": "≰",
    "\\ge": "≱",
    "\\geq": "≱",
    "\\mid": "∤",
    "\\parallel": "∦",
    "\\exists": "∄",
    "\\prec": "⊀",
    "\\succ": "⊁",
    "\\vdash": "⊬",
    "\\models": "⊭",
}

# 空白の命令と（出力する文字, 幅の em）。幅は Cambria Math の送り幅に合わせた。\! は何も出力しない。
SPACES: dict[str, tuple[str, float]] = {
    "\\,": (" ", 1 / 6),
    "\\thinspace": (" ", 1 / 6),
    "\\:": (" ", 4 / 18),
    "\\>": (" ", 4 / 18),
    "\\medspace": (" ", 4 / 18),
    "\\;": (" ", 1 / 4),
    "\\thickspace": (" ", 1 / 4),
    "\\enspace": (" ", 1 / 2),
    "\\quad": (" ", 1.0),
    "\\qquad": ("  ", 2.0),
    "\\ ": (" ", 0.22),
    "~": (" ", 0.22),
    "\\!": ("", 0.0),
}

# 記号付きの命令と、上に付ける記号（結合文字）。
ACCENTS: dict[str, str] = {
    "hat": "̂",
    "widehat": "̂",
    "tilde": "̃",
    "widetilde": "̃",
    "check": "̌",
    "breve": "̆",
    "acute": "́",
    "grave": "̀",
    "dot": "̇",
    "ddot": "̈",
    "vec": "⃗",
}

# 書体の命令と（字形の書式, Office Math の m:scr）。\boldsymbol・\bm は既定の書式を太字にする。
FONTS: dict[str, tuple[str, str | None]] = {
    "mathrm": ("upright", None),
    "mathit": ("italic", None),
    "mathbf": ("bold", None),
    "mathsf": ("upright", "sans-serif"),
    "mathtt": ("upright", "monospace"),
    "mathcal": ("upright", "script"),
    "mathscr": ("upright", "script"),
    "mathbb": ("upright", "double-struck"),
    "mathfrak": ("upright", "fraktur"),
}

# 対応していない命令と、エラーに添える助言（docs/math.md の「対応していない記法」）
_SIZE_ADVICE = "括弧の大きさは指定できない。伸ばす括弧は \\left・\\right で書く"
UNSUPPORTED: dict[str, str] = {
    "begin": "環境（\\begin…\\end）は対応していない",
    "end": "環境（\\begin…\\end）は対応していない",
    "big": _SIZE_ADVICE,
    "Big": _SIZE_ADVICE,
    "bigg": _SIZE_ADVICE,
    "Bigg": _SIZE_ADVICE,
    "bigl": _SIZE_ADVICE,
    "bigr": _SIZE_ADVICE,
    "Bigl": _SIZE_ADVICE,
    "Bigr": _SIZE_ADVICE,
    "biggl": _SIZE_ADVICE,
    "biggr": _SIZE_ADVICE,
    "Biggl": _SIZE_ADVICE,
    "Biggr": _SIZE_ADVICE,
    "displaystyle": "文字の大きさの段階は指定できない",
    "textstyle": "文字の大きさの段階は指定できない",
    "color": "色は要素の書式で指定する",
    "textcolor": "色は要素の書式で指定する",
    "hspace": "空白は \\,・\\:・\\;・\\quad などで書く",
    "mkern": "空白は \\,・\\:・\\;・\\quad などで書く",
    "over": "分数は \\frac で書く",
    "choose": "二項係数は \\binom で書く",
    "overrightarrow": "上の矢印は \\vec で書く",
    "overbrace": "対応していない",
    "underbrace": "対応していない",
    "substack": "対応していない",
    "tag": "式の番号は対応していない",
}

FRACTION = "分数と根号"
PLACEMENT = "添字の位置"
STRETCH = "伸びる括弧"
DECORATION = "記号付き"
TYPEFACE = "書体"
WORDS = "文章"
OPERATION = "演算"

STRUCTURES: dict[str, Structure] = {
    "frac": Structure(r"\frac{a}{b}", "分数", FRACTION),
    "dfrac": Structure(r"\dfrac{a}{b}", "分数（\\frac と同じ）", FRACTION),
    "tfrac": Structure(r"\tfrac{a}{b}", "分数（\\frac と同じ）", FRACTION),
    "binom": Structure(r"\binom{n}{k}", "二項係数", FRACTION, "伸びる丸括弧の中に線のない分数を置く"),
    "dbinom": Structure(r"\dbinom{n}{k}", "二項係数（\\binom と同じ）", FRACTION),
    "tbinom": Structure(r"\tbinom{n}{k}", "二項係数（\\binom と同じ）", FRACTION),
    "sqrt": Structure(r"\sqrt[3]{x}", "根号", FRACTION, "[…] の次数は省略できる。sqrt(x) と √x も根号になる"),
    "limits": Structure(
        r"\sum\limits_{i=1}^{n} a_i", "添字を記号の上下に置く", PLACEMENT, "大型演算子か関数名の直後に置く"
    ),
    "nolimits": Structure(r"\int\nolimits_0^1 f", "添字を記号の右に置く", PLACEMENT, "指定がないときと同じ"),
    "overset": Structure(r"\overset{?}{=}", "上に置く", PLACEMENT, "1 つ目の引数を 2 つ目の上に置く"),
    "stackrel": Structure(r"\stackrel{?}{=}", "上に置く（\\overset と同じ）", PLACEMENT),
    "underset": Structure(r"\underset{i}{=}", "下に置く", PLACEMENT, "1 つ目の引数を 2 つ目の下に置く"),
    "left": Structure(r"\left( \frac{a}{b} \right)", "伸びる括弧の開き側", STRETCH, "後に括弧の字句を置く"),
    "right": Structure(r"\left. x \right|", "伸びる括弧の閉じ側", STRETCH, "\\left と組で使う。. は見えない括弧"),
    "middle": Structure(
        r"\left\{ x \middle| x > 0 \right\}",
        "伸びる区切り",
        STRETCH,
        "\\left と \\right の間で使う。記号は 1 種類にする",
    ),
    "hat": Structure(r"\hat{x}", "記号付き（x̂）", DECORATION),
    "widehat": Structure(r"\widehat{xy}", "記号付き（幅の広い ^）", DECORATION),
    "tilde": Structure(r"\tilde{x}", "記号付き（x̃）", DECORATION),
    "widetilde": Structure(r"\widetilde{xy}", "記号付き（幅の広い ~）", DECORATION),
    "check": Structure(r"\check{x}", "記号付き（x̌）", DECORATION),
    "breve": Structure(r"\breve{x}", "記号付き（x̆）", DECORATION),
    "acute": Structure(r"\acute{x}", "記号付き（x́）", DECORATION),
    "grave": Structure(r"\grave{x}", "記号付き（x̀）", DECORATION),
    "dot": Structure(r"\dot{x}", "記号付き（ẋ）", DECORATION),
    "ddot": Structure(r"\ddot{x}", "記号付き（ẍ）", DECORATION),
    "vec": Structure(r"\vec{v}", "記号付き（v⃗）", DECORATION),
    "bar": Structure(r"\bar{x}", "上線", DECORATION),
    "overline": Structure(r"\overline{xy}", "上線", DECORATION),
    "underline": Structure(r"\underline{x}", "下線", DECORATION),
    "mathrm": Structure(r"\mathrm{d}", "立体", TYPEFACE, "引数の波括弧を省くと次の 1 文字が対象になる"),
    "mathit": Structure(r"\mathit{d}", "斜体", TYPEFACE),
    "mathbf": Structure(r"\mathbf{v}", "太字", TYPEFACE),
    "mathsf": Structure(r"\mathsf{A}", "サンセリフ体", TYPEFACE),
    "mathtt": Structure(r"\mathtt{A}", "等幅", TYPEFACE),
    "mathcal": Structure(r"\mathcal{O}", "筆記体", TYPEFACE),
    "mathscr": Structure(r"\mathscr{L}", "筆記体（\\mathcal と同じ）", TYPEFACE),
    "mathbb": Structure(r"\mathbb{R}", "黒板太字", TYPEFACE),
    "mathfrak": Structure(r"\mathfrak{g}", "ドイツ文字", TYPEFACE),
    "boldsymbol": Structure(r"\boldsymbol{\alpha}", "太字", TYPEFACE, "斜体の文字は太字の斜体になる"),
    "bm": Structure(r"\bm{x}", "太字（\\boldsymbol と同じ）", TYPEFACE),
    "operatorname": Structure(
        r"\operatorname{rank}(A)", "立体の関数名", TYPEFACE, "\\operatorname*{…} は添字を上下に置く"
    ),
    "text": Structure(
        r"\text{最大}", "文章", WORDS, "要素の文字の書体で書き出す。中に波括弧と命令は書くことができない"
    ),
    "not": Structure(r"\not\in", "打ち消し（∉）", OPERATION, "打ち消しに対応する記号は「打ち消し」の表のとおり"),
    "bmod": Structure(r"a \bmod b", "立体の mod", OPERATION, "前後に 4/18 em の空白（U+205F）を置く"),
    "pmod": Structure(r"a \pmod{n}", "空白と (mod n)", OPERATION),
}


def _is_name(value: str) -> bool:
    return re.fullmatch(r"[A-Za-z]+", value) is not None


@lru_cache(maxsize=1)
def command_names() -> frozenset[str]:
    """受け付ける英字の命令名（\\ を除いた名前）。記号・構造・関数名・括弧・空白の命令を含む。"""
    names = {name for name in (*SYMBOLS, *STRUCTURES) if _is_name(name)}
    names.update(FUNCTIONS)
    tokens = (*DELIMITERS, *SPACES, *NEGATIONS)
    names.update(token[1:] for token in tokens if token.startswith("\\") and _is_name(token[1:]))
    return frozenset(names)


GENERATED_NOTE = "この節は `python -m slidekit.math.commands --write docs/math.md` で作る。手で編集しない。"
BEGIN = "<!-- slidekit:math-commands:begin -->"
END = "<!-- slidekit:math-commands:end -->"


def _code(text: str) -> str:
    """表の中のコードを書く。縦線を含む字句は、表の区切りにならないよう HTML の要素で書く。"""
    if "|" in text or "`" in text:
        escaped = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return "<code>" + escaped.replace("|", "&#124;").replace("`", "&#96;") + "</code>"
    return "`" + text + "`"


def _plain(text: str) -> str:
    """表の中の文字を書く。Markdown の記号として解釈される文字を、文字参照か \\ で表す。"""
    replaced = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("|", "&#124;")
    replaced = re.sub(r"([*_`])", r"\\\1", replaced)
    if replaced.endswith("\\"):
        replaced = replaced[:-1] + "&#92;"
    return replaced


_SPACE_NOTES = {"\\ ": "\\ と空白 1 文字", "\\!": "何も出力しない"}


def _space_name(char: str) -> str:
    return "（なし）" if not char else "・".join(f"U+{ord(c):04X}" for c in char)


def _em(value: float) -> str:
    for denominator in (1, 2, 4, 6, 18):
        if abs(value * denominator - round(value * denominator)) < 1e-9:
            numerator = round(value * denominator)
            return f"{numerator} em" if denominator == 1 else f"{numerator}/{denominator} em"
    return f"{value:g} em"


def reference_markdown() -> str:
    """対応する命令の一覧を Markdown で返す（`docs/math.md` の印の間に入れる）。"""
    lines = [GENERATED_NOTE, ""]
    kinds = "・".join(KIND_NAMES.values())
    lines += [
        "### 記号",
        "",
        f"「種類」は前後の空きの決め方を表す（{kinds}）。斜体になる記号は備考に「斜体」と書く。",
        "",
    ]
    groups = list(dict.fromkeys(symbol.group for symbol in SYMBOLS.values()))
    for group in groups:
        lines += [f"#### {group}", "", "| 命令 | 出力 | 種類 | 備考 |", "|---|---|---|---|"]
        for name, symbol in SYMBOLS.items():
            if symbol.group != group:
                continue
            note = "、".join(part for part in ("斜体" if symbol.italic else "", symbol.note) if part)
            if group == OPERATOR:
                note = note or "後に引数が要る"
            command = "\\" + name
            lines.append(f"| {_code(command)} | {_plain(symbol.char)} | {KIND_NAMES[symbol.kind]} | {_plain(note)} |")
        lines.append("")
    lines += ["### 構造", ""]
    groups = list(dict.fromkeys(structure.group for structure in STRUCTURES.values()))
    for group in groups:
        lines += [f"#### {group}", "", "| 書き方 | 出力 | 備考 |", "|---|---|---|"]
        for structure in STRUCTURES.values():
            if structure.group == group:
                lines.append(f"| {_code(structure.syntax)} | {_plain(structure.output)} | {_plain(structure.note)} |")
        lines.append("")
    lines += [
        "### 括弧",
        "",
        "普通の括弧は伸びない。`\\left`・`\\right`・`\\middle` の後に置くと伸びる括弧になる。縦線は組にしない。",
        "",
        "| 字句 | 出力 | 普通の括弧として閉じる字句 |",
        "|---|---|---|",
    ]
    for token, char in DELIMITERS.items():
        closers = sorted(BRACKETS.get(token, ()), key=list(DELIMITERS).index)
        shown = _plain(char) if char else "（見えない括弧）"
        lines.append(f"| {_code(token)} | {shown} | {'、'.join(_code(closer) for closer in closers)} |")
    lines += ["", "### 空白", "", "| 命令 | 出力する文字 | 幅 | 備考 |", "|---|---|---|---|"]
    for token, (char, em) in SPACES.items():
        lines.append(f"| {_code(token)} | {_space_name(char)} | {_em(em)} | {_plain(_SPACE_NOTES.get(token, ''))} |")
    lines += ["", "### 打ち消し", "", "`\\not` は次の記号の打ち消しに対応する。", "", "| 書き方 | 出力 |", "|---|---|"]
    for token, char in NEGATIONS.items():
        negation = "\\not" + token
        lines.append(f"| {_code(negation)} | {_plain(char)} |")
    lines += ["", "### 関数名", "", "次の名前は `\\` の有無にかかわらず立体の関数名になる。", ""]
    lines.append("、".join(_code(name) for name in _FUNCTION_ORDER))
    lines += ["", "### 文字の変換", "", "| 文字 | 出力 | 種類 | 備考 |", "|---|---|---|---|"]
    for char, symbol in CHARACTERS.items():
        lines.append(f"| {_code(char)} | {_plain(symbol.char)} | {KIND_NAMES[symbol.kind]} | {_plain(symbol.note)} |")
    return "\n".join(lines).rstrip("\n") + "\n"


def _region(text: str) -> tuple[int, int] | None:
    """印の間（開始の印の次の行から、終了の印の直前まで）の位置を返す。印がなければ None。"""
    begin = text.find(BEGIN)
    if begin < 0:
        return None
    start = text.find("\n", begin + len(BEGIN))
    end = text.find(END, begin + len(BEGIN))
    if start < 0 or end < 0 or end <= start:
        return None
    return start + 1, end


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="python -m slidekit.math.commands", description="数式の命令の一覧を Markdown で出す。"
    )
    target = parser.add_mutually_exclusive_group()
    target.add_argument("--write", metavar="PATH", help="文書の印の間を一覧で書き換える")
    target.add_argument("--check", metavar="PATH", help="文書の印の間が一覧と一致するかを確かめる")
    arguments = parser.parse_args(argv)
    content = reference_markdown()
    if arguments.write is None and arguments.check is None:
        sys.stdout.write(content)
        return 0
    path = Path(arguments.write if arguments.write is not None else arguments.check)
    text = path.read_text(encoding="utf-8")
    region = _region(text)
    if region is None:
        print(f"{path} に印 {BEGIN} と {END} がない", file=sys.stderr)
        return 2
    start, end = region
    current = text[start:end]
    if arguments.write is not None:
        if current != content:
            path.write_text(text[:start] + content + text[end:], encoding="utf-8")
        return 0
    if current == content:
        return 0
    actual, expected = current.splitlines(), content.splitlines()
    index = next((i for i, (a, b) in enumerate(zip(actual, expected)) if a != b), min(len(actual), len(expected)))
    line = text[:start].count("\n") + 1 + index
    print(
        f"{path} の {line} 行目が一覧と違う。python -m slidekit.math.commands --write {path} で作り直す",
        file=sys.stderr,
    )
    return 1


if __name__ == "__main__":
    sys.exit(main())
