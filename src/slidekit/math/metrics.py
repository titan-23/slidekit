r"""OpenType の MATH 表による数式の寸法の計算。

`syntax.parse()` の構文木を、数式フォントの MATH 表の定数と字形（`font.MathFont`）で組み、幅・上下・描画範囲を pt で求める。
組み方は MathML Core の配置の規則と OpenType の MATH 表の仕様に従う。Office Math（PowerPoint）と違う点については、
PowerPoint で変換した PDF と照合し（`compare_reference()`）、`OFFICE` と `SPACING` の値を合わせる。

- 文字サイズの段階は最上位を 0 段とし、添字・上下の添え字・文中の分数の分子と分母で 1 段、根号の次数で 2 段増やす。
  1 段目は ScriptPercentScaleDown、2 段目以降は ScriptScriptPercentScaleDown を掛け、添字用の字形（GSUB の ssty）を使う。
- 独立した数式（``display=True``）では、最上位の分数・根号・大型演算子を独立した数式の組み方にする。
- 原子の間の空きは TeX と同じ規則（`SPACING`）で入れる。単位は 1/18 em で、添字の中では括弧付きの値を入れない。
  原子の種類は命令の表（`commands`）の値を使う。空白を挟まない数字の間のカンマとピリオドは数の一部とする
  （太字の数字を除く。`OFFICE`）。関数の適用の右側と横書きの分数は inner として扱い、上下に置くもののうち、土台が
  本体のない関数名だけのものは ord として扱う（`OFFICE`）。
- 添字の上下の位置は土台の並びの最後の原子で決める（`(\frac{a}{b})^2` の上付きは閉じ括弧の上付きと同じ位置。`OFFICE`）。
- 斜体の文字の後に立体の文字か空白が続くとき、および文中の数式の末尾が斜体の文字のとき、斜体補正を加える（`OFFICE`）。
- 添字は MathKern で詰める。``\text`` の中身は `TextFonts` を使い、書き出しと同じ大きさ（`office_math.text_size()`）で
  計測する（`TextFonts` がなければ数式フォントの字形で計測する）。

`MathMeasurer` は計測結果を保持し、同じ条件の 2 回目以降は式を解析し直さない。このモジュールは matplotlib を使わない。

PowerPoint の照合データ（`tests/data/powerpoint_math.json` の形）との比較には、試験と同じ `compare_reference()` を使う。
開発用のコマンド（合格 0、不合格 1、照合できなければ 2。2 は、照合データの書体のハッシュ値がない・フォントと違う場合と、
ファイル・照合データの形・引数の誤り、計測できない式、計測の内部の誤り）::

    python -m slidekit.math.metrics 照合データ.json --font /mnt/c/Windows/Fonts/cambria.ttc:1 \
        --latin /mnt/c/Windows/Fonts/arial.ttf --japanese /mnt/c/Windows/Fonts/YuGothM.ttc:1
"""

from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass, field, fields, replace
import json
import math
from pathlib import Path
from threading import RLock
from typing import Callable, Mapping, NamedTuple
import unicodedata

from ..files import file_hash
from ..measure import uses_east_asian_font
from .font import Assembly, Glyph, MathFont
from .office_math import text_size
from .syntax import (
    Accent,
    Bar,
    Fenced,
    Fraction,
    Function,
    Limits,
    Nary,
    Radical,
    Run,
    Scripts,
    Space,
    Text,
    function_names,
    parse,
)

__all__ = [
    "ENGINE_VERSION",
    "OFFICE",
    "REFERENCE_LIMITS",
    "SPACING",
    "Box",
    "MathMeasurer",
    "TextFonts",
    "compare_reference",
    "layout",
    "reference_font_matches",
]

# 組み方を変えたら上げる。FontRegistry.inventory() に入れ、計測の方法が変わったことを証拠に反映する。
ENGINE_VERSION = "opentype-4"

# Office Math に合わせる値。MathML Core・TeX・OpenType の MATH 表の規則を基に、PowerPoint 16.0.20430.20146 で
# 変換した PDF の字形の位置と照合して決めた（2026-10-12。照合データ tests/data/powerpoint_math.json、
# 文中と独立した数式の 100 式ずつ。PowerPoint を使う確認と、その結果による 1 回目の調整）。
# 2026-10-20 に、照合データにない式で PowerPoint 16.0.20430.20182 の PDF と違った規則（添字の土台、関数名の後の \{、
# 数字の間のカンマ、\bot の種類（commands の表）、関数の適用の右側、横書きの分数、本体のない関数名の上下に置くもの）を直し、
# 確かめた式を照合データに加えた。
# 値を変えたら、新しい MathMeasurer を作る（計測結果の保持は OFFICE を区別しない）。
OFFICE: dict[str, object] = {
    # 下付きを、土台の斜体補正だけ左へ寄せる（真）か、土台の送り幅の後に置く（偽）か。偽のとき MathKern で詰める
    "subscript_italic_shift": False,
    # 添字を、土台と添字の字形の MathKern で詰める（2 つの高さのうち、詰める量の小さい方）
    "math_kern": True,
    # 添字の上下の位置と MathKern を、土台の並びの最後の原子（括弧で囲んだ土台では閉じ括弧）で決める。偽なら土台全体の箱で
    # 決める（MathML Core・TeX と同じ）。PowerPoint は (\frac{a}{b})^2 と (\log x)^2 の形の式の上付きを、括弧の中の高さによらず
    # 通常の位置（20 pt で基線の上 7.32 pt）に置く（2026-10-20、校正用の資料と生成した資料の PDF）
    "script_last_atom": True,
    # 斜体の文字の後に立体の文字（演算子・括弧など）が続くと、斜体補正を送り幅に加える
    "italic_correction": True,
    # 数字の間のカンマとピリオド（1,2）を数の一部として扱い、後ろに空きを入れない。数を作るのは、間に空白がなく、太字でも
    # m:scr の書体でもない数字だけである。PowerPoint は (4,2) のカンマの後を空けず、太字の (0,0) と 1,\,2 のカンマの後には
    # 3/18 em を空ける（2026-10-20、生成した資料の PDF。太字の数字は数学用の英数字になり、数を作らない。m:scr の書体の
    # 数字も数学用の英数字になるので同じ扱いにした）
    "number_separators": True,
    # 添字・分数の分子と分母などに、添字用の字形（GSUB の ssty）を使う
    "script_alternates": True,
    # 斜体の文字の後に空白（\, や \quad など）が続くときも、斜体補正を加える。空白の幅は commands.SPACES のまま
    # （a\,b などと ab の差は、空白の幅に 𝑎 の斜体補正 0.0244 em を加えた値で、ずれは 0.006 em 以内）
    "italic_before_space": True,
    # 文中の数式の末尾が斜体の文字なら、斜体補正を送り幅に加える。添字・記号付き・上線と下線・根号の末尾の文字からも伝える
    "italic_trailing": True,
    # 関数名と引数の間の空き（em）。引数が function_space_brackets の括弧で始まるときは 0
    # （PowerPoint は \max( と同じく \max\{ の { の前も空けない。2026-10-20、生成した資料の PDF）
    "function_space_em": 3 / 18,
    "function_space_brackets": "([{",
    # 伸びる括弧の目標の長さに掛ける係数（TeX の \delimiterfactor の 901 と同じ）。長い括弧には短差の上限も適用する
    "delimiter_factor": 0.901,
    # 必要な長さより短くてよい量（em）。TeX の \delimitershortfall と同じ制限。PowerPoint の 17 ケースで求めた
    # 共通の範囲は 0.1509〜0.1736 em（2026-10-23、20 pt、Cambria Math）。その範囲内の 0.16 em を使う
    "delimiter_shortfall_em": 0.16,
    # この高さを軸からの距離が超える中身は、一つの大きな式として扱い、中身の中心に括弧を置く。Office は
    # DelimitedSubFormulaMinHeight と DisplayOperatorMinHeight を逆に読む（大型演算子の規則と対になる）
    "delimiter_subformula_min_height": "DisplayOperatorMinHeight",
    # 根号の字形が必要な長さより長いとき、余りの半分を中身の上の空きに加える（TeX・MathML Core と同じ）。偽なら余りを下に出す
    "radical_excess_above": True,
    # 関数の適用（本体のある関数名。書き出しは m:func）の右側を inner として扱う。後に普通の文字・句読点・開き括弧が続くと
    # （空白の命令を挟んでも）3/18 em を空け（文字サイズの段階が 0 のときだけ）、二項演算・関係・閉じ括弧の前は通常どおり。
    # 中身の最後の節が関数の適用である根号も同じにする。PowerPoint は \log x,y・\log x\,y・\log x(y)・\max(a,b)\,c・
    # \sqrt{x\log x},y で 3/18 em を空け、\sqrt{x},y・(\log x),y・\log x+y・f(x),y では空けない（2026-10-20、照合データと
    # 生成した資料の PDF）。偽なら関数の適用の右側は引数の最後の原子の種類のまま
    "function_apply_inner": True,
    # 横書きの分数（a/b）の斜線に使う文字。PowerPoint は U+2044（分数の斜線）の字形で描く。書体になければ / で計測する
    "fraction_slash": "\u2044",
    # 横書きの分数を inner として扱い、前の斜体の文字と分数の末尾の斜体の文字の斜体補正を加える。PowerPoint は x a/b・2a/b・
    # (x)a/b の分数の前と、a/b,x・a/b x・a/b(x) の分数の後に 3/18 em を空け、x+a/b・x(a/b) と添字の中では空けない。x a/b は x の、
    # a/b x は b の斜体補正だけさらに広い（2026-10-20、照合データ）。偽なら ord として扱う
    "linear_fraction_inner": True,
    # 上下に置くもののうち、土台が本体のない関数名だけのもの（\underset{i}{\max}。書き出しは m:limLow の土台に関数名の文字）を、
    # 両側とも ord として扱う。PowerPoint は x\underset{i}{\max} y の両側と \underset{i}{\max} x_i の後を空けない
    # （2026-10-20、照合データと校正用の資料）。偽なら関数名と同じ op として扱う
    "bare_function_limits_ord": True,
    # 文中の数式の分数で、分子と分母を 1 段縮める（分子・分母は 20 pt の 73 %）
    "inline_fraction_shrink": True,
    # 入れ子の分数の外側の線は、通常の送り幅より左右へ張り出す。内側の分数には半分の余白を加える
    # （20 pt の式では送り幅に 2 pt、線の描画幅に 4 pt。2026-10-23 の PowerPoint の PDF で照合）。
    "nested_fraction_padding_em": 0.1,
    # 独立した数式の大型演算子の字形。min_height は display_operator_min_height の定数以上の最初の候補、next は基本の次の候補
    "display_operator_variant": "min_height",
    # 独立した数式の大型演算子の最小の高さに使う MATH 表の定数。Office は DisplayOperatorMinHeight と
    # DelimitedSubFormulaMinHeight を逆に読み、Cambria Math の値はそれに合わせてある（Microsoft が認め、HarfBuzz 11.1 が補う）。
    # 独立した数式の ∑ と ∫ の高さが PowerPoint と合う（DisplayOperatorMinHeight では ∫ が 35.0 pt で、PowerPoint は 51.1 pt）
    "display_operator_min_height": "DelimitedSubFormulaMinHeight",
    # 文中の数式でも \limits の大型演算子の添字を上下に積む。偽なら右に置く（PowerPoint は文中では \limits を適用しない。
    # 関数名の \limits（\max\limits_i）は文中でも下に置く）
    "inline_nary_limits": False,
    # 文中の数式の積分以外の大型演算子で、添字の位置に演算子の字形の上下も使う（偽なら使わず、文字の添字と同じ位置にする）
    "inline_nary_script_drop": False,
    # 大型演算子の添字の後に SpaceAfterScript を入れる。偽なら入れず、添字の末尾の斜体補正だけを加える
    "nary_space_after_script": False,
    # 独立した数式で、添字を右に置く（か添字のない）大型演算子と引数の間に原子の間の空きを入れる。偽なら空けない
    # （上下に積むときは原子の間の空きを入れる）
    "display_nary_operand_space": False,
}

# 積分の記号。文中の数式でも、添字の位置に演算子の字形の上下を使う
INTEGRALS = frozenset("∫∬∭∮∯∰")

# 原子の間の空き。値は (1/18 em を単位とする幅, 文字サイズの段階が 0 のときだけ入れるか)。表にない組は 0。
# TeX の規則と同じ（括弧付きの値は添字の中では入れない）。種類は ord・op・bin・rel・open・close・punct・inner。
THIN, MEDIUM, THICK = 3, 4, 5
SPACING: dict[tuple[str, str], tuple[int, bool]] = {
    ("ord", "op"): (THIN, False),
    ("ord", "bin"): (MEDIUM, True),
    ("ord", "rel"): (THICK, True),
    ("ord", "inner"): (THIN, True),
    ("op", "ord"): (THIN, False),
    ("op", "op"): (THIN, False),
    ("op", "rel"): (THICK, True),
    ("op", "inner"): (THIN, True),
    ("bin", "ord"): (MEDIUM, True),
    ("bin", "op"): (MEDIUM, True),
    ("bin", "open"): (MEDIUM, True),
    ("bin", "inner"): (MEDIUM, True),
    ("rel", "ord"): (THICK, True),
    ("rel", "op"): (THICK, True),
    ("rel", "open"): (THICK, True),
    ("rel", "inner"): (THICK, True),
    ("close", "op"): (THIN, False),
    ("close", "bin"): (MEDIUM, True),
    ("close", "rel"): (THICK, True),
    ("close", "inner"): (THIN, True),
    ("punct", "ord"): (THIN, True),
    ("punct", "op"): (THIN, True),
    ("punct", "rel"): (THIN, True),
    ("punct", "open"): (THIN, True),
    ("punct", "close"): (THIN, True),
    ("punct", "punct"): (THIN, True),
    ("punct", "inner"): (THIN, True),
    ("inner", "ord"): (THIN, True),
    ("inner", "op"): (THIN, False),
    ("inner", "bin"): (MEDIUM, True),
    ("inner", "rel"): (THICK, True),
    ("inner", "open"): (THIN, True),
    ("inner", "punct"): (THIN, True),
    ("inner", "inner"): (THIN, True),
}

# 二項演算の記号を ord として扱う、直前の原子の種類（式の先頭も ord になる）
_BINARY_AFTER = frozenset({"bin", "op", "rel", "open", "punct"})
# 直前の二項演算の記号を ord にする、直後の原子の種類
_BINARY_BEFORE = frozenset({"rel", "close", "punct"})

# 計測結果を保持する件数
CACHE_SIZE = 4096


@dataclass(frozen=True)
class TextFonts:
    """\\text の中身を計測する文字用の書体。"""

    key: tuple  # 計測結果の保持に使う区別。例: ("Yu Gothic UI", "Arial")
    measure: Callable[[str, float, bool], tuple[float, ...]] = field(compare=False, hash=False)
    # 1 行の文字列の (送り幅, 基線より上, 基線より下) か、それに描画範囲の (左, 右) を加えた 5 つの値を pt で返す。
    # 引数は (文字列, 文字サイズ pt, 太字)。上下は字形の描画範囲（インク）にする（数式の字形と同じく、描画範囲と行の配置の
    # 両方に使う）。左右は文字列の左端を原点にした最初と最後の字形の描画範囲で、なければ送り幅の範囲（0, 送り幅）を使う


@dataclass(frozen=True)
class Box:
    """組んだ数式の寸法（pt。原点は式の左端の基線）。"""

    width: float  # 組版上の幅（送り幅）
    height: float  # 基線より上（行の配置に使う上側。根号・分数・上線の余白を含む）
    depth: float  # 基線より下（行の配置に使う下側）
    ink: tuple[float, float, float, float]  # 描画範囲 (左, 右, 基線より上, 基線より下)。インクがなければすべて 0
    italic: float = 0.0  # 末尾の字形の斜体補正
    kind: str = "ord"  # 先頭の原子の種類


# ---- 組み方の部品 ----


class _Style(NamedTuple):
    level: int  # 文字サイズの段階
    cramped: bool  # 詰めた形（上付きを低くする）
    display: bool  # 独立した数式の組み方


@dataclass
class _Glue:
    """空白。suppress が真なら、前後の原子の間の空き（SPACING）を入れない。"""

    width: float
    suppress: bool = False


@dataclass
class _Box:
    width: float = 0.0
    height: float = 0.0
    depth: float = 0.0
    ink: tuple[float, float, float, float] | None = None  # (左, 右, 基線より上, 基線より下)
    italic: float = 0.0
    left: str = "ord"  # 左側の原子の種類
    right: str = "ord"  # 右側の原子の種類
    first: tuple[Glyph, float] | None = None  # 先頭の字形と em から pt への倍率（文字の並びのとき）
    last: tuple[Glyph, float] | None = None  # 末尾の字形と倍率（文字の並びのとき）
    string: bool = False  # 文字の並び（添字の位置で「箱」として扱わない）
    single: bool = False  # 字形 1 つだけ
    extended: bool = False  # 伸びた字形（添字の位置で箱として扱う）
    upright_start: bool = False  # 立体の文字（演算子・括弧など）で始まる
    slanted_end: bool = False  # 斜体の文字で終わる
    text: str = ""  # 文字の並びの文字（数字の間のカンマの判定に使う）
    plain: bool = False  # 文字を書体の変換なしに書き出す（太字・m:scr の数字は数学用の英数字になり、数を作らない）


@dataclass(frozen=True)
class _Shape:
    """伸縮する字形の候補か組み立て（em。y は基線から上が正）。"""

    advance: float
    left: float
    right: float
    bottom: float
    top: float
    italic: float
    extended: bool


def _union(a, b):
    if a is None:
        return b
    if b is None:
        return a
    return min(a[0], b[0]), max(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3])


def _moved(ink, dx: float, dy: float):
    if ink is None:
        return None
    return ink[0] + dx, ink[1] + dx, ink[2] + dy, ink[3] - dy


def _place(target: _Box, item: _Box, dx: float, dy: float) -> None:
    """item を (dx, dy) に置き、target の描画範囲と上下を広げる。dy は上が正。"""
    target.ink = _union(target.ink, _moved(item.ink, dx, dy))
    target.height = max(target.height, item.height + dy)
    target.depth = max(target.depth, item.depth - dy)


def _top(box: _Box) -> float:
    return box.ink[2] if box.ink is not None else 0.0


def _bottom(box: _Box) -> float:
    return box.ink[3] if box.ink is not None else 0.0


def _effective(style: str, bold: bool) -> str:
    """要素の太字を字形の書式に反映する（Office Math の m:sty と同じ）。"""
    if not bold:
        return style
    return {"italic": "bold-italic", "upright": "bold"}.get(style, style)


def _opens_with_bracket(nodes, brackets: str) -> bool:
    """関数の引数が brackets のどれかの括弧で始まるか。"""
    node = nodes[0] if nodes else None
    while isinstance(node, Scripts) and node.base:
        node = node.base[0]
    if isinstance(node, Fenced):
        return bool(node.open) and node.open in brackets
    return isinstance(node, Run) and node.kind == "open" and node.text in brackets


def _ends_with_application(nodes) -> bool:
    """並びの最後の節（空白を除く）が関数の適用（本体のある関数名）か。"""
    nodes = [node for node in nodes if not isinstance(node, Space)]
    return bool(nodes) and isinstance(nodes[-1], Function) and bool(nodes[-1].body)


def _delimiter_body(nodes, *, all_scripts: bool) -> tuple:
    """括弧の寸法の判定用の中身。大型演算子の添字を除き、all_scripts なら普通の添字も除く。

    文中の大型演算子の添字は括弧を伸ばさない。大きな式への切り替えも添字を除いた本体で判定するので、
    分数の分子に上付きが付いたときは、その高さだけで括弧の中心を移さない。描画する中身は変更しない。
    """
    result = []
    for node in nodes:
        if all_scripts and isinstance(node, Scripts):
            result.extend(_delimiter_body(node.base, all_scripts=True))
            continue
        changes = {"sub": None, "sup": None} if isinstance(node, Nary) else {}
        for entry in fields(node):
            value = getattr(node, entry.name)
            if entry.name in changes or not isinstance(value, tuple) or not value:
                continue
            if isinstance(value[0], tuple):
                changed = tuple(_delimiter_body(part, all_scripts=all_scripts) for part in value)
            else:
                changed = _delimiter_body(value, all_scripts=all_scripts)
            if changed != value:
                changes[entry.name] = changed
        result.append(replace(node, **changes) if changes else node)
    return tuple(result)


def _number_separators(items: list) -> None:
    """数字の間のカンマとピリオド（1,2）を数の一部として ord にする（Office Math の数の扱い）。

    数を作るのは、空白を挟まずに隣り合う、書体の変換のない数字（`_Box.plain`）だけである。
    """
    for previous, atom, following in zip(items, items[1:], items[2:]):
        if not all(isinstance(item, _Box) for item in (previous, atom, following)):
            continue
        if (
            atom.text in (",", ".")
            and previous.plain
            and following.plain
            and previous.text[-1:].isdigit()
            and following.text[:1].isdigit()
        ):
            atom.left = atom.right = "ord"


def _convert_binary(atoms: list[_Box]) -> None:
    """二項演算の記号を、置かれた位置で ord にする（TeX の規則。単項の負号など）。"""
    previous = None
    for atom in atoms:
        if atom.left == "bin":
            if previous is None or previous.right in _BINARY_AFTER:
                atom.left = "ord"
                if atom.right == "bin":
                    atom.right = "ord"
        elif atom.left in _BINARY_BEFORE and previous is not None and previous.right == "bin":
            previous.right = "ord"
            if previous.left == "bin":
                previous.left = "ord"
        previous = atom
    if previous is not None and previous.right == "bin":
        previous.right = "ord"
        if previous.left == "bin":
            previous.left = "ord"


class _Engine:
    """1 回の組み立て。文字サイズ・太字・文字用の書体と、OFFICE の値の写しを持つ。"""

    def __init__(
        self, font: MathFont, size: float, bold: bool, text_fonts: TextFonts | None, *, display: bool = False
    ) -> None:
        self.font = font
        self.size = size
        self.bold = bold
        self.text_fonts = text_fonts
        self.display = display
        self.office = dict(OFFICE)
        script = font.constant("ScriptPercentScaleDown") or 0.8
        scriptscript = font.constant("ScriptScriptPercentScaleDown") or 0.6
        self.scales = (1.0, script, scriptscript)

    def em(self, level: int) -> float:
        return self.size * self.scales[min(max(level, 0), 2)]

    def constant(self, name: str, level: int) -> float:
        return self.font.constant(name) * self.em(level)

    # 並び

    def sequence(self, nodes, style: _Style) -> _Box:
        items: list[_Box | _Glue] = []
        for node in nodes:
            self.append(items, node, style)
        return self.pack(items, style.level)

    def append(self, items: list, node, style: _Style) -> None:
        """節を原子の列に加える。普通の括弧・大型演算子の引数・関数の引数は、外側の並びの原子として加える。"""
        if isinstance(node, Space):
            items.append(_Glue(node.em * self.em(style.level)))
        elif isinstance(node, Fenced) and not node.stretchy:
            items.append(self.character(node.open, "open", style.level))
            for child in node.parts[0] if node.parts else ():
                self.append(items, child, style)
            items.append(self.character(node.close, "close", style.level))
        elif isinstance(node, Nary):
            items.append(self.nary(node, style))
            if (
                node.body
                and style.display
                and node.limits != "undOvr"
                and not self.office["display_nary_operand_space"]
            ):
                # 独立した数式で添字を右に置く（か添字のない）大型演算子は、引数との間を空けない
                items.append(_Glue(0.0, suppress=True))
            for child in node.body:
                self.append(items, child, style)
        elif isinstance(node, Function):
            name = self.sequence(node.name, style)
            name.left = name.right = "op"
            items.append(name)
            if node.body:
                brackets = str(self.office["function_space_brackets"])
                gap = 0.0 if _opens_with_bracket(node.body, brackets) else float(self.office["function_space_em"])
                items.append(_Glue(gap * self.em(style.level), suppress=True))
                for child in node.body:
                    self.append(items, child, style)
                last = items[-1]
                if self.office["function_apply_inner"] and isinstance(last, _Box) and last.right in ("ord", "close"):
                    # 関数の適用の右側は inner（後の普通の文字・句読点・開き括弧との間に 3/18 em。OFFICE）
                    last.right = "inner"
        else:
            items.append(self.box(node, style))

    def pack(self, items: list, level: int, *, spacing: bool = True) -> _Box:
        """原子を横に並べる。原子が 1 つだけなら、その箱をそのまま返す。spacing が偽なら原子の間の空きを入れない。"""
        atoms = [item for item in items if isinstance(item, _Box)]
        if spacing:
            if self.office["number_separators"]:
                _number_separators(items)
            _convert_binary(atoms)
        if len(items) == 1 and atoms:
            return atoms[0]
        result = _Box()
        x = 0.0
        previous = None
        suppress = not spacing
        for index, item in enumerate(items):
            if isinstance(item, _Glue):
                if (
                    self.office["italic_before_space"]
                    and previous is not None
                    and items[index - 1] is previous
                    and previous.slanted_end
                    and item.width > 0
                    and not item.suppress
                ):
                    # 斜体の文字の後に空白が続くときも、斜体補正を加える（PowerPoint の PDF の空白の幅と合う）
                    x += previous.italic
                x += item.width
                suppress = suppress or item.suppress
                continue
            if previous is not None:
                if not suppress:
                    x += self.spacing(previous.right, item.left, level)
                if (
                    self.office["italic_correction"]
                    and previous is items[index - 1]
                    and previous.slanted_end
                    and item.upright_start
                ):
                    # 斜体の文字の後に立体の文字（演算子・括弧など）が続くと、斜体補正を送り幅に加える（OpenType の MATH 表）
                    x += previous.italic
            _place(result, item, x, 0.0)
            x += item.width
            previous = item
            suppress = not spacing
        result.width = x
        if atoms:
            result.left, result.right = atoms[0].left, atoms[-1].right
            if items[0] is atoms[0]:
                result.first, result.upright_start = atoms[0].first, atoms[0].upright_start
            if items[-1] is atoms[-1]:
                result.last, result.italic, result.slanted_end = atoms[-1].last, atoms[-1].italic, atoms[-1].slanted_end
        return result

    def spacing(self, left: str, right: str, level: int) -> float:
        value = SPACING.get((left, right))
        if value is None:
            return 0.0
        amount, outer_only = value
        if outer_only and level > 0:
            return 0.0
        return amount / 18 * self.em(level)

    def box(self, node, style: _Style) -> _Box:
        if isinstance(node, Run):
            return self.glyphs(node.text, _effective(node.style, self.bold), node.variant, node.kind, style.level)
        if isinstance(node, Text):
            return self.text(node, style)
        if isinstance(node, Fenced):
            return self.fenced(node, style)
        if isinstance(node, Fraction):
            return self.fraction(node, style)
        if isinstance(node, Radical):
            return self.radical(node, style)
        if isinstance(node, Scripts):
            return self.scripts(node, style)
        if isinstance(node, Limits):
            return self.limits(node, style)
        if isinstance(node, Accent):
            return self.accent(node, style)
        if isinstance(node, Bar):
            return self.bar(node, style)
        if isinstance(node, (Space, Nary, Function)):
            return self.sequence((node,), style)
        raise TypeError(f"未知の構文木の節: {node!r}")

    # 文字

    def glyphs(
        self, text: str, glyph_style: str, variant: str | None, kind: str, level: int, *, script: bool = True
    ) -> _Box:
        """文字の並び。送り幅を加算し、描画範囲は字形の範囲の和にする。"""
        s = self.em(level)
        slanted = glyph_style in ("italic", "bold-italic")
        box = _Box(
            left=kind,
            right=kind,
            string=True,
            single=len(text) == 1,
            upright_start=not slanted,
            slanted_end=slanted,
            text=text,
            plain=variant is None and glyph_style in ("upright", "italic"),
        )
        alternates = script and level > 0 and bool(self.office["script_alternates"])
        x = 0.0
        glyph = None
        for character in text:
            glyph = self.font.glyph_for(character, glyph_style, variant)
            if alternates:
                name = self.font.script_glyph(glyph.name, level)
                if name != glyph.name:
                    glyph = self.font.glyph(name)
            if glyph.bounds is not None:
                left, bottom, right, top = glyph.bounds
                box.ink = _union(box.ink, (x + left * s, x + right * s, top * s, -bottom * s))
            if box.first is None:
                box.first = (glyph, s)
            x += glyph.advance * s
        box.width = x
        if box.ink is not None:
            box.height, box.depth = max(0.0, box.ink[2]), max(0.0, box.ink[3])
        if glyph is not None:
            box.last = (glyph, s)
            box.italic = glyph.italic * s
        return box

    def character(self, char: str, kind: str, level: int) -> _Box:
        return self.glyphs(char, _effective("upright", self.bold), None, kind, level)

    def text(self, node: Text, style: _Style) -> _Box:
        """\\text の文章。書き出しと同じ大きさ（office_math.text_size()：段 1 は 73 %、段 2 以降は 60 % を 1/100 pt に
        丸める）で計測する。倍率は数式の書体によらない（数式の文字は書体の ScriptPercentScaleDown で縮める）。
        """
        if self.text_fonts is None:
            return self.glyphs(node.text, "upright", None, "ord", style.level, script=False)
        s = text_size(self.size, style.level)
        values = self.text_fonts.measure(node.text, s, self.bold)
        try:
            numbers = tuple(float(value) for value in values)
        except (TypeError, ValueError) as error:
            raise ValueError(
                "\\text の計測の値は (幅, 基線より上, 基線より下) か、描画の左右を加えた 5 つの数にする"
            ) from error
        if len(numbers) not in (3, 5):
            raise ValueError("\\text の計測の値は (幅, 基線より上, 基線より下) か、描画の左右を加えた 5 つの数にする")
        if not all(math.isfinite(value) for value in numbers) or numbers[0] < 0:
            raise ValueError("\\text の計測の値が不正: " + repr(values))
        width, ascent, descent = numbers[:3]
        left, right = numbers[3:] if len(numbers) == 5 else (0.0, width)
        box = _Box(width=width, height=max(0.0, ascent), depth=max(0.0, descent), string=True, upright_start=True)
        if node.text.strip():
            # 描画の左右は最初と最後の字形の描画範囲（独立した数式の描画の幅に影響する）
            box.ink = (left, right, ascent, descent)
        return box

    # 伸縮する字形

    @staticmethod
    def _shape(glyph: Glyph) -> _Shape:
        left, bottom, right, top = glyph.bounds if glyph.bounds is not None else (0.0, 0.0, glyph.advance, 0.0)
        return _Shape(glyph.advance, left, right, bottom, top, glyph.italic, False)

    def vertical(self, glyph: Glyph, target: float) -> _Shape:
        """縦に target em 以上の字形を選ぶ。基本の字形、候補の順に探し、なければ組み立てる（MathML Core の伸縮）。"""
        shape = self._shape(glyph)
        if shape.top - shape.bottom >= target:
            return shape
        last = shape
        for name, advance in self.font.vertical_variants(glyph.name):
            candidate = self._shape(self.font.glyph(name))
            last = _Shape(
                candidate.advance,
                candidate.left,
                candidate.right,
                candidate.bottom,
                candidate.top,
                candidate.italic,
                name != glyph.name or self.font.is_extended_shape(name),
            )
            if advance >= target:
                return last
        assembly = self.font.vertical_assembly(glyph.name)
        if assembly is not None:
            built = self.assemble(assembly, target)
            if built is not None:
                return built
        return last

    def assemble(self, assembly: Assembly, target: float) -> _Shape | None:
        """部品を、つなぎの重なりを等しくして target em に組み立てる（OpenType の MATH 表の手順）。"""
        overlap_min = self.font.constant("MinConnectorOverlap")
        extenders = [part for part in assembly.parts if part.extender]
        others = [part for part in assembly.parts if not part.extender]
        extender_advance = sum(part.advance for part in extenders)
        other_advance = sum(part.advance for part in others)
        repeats = 0
        if extenders and extender_advance - overlap_min * len(extenders) > 0:
            needed = (target - other_advance + overlap_min * (len(others) - 1)) / (
                extender_advance - overlap_min * len(extenders)
            )
            repeats = min(1000, max(0, math.ceil(needed)))
        sequence = [part for part in assembly.parts for _ in range(repeats if part.extender else 1)]
        if not sequence:
            return None
        total = sum(part.advance for part in sequence)
        if len(sequence) > 1:
            limit = min(min(a.end, b.start) for a, b in zip(sequence, sequence[1:]))
            overlap = max(overlap_min, min((total - target) / (len(sequence) - 1), limit))
            size = total - overlap * (len(sequence) - 1)
        else:
            size = total
        glyphs = [self.font.glyph(name) for name in dict.fromkeys(part.glyph for part in sequence)]
        bounds = [glyph.bounds for glyph in glyphs if glyph.bounds is not None]
        left = min((bound[0] for bound in bounds), default=0.0)
        right = max((bound[2] for bound in bounds), default=0.0)
        return _Shape(max(glyph.advance for glyph in glyphs), left, right, 0.0, size, assembly.italic, True)

    def horizontal(self, glyph: Glyph, target: float) -> Glyph:
        """横の候補から、幅が target em を超えない最も広い字形を選ぶ（TeX の \\widehat と同じ）。なければ基本の字形。"""
        chosen = glyph
        for name, advance in self.font.horizontal_variants(glyph.name):
            if advance > target:
                break
            chosen = self.font.glyph(name)
        return chosen

    def centered(self, shape: _Shape, level: int, kind: str) -> _Box:
        """字形を数式の軸の上下に等しく置いた箱。"""
        s = self.em(level)
        shift = self.constant("AxisHeight", level) - (shape.top + shape.bottom) / 2 * s
        top, bottom = shape.top * s + shift, -(shape.bottom * s + shift)
        return _Box(
            shape.advance * s,
            max(0.0, top),
            max(0.0, bottom),
            (shape.left * s, shape.right * s, top, bottom),
            shape.italic * s,
            kind,
            kind,
            single=True,
            extended=True,
            upright_start=True,
        )

    # 構造

    def fenced(self, node: Fenced, style: _Style) -> _Box:
        """\\left・\\right の括弧。通常は数式の軸を中心にし、背の高い中身では中身の中心に置く。"""
        level = style.level
        scale = self.em(level)
        parts = [self.sequence(part, style) for part in node.parts]
        axis = self.constant("AxisHeight", level)

        def bodies(*, all_scripts: bool) -> list[_Box]:
            result = []
            for nodes, original in zip(node.parts, parts):
                stripped = _delimiter_body(nodes, all_scripts=all_scripts)
                result.append(original if stripped == nodes else self.sequence(stripped, style))
            return result

        sized = parts if style.display else bodies(all_scripts=False)
        plain = bodies(all_scripts=True)
        top = max((part.height for part in sized), default=0.0)
        bottom = max((part.depth for part in sized), default=0.0)
        extent = max(top - axis, bottom + axis, 0.0)
        plain_extent = max([0.0] + [max(_top(part) - axis, _bottom(part) + axis) for part in plain])
        subformula = plain_extent > self.constant(str(self.office["delimiter_subformula_min_height"]), level)
        center = (top - bottom) / 2 if subformula else axis
        height = top + bottom if subformula else 2.0 * extent
        target = (
            max(
                height * float(self.office["delimiter_factor"]),
                height - scale * float(self.office["delimiter_shortfall_em"]),
            )
            / scale
        )

        def delimiter(char: str | None) -> _Box | None:
            if not char:
                return None
            box = self.centered(self.vertical(self.font.glyph_for(char, "upright", None), target), level, "inner")
            box.ink = _moved(box.ink, 0.0, center - axis)
            box.height, box.depth = max(0.0, _top(box)), max(0.0, _bottom(box))
            return box

        pieces = [delimiter(node.open)]
        for index, part in enumerate(parts):
            if index:
                pieces.append(delimiter(node.separator))
            pieces.append(part)
        pieces.append(delimiter(node.close))
        box = self.pack([piece for piece in pieces if piece is not None] or [_Glue(0.0)], level, spacing=False)
        box.left = box.right = "inner"
        box.extended, box.upright_start, box.slanted_end = True, bool(node.open), False
        box.first = box.last = None
        box.string = box.single = False
        return box

    def fraction(self, node: Fraction, style: _Style) -> _Box:
        level = style.level
        if node.kind == "lin":
            return self.linear_fraction(node, style)
        # 独立した数式では段 0 の入れ子の分数も縮めない。分子・分母の配置自体は文中の組み方にするので、
        # 文字サイズの段と、その分数の display の組み方は別に決める（書き出しの office_math と同じ段）。
        shrink = level > 0 or (not self.display and self.office["inline_fraction_shrink"])
        child = level + 1 if not style.display and shrink else level
        numerator = self.sequence(node.numerator, _Style(child, style.cramped, False))
        denominator = self.sequence(node.denominator, _Style(child, True, False))
        display = style.display
        axis = self.constant("AxisHeight", level)

        def c(name: str) -> float:
            return self.constant(name, level)

        if node.kind == "noBar":
            up = c("StackTopDisplayStyleShiftUp" if display else "StackTopShiftUp")
            down = c("StackBottomDisplayStyleShiftDown" if display else "StackBottomShiftDown")
            gap = (up - _bottom(numerator)) + (down - _top(denominator))
            gap_min = c("StackDisplayStyleGapMin" if display else "StackGapMin")
            if gap < gap_min:
                up += (gap_min - gap) / 2
                down += (gap_min - gap) / 2
            rule = None
        else:
            thickness = c("FractionRuleThickness")
            up = max(
                c("FractionNumeratorDisplayStyleShiftUp" if display else "FractionNumeratorShiftUp"),
                axis
                + thickness / 2
                + c("FractionNumDisplayStyleGapMin" if display else "FractionNumeratorGapMin")
                + _bottom(numerator),
            )
            down = max(
                c("FractionDenominatorDisplayStyleShiftDown" if display else "FractionDenominatorShiftDown"),
                thickness / 2
                + c("FractionDenomDisplayStyleGapMin" if display else "FractionDenominatorGapMin")
                + _top(denominator)
                - axis,
            )
            rule = (axis + thickness / 2, -(axis - thickness / 2))
        nested = node.kind != "noBar" and any(
            isinstance(child, Fraction) and child.kind != "lin" for child in (*node.numerator, *node.denominator)
        )
        padding = self.em(level) * float(self.office["nested_fraction_padding_em"]) if nested else 0.0
        width = max(numerator.width, denominator.width) + padding
        box = _Box(width=width, left="inner", right="inner")
        _place(box, numerator, (width - numerator.width) / 2, up)
        _place(box, denominator, (width - denominator.width) / 2, -down)
        if rule is not None:
            box.ink = _union(box.ink, (-padding, width, rule[0], rule[1]))
            box.height = max(box.height, rule[0])
            box.depth = max(box.depth, rule[1])
        return box

    def linear_fraction(self, node: Fraction, style: _Style) -> _Box:
        """横書きの分数（a/b）。分子、斜線、分母を同じ大きさで並べ、斜線の中心を数式の軸に置く。"""
        slash = str(self.office["fraction_slash"])
        if not self.font.has_character(slash):
            slash = "/"
        separator = self.character(slash, "ord", style.level)
        if separator.ink is not None:
            shift = self.constant("AxisHeight", style.level) - (_top(separator) - _bottom(separator)) / 2
            separator.ink = _moved(separator.ink, 0.0, shift)
            separator.height = max(0.0, separator.ink[2])
            separator.depth = max(0.0, separator.ink[3])
        items = [
            self.sequence(node.numerator, style),
            separator,
            self.sequence(node.denominator, style),
        ]
        box = self.pack(items, style.level, spacing=False)
        box.left = box.right = "ord"
        if self.office["linear_fraction_inner"]:
            # inner として扱い、前の斜体の文字の斜体補正（立体で始まる箱と同じ）と末尾の斜体補正を加える（OFFICE）
            box.left = box.right = "inner"
            box.upright_start = True
            if box.slanted_end:
                box.width += max(0.0, box.italic)
                box.italic, box.slanted_end = 0.0, False
        return box

    def radical(self, node: Radical, style: _Style) -> _Box:
        level = style.level
        s = self.em(level)
        body = self.sequence(node.body, _Style(level, True, style.display))
        gap = self.constant("RadicalDisplayStyleVerticalGap" if style.display else "RadicalVerticalGap", level)
        thickness = self.constant("RadicalRuleThickness", level)
        extra = self.constant("RadicalExtraAscender", level)
        target = _top(body) + _bottom(body) + gap + thickness
        shape = self.vertical(self.font.glyph_for("√", "upright", None), target / s)
        surd_height = (shape.top - shape.bottom) * s
        if self.office["radical_excess_above"] and surd_height > target:
            # 字形が必要な長さより長いときは、余りの半分を中身の上の空きに加える（TeX・MathML Core）
            gap += (surd_height - target) / 2
        top = _top(body) + gap + thickness  # 根号の字形の上端（横線の上端）
        bottom = surd_height - top  # 根号の字形の下端（基線より下が正）
        x = 0.0
        box = _Box(upright_start=not node.degree)
        if self.office["function_apply_inner"] and _ends_with_application(node.body):
            # 中身の最後の節が関数の適用なら、根号の右側も inner（OFFICE["function_apply_inner"]）
            box.right = "inner"
        if node.degree:
            degree = self.sequence(node.degree, _Style(level + 2, False, False))
            before = max(0.0, self.constant("RadicalKernBeforeDegree", level))
            after = max(-degree.width, self.constant("RadicalKernAfterDegree", level))
            raise_ = self.font.constant("RadicalDegreeBottomRaisePercent") * surd_height
            _place(box, degree, before, -bottom + raise_ + _bottom(degree))
            x = before + degree.width + after
        surd = _Box(
            shape.advance * s,
            max(0.0, top),
            max(0.0, bottom),
            (shape.left * s, shape.right * s, top, bottom),
            single=True,
            extended=shape.extended,
        )
        _place(box, surd, x, 0.0)
        x += surd.width
        _place(box, body, x, 0.0)
        box.ink = _union(box.ink, (x, x + body.width, top, -(top - thickness)))
        box.width = x + body.width
        box.height = max(box.height, top + extra)
        if self.office["italic_trailing"]:
            box.italic, box.slanted_end = body.italic, body.slanted_end
        return box

    def scripts(self, node: Scripts, style: _Style) -> _Box:
        level = style.level
        items: list[_Box | _Glue] = []
        for child in node.base:
            self.append(items, child, style)
        base = self.pack(items, level)
        anchor = base
        if self.office["script_last_atom"] and isinstance(items[-1] if items else None, _Box):
            # 添字の位置は土台の並びの最後の原子で決める（(\frac{a}{b})^2 は閉じ括弧に付けたときと同じ位置）
            anchor = items[-1]
        sub = self.sequence(node.sub, _Style(level + 1, True, False)) if node.sub is not None else None
        sup = self.sequence(node.sup, _Style(level + 1, style.cramped, False)) if node.sup is not None else None
        return self.attach(base, sub, sup, style, anchor=anchor)

    def attach(
        self,
        base: _Box,
        sub: _Box | None,
        sup: _Box | None,
        style: _Style,
        *,
        largeop: bool = False,
        integral: bool = False,
        anchor: _Box | None = None,
    ) -> _Box:
        """添字を置く（MathML Core の msub・msup・msubsup）。largeop は大型演算子、integral はそのうちの積分。

        anchor は添字の上下の位置と MathKern を決める原子（土台の並びの最後の原子。既定は土台の全体）。
        """
        level = style.level
        anchor = base if anchor is None else anchor

        def c(name: str) -> float:
            return self.constant(name, level)

        # 字形 1 つや文字の並びでない土台（箱）と伸びた字形では、土台の上下からの位置も確かめる（OpenType の MATH 表）
        boxed = largeop or not anchor.string or anchor.extended
        if largeop and not style.display and not integral and not self.office["inline_nary_script_drop"]:
            # 文中の数式の積分以外の大型演算子は、文字の添字と同じ位置に置く（演算子の字形の上下を使わない）
            boxed = False
        italic = base.italic
        shift_left = largeop or bool(self.office["subscript_italic_shift"])
        sub_x = base.width - (italic if shift_left else 0.0)
        sup_x = base.width + (0.0 if largeop else italic)
        sub_shift = sup_shift = 0.0
        if sub is not None:
            sub_shift = max(c("SubscriptShiftDown"), _top(sub) - c("SubscriptTopMax"))
            if boxed:
                sub_shift = max(sub_shift, _bottom(anchor) + c("SubscriptBaselineDropMin"))
        if sup is not None:
            sup_shift = max(
                c("SuperscriptShiftUpCramped" if style.cramped else "SuperscriptShiftUp"),
                _bottom(sup) + c("SuperscriptBottomMin"),
            )
            if boxed:
                sup_shift = max(sup_shift, _top(anchor) - c("SuperscriptBaselineDropMax"))
        if sub is not None and sup is not None:
            gap = (sub_shift - _top(sub)) + (sup_shift - _bottom(sup))
            gap_min = c("SubSuperscriptGapMin")
            if gap < gap_min:
                delta = c("SuperscriptBottomMaxWithSubscript") - (sup_shift - _bottom(sup))
                if delta > 0:
                    delta = min(delta, gap_min - gap)
                    sup_shift += delta
                    gap += delta
                if gap < gap_min:
                    sub_shift += gap_min - gap
        if self.office["math_kern"] and not largeop:
            if sup is not None:
                sup_x += self.superscript_kern(anchor, sup, sup_shift)
            if sub is not None:
                sub_x += self.subscript_kern(anchor, sub, sub_shift)
        box = _Box(left=base.left, right=base.right, upright_start=base.upright_start)
        _place(box, base, 0.0, 0.0)
        ends = [(base.width, base)]
        if sub is not None:
            _place(box, sub, sub_x, -sub_shift)
            ends.append((sub_x + sub.width, sub))
        if sup is not None:
            _place(box, sup, sup_x, sup_shift)
            ends.append((sup_x + sup.width, sup))
        if largeop and not self.office["nary_space_after_script"]:
            # 大型演算子の添字の後には SpaceAfterScript を入れず、添字の末尾の斜体補正を加える
            box.width = max(end + (max(0.0, part.italic) if part.slanted_end else 0.0) for end, part in ends)
            return box
        right, last = max(ends, key=lambda pair: pair[0])
        box.width = right + c("SpaceAfterScript")
        if self.office["italic_trailing"] and last is not base:
            # 右端の添字の末尾の斜体補正を、式の末尾に伝える
            box.italic, box.slanted_end = last.italic, last.slanted_end
        return box

    def _kern(self, entry: tuple[Glyph, float] | None, corner: str, height: float) -> float:
        """字形の MathKern（pt）。height は字形の基線からの高さ（pt）。"""
        if entry is None:
            return 0.0
        glyph, scale = entry
        return self.font.kern(glyph.name, corner, height / scale) * scale

    @staticmethod
    def _glyph_range(entry: tuple[Glyph, float] | None) -> tuple[float, float] | None:
        """字形の描画範囲の (下端, 上端)（pt、基線から上が正）。"""
        if entry is None or entry[0].bounds is None:
            return None
        glyph, scale = entry
        return glyph.bounds[1] * scale, glyph.bounds[3] * scale

    def superscript_kern(self, base: _Box, sup: _Box, shift: float) -> float:
        """上付きの MathKern。土台の右上と上付きの左下の値の和を 2 つの高さで求め、詰める量の小さい方（大きい値）を使う。

        OpenType の MATH 表の手順の「小さい方」を、詰める量の小さい方とした（PowerPoint の PDF の添字の位置と合う）。
        """
        base_entry = base.last if base.string and not base.extended else None
        sup_entry = sup.first if sup.string else None
        base_range, sup_range = self._glyph_range(base_entry), self._glyph_range(sup_entry)
        if base_entry is None and sup_entry is None:
            return 0.0
        sup_bottom = (sup_range[0] if sup_range else -_bottom(sup)) + shift
        base_top = base_range[1] if base_range else _top(base)
        values = [
            self._kern(base_entry, "TopRight", height) + self._kern(sup_entry, "BottomLeft", height - shift)
            for height in (sup_bottom, base_top)
        ]
        return max(values)

    def subscript_kern(self, base: _Box, sub: _Box, shift: float) -> float:
        """下付きの MathKern。土台の右下と下付きの左上の値の和を 2 つの高さで求め、詰める量の小さい方を使う。"""
        base_entry = base.last if base.string and not base.extended else None
        sub_entry = sub.first if sub.string else None
        base_range, sub_range = self._glyph_range(base_entry), self._glyph_range(sub_entry)
        if base_entry is None and sub_entry is None:
            return 0.0
        sub_top = (sub_range[1] if sub_range else _top(sub)) - shift
        base_bottom = base_range[0] if base_range else -_bottom(base)
        values = [
            self._kern(base_entry, "BottomRight", height) + self._kern(sub_entry, "TopLeft", height + shift)
            for height in (sub_top, base_bottom)
        ]
        return max(values)

    def limits(self, node: Limits, style: _Style) -> _Box:
        level = style.level
        base = self.sequence(node.base, style)
        over = self.sequence(node.over, _Style(level + 1, style.cramped, False)) if node.over is not None else None
        under = self.sequence(node.under, _Style(level + 1, True, False)) if node.under is not None else None
        box = self.stack(base, over, under, level)
        nodes = [item for item in node.base if not isinstance(item, Space)]
        if (
            self.office["bare_function_limits_ord"]
            and len(nodes) == 1
            and isinstance(nodes[0], Function)
            and not nodes[0].body
        ):
            # 土台が本体のない関数名だけ（\underset{i}{\max}）なら、両側とも ord（OFFICE）
            box.left = box.right = "ord"
        return box

    def stack(self, base: _Box, over: _Box | None, under: _Box | None, level: int, *, italic: float = 0.0) -> _Box:
        """上下の添え字を中央にそろえて置く（MathML Core の munderover。大型演算子の上下の添え字の定数）。"""

        def c(name: str) -> float:
            return self.constant(name, level)

        center = base.width / 2
        positions = [(base, 0.0, 0.0)]
        if over is not None:
            rise = _top(base) + max(c("UpperLimitBaselineRiseMin"), c("UpperLimitGapMin") + _bottom(over))
            positions.append((over, center + italic / 2 - over.width / 2, rise))
        if under is not None:
            drop = _bottom(base) + max(c("LowerLimitBaselineDropMin"), c("LowerLimitGapMin") + _top(under))
            positions.append((under, center - italic / 2 - under.width / 2, -drop))
        start = min(x for _, x, _ in positions)
        box = _Box(left=base.left, right=base.right, upright_start=base.upright_start)
        for item, x, y in positions:
            _place(box, item, x - start, y)
        box.width = max(x - start + item.width for item, x, _ in positions)
        return box

    def nary(self, node: Nary, style: _Style) -> _Box:
        """大型演算子と添字。字形を数式の軸の上下に等しく置く。文中の数式では \\limits でも添字を右に置く（OFFICE）。"""
        level = style.level
        limits = node.limits
        if limits == "undOvr" and not style.display and not self.office["inline_nary_limits"]:
            limits = "subSup"
        glyph = self.font.glyph_for(node.char, "upright", None)
        if style.display:
            if self.office["display_operator_variant"] == "next":
                names = [name for name, _ in self.font.vertical_variants(glyph.name) if name != glyph.name]
                shape = self._shape(self.font.glyph(names[0])) if names else self._shape(glyph)
            else:
                shape = self.vertical(glyph, self.font.constant(str(self.office["display_operator_min_height"])))
        else:
            shape = self._shape(glyph)
        operator = self.centered(shape, level, "op")
        sub = self.sequence(node.sub, _Style(level + 1, True, False)) if node.sub is not None else None
        sup = self.sequence(node.sup, _Style(level + 1, style.cramped, False)) if node.sup is not None else None
        if limits == "undOvr":
            box = self.stack(operator, sup, sub, level, italic=operator.italic)
        elif sub is None and sup is None:
            box = operator
        else:
            box = self.attach(operator, sub, sup, style, largeop=True, integral=node.char in INTEGRALS)
        box.left = box.right = "op"
        return box

    def accent(self, node: Accent, style: _Style) -> _Box:
        """記号付き。土台の取り付け位置（なければ中央）に記号の取り付け位置を合わせ、高い土台では持ち上げる。"""
        level = style.level
        s = self.em(level)
        base = self.sequence(node.body, _Style(level, True, style.display))
        upright_start, italic, slanted_end = base.upright_start, base.italic, base.slanted_end
        if base.single and base.string and base.last is not None:
            # i・j は点のない字形にする（GSUB の dtls）
            name = self.font.dotless_glyph(base.last[0].name)
            if name != base.last[0].name:
                base = self.glyphs_from(self.font.glyph(name), level, base.left)
        glyph = self.font.glyph_for(node.char, "upright", None)
        if not base.single:
            width = base.ink[1] - base.ink[0] if base.ink is not None else base.width
            glyph = self.horizontal(glyph, width / s)
        elif _top(base) > self.constant("FlattenedAccentBaseHeight", level):
            glyph = self.font.glyph(self.font.flattened_glyph(glyph.name))
        shift = max(0.0, _top(base) - self.constant("AccentBaseHeight", level))
        if base.single and base.last is not None and base.last[0].accent is not None:
            anchor = base.last[0].accent * base.last[1]
        else:
            anchor = base.width / 2
        if glyph.accent is not None:
            attachment = glyph.accent * s
        elif glyph.bounds is not None:
            attachment = (glyph.bounds[0] + glyph.bounds[2]) / 2 * s
        else:
            attachment = glyph.advance * s / 2
        box = _Box(width=base.width, height=base.height, depth=base.depth, ink=base.ink, upright_start=upright_start)
        if glyph.bounds is not None:
            x = anchor - attachment
            left, bottom, right, top = glyph.bounds
            mark = (x + left * s, x + right * s, top * s + shift, -(bottom * s + shift))
            box.ink = _union(box.ink, mark)
            box.height = max(box.height, mark[2])
        if self.office["italic_trailing"]:
            box.italic, box.slanted_end = italic, slanted_end
        return box

    def glyphs_from(self, glyph: Glyph, level: int, kind: str) -> _Box:
        """字形 1 つの箱（字形名から作る）。"""
        s = self.em(level)
        box = _Box(
            width=glyph.advance * s,
            italic=glyph.italic * s,
            left=kind,
            right=kind,
            first=(glyph, s),
            last=(glyph, s),
            string=True,
            single=True,
        )
        if glyph.bounds is not None:
            left, bottom, right, top = glyph.bounds
            box.ink = (left * s, right * s, top * s, -bottom * s)
            box.height, box.depth = max(0.0, top * s), max(0.0, -bottom * s)
        return box

    def bar(self, node: Bar, style: _Style) -> _Box:
        level = style.level

        def c(name: str) -> float:
            return self.constant(name, level)

        if node.position == "bot":
            base = self.sequence(node.body, style)
            top = -(_bottom(base) + c("UnderbarVerticalGap"))
            bottom = top - c("UnderbarRuleThickness")
            box = _Box(
                width=base.width,
                height=base.height,
                depth=max(base.depth, -bottom + c("UnderbarExtraDescender")),
                ink=_union(base.ink, (0.0, base.width, top, -bottom)),
            )
        else:
            base = self.sequence(node.body, _Style(level, True, style.display))
            bottom = _top(base) + c("OverbarVerticalGap")
            top = bottom + c("OverbarRuleThickness")
            box = _Box(
                width=base.width,
                height=max(base.height, top + c("OverbarExtraAscender")),
                depth=base.depth,
                ink=_union(base.ink, (0.0, base.width, top, -bottom)),
            )
        if self.office["italic_trailing"]:
            box.italic, box.slanted_end = base.italic, base.slanted_end
        return box


def layout(
    nodes, font: MathFont, size: float, *, display: bool, bold: bool = False, text_fonts: TextFonts | None = None
) -> Box:
    """構文木（`syntax.parse()` の結果）を組み、寸法を pt で返す。display は独立した数式の組み方。"""
    size = float(size)
    if not math.isfinite(size) or size <= 0:
        raise ValueError("文字サイズは正の有限値にする")
    if text_fonts is not None and not isinstance(text_fonts, TextFonts):
        raise TypeError("text_fonts は TextFonts で指定する")
    engine = _Engine(font, size, bool(bold), text_fonts, display=bool(display))
    box = engine.sequence(tuple(nodes), _Style(0, False, bool(display)))
    width = box.width
    if not display and engine.office["italic_trailing"] and box.slanted_end:
        # 文中の数式の末尾の斜体の文字の斜体補正は、送り幅に含まれる
        width += max(0.0, box.italic)
    ink = box.ink if box.ink is not None else (0.0, 0.0, 0.0, 0.0)
    return Box(
        float(width),
        float(box.height),
        float(box.depth),
        tuple(float(value) for value in ink),
        float(box.italic),
        box.left,
    )


class MathMeasurer:
    """指定した数式フォントで、数式の幅・全高・基線より下の高さを pt で返す。スレッド間で共有できる。"""

    def __init__(self, path, index=0, family="Cambria Math", functions=()):
        self.functions = function_names(functions)
        self.font = MathFont(path, index=index, family=family)
        self.path = self.font.path
        self.index = index
        self.family = family
        self._cache: OrderedDict[tuple, Box] = OrderedDict()
        self._lock = RLock()
        self._closed = False

    def box(self, expr, size=20, bold=False, *, display=False, text_fonts=None) -> Box:
        """組んだ数式の寸法（`Box`）。最大 4096 件の結果を保持し、保持に当たれば式を解析し直さない。"""
        if not isinstance(expr, str):
            raise TypeError("数式は文字列で指定する")
        size = float(size)
        if not math.isfinite(size) or size <= 0:
            raise ValueError("文字サイズは正の有限値にする")
        if text_fonts is not None and not isinstance(text_fonts, TextFonts):
            raise TypeError("text_fonts は TextFonts で指定する")
        key = (expr, size, bool(bold), bool(display), None if text_fonts is None else text_fonts.key)
        with self._lock:
            if self._closed:
                raise RuntimeError("閉じた数式計測器は使用できない")
            cached = self._cache.get(key)
            if cached is not None:
                self._cache.move_to_end(key)
                return cached
            # 保持に当たらないときだけ解析する（出力と同じ syntax.parse() で記法を検査する）
            nodes = parse(expr, self.functions)
            result = layout(nodes, self.font, size, display=display, bold=bold, text_fonts=text_fonts)
            self._cache[key] = result
            if len(self._cache) > CACHE_SIZE:
                self._cache.popitem(last=False)
            return result

    def measure(self, expr, size=20, bold=False, *, display=False, text_fonts=None) -> tuple[float, float, float]:
        """幅・全高・基線より下の高さ（pt）。全高と深さは基線を含む範囲（0 以上）で返す。"""
        result = self.box(expr, size, bold, display=display, text_fonts=text_fonts)
        depth = max(0.0, result.depth)
        return result.width, max(0.0, result.height) + depth, depth

    def ink(self, expr, size=20, bold=False, *, display=False, text_fonts=None) -> tuple[float, float, float, float]:
        """描画範囲 (左, 右, 基線より上, 基線より下)（pt）。PowerPoint との照合に使う。"""
        return self.box(expr, size, bold, display=display, text_fonts=text_fonts).ink

    def close(self) -> None:
        with self._lock:
            self._closed = True
            self._cache.clear()
            self.font.close()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()


# ---- PowerPoint の照合データとの比較 ----

# 照合の合格の基準（em）。幅の差が width_share の割合の式で width_em 以内、すべての式で width_max_em 以内、
# 描画の上下の差がすべての式で vertical_em 以内。
REFERENCE_LIMITS: Mapping[str, float] = {"width_share": 0.9, "width_em": 0.05, "width_max_em": 0.12, "vertical_em": 0.1}


def _reference_number(value, label: str) -> float:
    """照合データの数。真偽値・数でない値・非有限の値は ValueError。"""
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
        raise ValueError(f"照合データの {label} は有限の数にする: {value!r}")
    return float(value)


def _reference_entries(data) -> tuple[float, list[tuple[str, str, dict, str | None]]]:
    """照合データの文字サイズと、各式の (式, mode, PowerPoint の値, 比べない理由)。形が違えば ValueError。"""
    if not isinstance(data, Mapping) or not isinstance(data.get("items"), list):
        raise ValueError("照合データは、式の一覧 items を持つオブジェクトにする")
    size = _reference_number(data.get("size_pt"), "size_pt")
    if size <= 0:
        raise ValueError(f"照合データの size_pt は正の数にする: {size!r}")
    entries = []
    for index, item in enumerate(data["items"]):
        if not isinstance(item, Mapping) or not isinstance(item.get("expression"), str):
            raise ValueError(f"照合データの items[{index}] は、式 expression（文字列）を持つオブジェクトにする")
        expression, mode = item["expression"], item.get("mode", "inline")
        label = f"items[{index}]（{expression}）"
        if mode == "display":
            values = {
                key: _reference_number(item.get(key), f"{label} の {key}") for key in ("ink_width_pt", "ink_height_pt")
            }
        elif mode == "inline":
            values = {"advance_pt": _reference_number(item.get("advance_pt"), f"{label} の advance_pt")}
            ink = item.get("ink_pt")
            if ink is not None:
                if not isinstance(ink, list) or len(ink) != 4:
                    raise ValueError(
                        f"照合データの {label} の ink_pt は 4 つの数（左, 右, 基線より上, 基線より下）にする"
                    )
                values["ink_pt"] = [_reference_number(value, f"{label} の ink_pt") for value in ink]
        else:
            raise ValueError(f"照合データの {label} の mode は inline か display にする: {mode!r}")
        excluded = item.get("excluded")
        if excluded is not None and (not isinstance(excluded, str) or not excluded.strip()):
            raise ValueError(f"照合データの {label} の excluded は、比べない理由の文字列にする: {excluded!r}")
        entries.append((expression, mode, values, excluded))
    return size, entries


def compare_reference(
    data: Mapping,
    measurer: MathMeasurer,
    *,
    text_fonts: TextFonts | None = None,
    limits: Mapping[str, float] = REFERENCE_LIMITS,
) -> dict:
    """照合データ（`tests/data/powerpoint_math.json` の形）の各式を計測し、PowerPoint の値との差と合否を返す。

    照合データの size_pt は文字サイズ（pt）、items の各項目は expression・mode（inline か display）と次の値である。
    excluded（比べない理由の文字列）を持つ項目は比べずに skipped に入れる（書き出しを変えて値が古くなった式など）。
    文中の数式（inline）は送り幅 advance_pt（目印の [ の右端と ] の左端の差）と描画範囲 ink_pt
    （左, 右, 基線より上, 基線より下。原点は [ の右端と基線。描画がなければ null）、独立した数式（display）は
    描画範囲の幅と高さ ink_width_pt・ink_height_pt を持つ。形が違えば ValueError にする。
    幅の差は inline では送り幅、display では描画範囲の幅で求め、描画の上下の差は inline では上と下、display では高さで求める。
    差は「計測 − PowerPoint」を em で表す。計測できない式の ValueError には、その式を注記（add_note）に加える。
    \\text を含む式は text_fonts で計測する。text_fonts がなければ、その式を比べずに skipped に入れる。
    戻り値: {"items", "skipped", "count", "width_share", "width_max_em", "vertical_max_em", "passed"}。
    items の各項目は expression・mode・width_em・vertical_em・measured_pt（送り幅と描画範囲）を持ち、
    inline では描画範囲の幅の差 ink_width_em も持つ（参考）。
    """
    size, entries = _reference_entries(data)
    items, skipped = [], []
    for expression, mode, actual, excluded in entries:
        if excluded is not None:
            skipped.append({"expression": expression, "mode": mode, "reason": excluded})
            continue
        if text_fonts is None and "\\text" in expression:
            skipped.append({"expression": expression, "mode": mode, "reason": "\\text を測る文字用の書体がない"})
            continue
        try:
            box = measurer.box(expression, size, display=mode == "display", text_fonts=text_fonts)
        except ValueError as error:
            error.add_note(f"照合データの式: {expression}（{mode}）")
            raise
        ink_width = box.ink[1] - box.ink[0]
        row = {"expression": expression, "mode": mode}
        if mode == "display":
            width = ink_width - actual["ink_width_pt"]
            vertical = [(box.ink[2] + box.ink[3]) - actual["ink_height_pt"]]
        else:
            width = box.width - actual["advance_pt"]
            ink = actual.get("ink_pt")
            vertical = [box.ink[2] - ink[2], box.ink[3] - ink[3]] if ink else []
            if ink:
                row["ink_width_em"] = (ink_width - (ink[1] - ink[0])) / size
        row.update(
            width_em=width / size,
            vertical_em=max((abs(value) for value in vertical), default=0.0) / size,
            measured_pt=[box.width, *box.ink],
        )
        items.append(row)
    count = len(items)
    within = sum(abs(row["width_em"]) <= limits["width_em"] for row in items)
    share = within / count if count else 1.0
    width_max = max((abs(row["width_em"]) for row in items), default=0.0)
    vertical_max = max((row["vertical_em"] for row in items), default=0.0)
    passed = (
        count > 0
        and share >= limits["width_share"]
        and width_max <= limits["width_max_em"]
        and vertical_max <= limits["vertical_em"]
    )
    return {
        "items": items,
        "skipped": skipped,
        "count": count,
        "width_share": share,
        "width_max_em": width_max,
        "vertical_max_em": vertical_max,
        "passed": passed,
    }


def reference_font_matches(data: Mapping, path) -> bool:
    """照合データの書体のハッシュ値（font.sha256）が、そのフォントファイルと一致するか。ハッシュ値がなければ偽。"""
    font = data.get("font") if isinstance(data, Mapping) else None
    expected = font.get("sha256") if isinstance(font, Mapping) else None
    return isinstance(expected, str) and bool(expected) and expected == file_hash(path)


class _FileTextFonts:
    """照合のコマンドで使う \\text の簡易な計測（送り幅の和と、字形の描画範囲の上下と左右）。kern と太字は使わない。

    書体は文字の計測（measure.font_runs）と同じく、和文の文字と、PowerPoint が既定で和文の書体で描く記号
    （measure.uses_east_asian_font。和文の書体に字形があるとき）を和文の書体、ほかを欧文の書体にする。
    """

    def __init__(self, japanese: tuple[Path, int], latin: tuple[Path, int]) -> None:
        from fontTools.ttLib import TTFont

        self._faces = {}
        opened = []
        try:
            for role, (path, index) in (("japanese", japanese), ("latin", latin)):
                with path.open("rb") as stream:
                    collection = stream.read(4) == b"ttcf"
                face = TTFont(path, fontNumber=index if collection else -1, lazy=True)
                opened.append(face)
                outlines = face["glyf"] if "glyf" in face else None
                self._faces[role] = (
                    face,
                    face.getBestCmap() or {},
                    face["hmtx"].metrics,
                    face["head"].unitsPerEm,
                    outlines,
                )
        except Exception:
            for face in opened:
                face.close()
            raise

    def measure(self, text: str, size: float, bold: bool) -> tuple[float, float, float, float, float]:
        """(送り幅, 基線より上, 基線より下, 描画の左, 描画の右)（pt）。"""
        width, top, bottom = 0.0, 0.0, 0.0
        left = right = None
        for character in text:
            role = "japanese" if unicodedata.east_asian_width(character) in ("W", "F", "H") else "latin"
            if role == "latin" and uses_east_asian_font(character) and ord(character) in self._faces["japanese"][1]:
                role = "japanese"
            face, cmap, metrics, units, outlines = self._faces[role]
            name = cmap.get(ord(character))
            if name is None:
                raise ValueError(f"文字用の書体に字形がない: {character!r}")
            if outlines is not None:
                outline = outlines[name]
                if getattr(outline, "numberOfContours", 0) and hasattr(outline, "yMax"):
                    top = max(top, outline.yMax / units * size)
                    bottom = max(bottom, -outline.yMin / units * size)
                    if left is None:
                        left = width + outline.xMin / units * size
                    right = width + outline.xMax / units * size
            width += metrics[name][0] / units * size
        return width, top, bottom, 0.0 if left is None else left, width if right is None else right

    def close(self) -> None:
        for face, *_ in self._faces.values():
            face.close()


def _font_argument(value: str) -> tuple[Path, int]:
    """PATH か PATH:INDEX。"""
    path, _, index = value.rpartition(":")
    if path and index.isdigit():
        return Path(path), int(index)
    return Path(value), 0


def _compare_files(arguments) -> dict | None:
    """コマンドの引数のファイルで照合する。照合データの書体のハッシュ値がないか、フォントと違えば None。"""
    data = json.loads(Path(arguments.reference).read_text(encoding="utf-8"))
    _reference_entries(data)
    font_path, font_index = _font_argument(arguments.font)
    if not reference_font_matches(data, font_path):
        return None
    font = data.get("font")
    family = arguments.family or (font.get("family") if isinstance(font, Mapping) else None) or "Cambria Math"
    text_fonts = None
    files = None
    if arguments.latin and arguments.japanese:
        files = _FileTextFonts(_font_argument(arguments.japanese), _font_argument(arguments.latin))
        text_fonts = TextFonts(("files", arguments.japanese, arguments.latin), files.measure)
    try:
        with MathMeasurer(font_path, index=font_index, family=family) as measurer:
            return compare_reference(data, measurer, text_fonts=text_fonts)
    finally:
        if files is not None:
            files.close()


def main(argv: list[str] | None = None) -> int:
    """照合データと計測を比べる開発用のコマンド。

    合格なら 0、不合格なら 1、照合できなければ 2 を返す。2 は、照合データの書体のハッシュ値がない・フォントと違う場合と、
    ファイル・照合データの形・引数の誤り、計測できない式、計測の内部の誤りである（引数の誤りは argparse が終了コード 2 で終える）。
    照合できない理由は標準エラーに出す。1 は計測が基準に入らないときだけにする。
    """
    import argparse
    import sys
    import traceback

    from fontTools.ttLib import TTLibError

    parser = argparse.ArgumentParser(
        prog="python -m slidekit.math.metrics", description="PowerPoint の照合データと数式の計測を比べる。"
    )
    parser.add_argument("reference", help="照合データ（tests/data/powerpoint_math.json の形）")
    parser.add_argument("--font", required=True, help="数式フォントのファイル（PATH か PATH:INDEX）")
    parser.add_argument("--family", default=None, help="書体名（既定は照合データの font.family）")
    parser.add_argument(
        "--latin", help="\\text の欧文の書体のファイル（PATH か PATH:INDEX。--japanese と合わせて指定する）"
    )
    parser.add_argument(
        "--japanese", help="\\text の和文の書体のファイル（PATH か PATH:INDEX。--latin と合わせて指定する）"
    )
    parser.add_argument("--json", action="store_true", help="結果を JSON で出す")
    arguments = parser.parse_args(argv)
    if bool(arguments.latin) != bool(arguments.japanese):
        parser.error("--latin と --japanese は両方指定する（どちらもなければ \\text の式を比べない）")
    try:
        result = _compare_files(arguments)
    except (OSError, ValueError, KeyError, TTLibError) as error:
        notes = "".join("\n" + note for note in getattr(error, "__notes__", ()))
        print(f"照合できない: {type(error).__name__}: {error}{notes}", file=sys.stderr)
        return 2
    except Exception:
        # 計測の内部の誤りも不合格（1）と区別する
        traceback.print_exc()
        print("照合できない: 予期しない例外が起きた", file=sys.stderr)
        return 2
    if result is None:
        print(
            "照合できない: 照合データの書体のハッシュ値（font.sha256）がないか、フォントと違う。照合データを作り直す",
            file=sys.stderr,
        )
        return 2
    if arguments.json:
        print(json.dumps(result, ensure_ascii=False, indent=1))
    else:
        print(f"式 {result['count']} 件（比べなかった式 {len(result['skipped'])} 件）")
        print(
            f"幅の差が {REFERENCE_LIMITS['width_em']} em 以内の割合: {result['width_share']:.3f}（基準 {REFERENCE_LIMITS['width_share']}）"
        )
        print(f"幅の差の最大: {result['width_max_em']:.3f} em（基準 {REFERENCE_LIMITS['width_max_em']}）")
        print(f"描画の上下の差の最大: {result['vertical_max_em']:.3f} em（基準 {REFERENCE_LIMITS['vertical_em']}）")
        worst = sorted(result["items"], key=lambda row: -max(abs(row["width_em"]), row["vertical_em"]))[:10]
        for row in worst:
            print(
                f"  {row['mode']:7} 幅 {row['width_em']:+.3f} em  上下 {row['vertical_em']:.3f} em  {row['expression']}"
            )
        print("合格" if result["passed"] else "不合格")
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
