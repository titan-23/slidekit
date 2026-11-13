r"""短い数式を PowerPoint の編集可能な Office Math へ変換する。

``append_math(paragraph._p, r'O(\sqrt{m}\log N)', size=20)`` のように使う。本文との境界を示す ``$`` は呼び出し側で取り除く。
記法は ``syntax.parse()`` が解析し、このモジュールはその構文木から XML を作る（命令の一覧は ``docs/math.md``）。

- 普通の括弧は通常の文字として出力し、``\left``・``\right`` で囲んだ括弧だけを伸びる括弧（``m:d``）にする。
- 置き方は 2 つある。``display=True`` は独立した数式（``m:oMathPara`` と ``m:jc`` の揃え）、
  ``display=False``（既定）は文中の数式（``m:oMath`` の中身を改行しない箱 ``m:box`` で包む）である。
  ``math_mode()`` が要素から決める。中央揃えの独立した数式の後ろに PowerPoint が加える幅は ``display_trailing_width()`` で求める。
- ``\text{...}`` は通常の文章（``m:nor``）として、``text_fonts`` の書体で出力する。添字などの中では縮めた大きさ
  （``text_size()``。計測も同じ大きさを使う）を書く。
- log、min、max などは立体の関数名として扱い、資料に固有の名前は引数で追加する。
- 未知の命令や閉じていない括弧は ``syntax.MathSyntaxError`` にして誤記を検出する。

Office Math の DrawingML 表現に合わせ、文字の書式には a:rPr を使う。画像や通常文字で数式を代用しない。
"""

from __future__ import annotations

from dataclasses import dataclass, replace
import math
import re
from typing import Iterable, Mapping

from lxml import etree

from ..measure import uses_east_asian_font
from ..ooxml import NS, tag
from . import commands
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
    _east_asian,
    parse,
)

__all__ = [
    "COMMANDS",
    "DEFAULT_FUNCTIONS",
    "DISPLAY_TRAILING_EM",
    "NS",
    "TEXT_SCRIPT_SCALES",
    "append_math",
    "display_trailing_width",
    "math_element",
    "math_mode",
    "math_options",
    "supported_commands",
    "text_size",
]

# 立体で書く既定の関数名（commands.FUNCTIONS の別名）
DEFAULT_FUNCTIONS = commands.FUNCTIONS
# 命令名と出力の文字（記号・括弧・空白の命令）。既存の参照のために残す。
COMMANDS: dict[str, str] = {name: symbol.char for name, symbol in commands.SYMBOLS.items()}
COMMANDS.update({token[1:]: char for token, char in commands.DELIMITERS.items() if token.startswith("\\")})
COMMANDS.update({token[1:]: char for token, (char, _) in commands.SPACES.items() if token.startswith("\\")})
# 受け付ける英字の命令名（記述規則の数式の候補が使う）
supported_commands = commands.command_names

# 中央揃えの独立した数式では、PowerPoint は式の後ろに幅を加えてから中央に置く（式はその半分だけ左へずれる）。
# PowerPoint 16.0（build 20430）で確かめた規則：幅は「段落の末尾の書式（a:endParaRPr）の欧文書体の空白の送り幅（em）
# × 数式の文字サイズ」で、a:endParaRPr の sz にはよらない。書体は a:endParaRPr の a:latin、文字枠の a:lstStyle、
# テーマの本文の欧文書体の順に決まる。左揃えと右揃えでは位置がずれない。
# 書き出しは a:endParaRPr に要素の font_latin を書き、この幅（display_trailing_width）を a:pPr/@marL にして中心をそろえる。
# DISPLAY_TRAILING_EM は空白の送り幅に加える残りの幅（em）。確認では残りが 0.01 em 未満だったので 0 とした。
# 計測した幅（2026-10-12）：Calibri で 0.218〜0.233 em（空白の送り幅 0.2261 em）、Arial で 0.272〜0.277 em
# （0.2778 em）、Meiryo で 0.335〜0.339 em（0.3398 em）。補正した後の中央揃えの式の中心のずれは 0.15 pt 以内だった。
DISPLAY_TRAILING_EM = 0.0

# 添字・上下の添え字・文中の分数の中身と根号の次数にある \text（m:nor）の文字の大きさの倍率（段 0・1・2 以降）。
# PowerPoint は m:nor の文字を縮めずに a:rPr の sz のとおりに描くので、段に応じて縮めた大きさを書く。
# 値は Cambria Math の ScriptPercentScaleDown（73 %）と ScriptScriptPercentScaleDown（60 %）。添字の中で、和文と欧文の
# 書体を保ったまま計測と合うことを確かめた（PowerPoint 16.0.20430.20146、2026-10-12）。
# 数式の書体を替えても同じ倍率を使い、計測（math.metrics）も text_size() の大きさを使う（書き出しと計測を合わせる）。
TEXT_SCRIPT_SCALES = (1.0, 0.73, 0.60)

_JUSTIFICATION = {"left": "left", "center": "center", "right": "right", "justify": "left"}
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
# 書式と要素の太字から m:sty の値を決める
_STY = {
    ("italic", False): "i",
    ("italic", True): "bi",
    ("upright", False): "p",
    ("upright", True): "b",
    ("bold", False): "b",
    ("bold", True): "b",
    ("bold-italic", False): "bi",
    ("bold-italic", True): "bi",
}


@dataclass(frozen=True)
class _Options:
    size: float
    color: str
    bold: bool
    font: str
    text_fonts: tuple[str, str] | None
    level: int = 0  # 文字の大きさの段（\text の大きさに使う）
    display: bool = False  # 独立した数式か（段 0 の分数の中身を縮めない）


def text_size(size: float, level: int) -> float:
    """段 level の \\text の文字サイズ（pt）。TEXT_SCRIPT_SCALES を掛け、a:rPr の sz と同じく 1/100 pt に丸める。"""
    scale = TEXT_SCRIPT_SCALES[min(max(level, 0), len(TEXT_SCRIPT_SCALES) - 1)]
    return round(size * scale * 100) / 100


def _inner(options: _Options, levels: int = 1) -> _Options:
    """添字・上下の添え字・文中の分数の中身と根号の次数の書式。段を進める。"""
    return replace(options, level=options.level + levels)


def _el(name, **attrs):
    prefix, local = name.split(":")
    node = etree.Element(tag(prefix, local))
    for key, value in attrs.items():
        node.set(tag("m", key) if prefix == "m" else key, str(value))
    return node


def _drawing(options, *, italic=False, bold=False, latin=None, east_asian=None, lang="en-US", size=None):
    size = options.size if size is None else size
    properties = _el("a:rPr", lang=lang, sz=round(size * 100), b=int(options.bold or bold), i=int(italic), dirty="0")
    fill = _el("a:solidFill")
    fill.append(_el("a:srgbClr", val=options.color))
    properties.append(fill)
    latin = latin or options.font
    for script, family in (("latin", latin), ("ea", east_asian or options.font), ("cs", latin)):
        properties.append(_el("a:" + script, typeface=family))
    return properties


def _value(node, text):
    value = _el("m:t")
    value.text = text
    value.set(_XML_SPACE, "preserve")
    node.append(value)
    return node


def _run(text, style, variant, options):
    node = _el("m:r")
    properties = _el("m:rPr")
    if variant:
        properties.append(_el("m:scr", val=variant))
    properties.append(_el("m:sty", val=_STY[(style, options.bold)]))
    node.append(properties)
    node.append(_drawing(options, italic=style in ("italic", "bold-italic"), bold=style in ("bold", "bold-italic")))
    return _value(node, text)


def _text_run(text, options):
    # 通常の文章（m:nor）。m:nor と m:sty は同時に書くことができないので m:sty を置かない。中身は構文解析のとおりで、
    # - は解析でハイフン（U+2010）にしてある（PowerPoint は m:nor の - もマイナス記号で描く）。
    # PowerPoint は m:nor の文字を添字の中でも縮めないので、段に応じて縮めた大きさを書く（数式の文字は PowerPoint が縮める）。
    # 和文の文字か、PowerPoint が lang="ja-JP" で和文の書体（a:ea）で描く記号（“ ” ‘ ’ × ○ など。measure.uses_east_asian_font）を
    # 含めば lang="ja-JP" にする。計測（文字の計測）では、これらの記号に和文の書体を使う。lang="en-US" では、PowerPoint は
    # “ ” ‘ ’ を欧文の書体で描く
    node = _el("m:r")
    properties = _el("m:rPr")
    properties.append(_el("m:nor"))
    node.append(properties)
    japanese, latin = options.text_fonts or (options.font, options.font)
    japanese_lang = _east_asian(text) or any(uses_east_asian_font(char) for char in text)
    node.append(
        _drawing(
            options,
            latin=latin,
            east_asian=japanese,
            lang="ja-JP" if japanese_lang else "en-US",
            size=text_size(options.size, options.level),
        )
    )
    return _value(node, text)


def _construct(kind, arguments, settings, options):
    node = _el("m:" + kind)
    properties = _el("m:" + kind + "Pr")
    for key, value in settings:
        properties.append(_el("m:" + key, val=value))
    control = _el("m:ctrlPr")
    control.append(_drawing(options))
    properties.append(control)
    node.append(properties)
    for name, values in arguments:
        argument = _el("m:" + name)
        argument.extend(values)
        node.append(argument)
    return node


def _omml(nodes, options):
    """構文木を Office Math の要素の列にする。"""
    result = []
    for node in nodes:
        result.extend(_node(node, options))
    return result


def _node(node, options):
    if isinstance(node, Run):
        return [_run(node.text, node.style, node.variant, options)]
    if isinstance(node, Text):
        return [_text_run(node.text, options)]
    if isinstance(node, Space):
        return [_run(node.char, "upright", None, options)] if node.char else []
    if isinstance(node, Fenced):
        if not node.stretchy:
            # 普通の括弧は通常の文字にし、中身の高さに合わせて伸ばさない
            content = _omml(node.parts[0], options)
            return [_run(node.open, "upright", None, options), *content, _run(node.close, "upright", None, options)]
        settings = [("begChr", node.open)]
        if node.separator is not None:
            settings.append(("sepChr", node.separator))
        settings += [("endChr", node.close), ("grow", "1")]
        return [_construct("d", [("e", _omml(part, options)) for part in node.parts], settings, options)]
    if isinstance(node, Fraction):
        # 分子と分母の段：横書きは同じ段。独立した数式の段 0 の分数（入れ子の分数も）は同じ段で、それ以外は 1 段進める。
        # PowerPoint は独立した数式の入れ子の分数の分子と分母を縮めない（20 pt の式で 20.04 pt。文中の数式では縮める）ので、
        # \text も縮めない。計測（math.metrics）も同じ段で計測する。
        if node.kind == "lin" or (options.display and options.level == 0):
            inner = options
        else:
            inner = _inner(options)
        arguments = [("num", _omml(node.numerator, inner)), ("den", _omml(node.denominator, inner))]
        return [_construct("f", arguments, [("type", node.kind)], options)]
    if isinstance(node, Radical):
        arguments = [("deg", _omml(node.degree, _inner(options, 2))), ("e", _omml(node.body, options))]
        return [_construct("rad", arguments, [("degHide", "0" if node.degree else "1")], options)]
    if isinstance(node, Scripts):
        base = ("e", _omml(node.base, options))
        script = _inner(options)
        if node.sub is not None and node.sup is not None:
            arguments = [base, ("sub", _omml(node.sub, script)), ("sup", _omml(node.sup, script))]
            return [_construct("sSubSup", arguments, [], options)]
        if node.sub is not None:
            return [_construct("sSub", [base, ("sub", _omml(node.sub, script))], [], options)]
        return [_construct("sSup", [base, ("sup", _omml(node.sup, script))], [], options)]
    if isinstance(node, Limits):
        values = _omml(node.base, options)
        if node.under is not None:
            values = [_construct("limLow", [("e", values), ("lim", _omml(node.under, _inner(options)))], [], options)]
        if node.over is not None:
            values = [_construct("limUpp", [("e", values), ("lim", _omml(node.over, _inner(options)))], [], options)]
        return values
    if isinstance(node, Nary):
        script = _inner(options)
        arguments = [
            ("sub", _omml(node.sub or (), script)),
            ("sup", _omml(node.sup or (), script)),
            ("e", _omml(node.body, options)),
        ]
        settings = [
            ("chr", node.char),
            ("limLoc", node.limits),
            ("subHide", "0" if node.sub else "1"),
            ("supHide", "0" if node.sup else "1"),
        ]
        return [_construct("nary", arguments, settings, options)]
    if isinstance(node, Function):
        name = _omml(node.name, options)
        if not node.body:
            return name
        return [_construct("func", [("fName", name), ("e", _omml(node.body, options))], [], options)]
    if isinstance(node, Accent):
        return [_construct("acc", [("e", _omml(node.body, options))], [("chr", node.char)], options)]
    if isinstance(node, Bar):
        return [_construct("bar", [("e", _omml(node.body, options))], [("pos", node.position)], options)]
    raise TypeError(f"未知の構文木の節: {node!r}")


def _options(size, color, bold, font, text_fonts):
    if not math.isfinite(size) or size <= 0:
        raise ValueError("文字サイズは正の有限値にする")
    if not isinstance(font, str) or not font.strip():
        raise ValueError("数式用のフォント名を指定する")
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", color):
        raise ValueError("色は 6 桁の 16 進数で指定する")
    if text_fonts is not None:
        if (
            isinstance(text_fonts, str)
            or len(tuple(text_fonts)) != 2
            or not all(isinstance(family, str) and family.strip() for family in text_fonts)
        ):
            raise ValueError("\\text の書体は (日本語の書体, 欧文の書体) の組で指定する")
        text_fonts = tuple(text_fonts)
    return _Options(size, color, bool(bold), font, text_fonts)


def math_element(
    expr,
    size=20,
    color="17212D",
    bold=False,
    font="Cambria Math",
    functions=(),
    *,
    display=False,
    align="left",
    text_fonts=None,
):
    """数式に対応する a14:m 要素を作る。

    display=False は a14:m > m:oMath > m:box（文中の数式）、True は a14:m > m:oMathPara > m:oMath（独立した数式）を返す。
    文中の数式の m:box は m:noBreak を持ち、PowerPoint が行の残りに入らない式を途中の演算子の後で改行しないようにする
    （式ごと次の行へ送る。計測も数式の断片を折り返しで分けない）。空の式は包まない。
    独立した数式の揃え（m:jc）は align から決める（justify は left）。
    text_fonts は \\text に使う (日本語の書体, 欧文の書体)。None は font を使う。
    """
    if not isinstance(expr, str):
        raise TypeError("数式は文字列で指定する")
    options = replace(_options(size, color, bold, font, text_fonts), display=bool(display))
    if align not in _JUSTIFICATION:
        raise ValueError("数式の揃えは left・center・right・justify のいずれかにする")
    nodes = parse(expr, functions)
    wrapper = etree.Element(tag("a14", "m"), nsmap={"a14": NS["a14"], "m": NS["m"]})
    if display:
        paragraph = etree.SubElement(wrapper, tag("m", "oMathPara"))
        properties = etree.SubElement(paragraph, tag("m", "oMathParaPr"))
        properties.append(_el("m:jc", val=_JUSTIFICATION[align]))
        equation = etree.SubElement(paragraph, tag("m", "oMath"))
    else:
        equation = etree.SubElement(wrapper, tag("m", "oMath"))
        if nodes:
            box = etree.SubElement(equation, tag("m", "box"))
            properties = etree.SubElement(box, tag("m", "boxPr"))
            properties.append(_el("m:noBreak", val="1"))
            equation = etree.SubElement(box, tag("m", "e"))
    equation.extend(_omml(nodes, options))
    return wrapper


def append_math(
    paragraph,
    expr,
    size=20,
    color="17212D",
    bold=False,
    font="Cambria Math",
    functions=(),
    *,
    display=False,
    align="left",
    text_fonts=None,
):
    """DrawingML の a:p に数式を追加し、追加した a14:m を返す。a:endParaRPr があればその前に置く。"""
    paragraph_xml = getattr(paragraph, "_p", paragraph)
    wrapper = math_element(
        expr,
        size=size,
        color=color,
        bold=bold,
        font=font,
        functions=functions,
        display=display,
        align=align,
        text_fonts=text_fonts,
    )
    end = paragraph_xml.find(tag("a", "endParaRPr"))
    if end is None:
        paragraph_xml.append(wrapper)
    else:
        paragraph_xml.insert(paragraph_xml.index(end), wrapper)
    return wrapper


def display_trailing_width(space_advance_em: float, size: float) -> float:
    """中央揃えの独立した数式の後ろに PowerPoint が加える幅（pt）。

    space_advance_em は段落の末尾の書式の欧文書体の空白の送り幅（em）、size は数式の文字サイズ（pt）。
    """
    if (
        isinstance(space_advance_em, bool)
        or not isinstance(space_advance_em, (int, float))
        or not math.isfinite(space_advance_em)
        or space_advance_em < 0
    ):
        raise ValueError("空白の送り幅は 0 以上の有限値にする")
    if isinstance(size, bool) or not isinstance(size, (int, float)) or not math.isfinite(size) or size <= 0:
        raise ValueError("文字サイズは正の有限値にする")
    return (space_advance_em + DISPLAY_TRAILING_EM) * size


def math_mode(element: Mapping) -> str:
    """置き方を "display" か "inline" で返す。種類が equation で、空でない断片が数式 1 つだけなら "display"。"""
    if element.get("kind") != "equation":
        return "inline"
    content = element.get("content", ())
    if not isinstance(content, (list, tuple)):
        return "inline"
    fragments = [fragment for fragment in content if isinstance(fragment, Mapping) and str(fragment.get("text", ""))]
    return "display" if len(fragments) == 1 and fragments[0].get("kind") == "math" else "inline"


def math_options(element: Mapping, functions: Iterable[str] = (), *, style: Mapping | None = None) -> dict:
    """math_element()・append_math() に渡す引数を作る。

    戻り値: {"size", "color", "bold", "font", "functions", "display", "align", "text_fonts"}。
    style は断片の実効書式（省略時は element["style"]）。style にない項目は element["style"] の値を使う。
    """
    element_style = element.get("style", {}) or {}
    style = element_style if style is None else style

    def value(key, default):
        if key in style:
            return style[key]
        return element_style.get(key, default)

    return {
        "size": value("size_pt", 20),
        "color": str(value("color", "17212D")).removeprefix("#"),
        "bold": bool(value("bold", False)),
        "font": value("font_math", "Cambria Math"),
        "functions": tuple(functions),
        "display": math_mode(element) == "display",
        "align": value("align", "left"),
        "text_fonts": (value("font_japanese", "Yu Gothic UI"), value("font_latin", "Arial")),
    }
