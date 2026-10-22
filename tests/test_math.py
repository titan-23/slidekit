"""数式の構文と Office Math の構造、資料ごとの設定がほかの資料に影響しないことを確かめる。

数式の寸法の計算（math.font・math.metrics）は tests/test_math_metrics.py で確かめる。
"""

import pickle

from lxml import etree
import pytest

from slidekit.math import MathSyntaxError, append_math, math_element, parse
from slidekit.math import office_math
from slidekit.math.office_math import NS, TEXT_SCRIPT_SCALES, math_mode, math_options, text_size
from slidekit.math.syntax import Fraction, Nary, Run, Space, normalise, text_contents

M = "{%s}" % NS["m"]


def omath(expression, **kwargs):
    """数式の中身の要素。文中の数式は m:oMath > m:box（改行しない箱）> m:e の m:e を返す。"""
    root = math_element(expression, **kwargs).find(".//m:oMath", NS)
    box = root.find("m:box", NS)
    return root if box is None else box.find("m:e", NS)


def sty(run):
    node = run.find("m:rPr/m:sty", NS)
    return None if node is None else node.get(M + "val")


def runs(expression, **kwargs):
    """式の m:r を、(文字, m:sty の値) の列で返す。"""
    return [(run.find("m:t", NS).text, sty(run)) for run in omath(expression, **kwargs).iter(M + "r")]


def value(node, path):
    return node.find(path, NS).get(M + "val")


def test_office_math_keeps_structures_and_selected_font():
    root = math_element(r"\sum_{i=1}^{n}{\frac{x_i^2}{\sqrt{n}}}", font="Test Math")
    assert len(root.findall(".//m:nary", NS)) == 1
    assert len(root.findall(".//m:f", NS)) == 1
    assert len(root.findall(".//m:rad", NS)) == 1
    assert len(root.findall(".//m:sSubSup", NS)) == 1
    assert {node.get("typeface") for node in root.findall(".//a:latin", NS)} == {"Test Math"}


def test_functions_are_local_to_each_call():
    selected = math_element("LCE(x,y)", functions=["LCE"])
    ordinary = math_element("LCE(x,y)")
    assert len(selected.findall(".//m:func", NS)) == 1
    assert not ordinary.findall(".//m:func", NS)
    assert len(math_element("log n").findall(".//m:func", NS)) == 1
    with pytest.raises(ValueError, match="未対応"):
        math_element(r"\LCE(x,y)")


def test_text_is_normal_text_with_text_fonts():
    expression = r"\text{日本語 well-known}"
    run = omath(expression, text_fonts=("Yu Gothic UI", "Arial")).find("m:r", NS)
    assert run.find("m:rPr/m:nor", NS) is not None
    assert run.find("m:rPr/m:sty", NS) is None
    properties = run.find("a:rPr", NS)
    assert properties.find("a:ea", NS).get("typeface") == "Yu Gothic UI"
    assert properties.find("a:latin", NS).get("typeface") == "Arial"
    assert properties.get("lang") == "ja-JP" and properties.get("i") == "0"
    # 中身は置き換えない（ハイフンを負号にしない）。空白も保つ。
    assert run.find("m:t", NS).text == "日本語 well-known"
    plain = omath(r"\text{sample text}", font="Test Math").find("m:r", NS)
    assert plain.find("m:t", NS).text == "sample text"
    assert plain.find("a:rPr", NS).get("lang") == "en-US"
    assert {node.get("typeface") for node in plain.find("a:rPr", NS)[1:]} == {"Test Math"}
    # PowerPoint が lang="ja-JP" で和文の書体で描く記号を含む文章は ja-JP（計測でもこれらに和文の書体を使う）。→ は欧文の書体
    for text, lang in [("“ab”", "ja-JP"), ("a×b", "ja-JP"), ("‘a’ ○", "ja-JP"), ("a→b", "en-US"), ("a–b", "en-US")]:
        run = omath(r"\text{" + text + "}", text_fonts=("Yu Gothic UI", "Arial")).find("m:r", NS)
        assert run.find("a:rPr", NS).get("lang") == lang, text


@pytest.mark.parametrize(
    "expression",
    [
        r"\unknown{x}",
        r"\frac{x}",
        r"(x",
        r"x)",
        r"\left(x\right",
        r"\left x\right)",
        r"x_1_2",
        r"\lceil x",
        r"\rceil",
        "$x$",
        "x\ny",
    ],
)
def test_invalid_math_is_rejected(expression):
    with pytest.raises(ValueError):
        math_element(expression)


# 資料でよく使う記法の例。各行は、式、構造の要素（None は式全体）、m:val を確かめる子の要素と値、m:t に含まれる文字である。
# |x| の縦線は組にせず、通常の文字（立体）として出力する。
NOTATIONS = [
    (r"n \le m", None, {}, "≤"),
    (r"x \to \infty", None, {}, "∞"),
    (r"2\pi r + \delta", None, {}, "π"),
    (r"\alpha \Delta", None, {}, "Δ"),
    ("a/b", "m:f", {"m:fPr/m:type": "lin"}, "b"),
    ("f'(x)", "m:sSup", {}, "′"),
    ("x'_i", "m:sSubSup", {}, "′"),
    (r"\operatorname{rank}(A)", "m:func", {"m:fName/m:r/m:rPr/m:sty": "p"}, "rank"),
    (r"\mathbf{v}", "m:r", {"m:rPr/m:sty": "b"}, "v"),
    (r"\overline{x}", "m:bar", {"m:barPr/m:pos": "top"}, "x"),
    ("|x|", "m:r", {"m:rPr/m:sty": "p"}, "|"),
]


@pytest.mark.parametrize("expression, structure, values, text", NOTATIONS)
def test_common_notation_has_expected_structure(expression, structure, values, text):
    root = math_element(expression)
    node = root.find(".//" + structure, NS) if structure else root.find("m:oMath", NS)
    assert node is not None
    for path, expected in values.items():
        assert node.find(path, NS).get(M + "val") == expected
    assert text in [item.text for item in node.iter(M + "t")]


def test_bold_command_sets_bold_run():
    run = math_element(r"\mathbf{v}").find(".//m:r", NS)
    assert run.find("a:rPr", NS).get("b") == "1"


def test_append_math_precedes_paragraph_properties():
    paragraph = etree.Element("{%s}p" % NS["a"])
    end = etree.SubElement(paragraph, "{%s}endParaRPr" % NS["a"])
    added = append_math(paragraph, "x₂", size=18, color="AB0011", bold=True)
    assert list(paragraph) == [added, end]
    assert added.find(".//m:sSub", NS) is not None
    assert added.find(".//a:rPr", NS).get("sz") == "1800"


@pytest.mark.parametrize(
    "expression, bracket",
    [
        (r"O(\sqrt{n})", "()"),
        (r"\{x\}", "{}"),
        (r"\lceil x \rceil", "⌈⌉"),
        ("|x|", "||"),
        (r"\langle u, v \rangle", "⟨⟩"),
    ],
)
def test_plain_brackets_are_ordinary_characters(expression, bracket):
    node = omath(expression)
    assert node.find(".//m:d", NS) is None
    items = runs(expression)
    characters = [text for text, _ in items]
    assert bracket[0] in characters and bracket[1] in characters
    assert all(style == "p" for text, style in items if text in bracket)


def test_left_right_brackets_stretch():
    delimiters = omath(r"\left( \frac{a}{b} \right)").findall(".//m:d", NS)
    assert len(delimiters) == 1
    assert value(delimiters[0], "m:dPr/m:grow") == "1"
    assert value(delimiters[0], "m:dPr/m:begChr") == "("
    assert value(delimiters[0], "m:dPr/m:endChr") == ")"
    invisible = omath(r"\left. x \right|").find("m:d", NS)
    assert value(invisible, "m:dPr/m:begChr") == ""
    assert value(invisible, "m:dPr/m:endChr") == "|"
    reversed_brackets = omath(r"\left] a, b \right[").find("m:d", NS)
    assert (value(reversed_brackets, "m:dPr/m:begChr"), value(reversed_brackets, "m:dPr/m:endChr")) == ("]", "[")
    with pytest.raises(MathSyntaxError, match="伸縮括弧の始まりが不正"):
        math_element(r"\left{ x \right\}")


def test_middle_separator():
    delimiter = omath(r"\left\{ x \middle| x > 0 \right\}").find("m:d", NS)
    assert value(delimiter, "m:dPr/m:sepChr") == "|"
    assert (value(delimiter, "m:dPr/m:begChr"), value(delimiter, "m:dPr/m:endChr")) == ("{", "}")
    assert len(delimiter.findall("m:e", NS)) == 2
    with pytest.raises(MathSyntaxError) as outside:
        math_element(r"x \middle| y")
    assert outside.value.reason == r"\middle は \left と \right の間で使う"
    with pytest.raises(MathSyntaxError) as mixed:
        math_element(r"\left( a \middle| b \middle\| c \right)")
    assert mixed.value.reason == r"\middle の記号は 1 種類にする"
    assert mixed.value.token == r"\middle" and mixed.value.position == 20


def test_interval_brackets():
    for expression, opening, closing in [("[0, n)", "[", ")"), ("(0, n]", "(", "]")]:
        characters = [text for text, _ in runs(expression)]
        assert characters[0] == opening and characters[-1] == closing
    for expression, reason in [
        ("(x", "閉じ括弧がない"),
        ("x)", "対応する左側のない記号"),
        ("[x", "閉じ括弧がない"),
        (r"\{x]", "対応する左側のない記号"),
    ]:
        with pytest.raises(MathSyntaxError) as error:
            math_element(expression)
        assert error.value.reason == reason


def test_bars_are_not_paired():
    expression = r"\{ x \mid |x| < 1 \} \cup \{ y \mid |y| > 2 \}"
    node = omath(expression)
    assert node.find(".//m:d", NS) is None
    assert [text for text, _ in runs(expression)].count("|") == 4
    # |x|^2 の上付きは、閉じる側の縦線に付く
    script = omath("|x|^2").find("m:sSup", NS)
    assert [item.text for item in script.find("m:e", NS).iter(M + "t")] == ["|"]


@pytest.mark.parametrize(
    "expression, character, style",
    [
        (r"\zeta", "ζ", "i"),
        (r"\eta", "η", "i"),
        (r"\varphi", "φ", "i"),
        (r"\vartheta", "ϑ", "i"),
        (r"\Pi", "Π", "p"),
        (r"\Xi", "Ξ", "p"),
        (r"\varOmega", "Ω", "i"),
        (r"\ll", "≪", "p"),
        (r"\preceq", "⪯", "p"),
        (r"\coloneqq", "≔", "p"),
        (r"\models", "⊨", "p"),
        (r"\implies", "⟹", "p"),
        (r"\iff", "⟺", "p"),
        (r"\hookrightarrow", "↪", "p"),
        (r"\Downarrow", "⇓", "p"),
        (r"\setminus", "∖", "p"),
        (r"\otimes", "⊗", "p"),
        (r"\mp", "∓", "p"),
        (r"\nabla", "∇", "p"),
        (r"\aleph", "ℵ", "p"),
        (r"\dagger", "†", "p"),
        (r"\colon", ":", "p"),
        (r"\slash", "/", "p"),
        (r"\lvert", "|", "p"),
        (r"\|", "‖", "p"),
    ],
)
def test_added_symbols(expression, character, style):
    assert runs(expression) == [(character, style)]


@pytest.mark.parametrize(
    "expression, character",
    [
        (r"\int_0^1 f(x)\,dx", "∫"),
        (r"\oint f", "∮"),
        (r"\bigcup_i S_i", "⋃"),
        (r"\bigoplus_k V_k", "⨁"),
        (r"\coprod_i X", "∐"),
    ],
)
def test_added_large_operators(expression, character):
    operator = omath(expression).find("m:nary", NS)
    assert value(operator, "m:naryPr/m:chr") == character
    assert value(operator, "m:naryPr/m:limLoc") == "subSup"


@pytest.mark.parametrize(
    "expression, name", [(r"\cot x", "cot"), (r"\sec x", "sec"), ("Pr(X)", "Pr"), (r"\hom(A, B)", "hom")]
)
def test_added_functions(expression, name):
    function = omath(expression).find("m:func", NS)
    assert [(run.find("m:t", NS).text, sty(run)) for run in function.find("m:fName", NS).iter(M + "r")] == [(name, "p")]


def test_greek_and_symbol_styles():
    assert runs(r"\epsilon \varepsilon \phi \varphi") == [("ϵ", "i"), ("ε", "i"), ("ϕ", "i"), ("φ", "i")]
    assert runs(r"\Theta Θ \varTheta") == [("Θ", "p"), ("Θ", "p"), ("Θ", "i")]
    assert runs(r"\ell \hbar α") == [("ℓ", "p"), ("ℏ", "p"), ("α", "i")]
    assert runs(r"\alpha \Theta 2", bold=True) == [("α", "bi"), ("Θ", "b"), ("2", "b")]


def test_structures():
    binomial = omath(r"\binom{n}{k}").find("m:d", NS)
    assert (value(binomial, "m:dPr/m:begChr"), value(binomial, "m:dPr/m:endChr"), value(binomial, "m:dPr/m:grow")) == (
        "(",
        ")",
        "1",
    )
    assert value(binomial, "m:e/m:f/m:fPr/m:type") == "noBar"
    for expression, character in [
        (r"\hat{x}", "̂"),
        (r"\widehat{xy}", "̂"),
        (r"\tilde{x}", "̃"),
        (r"\vec{v}", "⃗"),
        (r"\dot{x}", "̇"),
        (r"\ddot{x}", "̈"),
    ]:
        assert value(omath(expression).find("m:acc", NS), "m:accPr/m:chr") == character
    assert [item.text for item in omath(r"\widehat{xy}").find("m:acc/m:e", NS).iter(M + "t")] == ["xy"]
    assert value(omath(r"\underline{x}").find("m:bar", NS), "m:barPr/m:pos") == "bot"
    upper = omath(r"\overset{?}{=}").find("m:limUpp", NS)
    assert [item.text for item in upper.find("m:e", NS).iter(M + "t")] == ["="]
    assert [item.text for item in upper.find("m:lim", NS).iter(M + "t")] == ["?"]
    lower = omath(r"\underset{i}{\max}").find("m:limLow", NS)
    assert [item.text for item in lower.find("m:e", NS).iter(M + "t")] == ["max"]
    assert value(omath(r"\sum\limits_{i=1}^{n} a_i").find("m:nary", NS), "m:naryPr/m:limLoc") == "undOvr"
    assert value(omath(r"\int\nolimits_0^1 f").find("m:nary", NS), "m:naryPr/m:limLoc") == "subSup"
    below = omath(r"\max\limits_{i} x_i").find("m:func/m:fName/m:limLow", NS)
    assert [item.text for item in below.find("m:e", NS).iter(M + "t")] == ["max"]
    assert [item.text for item in below.find("m:lim", NS).iter(M + "t")] == ["i"]
    both = omath(r"\max\limits_{i}^{n} x").find("m:func/m:fName/m:limUpp", NS)
    assert both.find("m:e/m:limLow", NS) is not None
    argmax = omath(r"\operatorname*{argmax}_x f").find("m:func/m:fName/m:limLow", NS)
    assert [(run.find("m:t", NS).text, sty(run)) for run in argmax.find("m:e", NS).iter(M + "r")] == [("argmax", "p")]
    assert runs(r"\not\in") == [("∉", "p")]
    assert runs(r"\not=") == [("≠", "p")]
    assert runs(r"\not\exists x") == [("∄", "p"), ("x", "i")]
    # \bmod は前後に \: の空白（U+205F）を置く（空白がないと PowerPoint は「amodb」と詰めて描く）
    assert runs(r"a \bmod b") == [("a", "i"), ("\u205f", "p"), ("mod", "p"), ("\u205f", "p"), ("b", "i")]
    assert runs(r"a \pmod{n}") == [("a", "i"), (" ", "p"), ("(", "p"), ("mod", "p"), (" ", "p"), ("n", "i"), (")", "p")]
    fraktur = omath(r"\mathfrak{g}").find("m:r", NS)
    assert value(fraktur, "m:rPr/m:scr") == "fraktur" and sty(fraktur) == "p"
    assert value(omath(r"\mathscr{L}").find("m:r", NS), "m:rPr/m:scr") == "script"
    bold = omath(r"\boldsymbol{\alpha}").find("m:r", NS)
    assert sty(bold) == "bi"
    assert (bold.find("a:rPr", NS).get("b"), bold.find("a:rPr", NS).get("i")) == ("1", "1")
    assert runs(r"\bm{x+2}") == [("x", "bi"), ("+", "b"), ("2", "b")]
    with pytest.raises(MathSyntaxError, match="打ち消せない"):
        math_element(r"\not\alpha")
    with pytest.raises(MathSyntaxError, match="直後"):
        math_element(r"x\limits_1")


@pytest.mark.parametrize(
    "expression, character",
    [
        (r"a\,b", " "),
        (r"a\thinspace b", " "),
        (r"a\:b", " "),
        (r"a\>b", " "),
        (r"a\medspace b", " "),
        (r"a\;b", " "),
        (r"a\thickspace b", " "),
        (r"a\enspace b", " "),
        (r"a\quad b", " "),
        (r"a\qquad b", "  "),
        (r"a\ b", " "),
        ("a~b", " "),
    ],
)
def test_spaces(expression, character):
    assert runs(expression) == [("a", "i"), (character, "p"), ("b", "i")]


def test_negative_space_outputs_nothing():
    assert runs(r"a\!b") == [("a", "i"), ("b", "i")]


def test_character_conversions():
    assert [text for text, _ in runs("a*b")] == ["a", "∗", "b"]
    assert [text for text, _ in runs("a-b")] == ["a", "−", "b"]
    assert [text for text, _ in runs("a−b")] == ["a", "−", "b"]
    assert [text for text, _ in runs(r"a \mid b")] == ["a", "∣", "b"]
    assert runs("f'") == [("f", "i"), ("′", "p")]


def test_bare_words_are_not_commands():
    for expression in [r"\mathrm{bar}", r"\mathrm{text}", "sum + 1", "frac12", "overline x"]:
        node = omath(expression)
        assert (
            node.find(".//m:nary", NS) is None and node.find(".//m:bar", NS) is None and node.find(".//m:f", NS) is None
        )
    assert runs(r"\mathrm{bar}") == [("bar", "p")]
    assert runs("sum + 1") == [("sum", "i"), ("+", "p"), ("1", "p")]
    assert omath("log n").find("m:func", NS) is not None
    assert omath("sqrt(x)").find("m:rad", NS) is not None


def test_japanese_outside_text_is_rejected():
    with pytest.raises(MathSyntaxError) as error:
        math_element("x の値")
    assert r"\text" in str(error.value)
    assert error.value.reason == r"日本語は \text{...} の中に書く"
    assert (error.value.token, error.value.position) == ("の", 2)
    assert [text for text, _ in runs(r"x_{\text{最大}}")] == ["x", "最大"]


def test_errors_have_position_and_hint():
    with pytest.raises(MathSyntaxError) as error:
        math_element(r"x+\alpah")
    assert error.value.code == "math_syntax_error"
    assert (error.value.reason, error.value.token, error.value.position) == ("未対応の数式命令", r"\alpah", 2)
    assert "3 文字目" in str(error.value) and r"\alpha" in str(error.value)
    assert str(error.value).startswith(r"未対応の数式命令: \alpah（3 文字目）。近い命令: \alpha")
    # 位置は Unicode の添字を正規化する前の、元の式の位置で示す
    with pytest.raises(MathSyntaxError) as normalised:
        math_element(r"x₂+\foo")
    assert normalised.value.position == 3 and normalised.value.expression == r"x₂+\foo"
    with pytest.raises(MathSyntaxError) as missing:
        math_element(r"\frac{x}")
    assert (missing.value.reason, missing.value.token, missing.value.position) == ("引数のない数式", r"\frac", 0)


@pytest.mark.parametrize(
    "expression",
    [
        r"\begin{cases} x \end{cases}",
        r"\big(",
        r"\displaystyle x",
        r"\textstyle x",
        r"\color{red} x",
        r"\textcolor{red}{x}",
        r"\hspace{1em}",
        r"\mkern3mu",
        r"a \over b",
        r"n \choose k",
        r"\overrightarrow{AB}",
        r"\overbrace{x}",
        r"\underbrace{x}",
        r"\substack{a}",
        r"x \tag{1}",
    ],
)
def test_unsupported_notations_are_rejected(expression):
    with pytest.raises(MathSyntaxError, match="未対応"):
        math_element(expression)


def test_display_math_writes_paragraph_alignment():
    for align, expected in [("left", "left"), ("center", "center"), ("right", "right"), ("justify", "left")]:
        root = math_element("x+1", display=True, align=align)
        assert value(root, "m:oMathPara/m:oMathParaPr/m:jc") == expected
        assert root.find("m:oMathPara/m:oMath", NS) is not None
    inline = math_element("x+1", align="center")
    assert inline.find("m:oMathPara", NS) is None and inline.find("m:oMath", NS) is not None
    with pytest.raises(ValueError, match="揃え"):
        math_element("x", display=True, align="middle")
    with pytest.raises(ValueError, match=r"\\text"):
        math_element("x", text_fonts=("Yu Gothic UI",))
    assert office_math.DISPLAY_TRAILING_EM == 0


def test_math_mode_and_options():
    style = {
        "size_pt": 30,
        "color": "#183153",
        "bold": True,
        "font_math": "Test Math",
        "align": "center",
        "font_japanese": "Test Gothic",
        "font_latin": "Test Sans",
    }
    equation = {"kind": "equation", "content": [{"kind": "math", "text": "x+1"}], "style": style}
    assert math_mode(equation) == "display"
    assert (
        math_mode({**equation, "content": [{"kind": "text", "text": ""}, {"kind": "math", "text": "x"}]}) == "display"
    )
    assert (
        math_mode({**equation, "content": [{"kind": "math", "text": "x"}, {"kind": "text", "text": " "}]}) == "inline"
    )
    assert math_mode({**equation, "kind": "text"}) == "inline"
    options = math_options(equation, ["rank"])
    assert options == {
        "size": 30,
        "color": "183153",
        "bold": True,
        "font": "Test Math",
        "functions": ("rank",),
        "display": True,
        "align": "center",
        "text_fonts": ("Test Gothic", "Test Sans"),
    }
    fragment = math_options(equation, style={"size_pt": 24, "bold": False})
    assert (fragment["size"], fragment["bold"], fragment["align"], fragment["font"]) == (
        24,
        False,
        "center",
        "Test Math",
    )
    paragraph = etree.Element("{%s}p" % NS["a"])
    added = append_math(paragraph, r"\text{rank} A", **options)
    assert value(added, "m:oMathPara/m:oMathParaPr/m:jc") == "center"
    assert added.find(".//m:nor/../../a:rPr/a:latin", NS).get("typeface") == "Test Sans"
    assert {"frac", "alpha", "log", "quad", "langle", "middle", "boldsymbol"} <= office_math.supported_commands()


def test_parse_shares_tree_and_reports_text():
    nodes = parse(r"x_{\text{最大}} + \text{個}")
    assert parse(r"x_{\text{最大}} + \text{個}") is nodes
    assert text_contents(nodes) == ["最大", "個"]
    assert parse("f(x)", ["f"]) != parse("f(x)")
    with pytest.raises(ValueError, match="英字"):
        parse("x", ["f1"])
    with pytest.raises(TypeError):
        parse(None)


def test_normalise_maps_original_positions():
    text, positions = normalise(" x₂ − y ")
    assert text == "x_{2} - y"
    assert len(positions) == len(text) and positions[0] == 1 and positions[text.index("2")] == 2 and positions[-1] == 6
    assert normalise(r"\text{a−b₂} − c")[0] == r"\text{a−b₂} - c"


def test_command_tables_are_exported():
    # 命令の表の別名（既存の参照のために残す）。 0.1 の計測（matplotlib）の互換の関数は、計測とともに削除した。
    assert {"log", "max"} <= office_math.DEFAULT_FUNCTIONS
    assert office_math.COMMANDS["alpha"] == "α" and office_math.COMMANDS["quad"] == "\u2003"
    assert office_math.COMMANDS["iff"] == "⟺" and office_math.COMMANDS["slash"] == "/"
    assert not hasattr(office_math, "_normalise") and "function_names" not in office_math.__all__


def test_syntax_error_keeps_details_after_pickling():
    with pytest.raises(MathSyntaxError) as error:
        parse(r"\left( x")
    copied = pickle.loads(pickle.dumps(error.value))
    assert (copied.reason, copied.token, copied.position, copied.expression) == (
        "閉じ括弧がない",
        r"\left",
        0,
        r"\left( x",
    )
    assert str(copied) == str(error.value)


def test_symbol_kinds_follow_the_command_table():
    """原子の種類（計測の前後の空き）は命令の表の値。: は句読点、… と ⋯ は内部、ピリオドは句読点、\\iff は関係。"""

    def kinds(expression):
        return [(node.text, node.kind) for node in parse(expression) if isinstance(node, Run)]

    assert kinds("a:b") == [("a", "ord"), (":", "punct"), ("b", "ord")]
    assert kinds(r"\colon \ldots \dots \cdots \vdots") == [
        (":", "punct"),
        ("…", "inner"),
        ("…", "inner"),
        ("⋯", "inner"),
        ("⋮", "ord"),
    ]
    assert kinds("p.q 3.14 …") == [("p", "ord"), (".", "punct"), ("q", "ord"), ("3.14", "ord"), ("…", "inner")]
    assert kinds(r"p \iff q") == [("p", "ord"), ("⟺", "rel"), ("q", "ord")]


def test_slash_command_is_a_symbol():
    # / は横書きの分数。記号の斜線は \slash か \text{/} で書く。
    assert [type(node) for node in parse("a/b")] == [Fraction]
    assert runs(r"a \slash b") == [("a", "i"), ("/", "p"), ("b", "i")]
    assert omath(r"a \slash b").find(".//m:f", NS) is None
    assert [text for text, _ in runs(r"\text{/}x")] == ["/", "x"]


def test_inline_math_is_wrapped_in_unbreakable_box():
    # 文中の数式は m:oMath の中身を m:box で包み、m:noBreak で途中の改行を止める（PowerPoint は行の残りに入らない式を
    # 演算子の後で改行する。包むと式ごと次の行へ送る）。独立した数式と空の式は包まない。
    equation = math_element("x+1").find("m:oMath", NS)
    assert [child.tag for child in equation] == [M + "box"]
    box = equation.find("m:box", NS)
    assert [child.tag for child in box] == [M + "boxPr", M + "e"]
    assert value(box, "m:boxPr/m:noBreak") == "1"
    assert [item.text for item in box.find("m:e", NS).iter(M + "t")] == ["x", "+", "1"]
    display = math_element("x+1", display=True).find("m:oMathPara/m:oMath", NS)
    assert display.find("m:box", NS) is None and [item.text for item in display.iter(M + "t")] == ["x", "+", "1"]
    assert list(math_element("").find("m:oMath", NS)) == []
    paragraph = etree.Element("{%s}p" % NS["a"])
    assert append_math(paragraph, "x_i").find("m:oMath/m:box/m:e/m:sSub", NS) is not None


def text_sizes(expression, **kwargs):
    """\\text の文字（m:nor の m:r）の a:rPr/@sz を、出現順に返す。"""
    return [
        run.find("a:rPr", NS).get("sz")
        for run in omath(expression, **kwargs).iter(M + "r")
        if run.find("m:rPr/m:nor", NS) is not None
    ]


def test_text_in_scripts_uses_reduced_size():
    # PowerPoint は m:nor の文字を添字の中でも縮めずに sz のとおりに描くので、段に応じて縮めた大きさを書く
    # （1 段目は 73 %、2 段目以降は 60 %。1/100 pt に丸める）。数式の文字の sz は要素の大きさのまま（PowerPoint が縮める）。
    # 倍率は数式の書体によらない（計測にも text_size() の大きさを使う）。
    assert TEXT_SCRIPT_SCALES == (1.0, 0.73, 0.60)
    assert [text_size(20, level) for level in range(4)] == [20.0, 14.6, 12.0, 12.0]
    assert text_size(11, 1) == 8.03 and text_size(13.333, 0) == 13.33
    assert text_sizes(r"x_{\text{a}}", size=11) == ["803"]
    assert text_sizes(r"x_{\text{a}}", font="DejaVu Serif") == ["1460"]
    assert text_sizes(r"x_{\text{最大}}") == ["1460"]
    assert {
        run.find("a:rPr", NS).get("sz")
        for run in omath(r"x_{\text{最大}}").iter(M + "r")
        if run.find("m:rPr/m:nor", NS) is None
    } == {"2000"}
    assert text_sizes(r"x_{y_{\text{a}}}") == ["1200"]
    assert text_sizes(r"x^{\text{a}} \text{b}", size=24) == ["1752", "2400"]
    assert text_sizes(r"\sum_{\text{a}}^{\text{b}} \text{c}") == ["1460", "1460", "2000"]
    assert text_sizes(r"\max\limits_{\text{a}} x") == ["1460"]
    assert text_sizes(r"\sqrt[\text{a}]{\text{b}}") == ["1200", "2000"]
    # 分数の分子と分母：文中の数式では 1 段進める。独立した数式では、段 0 の分数（入れ子の分数も）は進めない
    # （PowerPoint は独立した数式の入れ子の分数の分子と分母を縮めない）。添字の中の分数は進める。
    assert text_sizes(r"\frac{\text{a}}{\text{b}}") == ["1460", "1460"]
    assert text_sizes(r"\frac{\text{a}}{\frac{\text{b}}{c}}") == ["1460", "1200"]
    assert text_sizes(r"\frac{\text{a}}{\frac{\text{b}}{c}}", display=True) == ["2000", "2000"]
    assert text_sizes(r"\left(\frac{1}{\frac{\text{a}}{b}}\right)", display=True) == ["2000"]
    assert text_sizes(r"x_{\frac{\text{a}}{b}}", display=True) == ["1200"]
    assert text_sizes(r"\text{a}/\text{b}") == ["2000", "2000"]
    assert text_sizes(r"\left(\text{a}\right)_{\text{b}}", display=True) == ["2000", "1460"]


def operand(expression):
    """最初の大型演算子の引数（m:e）の m:t の文字。"""
    return [item.text for item in omath(expression).find(".//m:nary/m:e", NS).iter(M + "t")]


def test_large_operator_operand_extends_to_next_operator():
    # 大型演算子の引数は、次のまとまりから、同じ段の次の二項演算・関係・句読点の前まで（なければ群の終わりまで）。
    # PowerPoint は引数と後ろの原子の間に細い空きを入れるので、続く原子も引数に入れる。
    assert operand(r"\int_0^1 f(x)\,dx") == ["f", "(", "x", ")", "\u2009", "dx"]
    assert operand(r"\sum_i a_i b_i = 0") == ["a", "i", "b", "i"]
    assert operand(r"\sum_i a_i + b") == ["a", "i"]
    assert operand(r"\sum_i a_i, b") == ["a", "i"]
    assert operand(r"\sum_i a_i : b") == ["a", "i"]
    assert operand(r"\sum_i a_i \le b") == ["a", "i"]
    assert operand(r"\sum_i a_i \not\in S") == ["a", "i"]
    assert operand(r"\sum_i a_i \not\exists") == ["a", "i", "∄"]
    assert operand(r"\left( \sum_i a_i \right) c") == ["a", "i"]
    assert operand(r"\sum_i \frac{a_i}{2} x/y") == ["a", "i", "2", "x", "y"]
    assert operand(r"\sum_i -a_i - b") == ["−", "a", "i"]
    assert operand(r"\prod_{i=1}^{n} \sum_j a_{ij} x_j") == ["j", "a", "ij", "x", "j"]
    assert operand(r"\sum_i a_i \bmod n") == ["a", "i", "\u205f", "mod", "\u205f", "n"]
    tree = parse(r"\int f\,dx + 1")
    assert isinstance(tree[0], Nary) and len(tree) == 3 and isinstance(tree[0].body[1], Space)
    assert [type(node) for node in parse(r"\sum_i a_i/b")[0].body] == [Fraction]
    with pytest.raises(MathSyntaxError, match="引数のない数式"):
        parse(r"\sum_i")
