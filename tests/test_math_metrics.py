"""数式フォントの MATH 表による数式の計測（math.font・math.metrics）を検査する。

DejaVu Serif（MATH 表がある）があればそれで、なければ Cambria Math で計測する。どちらもなければ省略する。
DejaVu Serif には斜体補正がないので、斜体補正の規則は字形の値を加えた計測器で確かめる。
PowerPoint の照合データ（tests/data/powerpoint_math.json）との比較は、Windows の Cambria Math・Yu Gothic UI・Arial で行う。
"""

from concurrent.futures import ThreadPoolExecutor
import hashlib
import importlib.util
import json
import logging
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import tomllib
from types import SimpleNamespace

from fontTools.ttLib import TTCollection, TTFont
import pytest

from slidekit.math import MathMeasurer, MathSyntaxError, math_element
from slidekit.math import commands, metrics
from slidekit.math.font import MathFont, math_character
from slidekit.math.metrics import Box, TextFonts, compare_reference, layout, reference_font_matches
from slidekit.math.syntax import parse

DATA = Path(__file__).resolve().parent / "data" / "powerpoint_math.json"
# 資料でよく使う記法の例（tests/test_math.py の NOTATIONS と同じ式）。出力と同じく計測できることを確かめる。
NOTATIONS = [
    r"n \le m",
    r"x \to \infty",
    r"2\pi r + \delta",
    r"\alpha \Delta",
    "a/b",
    "f'(x)",
    "x'_i",
    r"\operatorname{rank}(A)",
    r"\mathbf{v}",
    r"\overline{x}",
    "|x|",
]


@pytest.fixture
def math_font():
    path = Path("/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf")
    if path.is_file():
        return path, 0, "DejaVu Serif"
    path = Path("/mnt/c/Windows/Fonts/cambria.ttc")
    if path.is_file():
        return path, 1, "Cambria Math"
    pytest.skip("計測に使う実フォントが導入されていない")


@pytest.fixture
def measurer(math_font):
    path, index, family = math_font
    with MathMeasurer(path, index=index, family=family) as value:
        yield value


def em(value, size=20):
    """pt の値を em にする。"""
    return value / size


def width(measurer, expression, size=20, **options):
    return measurer.measure(expression, size, **options)[0]


def test_measurement_shape_and_scaling(math_font):
    path, index, family = math_font
    with MathMeasurer(path, index=index, family=family, functions=["Query"]) as measurer:
        for expression in [
            r"x_1^2",
            r"\frac{a}{b}",
            r"sqrt(n)",
            r"\sqrt[3]{x}",
            r"\sum_{i=1}^{n}{a_i}",
            r"\mathcal U",
            r"\mathbb R",
            r"\left(x\right)",
            r"\text{sample text}",
            r"Query(x)",
        ]:
            math_element(expression, font=family, functions=["Query"])
            value, height, depth = measurer.measure(expression, size=20)
            assert value > 0 and height > 0 and 0 <= depth <= height
        small = measurer.measure("x_1", size=16)
        large = measurer.measure("x_1", size=32)
        assert large == pytest.approx(tuple(value * 2 for value in small), rel=1e-9)
        assert measurer.measure("") == (0.0, 0.0, 0.0)
        assert measurer.ink("") == (0.0, 0.0, 0.0, 0.0)
        with pytest.raises(ValueError, match="未対応"):
            measurer.measure(r"\unsupported{x}")
        with pytest.raises(ValueError, match="文字サイズ"):
            measurer.measure("x", size=0)
        with pytest.raises(TypeError):
            measurer.measure(None)


@pytest.mark.parametrize("expression", NOTATIONS)
def test_common_notation_is_measurable(measurer, expression):
    value, height, depth = measurer.measure(expression, size=20)
    assert value > 0 and height > 0 and 0 <= depth <= height
    left, right, top, bottom = measurer.ink(expression, size=20)
    assert left < right and top > -bottom


def test_spacing_rules(math_font, monkeypatch):
    # 原子の間の空きだけを確かめるため、斜体補正と添字用の字形を使わない
    monkeypatch.setitem(metrics.OFFICE, "italic_correction", False)
    monkeypatch.setitem(metrics.OFFICE, "italic_trailing", False)
    monkeypatch.setitem(metrics.OFFICE, "script_alternates", False)
    path, index, family = math_font
    with MathMeasurer(path, index=index, family=family) as measurer:
        a, b, plus, equal, minus, comma = (width(measurer, value) for value in ("a", "b", "+", "=", "-", ","))
        assert em(width(measurer, "a+b") - (a + plus + b)) == pytest.approx(2 * 4 / 18, abs=0.01)
        assert em(width(measurer, "a=b") - (a + equal + b)) == pytest.approx(2 * 5 / 18, abs=0.01)
        assert em(width(measurer, "-a") - (minus + a)) == pytest.approx(0, abs=0.01)
        assert em(width(measurer, "a,b") - (a + comma + b)) == pytest.approx(3 / 18, abs=0.01)
        # 添字の中には中の空き（二項演算の前後の空き）がない
        scale = measurer.font.constant("ScriptPercentScaleDown")
        inside = width(measurer, "x_{a+b}") - width(measurer, "x_{ab}")
        assert em(inside - plus * scale) == pytest.approx(0, abs=0.01)


def test_number_separators_and_inner_ellipsis(measurer):
    """数字の間のカンマとピリオドは数の一部（空きなし）で、ほかのピリオドは句読点。\\ldots は inner で、カンマとの間にも細い空きが入る。"""
    one, comma, two = (width(measurer, value) for value in ("1", ",", "2"))
    assert em(width(measurer, "1,2") - (one + comma + two)) == pytest.approx(0, abs=0.01)
    assert em(width(measurer, "1..2") - (width(measurer, "1.") + width(measurer, ".") + two)) == pytest.approx(
        2 * 3 / 18, abs=0.01
    )
    assert em(width(measurer, "x,2") - (width(measurer, "x,") + two)) == pytest.approx(3 / 18, abs=0.01)
    assert em(width(measurer, "x.2") - (width(measurer, "x.") + two)) == pytest.approx(3 / 18, abs=0.01)
    # 空白を挟むと数を作らない（1,\,2 はカンマの後に句読点の空きと \, が入る。PowerPoint の PDF と同じ）
    assert em(width(measurer, r"1,\,2") - (one + comma + two)) == pytest.approx(3 / 18 + 1 / 6, abs=0.01)
    dots = width(measurer, r"\ldots")
    assert em(width(measurer, r"1,\ldots") - (one + comma + dots)) == pytest.approx(3 / 18, abs=0.01)
    # \dots は \ldots と同じ扱いで、\cdots も inner（DejaVu Serif には ⋯ がない）
    assert width(measurer, r"1,\dots") == pytest.approx(width(measurer, r"1,\ldots"))
    if measurer.font.has_character("⋯"):
        assert em(width(measurer, r"1,\cdots") - (one + comma + width(measurer, r"\cdots"))) == pytest.approx(
            3 / 18, abs=0.01
        )


def switched(math_font, monkeypatch, key, expressions):
    """OFFICE[key] を真にしたときと偽にしたときの文中の数式の送り幅の差（em）。計測器は設定ごとに作り直す。"""
    path, index, family = math_font

    def widths():
        with MathMeasurer(path, index=index, family=family) as measurer:
            return [width(measurer, expression) for expression in expressions]

    on = widths()
    with monkeypatch.context() as patch:
        patch.setitem(metrics.OFFICE, key, False)
        off = widths()
    return [em(a - b) for a, b in zip(on, off)]


def test_function_application_is_inner(math_font, monkeypatch):
    """関数の適用（本体のある関数名）の右側は inner として扱う（OFFICE['function_apply_inner']）。

    後に普通の文字・句読点・開き括弧が続くと 3/18 em を空け（空白の命令を挟んでも）、二項演算・関係・閉じ括弧の前と、
    添字の中では空けない。中身の最後の節が関数の適用である根号も同じ（PowerPoint の \\log x,y・\\log x\\,y・\\log x(y)・
    \\max(a,b)\\,c・\\sqrt{x\\log x},y と同じ。\\sqrt{x},y・(\\log x),y・\\log x+y・f(x),y は空けない）。
    """
    cases = {
        r"\log x,y": 3,
        r"\log x\,y": 3,
        r"\log x\quad y": 3,
        r"\log x(y)": 3,
        r"\log(x)y": 3,
        r"\max(a,b)\,c": 3,
        r"\min\{a,b\},c": 3,
        r"\max_i x_i,y": 3,
        r"\sqrt{x\log x},y": 3,
        r"\sqrt{\log x},y": 3,
        r"\log x+y": 0,
        r"\log x=y": 0,
        r"(\log x),y": 0,
        r"f(x),y": 0,
        r"\sqrt{x},y": 0,
        r"\log x": 0,
        r"\log x\sum_i a_i": 0,
        r"z_{\log x,y}": 0,
    }
    added = switched(math_font, monkeypatch, "function_apply_inner", list(cases))
    assert added == pytest.approx([value / 18 for value in cases.values()], abs=1e-6)


def test_linear_fraction_is_inner(math_font, monkeypatch, measurer):
    """横書きの分数は inner として扱い、前の斜体の文字と末尾の斜体の文字の斜体補正を加える（OFFICE['linear_fraction_inner']）。

    PowerPoint は x a/b・2a/b・(x)a/b の分数の前と a/b,x・a/b x・a/b(x) の分数の後に 3/18 em を空け、x+a/b・a/b+x・x(a/b) と
    添字の中では空けない。x a/b は x の、a/b x は b の斜体補正だけさらに広い。
    """
    x, b = (em(measurer.box(letter).italic) for letter in "xb")
    cases = {
        "x a/b": 3 / 18 + x,
        "2a/b": 3 / 18,
        "(x)a/b": 3 / 18,
        "a/b,x": 3 / 18,
        "a/b x": 3 / 18 + b,
        "a/b(x)": 3 / 18,
        r"x\,a/b": 3 / 18,
        "x+a/b": 0,
        "a/b+x": 0,
        "x(a/b)": 0,
        "a/b": 0,
        "z^{2a/b}": 0,
    }
    added = switched(math_font, monkeypatch, "linear_fraction_inner", list(cases))
    assert added == pytest.approx(list(cases.values()), abs=1e-6)


def test_limits_on_bare_function_name_are_ordinary(math_font, monkeypatch):
    """上下に置くもののうち、土台が本体のない関数名だけのもの（\\underset{i}{\\max}）は、両側とも ord（OFFICE['bare_function_limits_ord']）。

    PowerPoint は \\underset{i}{\\max} x_i の後と x\\underset{i}{\\max} y の両側を空けない。土台が関数の適用のもの
    （\\underset{x\\in S}{\\arg\\min} f(x)）と、\\limits を付けた関数名（\\max\\limits_i x_i）は関数名のまま。
    """
    cases = {
        r"\underset{i}{\max} x_i": -3,
        r"x\underset{i}{\max} y": -6,
        r"\underset{i}{\max}": 0,
        r"\underset{x\in S}{\arg\min} f(x)": 0,
        r"\max\limits_i x_i": 0,
    }
    added = switched(math_font, monkeypatch, "bare_function_limits_ord", list(cases))
    assert added == pytest.approx([value / 18 for value in cases.values()], abs=1e-6)


def test_bot_is_a_relation(measurer):
    """\\bot は関係（PowerPoint は \\text{/}\\bot の ⊥ の前に関係の空き 5/18 em を入れる）。\\perp と同じ空きになる。"""
    assert commands.SYMBOLS["bot"].kind == "rel"
    gap = width(measurer, r"\text{/}\bot") - width(measurer, r"\text{/}") - width(measurer, r"\bot")
    assert em(gap) == pytest.approx(5 / 18, abs=0.01)
    assert width(measurer, r"x\bot y") == pytest.approx(width(measurer, r"x\perp y"))


def test_bold_digits_do_not_form_numbers(measurer):
    """太字の数字（数学用の太字の数字になる）の間のカンマは句読点で、後ろに細い空きが入る（PowerPoint の PDF と同じ）。"""

    def gap(expression, joined, comma, bold=False):
        # カンマの後の空き：式の幅から、カンマを除いた式（00 は 1 つの数の字句）とカンマの幅を引く
        return em(
            width(measurer, expression, bold=bold)
            - width(measurer, joined, bold=bold)
            - width(measurer, comma, bold=bold)
        )

    assert gap("(0,0)", "(00)", ",", bold=True) == pytest.approx(3 / 18, abs=0.01)
    assert gap(r"\mathbf{(0,0)}", r"\mathbf{(00)}", r"\mathbf{,}") == pytest.approx(3 / 18, abs=0.01)
    # 書体の変換のない数字だけが数を作る
    assert gap("(0,0)", "(00)", ",") == pytest.approx(0, abs=0.01)


def test_function_spacing(math_font, monkeypatch):
    """関数名と引数の間に OFFICE['function_space_em'] の空きを入れ、引数が (・[・\\{ で始まれば 0 にする（PowerPoint の \\max\\{）。"""
    path, index, family = math_font
    cases = [
        (r"\log n", r"\log", "n"),
        (r"\log(n)", r"\log", "(n)"),
        (r"\log[n]", r"\log", "[n]"),
        (r"\log\{n\}", r"\log", r"\{n\}"),
        (r"\max\{1,2\}", r"\max", r"\{1,2\}"),
    ]

    def gaps():
        # 新しい計測器で計測する（計測結果の保持は OFFICE を区別しない）
        with MathMeasurer(path, index=index, family=family) as measurer:
            return [
                em(width(measurer, expression) - width(measurer, name) - width(measurer, argument))
                for expression, name, argument in cases
            ]

    assert gaps() == pytest.approx([3 / 18, 0, 0, 0, 0], abs=0.01)
    # 空きは原子の間の空き（SPACING）ではなく、OFFICE の値で決まる
    monkeypatch.setitem(metrics.OFFICE, "function_space_em", 0.5)
    monkeypatch.setitem(metrics.OFFICE, "function_space_brackets", "([")
    assert gaps() == pytest.approx([0.5, 0, 0, 0.5, 0.5], abs=0.01)


def test_math_kern_uses_the_larger_sum(math_font, monkeypatch):
    """添字の MathKern は土台と添字の値の和を 2 つの高さで求め、大きい方（詰める量の小さい方）を使う。

    OpenType の仕様の文言の「小さい方」から変えた点である（PowerPoint の PDF の添字の位置と合う）。合成の表で確かめる。
    土台は b にする（DejaVu Serif では SuperscriptBottomMin が x の高さと同じで、x^2 の 2 つの高さが一致する）。
    """
    monkeypatch.setitem(metrics.OFFICE, "script_alternates", False)
    path, index, family = math_font
    with MathFont(path, index=index, family=family) as font:
        unit, scale = font.units_per_em, font.constant("ScriptPercentScaleDown")
        base, script = font.glyph_for("b", "italic").name, font.glyph_for("2").name

        def table(*values, split=None):
            # MathKern の表（em）。split があれば、それより下は values[0]、上は values[1]
            heights = [] if split is None else [SimpleNamespace(Value=round(split * unit))]
            return SimpleNamespace(
                CorrectionHeight=heights, KernValue=[SimpleNamespace(Value=round(value * unit)) for value in values]
            )

        def measured(expression, kerns):
            monkeypatch.setattr(font, "_kerns", kerns)
            return layout(parse(expression), font, 20, display=False).width

        # 区切りは添字の字形の端の高さと土台の字形の端の高さの間に置く（下付きは 0.15 em、上付きは 0.6 em）
        cases = [("b_2", "BottomRight", "TopLeft", 0.15), ("b^2", "TopRight", "BottomLeft", 0.6)]
        for expression, corner, script_corner, split in cases:
            plain = measured(expression, {})
            for values in [(-0.2, -0.05), (-0.05, -0.2)]:
                kerns = {
                    base: SimpleNamespace(**{corner + "MathKern": table(*values, split=split)}),
                    script: SimpleNamespace(**{script_corner + "MathKern": table(-0.01)}),
                }
                # どちらの表でも、和の大きい方（土台の −0.05 em と添字の −0.01 em。添字の値は添字の大きさの em）だけ
                # 添字が左へ寄り、式の幅が縮む
                assert em(measured(expression, kerns) - plain) == pytest.approx(-0.05 - 0.01 * scale, abs=0.002)
            with monkeypatch.context() as patch:
                patch.setitem(metrics.OFFICE, "math_kern", False)
                assert measured(expression, kerns) == pytest.approx(plain)


def test_scripts_follow_the_last_atom_of_the_base(measurer, monkeypatch):
    """添字の上下の位置と MathKern は、土台の並びの最後の原子で決める（OFFICE['script_last_atom']）。

    PowerPoint は (\\frac{a}{b})^2 の上付きを、括弧の中の高さによらず、閉じ括弧に付けたときと同じ通常の位置に置く
    （校正用の資料と生成した資料の PDF）。土台が分数か伸びる括弧だけなら、その高さに合わせて上げる。
    """
    fraction = measurer.ink(r"(\frac{a}{b})")
    for expression, plain in [(r"(\frac{a}{b})^2", "(x)^2"), (r"(\frac{a}{b})_i", "(x)_i")]:
        # 添字は (x) に付けたときと同じ位置にあり、描画の上下は括弧の中の分数と添字の大きい方になる
        ink, reference = measurer.ink(expression), measurer.ink(plain)
        assert ink[2] == pytest.approx(max(fraction[2], reference[2])), expression
        assert ink[3] == pytest.approx(max(fraction[3], reference[3])), expression
        added = width(measurer, expression) - width(measurer, r"(\frac{a}{b})")
        assert added == pytest.approx(width(measurer, plain) - width(measurer, "(x)")), expression
    # 土台が分数だけ・伸びる括弧のときは、上付きを土台の高さに合わせて上げる
    lowered = measurer.ink(r"(\frac{a}{b})^2")[2]
    assert measurer.ink(r"{\frac{a}{b}}^2")[2] > lowered + 1.0
    assert measurer.ink(r"\left(\frac{a}{b}\right)^2")[2] > lowered + 1.0
    # 規則を止めると、土台の全体（括弧の中の分数）の高さに合わせて上げる
    monkeypatch.setitem(metrics.OFFICE, "script_last_atom", False)
    measurer._cache.clear()
    assert measurer.ink(r"(\frac{a}{b})^2")[2] > lowered + 1.0
    assert measurer.ink(r"(\frac{a}{b})_i")[3] > measurer.ink(r"(x)_i")[3] + 1.0


def test_quad_and_spaces(measurer):
    # 斜体の文字の後の空白の前には、その文字の斜体補正が入る（DejaVu Serif では 0）
    base = width(measurer, "ab") + measurer.box("a").italic
    assert em(width(measurer, r"a\quad b") - base) == pytest.approx(1, abs=0.01)
    assert em(width(measurer, r"a\,b") - base) == pytest.approx(1 / 6, abs=0.01)
    assert em(width(measurer, r"a\!b") - width(measurer, "ab")) == pytest.approx(0, abs=0.01)


def test_display_and_inline_styles(measurer):
    for expression in [r"\frac{a}{b}", r"\sum_{i=1}^{n} a_i"]:
        inline = measurer.measure(expression, display=False)
        display = measurer.measure(expression, display=True)
        assert display[1] > inline[1]


def test_stretchy_delimiters(measurer):
    stretched = measurer.measure(r"\left(\frac{a}{b}\right)")
    plain = measurer.measure(r"(\frac{a}{b})")
    assert stretched[1] > plain[1] and stretched[0] > plain[0]
    assert em(width(measurer, "(x)") - width(measurer, r"\left(x\right)")) == pytest.approx(0, abs=0.01)
    # \left. の見えない括弧は幅を持たない
    assert width(measurer, r"\left. x \right|") < width(measurer, r"\left| x \right|")


def test_limits_placement(measurer, monkeypatch):
    # 独立した数式では \limits の添字を上下に積む
    limits = measurer.measure(r"\sum\limits_{i=1}^{n} a_i", display=True)
    scripts = measurer.measure(r"\sum_{i=1}^{n} a_i", display=True)
    assert limits[0] < scripts[0] and limits[1] > scripts[1]
    # 文中の数式では、PowerPoint と同じく \limits でも添字を右に置く（関数名の \limits は文中でも下に置く）
    assert measurer.measure(r"\sum\limits_{i=1}^{n} a_i") == measurer.measure(r"\sum_{i=1}^{n} a_i")
    assert width(measurer, r"\max\limits_{i} x_i") < width(measurer, r"\max_{i} x_i")
    monkeypatch.setitem(metrics.OFFICE, "inline_nary_limits", True)
    measurer._cache.clear()
    assert width(measurer, r"\sum\limits_{i=1}^{n} a_i") < width(measurer, r"\sum_{i=1}^{n} a_i")


def test_structures_extend_the_ink(measurer):
    """根号・上線・下線・記号付き・分数線は、土台の描画範囲の外に描かれる。"""
    x = measurer.ink("x")
    for expression in [r"\sqrt{x}", r"\overline{x}", r"\hat{x}"]:
        assert measurer.ink(expression)[2] > x[2]
    assert measurer.ink(r"\underline{x}")[3] > x[3]
    assert measurer.ink(r"\sqrt[3]{x}")[1] > measurer.ink(r"\sqrt{x}")[1]
    fraction = measurer.ink(r"\frac{x}{x}")
    assert fraction[2] > x[2] and fraction[3] > x[3]
    # 横書きの分数は分子、分数の斜線（U+2044）、分母を同じ大きさで並べる（斜体の分子の後と式の末尾には斜体補正が入る。
    # 文中の数式の a と b の幅も末尾の斜体補正を含む）
    assert metrics.OFFICE["fraction_slash"] == "\u2044"
    slash = measurer.font.glyph_for("\u2044").advance * 20
    assert width(measurer, "a/b") == pytest.approx(width(measurer, "a") + slash + width(measurer, "b"), abs=1e-6)


def test_text_uses_text_fonts(measurer):
    calls = []

    def measure_text(text, size, bold):
        calls.append((text, size, bold))
        return size * len(text), size * 0.8, size * 0.2

    fonts = TextFonts(("Test Gothic", "Test Sans"), measure_text)
    assert measurer.measure(r"\text{日本語}", 20, text_fonts=fonts) == pytest.approx((60.0, 20.0, 4.0))
    # 描画の左右を返さない計測では、描画範囲の左右を送り幅の範囲にする
    assert measurer.ink(r"\text{日本語}", 20, text_fonts=fonts) == pytest.approx((0.0, 60.0, 16.0, 4.0))
    # 添字の中では書き出しと同じ大きさ（office_math.text_size()。数式の書体によらず 73 %・60 %）で計測する
    measurer.measure(r"x_{\text{最大}}", 20, bold=True, text_fonts=fonts)
    assert calls[-1] == ("最大", 14.6, True)
    measurer.measure(r"x_{y_{\text{最大}}}", 20, text_fonts=fonts)
    assert calls[-1] == ("最大", 12.0, False)
    with pytest.raises(ValueError, match="字形がない"):
        measurer.measure(r"\text{日本語}", 20)
    for values in [(float("nan"), 0, 0), (1.0, 0, 0, 0), ("a", 0, 0), (1.0, 0, 0, 0, float("inf")), (-1.0, 0, 0)]:
        broken = TextFonts(("broken", repr(values)), lambda text, size, bold, values=values: values)
        with pytest.raises(ValueError, match=r"\\text"):
            measurer.measure(r"\text{abc}", 20, text_fonts=broken)


def written_text_sizes(expression, size, display, font):
    """書き出した \\text の文字（m:nor の m:r）の a:rPr/@sz を pt にして返す。"""
    from slidekit.math.office_math import NS

    element = math_element(expression, size=size, display=display, font=font)
    return [
        int(run.find("a:rPr", NS).get("sz")) / 100
        for run in element.iter("{%s}r" % NS["m"])
        if run.find("m:rPr/m:nor", NS) is not None
    ]


def test_text_size_matches_written_size(measurer):
    """\\text を計測する大きさは、書き出す a:rPr/@sz と同じ（数式の書体の倍率によらず、同じ倍率と同じ丸め）。

    数式の書体が Cambria Math 以外（DejaVu Serif の ScriptPercentScaleDown は 80 %）でも合うことを確かめる。
    """
    sizes = []

    def measure_text(text, size, bold):
        sizes.append(size)
        return size * len(text), size * 0.7, size * 0.2

    fonts = TextFonts(("Size",), measure_text)
    for expression, size, display in [
        (r"x_{\text{A}}", 20, False),
        (r"x_{\text{A}}", 20, True),
        (r"x^{y_{\text{A}}} \text{B}", 20, False),
        (r"\frac{\text{a}}{b}", 20, False),
        (r"\frac{\text{a}}{\text{b}}", 20, True),
        (r"\sqrt[\text{n}]{\text{x}}", 20, False),
        (r"\sum_{\text{a}}^{\text{b}} \text{c}", 18, True),
        (r"x_{\text{A}}", 11, False),
        (r"\text{A}", 13.333, False),
        (r"\text{a}/\text{b}", 20, False),
        (r"x_{\frac{\text{a}}{b}}", 20, True),
    ]:
        sizes.clear()
        measurer.measure(expression, size, display=display, text_fonts=fonts)
        # 計測の順（根号は中身が先）は書き出しの順と違うので、並べ替えて比べる
        assert sorted(sizes) == sorted(written_text_sizes(expression, size, display, measurer.family)), (
            expression,
            display,
        )
    # 残る差：独立した数式の入れ子の分数の中は、書き出しは縮めない（PowerPoint は縮めない）が、計測では 1 段縮める
    sizes.clear()
    measurer.measure(r"\frac{\text{a}}{\frac{\text{b}}{c}}", 20, display=True, text_fonts=fonts)
    assert sizes == [20.0, 14.6]
    assert written_text_sizes(r"\frac{\text{a}}{\frac{\text{b}}{c}}", 20, True, measurer.family) == [20.0, 20.0]


def test_missing_glyph_is_reported(measurer):
    # DejaVu Serif には ≪ がない（Cambria Math にはある）
    if not measurer.font.has_character("≪"):
        with pytest.raises(ValueError, match="字形がない"):
            measurer.measure(r"a \ll b")
    with pytest.raises(ValueError, match="字形がない"):
        measurer.measure(r"x + \text{語}")


def test_collection_face_is_measured_without_extraction(math_font, tmp_path, monkeypatch):
    path, index, family = math_font
    source = tmp_path / "collection.ttc"
    with TTFont(path, fontNumber=index) as first, TTFont(path, fontNumber=index) as second:
        collection = TTCollection()
        collection.fonts = [first, second]
        collection.save(source)
    temporary = tmp_path / "temporary"
    temporary.mkdir()
    monkeypatch.setattr(tempfile, "tempdir", str(temporary))
    with MathMeasurer(path, index=index, family=family) as direct:
        expected = direct.measure(r"x_1^2 + \sqrt{y}", 20)
    with MathMeasurer(source, index=1, family=family) as measurer:
        assert measurer.measure(r"x_1^2 + \sqrt{y}", 20) == pytest.approx(expected)
    assert not any(temporary.iterdir())
    assert sorted(item.name for item in tmp_path.iterdir()) == ["collection.ttc", "temporary"]
    with pytest.raises(RuntimeError, match="閉じた"):
        measurer.measure("x")
    single = tmp_path / "single.ttf"
    with TTFont(path, fontNumber=index) as face:
        face.save(single)
    with pytest.raises(ValueError, match="0 を指定"):
        MathMeasurer(single, index=1, family=family)
    with pytest.raises(ValueError, match="読めない"):
        MathMeasurer(source, index=5, family=family)


def test_no_fonttools_warning_when_measuring_collection(math_font, tmp_path, caplog, capfd):
    path, index, family = math_font
    caplog.set_level(logging.WARNING, logger="fontTools")
    source = tmp_path / "flagged.ttc"
    with TTFont(path, fontNumber=index) as face:
        table = face["OS/2"]
        if table.version < 2:
            table.sxHeight, table.sCapHeight, table.usDefaultChar, table.usBreakChar, table.usMaxContext = (
                0,
                0,
                0,
                32,
                0,
            )
        table.version = 3
        table.fsSelection |= 1 << 7
        collection = TTCollection()
        collection.fonts = [face]
        collection.save(source)
    # 書体を保存し直すと fontTools が警告する書体である（0.1 の計測は、取り出した書体を保存して警告を出していた）
    assert any("fsSelection" in record.getMessage() for record in caplog.records)
    caplog.clear()
    capfd.readouterr()
    with MathMeasurer(source, index=0, family=family) as measurer:
        assert measurer.measure(r"\frac{a}{b} + x_i", 20)[0] > 0
    assert not [record for record in caplog.records if record.name.startswith("fontTools")]
    output = capfd.readouterr()
    assert output.out == "" and output.err == ""


def test_font_requires_matching_family_and_math_table(math_font, tmp_path):
    path, index, family = math_font
    with pytest.raises(ValueError, match="書体"):
        MathMeasurer(path, index=index, family="Wrong Family")
    modified = tmp_path / "no-math.ttf"
    with TTFont(path, fontNumber=index) as face:
        del face["MATH"]
        face.save(modified)
    with pytest.raises(ValueError, match="MATH"):
        MathMeasurer(modified, family=family)
    with pytest.raises(ValueError, match="英字"):
        MathMeasurer(path, index=index, family=family, functions=["f1"])


def test_cache_distinguishes_mode_and_fonts(math_font, monkeypatch):
    path, index, family = math_font
    parsed = []
    original = metrics.parse

    def counting_parse(expression, functions=()):
        parsed.append(expression)
        return original(expression, functions)

    monkeypatch.setattr(metrics, "parse", counting_parse)
    first = TextFonts(("A",), lambda text, size, bold: (size * len(text), size * 0.7, size * 0.2))
    second = TextFonts(("B",), lambda text, size, bold: (2 * size * len(text), size * 0.7, size * 0.2))
    with MathMeasurer(path, index=index, family=family) as measurer:
        inline = measurer.measure(r"\frac{a}{b}")
        assert measurer.measure(r"\frac{a}{b}") == inline
        measurer.ink(r"\frac{a}{b}")
        # 保持に当たった計測は、式を解析し直さない（出力と同じ解析で検証済み）
        assert parsed == [r"\frac{a}{b}"]
        display = measurer.measure(r"\frac{a}{b}", display=True)
        assert display != inline and len(parsed) == 2
        one = measurer.measure(r"\text{ab}", text_fonts=first)
        two = measurer.measure(r"\text{ab}", text_fonts=second)
        assert two[0] == pytest.approx(2 * one[0]) and len(parsed) == 4
        # TextFonts の key が同じなら、計測の関数が違っても保持した結果を使う
        assert measurer.measure(r"\text{ab}", text_fonts=TextFonts(("A",), second.measure)) == one
        assert len(parsed) == 4


def test_repeated_measurement_is_reused_only_for_same_size(measurer):
    first = measurer.measure(r"\alpha_i + \beta", size=20)
    measurer.measure(r"\gamma", size=20)
    assert measurer.measure(r"\alpha_i + \beta", size=20) == first
    larger = measurer.measure(r"\alpha_i + \beta", size=40)
    assert larger[0] == pytest.approx(first[0] * 2)


def test_measurer_can_be_shared_between_threads(measurer):
    expressions = [r"x_{%d}^{2} + \frac{a}{%d}" % (number, number) for number in range(40)]
    serial = [measurer.measure(expression, 20) for expression in expressions]
    measurer._cache.clear()
    with ThreadPoolExecutor(max_workers=4) as pool:
        parallel = list(pool.map(lambda expression: measurer.measure(expression, 20), expressions))
    assert parallel == serial


@pytest.mark.parametrize(
    "expression",
    [
        r"\frac12",
        r"\sqrt[3]{x}",
        r"\text{日本語}",
        "[0, n)",
        r"\binom{n}{k}",
        r"\left\{ x \middle| x > 0 \right\}",
        r"\hat{x}",
        r"\not\in",
        r"a \pmod{n}",
        r"\mathfrak{g}",
        r"\overset{?}{=}",
        r"\int_0^1 f(x)\,dx",
        r"\max\limits_{i} x_i",
        r"\operatorname*{argmax}_x f",
        r"\lfloor x \rfloor",
        r"\boldsymbol{\alpha}",
        "x₂+y²",
    ],
)
def test_output_and_measurement_accept_the_same_notation(measurer, expression):
    fonts = TextFonts(("Test Gothic", "Test Sans"), lambda text, size, bold: (size * len(text), size * 0.8, size * 0.2))
    math_element(expression, text_fonts=("Test Gothic", "Test Sans"))
    assert measurer.measure(expression, 20, text_fonts=fonts)[0] > 0


@pytest.mark.parametrize(
    "expression", [r"\unknown", r"\unknown{x}", "(x", "x)", "x の値", r"\frac{x}", r"\left( x", "$x$"]
)
def test_output_and_measurement_reject_the_same_notation(measurer, expression):
    with pytest.raises(MathSyntaxError) as output:
        math_element(expression)
    with pytest.raises(MathSyntaxError) as measured:
        measurer.measure(expression, 20)
    assert (measured.value.reason, measured.value.position) == (output.value.reason, output.value.position)


def test_layout_returns_box_from_tree(math_font):
    path, index, family = math_font
    with MathFont(path, index=index, family=family) as font:
        box = layout(parse("x^2"), font, 20, display=False)
        assert isinstance(box, Box) and box.width > 0 and box.kind == "ord"
        left, right, top, bottom = box.ink
        assert left < right and box.height >= top > 0
        with pytest.raises(ValueError, match="文字サイズ"):
            layout(parse("x"), font, -1, display=False)
        with pytest.raises(TypeError):
            layout(parse(r"\text{a}"), font, 20, display=False, text_fonts=("Gothic", "Sans"))


def test_math_characters_map_styles_and_exceptions():
    assert math_character("x", "italic") == "𝑥"
    assert math_character("h", "italic") == "ℎ"
    assert math_character("x", "upright") == "x"
    assert math_character("α", "bold-italic") == "𝜶"
    assert math_character("ϵ", "italic") == "𝜖"
    assert math_character("2", "bold") == "𝟐"
    assert math_character("2", "italic") == "2"
    assert math_character("+", "bold") == "+"
    assert math_character("R", "upright", "double-struck") == "ℝ"
    assert math_character("A", "upright", "double-struck") == "𝔸"
    assert math_character("L", "upright", "script") == "ℒ"
    assert math_character("L", "bold", "script") == "𝓛"
    assert math_character("C", "upright", "fraktur") == "ℭ"
    assert math_character("g", "upright", "fraktur") == "𝔤"
    assert math_character("A", "upright", "sans-serif") == "𝖠"
    assert math_character("A", "bold-italic", "sans-serif") == "𝘼"
    assert math_character("A", "bold", "monospace") == "𝙰"


def test_math_font_values(math_font):
    path, index, family = math_font
    with MathFont(path, index=index, family=family) as font:
        assert 0 < font.constant("ScriptScriptPercentScaleDown") < font.constant("ScriptPercentScaleDown") <= 1
        assert 0 < font.constant("AxisHeight") < 1 and font.constant("MinConnectorOverlap") >= 0
        with pytest.raises(KeyError, match="定数"):
            font.constant("NoSuchConstant")
        x = font.glyph_for("x", "italic")
        assert x.advance > 0 and x.bounds[0] < x.bounds[2] and x.bounds[1] < x.bounds[3]
        assert font.glyph_for(" ").bounds is None
        parenthesis = font.glyph_for("(").name
        assert font.vertical_variants(parenthesis) or font.vertical_assembly(parenthesis) is not None
        assert font.script_glyph(x.name, 0) == x.name
        assert font.kern(font.glyph_for(" ").name, "TopRight", 0.5) == 0
        with pytest.raises(ValueError, match="角"):
            font.kern(x.name, "Top", 0.5)
        # MathKern は高さの区切りで値を選ぶ（最初の区切りより下は最初の値、最後の区切り以上は最後の値）
        table = SimpleNamespace(
            CorrectionHeight=[
                SimpleNamespace(Value=font.units_per_em // 4),
                SimpleNamespace(Value=font.units_per_em // 2),
            ],
            KernValue=[SimpleNamespace(Value=-100), SimpleNamespace(Value=-50), SimpleNamespace(Value=10)],
        )
        font._kerns["test"] = SimpleNamespace(TopRightMathKern=table)
        unit = font.units_per_em
        assert font.kern("test", "TopRight", 0.1) == pytest.approx(-100 / unit)
        assert font.kern("test", "TopRight", 0.25) == pytest.approx(-50 / unit)
        assert font.kern("test", "TopRight", 0.4) == pytest.approx(-50 / unit)
        assert font.kern("test", "TopRight", 0.6) == pytest.approx(10 / unit)
        assert font.kern("test", "BottomLeft", 0.6) == 0
    with pytest.raises(RuntimeError, match="閉じた"):
        font.glyph_for("x")


def reference(measurer, size=20, *, advance=0.0, height=0.0):
    """計測の値に差を加えた照合データ（tests/data/powerpoint_math.json の形）。"""
    items = []
    for expression in ["x_i^2", "a+b", r"\sqrt{n}", "f(x)", r"\text{abc}"]:
        box = measurer.box(expression, size)
        items.append(
            {
                "expression": expression,
                "mode": "inline",
                "advance_pt": box.width - advance * size,
                "ink_pt": [box.ink[0], box.ink[1], box.ink[2] - height * size, box.ink[3]],
            }
        )
    box = measurer.box(r"\frac{a}{b}", size, display=True)
    items.append(
        {
            "expression": r"\frac{a}{b}",
            "mode": "display",
            "ink_width_pt": box.ink[1] - box.ink[0],
            "ink_height_pt": box.ink[2] + box.ink[3],
        }
    )
    return {"schema_version": 1, "size_pt": size, "font": {"family": measurer.family}, "items": items}


def test_text_ink_uses_glyph_sides(measurer):
    # 文字の計測が描画の左右も返すときは、描画範囲の左右をその値にする（独立した数式の描画の幅に影響する）
    fonts = TextFonts(
        ("Sides",), lambda text, size, bold: (size * len(text), size * 0.7, size * 0.1, 1.5, size * len(text) - 2.0)
    )
    assert measurer.ink(r"\text{ab}", 20, text_fonts=fonts) == pytest.approx((1.5, 38.0, 14.0, 2.0))
    assert measurer.measure(r"\text{ab}", 20, text_fonts=fonts)[0] == pytest.approx(40.0)
    # 空白だけの文章は描画しない
    assert measurer.ink(r"\text{ }", 20, text_fonts=fonts) == (0.0, 0.0, 0.0, 0.0)


def italic_measurer(math_font, characters, amount=0.1):
    """characters の斜体の字形（添字用の字形を含む）に、斜体補正 amount em（設計単位に丸める）を加えた計測器。

    DejaVu Serif には斜体補正がないので、斜体補正の規則はこの計測器で確かめる。
    """
    path, index, family = math_font
    measurer = MathMeasurer(path, index=index, family=family)
    font = measurer.font
    for character in characters:
        name = font._cmap[ord(math_character(character, "italic"))]
        for level in (0, 1, 2):
            font._italics[font.script_glyph(name, level)] = round(amount * font.units_per_em)
    return measurer


def test_italic_correction_before_spaces_and_at_the_end(math_font, monkeypatch):
    """斜体の文字の後に空白が続くときと、文中の数式の末尾が斜体の文字のとき、斜体補正を加える。"""
    with italic_measurer(math_font, "abxi") as measurer:
        italic = measurer.font.glyph_for("a", "italic").italic * 20
        assert italic == pytest.approx(2.0, abs=0.01)
        plain = measurer.font.glyph_for("a", "italic").advance * 20
        # 末尾の斜体補正は文中の数式の送り幅に含まれ、独立した数式の組版上の幅には含まれない
        assert width(measurer, "a") == pytest.approx(plain + italic)
        assert width(measurer, "a", display=True) == pytest.approx(plain)
        assert measurer.box("a").italic == pytest.approx(italic)
        # 斜体の文字が続くときは入れない。空白の前と、立体の文字（演算子・括弧）の前には入れる
        assert width(measurer, r"a\,b") - width(measurer, "ab") == pytest.approx(20 / 6 + italic)
        assert width(measurer, r"a\quad b") - width(measurer, "ab") == pytest.approx(20 + italic)
        assert width(measurer, r"a\!b") == pytest.approx(width(measurer, "ab"))
        assert width(measurer, "a(b)") == pytest.approx(width(measurer, "a") + width(measurer, "(b)"))
        # 添字・記号付き・根号の末尾の斜体の文字からも伝える
        scale = measurer.font.constant("ScriptPercentScaleDown")
        assert width(measurer, "x_i") - width(measurer, "x_i", display=True) == pytest.approx(italic * scale, abs=1e-6)
        for expression in [r"\hat{x}", r"\sqrt{x}", r"\bar{x}", r"\underline{x}"]:
            assert width(measurer, expression) - width(measurer, expression, display=True) == pytest.approx(
                italic
            ), expression
        assert width(measurer, "x^2") == pytest.approx(width(measurer, "x^2", display=True))
        # 規則を止めると入れない
        monkeypatch.setitem(metrics.OFFICE, "italic_before_space", False)
        monkeypatch.setitem(metrics.OFFICE, "italic_trailing", False)
        measurer._cache.clear()
        assert width(measurer, r"a\,b") - width(measurer, "ab") == pytest.approx(20 / 6)
        assert width(measurer, "a") == pytest.approx(plain)


def test_large_operator_scripts_and_operand(math_font, monkeypatch):
    """大型演算子の添字の後に SpaceAfterScript を入れず、添字の末尾の斜体補正を加える。独立した数式で添字を右に置くときは
    引数との間を空けず、上下に積むときは細い空きを入れる。文中の数式では添字を文字の添字と同じ位置に置く（積分を除く）。"""
    with italic_measurer(math_font, "i") as measurer:
        size = 20

        def scripted(expression, **options):
            return measurer.box(expression, size, **options)

        # 引数のない大型演算子（空の群 {} を引数にする）の幅と比べる
        operator = scripted(r"\sum{}", display=True).width
        a = width(measurer, "a", display=True)
        thin = 3 / 18 * size
        # 添字を右に置く独立した数式：演算子と引数の間を空けない
        alone = scripted(r"\sum_i{}", display=True).width
        assert scripted(r"\sum_i a", display=True).width == pytest.approx(alone + a)
        # 添字の後には SpaceAfterScript を入れず、添字 i の斜体補正（添字の大きさ）を加える
        scale = measurer.font.constant("ScriptPercentScaleDown")
        with monkeypatch.context() as patch:
            patch.setitem(metrics.OFFICE, "nary_space_after_script", True)
            measurer._cache.clear()
            spaced = scripted(r"\sum_i{}", display=True).width
        measurer._cache.clear()
        space_after = measurer.font.constant("SpaceAfterScript") * size
        italic = measurer.font.glyph(measurer.font.script_glyph(measurer.font.glyph_for("i", "italic").name, 1)).italic
        assert spaced - alone == pytest.approx(space_after - italic * size * scale, abs=1e-6)
        # 上下に積む独立した数式：演算子と引数の間に細い空き
        stacked = scripted(r"\sum\limits_i{}", display=True).width
        assert scripted(r"\sum\limits_i a", display=True).width == pytest.approx(stacked + thin + a)
        # 添字のない独立した数式の演算子も、引数との間を空けない
        assert scripted(r"\sum a", display=True).width == pytest.approx(operator + a)
        with monkeypatch.context() as patch:
            patch.setitem(metrics.OFFICE, "display_nary_operand_space", True)
            measurer._cache.clear()
            assert scripted(r"\sum_i a", display=True).width == pytest.approx(alone + thin + a)
        measurer._cache.clear()
        # 文中の数式では演算子と引数の間に細い空き（TeX と同じ）
        inline = scripted(r"\sum_i{}").width
        assert scripted(r"\sum_i a").width == pytest.approx(inline + thin + width(measurer, "a"), abs=1e-6)
        # 文中の数式の積分以外の大型演算子は、添字を文字の添字と同じ位置に置く（演算子の字形の上下を使わない）
        low = measurer.ink(r"\sum_{i=1}^{n} a_i")
        with monkeypatch.context() as patch:
            patch.setitem(metrics.OFFICE, "inline_nary_script_drop", True)
            measurer._cache.clear()
            dropped = measurer.ink(r"\sum_{i=1}^{n} a_i")
            integral = measurer.ink(r"\int_0^1 f")
        measurer._cache.clear()
        assert low[2] < dropped[2] and low[3] != dropped[3]
        assert measurer.ink(r"\int_0^1 f") == pytest.approx(integral)


def test_radical_excess_goes_above(measurer, monkeypatch):
    """根号の字形が必要な長さより長いとき、余りの半分を中身の上の空きに加える（TeX・MathML Core）。"""
    expressions = [r"\sqrt{n}", r"\sqrt{x}", r"\sqrt{\frac{a}{b}}", r"\sqrt[3]{x}", r"\sqrt{x+y}"]
    above = [measurer.ink(expression) for expression in expressions]
    monkeypatch.setitem(metrics.OFFICE, "radical_excess_above", False)
    measurer._cache.clear()
    below = [measurer.ink(expression) for expression in expressions]
    for upper, lower, expression in zip(above, below, expressions):
        # 字形の長さは同じで、位置だけが上へ動く
        assert upper[2] + upper[3] == pytest.approx(lower[2] + lower[3]), expression
        assert upper[2] >= lower[2] - 1e-9, expression
    assert any(upper[2] > lower[2] + 0.1 for upper, lower in zip(above, below))


def test_delimiter_factor_selects_shorter_variants(measurer, monkeypatch):
    """伸びる括弧は中身の軸からの上下の 2 倍に 0.901 を掛けた長さ以上の最初の候補にする（TeX の \\delimiterfactor）。"""
    assert metrics.OFFICE["delimiter_factor"] == 0.901
    expressions = [
        r"\binom{n}{k}",
        r"\left(\frac{a}{b}\right)",
        r"\left(x^2\right)",
        r"\left(\frac{a^2}{b}\right)",
        r"\left(x_i\right)",
        r"\left(\sqrt{x}\right)",
    ]
    cases = [(expression, display) for display in (False, True) for expression in expressions]

    def heights():
        inks = {case: measurer.ink(case[0], display=case[1]) for case in cases}
        return {case: ink[2] + ink[3] for case, ink in inks.items()}

    factored = heights()
    monkeypatch.setitem(metrics.OFFICE, "delimiter_factor", 1.0)
    measurer._cache.clear()
    full = heights()
    measurer._cache.clear()
    # 係数を掛けると、同じか短い候補になる
    assert all(factored[case] <= full[case] + 1e-9 for case in cases)
    # 係数を掛けた長さで、短い候補が選ばれる式がある（DejaVu Serif でも Cambria Math でも、独立した数式の
    # \left(\frac{a^2}{b}\right) は中身の上下の 2 倍より少し短い候補になる）
    case = (r"\left(\frac{a^2}{b}\right)", True)
    assert factored[case] < full[case] - 1.0


def test_colon_and_mod_spacing(measurer):
    """: は ; と同じ句読点（後ろだけが空く）。\\bmod は前後に \\: の空白を置いた立体の mod。"""
    assert width(measurer, "a:b") - width(measurer, "a;b") == pytest.approx(width(measurer, ":") - width(measurer, ";"))
    medium = 4 / 18 * 20
    mod = width(measurer, r"\mathrm{mod}")
    assert width(measurer, r"a \bmod b") == pytest.approx(
        width(measurer, "a") + medium + mod + medium + width(measurer, "b")
    )
    # 添字の中でも空白は残る（空白の節は書き出しと計測の両方にある）
    assert width(measurer, r"x_{a \bmod b}") > width(measurer, r"x_{a\mathrm{mod}b}")


def test_compare_reference_reports_errors_and_limits(measurer):
    exact = compare_reference(reference(measurer), measurer)
    assert exact["passed"] and exact["count"] == 5 and exact["width_max_em"] == pytest.approx(0, abs=1e-9)
    assert [item["expression"] for item in exact["skipped"]] == [r"\text{abc}"]
    fonts = TextFonts(("Test",), lambda text, size, bold: (size * 0.5 * len(text), size * 0.7, size * 0.2))
    assert compare_reference(reference(measurer), measurer, text_fonts=fonts)["count"] == 6
    # 幅の差 0.04 em は基準の内、0.08 em は 9 割の基準を外れ、0.15 em は全件の基準も外れる
    assert compare_reference(reference(measurer, advance=0.04), measurer)["passed"]
    shifted = compare_reference(reference(measurer, advance=0.08), measurer)
    assert not shifted["passed"] and shifted["width_share"] == pytest.approx(1 / 5)
    assert shifted["items"][0]["width_em"] == pytest.approx(0.08)
    assert compare_reference(reference(measurer, advance=0.15), measurer)["width_max_em"] == pytest.approx(0.15)
    vertical = compare_reference(reference(measurer, height=0.11), measurer)
    assert not vertical["passed"] and vertical["vertical_max_em"] == pytest.approx(0.11)
    # 描画を計測できなかった文中の数式（ink_pt が null）は、送り幅だけを比べる
    plain = reference(measurer)
    plain["items"][1]["ink_pt"] = None
    row = compare_reference(plain, measurer)["items"][1]
    assert row["expression"] == "a+b" and row["vertical_em"] == 0 and "ink_width_em" not in row
    # excluded（比べない理由）を持つ項目は比べずに skipped に入れる
    marked = reference(measurer, advance=0.5)
    for item in marked["items"]:
        item["excluded"] = "書き出しを変えたので作り直す"
    marked["items"][0].pop("excluded")
    marked["items"][0]["advance_pt"] = measurer.box("x_i^2", 20).width
    result = compare_reference(marked, measurer)
    assert result["passed"] and result["count"] == 1 and len(result["skipped"]) == 5
    assert result["skipped"][0] == {"expression": "a+b", "mode": "inline", "reason": "書き出しを変えたので作り直す"}
    # 計測できない式は、その式を注記に加えて例外にする
    with pytest.raises(MathSyntaxError) as error:
        compare_reference({"size_pt": 20, "items": [{"expression": r"\unknown", "advance_pt": 1.0}]}, measurer)
    assert r"照合データの式: \unknown（inline）" in error.value.__notes__


@pytest.mark.parametrize(
    "data",
    [
        [],
        {"items": []},
        {"size_pt": 20, "items": 3},
        {"size_pt": 0, "items": []},
        {"size_pt": True, "items": []},
        {"size_pt": 20, "items": ["x"]},
        {"size_pt": 20, "items": [{"expression": "x"}]},
        {"size_pt": 20, "items": [{"expression": "x", "advance_pt": float("nan")}]},
        {"size_pt": 20, "items": [{"expression": "x", "advance_pt": 1.0, "ink_pt": [0, 1]}]},
        {"size_pt": 20, "items": [{"expression": "x", "mode": "block", "advance_pt": 1.0}]},
        {"size_pt": 20, "items": [{"expression": "x", "mode": "display", "ink_width_pt": 1.0}]},
        {"size_pt": 20, "items": [{"expression": "x", "advance_pt": 1.0, "excluded": ""}]},
        {"size_pt": 20, "items": [{"expression": "x", "advance_pt": 1.0, "excluded": True}]},
    ],
)
def test_compare_reference_rejects_malformed_data(measurer, data):
    with pytest.raises(ValueError, match="照合データ"):
        compare_reference(data, measurer)


def test_reference_command(math_font, tmp_path, capsys, monkeypatch):
    path, index, family = math_font
    with MathMeasurer(path, index=index, family=family) as measurer:
        data = reference(measurer)
    data["font"]["sha256"] = hashlib.sha256(path.read_bytes()).hexdigest()
    target = tmp_path / "reference.json"
    target.write_text(json.dumps(data), encoding="utf-8")
    assert reference_font_matches(data, path)
    font = f"{path}:{index}"
    assert metrics.main([str(target), "--font", font, "--family", family]) == 0
    output = capsys.readouterr().out
    assert "合格" in output and "不合格" not in output and "比べなかった式 1 件" in output
    assert metrics.main([str(target), "--font", font, "--family", family, "--json"]) == 0
    result = json.loads(capsys.readouterr().out)
    assert result["count"] == 5 and result["passed"] and len(result["skipped"]) == 1
    # 文字用の書体を指定すると \text の式も比べる
    metrics.main(
        [str(target), "--font", font, "--family", family, "--json", "--latin", str(path), "--japanese", str(path)]
    )
    assert json.loads(capsys.readouterr().out)["count"] == 6
    with MathMeasurer(path, index=index, family=family) as measurer:
        failing = reference(measurer, advance=0.2)
    failing["font"]["sha256"] = data["font"]["sha256"]
    target.write_text(json.dumps(failing), encoding="utf-8")
    assert metrics.main([str(target), "--font", font, "--family", family]) == 1
    assert "不合格" in capsys.readouterr().out
    failing["font"]["sha256"] = "0" * 64
    target.write_text(json.dumps(failing), encoding="utf-8")
    assert metrics.main([str(target), "--font", font, "--family", family]) == 2
    assert "作り直す" in capsys.readouterr().err
    # ファイル・照合データの形・引数の誤りも、不合格（1）と区別して 2 にし、理由を標準エラーに出す
    target.write_text(json.dumps(data), encoding="utf-8")
    missing = str(tmp_path / "missing.ttf")
    broken = tmp_path / "broken.json"
    broken.write_text("{", encoding="utf-8")
    malformed = tmp_path / "malformed.json"
    malformed.write_text(json.dumps({**data, "items": [{"expression": "x"}]}), encoding="utf-8")
    unmeasurable = tmp_path / "unmeasurable.json"
    unmeasurable.write_text(
        json.dumps({**data, "items": [{"expression": r"\unknown", "advance_pt": 1.0}]}), encoding="utf-8"
    )
    fonts = ["--latin", str(path), "--japanese"]
    for arguments, reason in [
        ([str(target), "--font", missing], "FileNotFoundError"),
        ([str(tmp_path / "missing.json"), "--font", font], "FileNotFoundError"),
        ([str(target), "--font", font, "--family", family, *fonts, missing], "FileNotFoundError"),
        ([str(target), "--font", font, "--family", family, *fonts, str(target)], "TTLibError"),
        ([str(target), "--font", font, "--family", "Wrong Family"], "書体"),
        ([str(broken), "--font", font], "JSONDecodeError"),
        ([str(malformed), "--font", font, "--family", family], "advance_pt"),
        ([str(unmeasurable), "--font", font, "--family", family], r"照合データの式: \unknown"),
    ]:
        assert metrics.main(arguments) == 2, arguments
        error = capsys.readouterr().err
        assert "照合できない" in error and reason in error and "Traceback" not in error, arguments
    for arguments in ([], [str(target), "--font", font, "--latin", str(path)]):
        with pytest.raises(SystemExit) as exit_info:
            metrics.main(arguments)
        assert exit_info.value.code == 2
    capsys.readouterr()

    # 計測の内部の誤りも 2 にし、traceback を出す
    def broken_compare(*args, **kwargs):
        raise RuntimeError("計測の内部の誤り")

    monkeypatch.setattr(metrics, "compare_reference", broken_compare)
    assert metrics.main([str(target), "--font", font, "--family", family]) == 2
    error = capsys.readouterr().err
    assert "Traceback" in error and "予期しない例外" in error


def test_reference_command_text_fonts_follow_text_measurement(tmp_path):
    """照合のコマンドの \\text の書体は、文字の計測と同じく、PowerPoint が既定で和文の書体で描く記号（“ ×）を和文の書体で計測する。"""
    serif, sans = (Path("/usr/share/fonts/truetype/dejavu") / name for name in ("DejaVuSerif.ttf", "DejaVuSans.ttf"))
    if not (serif.is_file() and sans.is_file()):
        pytest.skip("DejaVu Serif と DejaVu Sans が導入されていない")
    fonts = metrics._FileTextFonts((sans, 0), (serif, 0))
    try:

        def advance(path, character):
            with TTFont(path) as face:
                return face["hmtx"][face.getBestCmap()[ord(character)]][0] / face["head"].unitsPerEm * 20

        for character, path in [("“", sans), ("×", sans), ("a", serif), ("→", serif)]:
            assert fonts.measure(character, 20, False)[0] == pytest.approx(advance(path, character)), character
    finally:
        fonts.close()


# 照合の規則（添字の中の文章、大型演算子の引数と添字、伸びる括弧、添字の土台の最後の原子、関数名の後の \{、数字の間の
# カンマ、\bot の種類、関数の適用の右側、横書きの分数、本体のない関数名の上下に置くもの、\text の中の和文の書体の記号）を
# 確かめる式。照合データの道具の式の一覧に入れる
RULE_EXPRESSIONS = {
    r"x_{\text{最大}}",
    r"\sum_i a_i b_i",
    r"\int f\,dx",
    r"\sum^n a_i",
    r"\prod_{i=1}^{n} x_i",
    r"\oint_C f",
    r"\left(x^2\right)",
    r"\left(\frac{a^2}{b}\right)",
    r"(\frac{a}{b})^2",
    r"{\frac{a}{b}}^2",
    r"\left(\frac{a}{b}\right)^2",
    r"(\sqrt{x})^2",
    r"(a+b)_i",
    r"\max\{1,2\}",
    r"\min\{w,p-j\}",
    "(0,0)",
    "f(3,2)",
    r"\mathbf{(0,0)}",
    r"\text{/}\bot",
    r"x\bot y",
    r"\sqrt{x},y",
    r"\log x,y",
    r"\log x\,y",
    r"\log(x)y",
    r"\max(a,b)\,c",
    r"\log x(y)",
    r"\sqrt{x\log x},y",
    "x a/b",
    "2a/b",
    "a/b,x",
    "x+a/b",
    "x(a/b)",
    "a/b x",
    "a/b(x)",
    "(x)a/b",
    r"x\underset{i}{\max} y",
    r"\text{“ab”}",
    r"\text{a×b}",
}


def test_math_metrics_match_powerpoint(windows_fonts_environment):
    """照合データ（PowerPoint で変換した PDF の値）と計測を、合格の基準（metrics.REFERENCE_LIMITS）で比べる。

    基準は、幅の差（文中は送り幅、独立は描画の幅）が 9 割の式で 0.05 em 以内、すべての式で 0.12 em 以内、描画の上下の差が
    すべての式で 0.1 em 以内。数式の書体のハッシュ値が照合データと違えば省略する（照合データを作り直す）。
    \\text は資料と同じく、文字の書体（Yu Gothic UI・Arial）で、FontRegistry の文字の計測を使って計測する。
    """
    from slidekit.fonts import FontRegistry, default_font

    data = json.loads(DATA.read_text(encoding="utf-8"))
    environment = tomllib.loads(windows_fonts_environment.read_text(encoding="utf-8"))
    family = data["font"]["family"]
    if not reference_font_matches(data, default_font(family, environment)["path"]):
        pytest.skip(
            "照合データの数式の書体のハッシュ値がフォントと違う。tests/tools/powerpoint_math_data.py で照合データを作り直す"
        )
    fonts = data["text_fonts"]
    project = SimpleNamespace(
        theme={"fonts": {"japanese": fonts["japanese"], "latin": fonts["latin"], "math": family}},
        environment=environment,
        config={"math": {}},
    )
    with FontRegistry(project) as registry:
        result = compare_reference(data, registry._math_measurer(), text_fonts=registry._text_fonts())
    worst = sorted(result["items"], key=lambda row: -max(abs(row["width_em"]) / 0.05, row["vertical_em"] / 0.1))[:5]
    assert result["passed"], [
        (row["mode"], row["expression"], round(row["width_em"], 4), round(row["vertical_em"], 4)) for row in worst
    ]
    # 比べなかった式は、照合データで excluded（比べない理由）を付けた項目だけ
    assert sorted((row["expression"], row["mode"]) for row in result["skipped"]) == sorted(
        (item["expression"], item["mode"]) for item in data["items"] if "excluded" in item
    )
    assert result["count"] == len(data["items"]) - len(result["skipped"]) >= 190


def reference_tool():
    """照合データの道具（tests/tools/powerpoint_math_data.py）のモジュール。"""
    spec = importlib.util.spec_from_file_location(
        "powerpoint_math_data", DATA.parents[1] / "tools" / "powerpoint_math_data.py"
    )
    tool = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(tool)
    return tool


def test_reference_tool_counts_only_ink_connected_to_the_frame(tmp_path):
    """照合データの道具は、枠の上端より上の描画を、枠の中の描画に続く分だけ含める（上の行の式のはみ出しを拾わない）。"""
    from pdf_tools import Document

    from slidekit.pdf import Box, open_pdf

    tool = reference_tool()
    black = (0, 0, 0)
    pdf = tmp_path / "ink.pdf"
    with Document() as document:
        page = document.new_page(200, 120)
        # 枠の上端は y = 50。上の行の式のはみ出し（47〜49）と、間を空けて枠の中に描かれた式（56〜70）
        page.draw_rect((20, 47, 60, 49), color=None, fill=black)
        page.draw_rect((20, 56, 60, 70), color=None, fill=black)
        # 枠の上端をまたいで描かれた式（48〜70）
        page.draw_rect((100, 48, 140, 70), color=None, fill=black)
        document.save(pdf)
    with open_pdf(pdf) as document:
        page = document[0]
        separate = tool.connected_ink_box(page, Box(10, 46, 90, 80), 50)
        crossing = tool.connected_ink_box(page, Box(90, 46, 150, 80), 50)
        everything = tool.ink_box(page, Box(10, 46, 90, 80))
        empty = tool.connected_ink_box(page, Box(160, 46, 190, 80), 50)
    assert separate == pytest.approx((20, 56, 60, 70), abs=0.2)
    assert crossing == pytest.approx((100, 48, 140, 70), abs=0.2)
    assert everything == pytest.approx((20, 47, 60, 70), abs=0.2)
    assert empty is None


def test_reference_data_is_well_formed():
    """照合データは汎用の式だけで、パスを含まない。式は照合データの道具の一覧の先頭からの部分で、文中と独立の両方がある。

    照合データを道具で作り直すと、式の一覧と同じになる。excluded（比べない理由）は文字列で、文中と独立の両方に付ける。
    """
    data = json.loads(DATA.read_text(encoding="utf-8"))
    assert data["schema_version"] == 1 and data["size_pt"] == 20
    assert set(data["font"]) == {"family", "file", "index", "version", "sha256"} and "/" not in data["font"]["file"]
    expressions = [item["expression"] for item in data["items"]]
    inline = [expression for expression, item in zip(expressions, data["items"]) if item["mode"] == "inline"]
    display = [expression for expression, item in zip(expressions, data["items"]) if item["mode"] == "display"]
    assert inline == display and len(set(inline)) == len(inline)
    excluded = {(item["expression"], item["mode"]): item["excluded"] for item in data["items"] if "excluded" in item}
    assert all(isinstance(reason, str) and reason.strip() for reason in excluded.values())
    assert {key for key in excluded} == {
        (expression, mode) for expression, _ in excluded for mode in ("inline", "display")
    }
    # 照合データの式は道具の式の一覧の先頭からの部分（道具に加えた式は、作り直すと照合データに入る）。
    # 道具の一覧は照合の規則を確かめた式を含み、すべて書き出すことができる
    tool = reference_tool()
    assert list(tool.EXPRESSIONS[: len(inline)]) == inline
    assert RULE_EXPRESSIONS <= set(tool.EXPRESSIONS) and len(set(tool.EXPRESSIONS)) == len(tool.EXPRESSIONS)
    assert tool.supported() == (list(tool.EXPRESSIONS), {})


def test_matplotlib_is_not_imported(math_font):
    # 数式の計測は数式フォントの MATH 表で行い、matplotlib を読み込まない（0.1 の math.preview は削除した）
    import slidekit

    path, index, family = math_font
    script = (
        "import sys, importlib.util\n"
        "import slidekit.math, slidekit.fonts\n"
        "from slidekit.math import MathMeasurer\n"
        f"with MathMeasurer({str(path)!r}, index={index}, family={family!r}) as measurer:\n"
        r"    assert measurer.measure(r'\frac{x_i^2}{\sqrt{n}} + \sum_{k} a_k', 20)[0] > 0"
        "\n"
        "assert importlib.util.find_spec('slidekit.math.preview') is None\n"
        "print('matplotlib' in sys.modules)\n"
    )
    environment = {**os.environ, "PYTHONDONTWRITEBYTECODE": "1", "PYTHONPATH": str(Path(slidekit.__file__).parents[1])}
    result = subprocess.run(
        [sys.executable, "-c", script], capture_output=True, text=True, env=environment, check=False
    )
    assert (result.returncode, result.stdout.strip(), result.stderr) == (0, "False", "")
