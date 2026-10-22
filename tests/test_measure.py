"""文字と数式の寸法・折り返し・配置の計算（measure）を、偽のフォントと合成フォントで確かめる。"""

from __future__ import annotations

import json
import math
import sys
import types

import pytest

from fake_fonts import FakeFonts
from slidekit import fragments, measure, styles
from slidekit.errors import InputError
from slidekit.layout import Rect
from slidekit.measure import (
    MEASURE_VERSION,
    MODEL,
    NO_LINE_END,
    NO_LINE_START,
    ElementLayout,
    Insets,
    MathBox,
    MeasureError,
    fit_findings,
    font_runs,
    frame_size,
    layout_element,
    layout_record,
    layout_scene,
    layout_text,
    missing_glyphs,
    normalize_padding,
    padding_insets,
    text_area,
)
from synthetic_fonts import SyntheticFonts, make_font, make_standard_fonts


def natural(*faces):
    """行の書体の組の（上側、最後の基線より下）の割合（em）。faces は書体ごとの（usWinAscent, usWinDescent）の em。

    書体ごとに A = 1.2 × ascent ÷（ascent＋descent）、D = 1.2 − A とし、組の値は 1.2 × max A ÷（max A＋max D）と
    1.2 からその値を引いたもの（校正用の資料の PowerPoint の配置から求めたモデル）。
    """
    ascents = [1.2 * ascent / (ascent + descent) for ascent, descent in faces]
    descents = [1.2 - value for value in ascents]
    top = 1.2 * max(ascents) / (max(ascents) + max(descents))
    return top, 1.2 - top


def line_top(ascent, spacing):
    """行の上側（em）。行間がちょうど 1 なら書体の値、1 より大きければ行間 × min(書体の値, 0.9035)、
    1 未満は 書体の値 − 1.2 × (1 − 行間) と 行間 × min(書体の値, 0.9035) の大きいほう。"""
    if spacing == 1.0:
        return ascent
    proportional = spacing * min(ascent, 0.9035)
    if spacing > 1.0:
        return proportional
    return max(ascent - 1.2 * (1 - spacing), proportional)


FAKE_BOTTOM = 0.3  # FakeFonts の最後の基線より下（em。usWin 0.9・0.3 の組）
LATIN = (0.9, 0.2)  # 合成フォント Test Latin の usWin（em）
JAPANESE = (1.08, 0.25)  # 合成フォント Test Japanese の usWin（em）
SYNTHETIC_TOP, SYNTHETIC_BOTTOM = natural(LATIN, JAPANESE)  # 和文を含む行の上側と最後の基線より下（em）
LATIN_TOP, LATIN_BOTTOM = natural(LATIN)  # 欧文だけの行の上側と最後の基線より下（em）


def style(**overrides):
    values = {
        "size_pt": 20,
        "font_latin": "Latin",
        "font_japanese": "Japanese",
        "font_math": "Math",
        "line_spacing": 1.0,
        "word_wrap": True,
        "bold": False,
        "align": "left",
        "valign": "top",
        "padding_pt": 0,
    }
    return {**values, **overrides}


def synthetic_style(**overrides):
    return style(font_latin="Test Latin", font_japanese="Test Japanese", **overrides)


def content(*parts):
    """文字列は通常文字の断片、(種類, 文字[, 断片の書式]) はその種類の断片にする。"""
    result = []
    for part in parts:
        if isinstance(part, str):
            result.append({"kind": "text", "text": part})
        else:
            fragment = {"kind": part[0], "text": part[1]}
            if len(part) > 2:
                fragment["style"] = part[2]
            result.append(fragment)
    return result


def element(*parts, id="body", kind="text", area=(100, 50, 200, 100), shape=None, **overrides):
    data = {
        "id": id,
        "kind": kind,
        "area": list(area),
        "role": "body",
        "style": style(**overrides),
        "content": content(*parts),
    }
    if shape is not None:
        data["shape"] = shape
    return data


def texts(layout):
    return [line.text for line in layout.lines]


def widths(layout):
    return [line.width for line in layout.lines]


@pytest.fixture(scope="module")
def synthetic_environment(tmp_path_factory):
    return make_standard_fonts(tmp_path_factory.mktemp("fonts"))


@pytest.fixture(scope="module")
def synthetic(synthetic_environment):
    return SyntheticFonts.from_environment(synthetic_environment)


@pytest.fixture
def math_mode(monkeypatch):
    """office_math.math_mode の代わり（equation で、空でない断片が数式 1 つだけなら display）。

    計測が math_mode の判定で数式の置き方を決めることを、判定の実装と切り離して確かめる。
    """
    seen = []

    def fake(element):
        seen.append(element.get("id"))
        filled = [fragment for fragment in element.get("content", []) if str(fragment.get("text", ""))]
        single = len(filled) == 1 and filled[0].get("kind") == "math"
        return "display" if element.get("kind") == "equation" and single else "inline"

    module = types.ModuleType("slidekit.math.office_math")
    module.math_mode = fake
    monkeypatch.setitem(sys.modules, "slidekit.math.office_math", module)
    return seen


# ---- 折り返し ----


def test_wrap_breaks_between_japanese_characters_and_at_spaces():
    fonts = FakeFonts()  # 20 pt で 1 文字 10 pt
    assert texts(layout_text(content("あいうえおかきく"), style(), fonts, width=50)) == ["あいうえお", "かきく"]
    spaced = layout_text(content("ab cd ef"), style(), fonts, width=55)
    assert texts(spaced) == ["ab cd ", "ef"]
    assert widths(spaced) == [50, 20]  # 行末の空白は前の行に残し、幅に含めない
    assert texts(layout_text(content("abcあいう"), style(), fonts, width=45)) == ["abcあ", "いう"]
    assert texts(layout_text(content("あいabc"), style(), fonts, width=35)) == ["あい", "abc"]


def test_wrap_keeps_latin_words_and_math_fragments_whole():
    fonts = FakeFonts()
    assert texts(layout_text(content("alpha beta"), style(), fonts, width=70)) == ["alpha ", "beta"]
    mixed = layout_text(content("ab ", ("math", "x+y+z"), " cd"), style(), fonts, width=60)
    assert texts(mixed) == ["ab ", "$x+y+z$ ", "cd"]
    joined = layout_text(content("あい", ("math", "xyz"), "うえ"), style(), fonts, width=45)
    assert texts(joined) == ["あい", "$xyz$う", "え"]


def test_kinsoku_pushes_previous_character_before_closing_punctuation():
    fonts = FakeFonts()
    for closing in "。」、":
        lines = texts(layout_text(content(f"あいうえお{closing}か"), style(), fonts, width=50))
        assert lines == ["あいうえ", f"お{closing}か"]
    assert texts(layout_text(content("あいうえ「お"), style(), fonts, width=50)) == ["あいうえ", "「お"]
    # 数式の直後の「、」も行頭にしない。数式の前で分ける。
    assert texts(layout_text(content("あい", ("math", "xy"), "、う"), style(), fonts, width=45)) == ["あい", "$xy$、う"]
    # Office の「標準」の禁則に、PowerPoint が行頭に置かない小書きの仮名と長音記号を加えた表。中点も行頭に置かない。
    assert {"、", "。", "」", "・", "%"} <= NO_LINE_START and {"「", "（", "$"} <= NO_LINE_END
    assert set("ぁっょァッョヵヶㇰㇿーｧｯｰ") <= NO_LINE_START and not set("ぁょー") & NO_LINE_END
    # 全角空白は禁則の表に入れず、空白として扱う。
    assert "\u3000" not in NO_LINE_START | NO_LINE_END


@pytest.mark.parametrize("small", ["ょ", "っ", "ー", "ャ", "ｯ"])
def test_small_kana_and_long_vowel_push_previous_character(small):
    """PowerPoint は小書きの仮名と長音記号を行頭に置かず、直前の文字ごと次の行へ送る（追い出し。追い込みはしない）。"""
    lines = layout_text(content(f"あいうえお{small}か"), style(), FakeFonts(), width=50).lines
    assert [line.text for line in lines] == ["あいうえ", f"お{small}か"]
    assert not any(line.near_limit for line in lines)  # 追い込みがないので、誤差の範囲にする規則は加えない


def test_ideographic_space_is_treated_like_space():
    """全角空白（U+3000）は半角空白と同じく、前で分けず、前の行の末尾に残し、行末では幅に含めない。"""
    fonts = FakeFonts()
    layout = layout_text(content("あいうえ\u3000かきく"), style(), fonts, width=40)
    assert texts(layout) == ["あいうえ\u3000", "かきく"]  # 空白は内側の幅を超えて前の行に残る
    assert widths(layout) == [40, 30]
    # 行頭の全角空白は幅に含め、行末の全角空白だけを除く。中央揃えも行末の空白を除いた幅で置く。
    assert widths(layout_text(content("\u3000あ\u3000"), style(), fonts)) == [20]
    centered = layout_element(element("\u3000あ\u3000", align="center", area=(0, 0, 100, 40)), fonts)
    assert centered.line_left == pytest.approx((40,))
    # 全角空白の後では分けることができる。
    assert texts(layout_text(content("ab\u3000cd"), style(), fonts, width=35)) == ["ab\u3000", "cd"]


def test_explicit_newlines_are_wrapped_independently():
    fonts = FakeFonts()
    layout = layout_text(content("あいうえおか\nきく"), style(), fonts, width=50)
    assert texts(layout) == ["あいうえお", "か", "きく"]
    assert [line.hard_break for line in layout.lines] == [False, True, False]
    empty = layout_text(content("あ\n\nい\n"), style(), fonts, width=50)
    assert texts(empty) == ["あ", "", "い", ""]
    assert [line.hard_break for line in empty.lines] == [True, True, True, False]
    # 空の行も行として高さに含める（行送り 24 pt）。
    assert empty.content_height == pytest.approx(18 + 3 * 24 + FAKE_BOTTOM * 20)


def test_long_word_is_split_inside_like_powerpoint():
    """内側より長い語は、PowerPoint と同じく、収まる最後の文字の後で語の途中で分ける（1 行に 1 文字以上）。
    分ける位置のない部分の幅は unbreakable_width に残す（はみ出しの判定で text_overflow にする）。"""
    fonts = FakeFonts()
    layout = layout_text(content("a verylongword b"), style(), fonts, width=60)
    assert texts(layout) == ["a ", "verylo", "ngword ", "b"]
    assert [line.split for line in layout.lines] == [False, True, False, False]
    assert layout.unbreakable_width == pytest.approx(120) and max(widths(layout)) == 60
    assert layout.available_width == 60
    # 語の途中で分ける位置も禁則を守る（行頭の , と行末の ( を避けて前へずらす）。
    assert texts(layout_text(content("abcde,fgh"), style(), fonts, width=50)) == ["abcd", "e,fgh"]
    assert texts(layout_text(content("abcd(efgh"), style(), fonts, width=50)) == ["abcd", "(efgh"]
    # 1 文字より狭い幅でも、1 行に 1 文字を置く。
    assert texts(layout_text(content("abc"), style(), fonts, width=5)) == ["a", "b", "c"]
    # 内側より広い数式は途中で分けない（PowerPoint にも分けさせない）。直後の空白は前の行に残る。
    wide = layout_text(content("和 ", ("math", "a+b+c+d+e+f"), " です"), style(), fonts, width=50)
    assert texts(wide) == ["和 ", "$a+b+c+d+e+f$ ", "です"]
    assert not any(line.split for line in wide.lines) and wide.unbreakable_width == pytest.approx(110)


def test_word_wrap_false_keeps_explicit_lines():
    fonts = FakeFonts()
    for wrap in ({"word_wrap": False}, {"word_wrap": None}):
        layout = layout_text(content("alpha beta gamma\nx"), style(**wrap), fonts, width=50)
        assert texts(layout) == ["alpha beta gamma", "x"]
        assert layout.available_width is None and not layout.word_wrap
        assert not any(line.near_limit for line in layout.lines)
    # word_wrap を持たない古い版の場面も折り返さない。
    old = {key: value for key, value in style().items() if key != "word_wrap"}
    assert texts(layout_text(content("alpha beta"), old, fonts, width=50)) == ["alpha beta"]


@pytest.mark.parametrize(
    "width, lines, near",
    [
        (100, ["aaaa bbbbb"], [True]),  # 幅 100 の行が内側の幅 100 の帯（0.65 pt）に入る
        (102, ["aaaa bbbbb"], [False]),  # 帯の外
        (99.5, ["aaaa ", "bbbbb"], [True, False]),  # 次の語を含めた幅 100 が 99.5 + 帯 0.65 以下
        (99, ["aaaa ", "bbbbb"], [False, False]),  # 99 + 0.65 より広い
    ],
)
def test_near_limit_lines_are_flagged(width, lines, near):
    layout = layout_text(content("aaaa bbbbb"), style(), FakeFonts(), width=width)
    assert texts(layout) == lines
    assert [line.near_limit for line in layout.lines] == near


def test_hyphen_break_follows_rule():
    fonts = FakeFonts()
    assert texts(layout_text(content("well-known"), style(), fonts, width=55)) == ["well-", "known"]
    assert texts(layout_text(content("well‐known"), style(), fonts, width=55)) == ["well‐", "known"]
    # ハイフンの後が英数字でなければ分けない（語の全体を次の行へ送る）。
    assert texts(layout_text(content("ab well-(known)"), style(), fonts, width=125)) == ["ab ", "well-(known)"]


# ---- 文字の幅と行の配置 ----


def test_width_has_no_gap_at_script_boundary(synthetic):
    for value in ("AB漢字", "漢字AB", "A漢B字"):
        assert layout_text(content(value), synthetic_style(), synthetic).content_width == pytest.approx(60)


def test_latin_kerning_uses_kern_table_from_12pt(synthetic):
    def width(*parts, size=20, **overrides):
        return layout_text(content(*parts), synthetic_style(size_pt=size, **overrides), synthetic).content_width

    assert width("TAV", size=12) == pytest.approx((1500 - 80 - 60) / 1000 * 12)
    assert width("TAV", size=11.9) == pytest.approx(1.5 * 11.9)  # 12 pt 未満は詰めない
    assert width("あい") == pytest.approx(40)  # 和文の書体の kern 表は使わない
    assert width("T漢") == pytest.approx(30)  # 書体の境界では詰めない
    assert width("TA", bold=True) == pytest.approx((520 + 520 - 80) / 1000 * 20)


def test_kerning_crosses_fragment_boundaries_with_same_font(synthetic):
    """PowerPoint は断片（a:r）の境界をまたぐ欧文の対も kern 表のとおりに詰める（書体・大きさ・太さと forced が同じ境界）。
    大きさか太さが違う境界は詰めない（校正用の資料で確かめた）。"""

    def width(*parts):
        return layout_text(content(*parts), synthetic_style(), synthetic).content_width

    assert width("T", "A") == width("TA") == pytest.approx((1000 - 80) / 1000 * 20)
    assert width(("emphasis", "T", {"bold": False}), "A") == pytest.approx(18.4)  # 強調でも書体・大きさ・太さが同じ
    # 大きさか太さが違う境界、数式と明示改行をはさむ対は詰めない。
    assert width(("text", "T", {"size_pt": 30}), "A") == pytest.approx(15 + 10)
    assert width(("text", "T", {"bold": True}), "A") == pytest.approx(10.4 + 10)
    assert width("T", ("math", "x"), "A") == pytest.approx(10 + 10 + 10)
    assert layout_text(content("T", "\nA"), synthetic_style(), synthetic).content_width == pytest.approx(10)


def test_kerning_stops_at_forced_boundary():
    """欧文の書体で計測する幅の曖昧な記号（forced。書き出しは別の a:r にする）と通常の欧文の文字の境界は、同じ断片の中でも
    詰めない（PowerPoint は forced の区切りの境界をまたぐ対を詰めない。MEASURE_VERSION 5 の校正用の資料の A’・L’ で
    確かめた。’ は 6 から和文の書体で計測する）。forced の文字どうし（同じ区切りの中）は詰める。和文の書体で計測する記号
    （欧文の書体にない ※ と、PowerPoint が既定で和文の書体で描く “ ”）の境界は、書体が違うので詰めない。"""
    kerning = {
        ("T", "→"): -0.1,
        ("…", "A"): -0.2,
        ("→", "…"): -0.15,
        ("T", "A"): -0.05,
        ("T", "※"): -0.3,
        ("※", "A"): -0.3,
        ("T", "“"): -0.1,
        ("”", "A"): -0.2,
        ("“", "”"): -0.15,
    }
    fonts = FakeFonts(missing={"Latin": "※"}, kerning=kerning)  # 20 pt で 1 文字 10 pt。kern は書体によらず返す

    def width(value):
        return layout_text(content(value), style(), fonts).content_width

    runs = font_runs("T→A…A※A“”", style(), fonts, bold=False)
    assert [(run.text, run.family, run.forced) for run in runs] == [
        ("T", "Latin", False),
        ("→", "Latin", True),
        ("A", "Latin", False),
        ("…", "Latin", True),
        ("A", "Latin", False),
        ("※", "Japanese", True),
        ("A", "Latin", False),
        ("“”", "Japanese", True),
    ]
    assert width("T→") == width("…A") == pytest.approx(20)
    assert width("T→A…A") == pytest.approx(50)
    assert width("→…") == pytest.approx(20 - 3)  # 同じ forced の区切りの中の対
    assert width("TA") == pytest.approx(20 - 1)  # forced でない文字どうしは詰める
    # 和文の書体で計測する記号との境界は、書体が違うので詰めない（kern 表に値があっても）。和文の書体の文字どうしも詰めない。
    assert width("T※") == width("※A") == width("T“") == width("”A") == width("“”") == pytest.approx(20)
    assert width("T※A") == pytest.approx(30)


@pytest.mark.parametrize("size", [14, 20, 28])
@pytest.mark.parametrize("spacing", [1.0, 1.15, 1.5, 0.9])
def test_line_pitch_and_first_baseline_follow_powerpoint_model(synthetic, size, spacing):
    layout = layout_text(content("漢字\n漢字\n漢字"), synthetic_style(size_pt=size, line_spacing=spacing), synthetic)
    # 和文を含む行は欧文と和文の書体の組。行の上側は行間がちょうど 1 なら組の値、1 より大きければ行間 × 上限 0.9035
    # （組の値 0.976 のほうが大きい）、1 未満は 1.2 × (1 − 行間) を除く（行間 × 0.9035 より大きい）。
    # 行送りは 1.2 × 行間 × 文字サイズ。
    first = {1.0: SYNTHETIC_TOP, 1.15: 1.15 * 0.9035, 1.5: 1.5 * 0.9035, 0.9: SYNTHETIC_TOP - 0.12}[spacing] * size
    assert first == pytest.approx(line_top(SYNTHETIC_TOP, spacing) * size)
    baselines = [first, first + 1.2 * spacing * size, first + 2 * 1.2 * spacing * size]
    assert [line.baseline for line in layout.lines] == pytest.approx(baselines)
    # 最後の基線より下は行間によらない。
    assert layout.content_height == pytest.approx(baselines[-1] + SYNTHETIC_BOTTOM * size)


@pytest.fixture(scope="module")
def unequal(tmp_path_factory):
    """欧文と和文で usWin の割合が大きく違う合成フォント（和文 0.8・0.4 em で、和文の行の下が深い）。"""
    directory = tmp_path_factory.mktemp("unequal")
    latin = {char: (500, 0, 700) for char in "Aabc"} | {" ": (250, 0, 0)}
    japanese = {char: (1000, -80, 820) for char in "漢字あ"} | {"\u3000": (1000, 0, 0)}
    faces = {
        ("Test Latin", False): make_font(directory / "latin.ttf", "Test Latin", latin, ascent=900, descent=200),
        ("Test Japanese", False): make_font(
            directory / "japanese.ttf",
            "Test Japanese",
            japanese,
            ascent=880,
            descent=120,
            win_ascent=800,
            win_descent=400,
        ),
    }
    for family in ("Test Latin", "Test Japanese"):
        faces[(family, True)] = faces[(family, False)]
    return SyntheticFonts(faces)


def test_line_font_set_decides_line_position(unequal):
    """行の上下は行の書体の組で決まる。欧文だけの行と空の行（U+0020 だけの行を含む）は欧文の書体だけ、和文を含む行は
    欧文と和文の書体。書体の組が違う行が続くと、行送りは 前の行の下側 ＋ 次の行の上側 になる（1.2 × 行間 × 文字サイズと違う）。"""
    latin_top, latin_bottom = natural((0.9, 0.2))
    mixed_top, mixed_bottom = natural((0.9, 0.2), (0.8, 0.4))
    assert mixed_top < latin_top and mixed_bottom > latin_bottom
    # 和文の書体の usWin は hhea と違う値にしてある。計測は usWin を使う。
    assert unequal.face_metrics("Test Japanese").ascent == pytest.approx(0.8)
    one = layout_text(content("Aa\n漢字\n\nbc"), synthetic_style(), unequal)
    tops = [latin_top, mixed_top, latin_top, latin_top]  # 3 行目の空の行は欧文の書体だけ
    baselines = [tops[0] * 20]
    for previous, top in zip(tops, tops[1:]):
        baselines.append(baselines[-1] + (1.2 - previous) * 20 + top * 20)
    assert [line.baseline for line in one.lines] == pytest.approx(baselines)
    assert baselines[3] - baselines[2] == pytest.approx(24) and baselines[2] - baselines[1] != pytest.approx(24)
    assert one.content_height == pytest.approx(baselines[-1] + latin_bottom * 20)
    # 行間が 1 より大きいと、上側は行間 × min(組の値, 0.9035)。最後の行の下は行間によらない。
    spaced = layout_text(content("漢字\nAa"), synthetic_style(line_spacing=1.5), unequal)
    first = 1.5 * min(mixed_top, 0.9035) * 20
    second = first + (1.2 * 1.5 - 1.5 * min(mixed_top, 0.9035)) * 20 + 1.5 * min(latin_top, 0.9035) * 20
    assert [line.baseline for line in spaced.lines] == pytest.approx([first, second])
    assert spaced.content_height == pytest.approx(second + latin_bottom * 20)
    # 全角空白だけの行は、全角空白を計測した和文の書体を含む（和文の文字の行と同じ）。
    blank = layout_text(content("\u3000"), synthetic_style(), unequal)
    assert blank.lines[0].baseline == pytest.approx(mixed_top * 20)
    # 最後の行が空の行か U+0020 だけの行でも、文字ブロックの下端は欧文の書体だけの組で決まり、欧文の文字の最後の行と
    # 同じ高さになる（PowerPoint で、和文の書体が Meiryo の要素の末尾の空の行と空白だけの行で確かめた）。
    latin_last = layout_text(content("Aa\nA"), synthetic_style(line_spacing=1.5), unequal)
    assert latin_last.content_height == pytest.approx((1.8 + 1.5 * min(latin_top, 0.9035) + latin_bottom) * 20)
    for value in ("Aa\n", "Aa\n ", "漢字\n"):
        trailing = layout_text(content(value), synthetic_style(line_spacing=1.5), unequal)
        assert trailing.lines[-1].baseline - trailing.lines[-2].baseline == pytest.approx(
            (1.2 * 1.5 - 1.5 * min(latin_top if value[0] == "A" else mixed_top, 0.9035) + 1.5 * min(latin_top, 0.9035))
            * 20
        ), value
        assert trailing.content_height == pytest.approx(trailing.lines[-1].baseline + latin_bottom * 20), value
    assert layout_text(content("Aa\n"), synthetic_style(line_spacing=1.5), unequal).content_height == pytest.approx(
        latin_last.content_height
    )


def test_line_spacing_below_one_uses_larger_of_two_tops(unequal):
    """行間 1 未満の行の上側は、書体の値 − 1.2 ×（1 − 行間）と、行間 × min(書体の値, 0.9035) の大きいほう。
    PowerPoint の行間 0.9 の観測では、書体の値が上限より大きい組（Arial など）は前者に、上限より小さい組
    （Meiryo と Arial の和文の行）は後者に合った。"""
    latin_top, latin_bottom = natural((0.9, 0.2))
    mixed_top, _mixed_bottom = natural((0.9, 0.2), (0.8, 0.4))
    assert mixed_top < MODEL.line_top_limit_em < latin_top
    layout = layout_text(content("漢字\nAa"), synthetic_style(line_spacing=0.9), unequal)
    first = 0.9 * mixed_top  # 和文の行：mixed_top − 0.12 より大きい
    second_top = latin_top - 1.2 * (1 - 0.9)  # 欧文の行：0.9 × 0.9035 より大きい
    assert first > mixed_top - 0.12 and second_top > 0.9 * 0.9035
    second = first + (1.2 * 0.9 - first) + second_top
    assert [line.baseline for line in layout.lines] == pytest.approx([first * 20, second * 20])
    assert layout.content_height == pytest.approx((second + latin_bottom) * 20)  # 最後の行の下は行間によらない
    # 行間を 1 に近づけると、どちらの組も書体の値に近づく（1 から下では値が飛ばない）。
    for value, top in (("漢字", mixed_top), ("Aa", latin_top)):
        near = layout_text(content(value), synthetic_style(line_spacing=0.9999), unequal)
        assert near.lines[0].baseline == pytest.approx(top * 20, abs=0.01), value


def test_line_spacing_one_is_exact(synthetic):
    """行の上側は行間がちょうど 1 のときだけ書体の値になり、1 を少し超えると上限で抑えた値に飛ぶ。
    line_spacing_pt は 1.2 × 文字サイズで割って行間にするので、24 pt（20 pt の文字）は行間 1 と同じ。"""

    def first(**options):
        return layout_text(content("Aa"), synthetic_style(**options), synthetic).lines[0].baseline

    assert first(line_spacing=1.0) == first(line_spacing_pt=24) == pytest.approx(LATIN_TOP * 20)
    assert first(line_spacing=1.000001) == pytest.approx(LATIN_TOP * 20)  # 書き出しの行間（1/100000 単位）の丸めの範囲
    assert first(line_spacing=1.02) == pytest.approx(1.02 * 0.9035 * 20)
    assert MODEL.line_top_limit_em == 0.9035


def test_math_only_line_follows_display_mode():
    """独立した数式だけの行は数式だけの行（最後の基線から行の下側まで）として、文中の数式だけの行（書き出しが U+200B を
    加える段落）は欧文の書体の文字の行として置く。"""
    fonts = FakeFonts()
    display = layout_text(content(("math", "x")), style(line_spacing=1.2), fonts, display=True)
    inline = layout_text(content(("math", "x")), style(line_spacing=1.2), fonts)
    text_only = layout_text(content("x"), style(line_spacing=1.2), fonts)
    assert (
        display.lines[0].baseline
        == inline.lines[0].baseline
        == text_only.lines[0].baseline
        == pytest.approx(1.2 * 0.9 * 20)
    )
    assert display.content_height == pytest.approx(1.2 * 1.2 * 20)
    assert inline.content_height == text_only.content_height == pytest.approx(1.2 * 0.9 * 20 + FAKE_BOTTOM * 20)


def test_tall_math_extends_line(synthetic_environment):
    # 規則に当たらない背の高い数式（アクセントの式）。行間 1 では、数式の箱の行の上側を書体の組の値と比べる
    # （合成フォントの欧文の行は 0.982 em）。
    tall_box = (1.5, 1.1, 0.3, 1.3, 0.4)  # 行の上 1.3 em、下 0.4 em
    boxes = {"x": (0.5, 0.7, 0.1, 0.95, 0.1), r"\hat{T}": tall_box}
    exact = SyntheticFonts.from_environment(synthetic_environment, math=FakeFonts(math_boxes=boxes))
    assert layout_text(content("a", ("math", "x")), synthetic_style(), exact).lines[0].baseline == pytest.approx(
        LATIN_TOP * 20
    )
    assert layout_text(content("a", ("math", r"\hat{T}")), synthetic_style(), exact).lines[0].baseline == 26
    fonts = FakeFonts(math_boxes={r"\hat{T}": tall_box})
    tall = layout_text(content("a", ("math", r"\hat{T}")), style(), fonts)
    assert tall.lines[0].baseline == pytest.approx(26)
    assert tall.content_height == pytest.approx(26 + FAKE_BOTTOM * 20 + (8 - 6))
    assert tall.lines[0].ink_top == pytest.approx(22) and tall.lines[0].ink_bottom == pytest.approx(6)
    plain = layout_text(content("a", ("math", "x")), style(), fonts)
    assert plain.content_height == pytest.approx(18 + FAKE_BOTTOM * 20)
    two = layout_text(content("a", ("math", r"\hat{T}"), "\nb"), style(), fonts)
    assert [line.baseline for line in two.lines] == pytest.approx([26, 26 + 8 + 18])
    below = layout_text(content("a\n", ("math", r"\hat{T}")), style(), fonts)
    assert [line.baseline for line in below.lines] == pytest.approx([18, 18 + 6 + 26])
    assert below.content_height == pytest.approx(50 + 8)
    spaced = layout_text(content("a", ("math", r"\hat{T}")), style(line_spacing=1.5), fonts)
    assert spaced.lines[0].baseline == pytest.approx(1.5 * 26)  # 数式の箱にも行間を掛ける


STACKED = (
    r"\frac{a}{b}",
    r"\dfrac{1}{n}",
    r"\binom{n}{k}",
    r"\left(x\right)",
    r"\left.x\right|",
    r"\sum_i a_i",
    r"\sum\limits_{i} a_i",
    r"\int f",
    r"\sqrt{\frac{a}{b}}",
    r"(\frac{a}{b})^2",
)
UNDER = (r"\underset{i}{\frac{a}{b}}", r"\lim\limits_{n} a_n", r"\underset{i}{\max} x_i", r"(\underset{i}{x})^2")
RADICAL = (r"\sqrt{n}", r"O(\sqrt{n\log n})", r"\sqrt[3]{x^2}", r"\sqrt{x_0}+y_i")
NOT_STACKED = (
    "x^2",
    "(x)^2",
    "a/b",
    r"\hat{x}",
    r"\overline{x}",
    r"\lim_{n} a_n",
    r"\text{abc}",
    r"\overset{n}{x}",
    r"x^{\frac{1}{2}}",
    r"x_{\frac{1}{2}}",
    r"\sqrt{n}^2",
    r"\sqrt{n}+x^2",
    r"\hat{x}+\sqrt{y}",
    r"\overset{n}{\sqrt{x}}",
)


@pytest.mark.parametrize("expression", STACKED + UNDER + RADICAL + NOT_STACKED)
def test_stacked_math_line_top_follows_ink(expression):
    """分数・\\binom・伸びる括弧・大型演算子・下に置くもの（\\underset・\\limits の下）を含み、インクの上下の和が文字サイズを
    超える式は、背の高い数式の規則で行の上側を決める（和が 1.3 em 未満の式は 1.2 × 行間 × インクの上以上。下に置くものを
    含む式は、和によらず a・d・h から決める）。それらを含まず根号を含む文中の式は、根号の規則（行間 1 より大きいとき
    1.17 × 行間 × インクの上以上、ちょうど 1 のとき 1.22 × インクの上以上）にする。根号の外に上へ出る部分
    （上付きの添字・上に置くもの・アクセント・上線）がある式は、インクの上が根号の上とは限らないので当てない。
    分数などが添字の中（添字の土台の中は除く）か、上に置くもの（\\overset）の中にあるだけの式と、普通の括弧・横書きの分数
    などの式は、背が高くても数式の計測の上側に行間を掛けた値にする（PowerPoint は添字の中の分数と上に置くものの行に
    規則を使わない）。行間がちょうど 1 の行では、その上側をインクの上 + 0.147 em より小さくしない。"""
    fonts = FakeFonts(math_ascent=1.0, math_descent=0.1)  # インクの上下の和 1.1 em。行の上側 1.0 em
    # 下に置くもの：a = 1.147 em（書体の組の行間 1.0 の上側 0.9 em より大きい）、d = 0.1 em、h = 1.247 em。
    # 行間 1.2 で増える分 0.2 h は 1.17 d より大きいので、4 分の 3 を上に置く。
    under = 1.147 + 0.75 * 0.2 * 1.247
    tops = {1.0: (1.2, 1.147, 1.22, 1.147), 1.2: (1.2 * 1.2, under, 1.17 * 1.2, 1.2)}
    for spacing, (stacked, below, radical, other) in tops.items():
        layout = layout_text(content(("math", expression)), style(line_spacing=spacing), fonts)
        top = (
            stacked
            if expression in STACKED
            else below if expression in UNDER else radical if expression in RADICAL else other
        )
        assert layout.lines[0].baseline == pytest.approx(top * 20), spacing
        # 独立した数式の根号は規則に当てない（独立した数式の背の高い根号は、行間を掛けた計測の上下で合った）。
        display = layout_text(content(("math", expression)), style(line_spacing=spacing), fonts, display=True)
        top = stacked if expression in STACKED else below if expression in UNDER else other
        assert display.lines[0].baseline == pytest.approx(top * 20), spacing
    # インクの上下の和が文字サイズ以下なら、どの式も規則に当てない（行間 1.0 でも、インクの上 + 0.147 em は 0.9 em より低い）。
    for spacing in (1.0, 1.2):
        low = layout_text(
            content(("math", expression)), style(line_spacing=spacing), FakeFonts(math_ascent=0.75, math_descent=0.1)
        )
        assert low.lines[0].baseline == pytest.approx(spacing * 0.9 * 20), spacing


# 背の高い数式の規則の確かめに使う式の (幅, インクの上, インクの下, 行の上, 行の下)（em）。
STACK_BOXES = {
    r"\frac{a}{b}": (0.6, 1.2, 0.7, 1.2, 0.7),  # 和 1.9 em。a = 1.347、h = 2.047 em
    r"\sqrt{\frac{a}{b}}": (1.2, 1.4, 0.8, 1.5, 0.8),  # 和 2.2 em。a = 1.547、h = 2.347 em
    r"\left(x^2\right)": (0.9, 0.8, 0.45, 0.8, 0.45),  # 和 1.25 em（1.3 em 未満）
    r"\sum_i a_i": (1.5, 1.0, 0.2, 1.3, 0.2),  # 同。計測の行の上側 1.3 em が 1.2 × 行間 × インクの上より大きい
    r"\dfrac{1}{n}": (0.5, 0.9, 0.1, 0.9, 0.1),  # 和 1.0 em（文字サイズを超えない）：規則に当たらない
    r"\left(x\right)": (0.9, 0.75, 0.2, 0.75, 0.2),  # 和 0.95 em：同
    r"x_{\frac{1}{2}}": (0.5, 0.6, 0.6, 0.6, 0.6),  # 添字の中の分数：規則に当たらない
    r"\int_0^1 f": (1.2, 1.4, 0.3, 1.4, 0.3),  # 和 1.7 em で下が浅い。a = 1.547、h = 1.847 em
}


def test_stacked_math_line_follows_ink_and_spacing():
    """背の高い数式の行の上側と下側（PowerPoint 16.0.20430.20182 の校正用の資料の math_stack・math_stack_more・
    math_stack_spacing・math_limits のページで確かめた規則）。

    a = インクの上 + 0.147 em、d = インクの下、h = a + d とする。インクの上下の和が 1.3 em 以上の式は、行間で増える分
    e = (sp − 1) × h が 1.17 d 以下ならすべてを行の上に、超えれば 4 分の 3 を行の上に置き、残りを行の下に置く
    （切り替わる行間は d と h で決まる。校正用の資料の 20 pt の独立した分数は行間 1.3 まですべて上で 1.4 から 4 分の 3、
    総和は行間 1.4 まですべて上で 1.5 から 4 分の 3）。
    1.0〜1.3 em の式は、上側 1.2 × sp × インクの上、下側 d。どちらも上側は数式の計測の行の上 × sp より小さくせず、
    d に行間を掛けない。
    """
    assert (
        MODEL.math_stack_gap_em,
        MODEL.math_stack_full_em,
        MODEL.math_stack_above_share,
        MODEL.math_stack_above_full,
        MODEL.math_stack_top_ratio,
    ) == (0.147, 1.3, 0.75, 1.17, 1.2)
    fonts = FakeFonts(math_boxes=STACK_BOXES)

    def first(expression, spacing, **options):
        return layout_text(content(("math", expression)), style(line_spacing=spacing), fonts, **options)

    a, h = 1.347, 2.047  # \frac{a}{b}（d = 0.7 em。1.17 d = 0.819 em は行間 1.4 の e = 0.819 em をわずかに上回る）
    spacings = (1.0, 1.15, 1.2, 1.25, 1.3, 1.45, 1.5, 2.0)
    tops = {spacing: first(r"\frac{a}{b}", spacing).lines[0].baseline for spacing in spacings}
    assert tops == pytest.approx(
        {
            1.0: a * 20,
            1.15: (a + 0.15 * h) * 20,
            1.2: (a + 0.2 * h) * 20,
            1.25: (a + 0.25 * h) * 20,
            1.3: (a + 0.3 * h) * 20,
            1.45: (a + 0.75 * 0.45 * h) * 20,
            1.5: (a + 0.375 * h) * 20,
            2.0: (a + 0.75 * h) * 20,
        }
    )
    # 下が浅い式は、同じ行間でも早く 4 分の 3 に切り替わる（1.17 d = 0.351 em。行間 1.15 の e は 0.277 em、1.2 は 0.369 em）。
    a, h = 1.547, 1.847  # \int_0^1 f
    shallow = {spacing: first(r"\int_0^1 f", spacing).lines[0].baseline for spacing in (1.15, 1.2)}
    assert shallow == pytest.approx({1.15: (a + 0.15 * h) * 20, 1.2: (a + 0.75 * 0.2 * h) * 20})
    assert first(r"\int_0^1 f", 1.2, display=True).content_height - shallow[1.2] == pytest.approx(
        (0.3 + 0.25 * 0.2 * h) * 20
    )
    # 独立した数式だけの行の文字ブロックは、上側と下側の和（行間 sp では sp × h）。
    a, h = 1.347, 2.047
    for spacing in spacings:
        display = first(r"\frac{a}{b}", spacing, display=True)
        assert display.content_height == pytest.approx(spacing * h * 20), spacing
    baselines = {expression: first(expression, 1.2).lines[0].baseline for expression in STACK_BOXES}
    assert baselines == pytest.approx(
        {
            r"\frac{a}{b}": (a + 0.2 * h) * 20,
            r"\sqrt{\frac{a}{b}}": (1.547 + 0.2 * 2.347) * 20,
            r"\left(x^2\right)": 1.2 * 1.2 * 0.8 * 20,
            r"\sum_i a_i": 1.2 * 1.3 * 20,
            # 規則に当たらない式は、書体の組の上側（1.2 × 0.9 em）と数式の計測の行の上 × 行間の大きいほう。
            r"\dfrac{1}{n}": 1.2 * 0.9 * 20,
            r"\left(x\right)": 1.2 * 0.9 * 20,
            # 下が書体の組の下（0.3 em）より深い分の 0.5 × 行間を、行の上側に加える（test_deep_math_lowers_spaced_line）。
            r"x_{\frac{1}{2}}": (1.2 * 0.9 + 0.5 * 1.2 * (0.6 - 0.3)) * 20,
            r"\int_0^1 f": (1.547 + 0.75 * 0.2 * 1.847) * 20,
        }
    )


def test_stacked_math_line_bottom_has_no_spacing():
    """背の高い数式の規則で置いた行の最後の行の下側は、行間を掛けずに、文字の段落の行では d、独立した数式だけの行では
    d と行間で増えた分の残りにする（下揃えの文字ブロックの下端に現れる）。分数の行の次の行の行送りは、分数の行の下側
    （d と残り）と次の行の上側の和。規則に当たらない数式の最後の行の下側は、行間が 1 より大きければ d（書体の組の下より深いとき）。"""
    fonts = FakeFonts(math_boxes=STACK_BOXES)
    a, h = 1.347, 2.047
    # 下揃え：最後の基線から文字ブロックの下端（内側の下端）までが d。
    for spacing in (1.0, 1.2, 1.5, 2.0):
        bottom = layout_element(
            element(("math", r"\frac{a}{b}"), area=(0, 0, 200, 100), line_spacing=spacing, valign="bottom"), fonts
        )
        assert bottom.block_top + bottom.text.lines[0].baseline == pytest.approx(100 - 0.7 * 20), spacing
    # 独立した数式だけの行は d と残り（行間 1.5 で 0.125 h、2.0 で 0.25 h。1.2 以下は残りがない）。
    for spacing, rest in ((1.2, 0.0), (1.5, 0.125), (2.0, 0.25)):
        display = layout_text(content(("math", r"\frac{a}{b}")), style(line_spacing=spacing), fonts, display=True)
        assert display.content_height - display.lines[0].baseline == pytest.approx((0.7 + rest * h) * 20), spacing
    # 文字と分数の段落の最後の行も d（行間 1.5）。
    text = layout_text(content("a\nb ", ("math", r"\frac{a}{b}")), style(line_spacing=1.5), fonts)
    assert text.content_height - text.lines[-1].baseline == pytest.approx(0.7 * 20)
    middle = layout_text(content(("math", r"\left(x^2\right)")), style(line_spacing=1.2), fonts)
    assert middle.content_height - middle.lines[0].baseline == pytest.approx(0.45 * 20)  # 文字の行の 0.3 em より深い
    # 分数の行の次の行：行送りは分数の行の下側（d と残り）と次の行の上側（行間 × 0.9 em）の和。
    for spacing, rest in ((1.2, 0.0), (1.5, 0.125)):
        two = layout_text(content(("math", r"\frac{a}{b}"), "\nb"), style(line_spacing=spacing), fonts)
        first = (a + (0.2 if spacing == 1.2 else 0.375) * h) * 20
        assert [line.baseline for line in two.lines] == pytest.approx(
            [first, first + (0.7 + rest * h) * 20 + spacing * 0.9 * 20]
        ), spacing
    # 規則に当たらない添字の中の分数（d = 0.6 em）：最後の行の下側は d。行の上側には 0.5 × 行間 ×（d − 0.3 em）を加える。
    script = layout_text(content(("math", r"x_{\frac{1}{2}}")), style(line_spacing=1.2), fonts)
    assert script.content_height == pytest.approx((1.2 * 0.9 + 0.5 * 1.2 * 0.3 + 0.6) * 20)


def test_deep_math_lowers_spaced_line():
    """行間が 1 より大きい文字の段落の行で、規則に当たらない数式のインクの下 d が書体の組の下 D × s より深いときは、
    行の上側に 0.5 × 行間 ×（d − D × s）を加え、最後の行の下側を d にする（実際の資料の PDF の、和文と y_j のような添字の行と
    添字のある根号の行の、行間 1.2 で確かめた。PowerPoint の行は、上揃えで (行間 − 1) × d に近い分だけ下がる）。
    行間 1 ちょうどと 1 未満、独立した数式だけの行、d が D × s より浅い行は、今までどおり。"""
    deep = FakeFonts(math_boxes={"y_j": (0.9, 0.7, 0.35, 0.7, 0.35), "a_i": (0.9, 0.7, 0.2, 0.7, 0.2)})

    def placed(*parts, spacing, **options):
        return layout_text(content(*parts), style(line_spacing=spacing), deep, **options)

    assert MODEL.math_descent_above == 0.5
    line = placed("和文の ", ("math", "y_j"), " の行", spacing=1.2)
    excess = (0.35 - FAKE_BOTTOM) * 20
    assert line.lines[0].baseline == pytest.approx(1.2 * 0.9 * 20 + 0.5 * 1.2 * excess)
    assert line.content_height - line.lines[0].baseline == pytest.approx(0.35 * 20)
    # 次の行の行送りは、行間を掛けた数式の下（1.2 × 0.35 em）と次の行の上側の和のまま。
    two = placed(("math", "y_j"), " の行\n次の行", spacing=1.2)
    assert two.lines[1].baseline - two.lines[0].baseline == pytest.approx((1.2 * 0.35 + 1.2 * 0.9) * 20)
    # 行間 1.5（確かめていない）も同じ式。
    assert placed(("math", "y_j"), spacing=1.5).lines[0].baseline == pytest.approx(1.5 * 0.9 * 20 + 0.5 * 1.5 * excess)
    # 行間 1 ちょうど、1 未満、独立した数式だけの行、浅い添字の行は、上側を変えない。
    assert placed(("math", "y_j"), spacing=1.0).lines[0].baseline == pytest.approx(0.9 * 20)
    assert placed(("math", "y_j"), spacing=0.9).lines[0].baseline == pytest.approx(0.9 * 0.9 * 20)
    assert placed(("math", "y_j"), spacing=1.2, display=True).lines[0].baseline == pytest.approx(1.2 * 0.9 * 20)
    shallow = placed("和文の ", ("math", "a_i"), spacing=1.2)
    assert shallow.lines[0].baseline == pytest.approx(1.2 * 0.9 * 20)
    assert shallow.content_height - shallow.lines[0].baseline == pytest.approx(FAKE_BOTTOM * 20)


def test_inline_radical_line_follows_ink():
    """文中の背の高い根号（インクの上下の和が文字サイズを超え、分数などを含まない式）は、行間 1 以上で次のように置く
    （実際の資料の PDF の O(\\sqrt{n\\log n}) の形の式の 14・18・20 pt、行間 1.2 の上揃え・中揃え・下揃えと、校正用の資料の
    18・20・24 pt、行間 1.0・1.2・1.5 で確かめた）。独立した数式の根号は規則に当てない。

    - 行の上側：行間がちょうど 1 なら T1 = 1.22 × インクの上、1 より大きければ 1.17 × 行間 × インクの上（どちらも計測の
      行の上 × 行間以上）。
    - 次の行への行送りに使う下側：行間 × (T1 + d) − 上側（行間 1.0 の行の高さに行間を掛けた値から上側を引いた値）。
      最後の行の下側は d。
    """
    root = r"O(\sqrt{n\log n})"
    fonts = FakeFonts(math_boxes={root: (5.6, 0.934, 0.308, 1.0, 0.308)})  # インクの上下の和 1.24 em
    assert (MODEL.math_radical_top_ratio, MODEL.math_radical_single_ratio) == (1.17, 1.22)
    single = 1.22 * 0.934
    for spacing in (1.0, 1.2, 1.5):
        line = layout_text(content("計算は ", ("math", root), " である"), style(line_spacing=spacing), fonts)
        top = single if spacing == 1.0 else max(spacing * 1.0, 1.17 * spacing * 0.934)
        assert line.lines[0].baseline == pytest.approx(top * 20), spacing
        assert line.content_height - line.lines[0].baseline == pytest.approx(0.308 * 20), spacing
    # 次の行の行送りは、根号の行の下側（行間 1.2 で 1.2 × (T1 + d) − 上側。書体の組の下側 0.36 em より深い）と、
    # 次の行の上側の和。行間 1.0 の下側は d（書体の組の下側 0.3 em より深い）。
    two = layout_text(content(("math", root), "\n次の行"), style(line_spacing=1.2), fonts)
    lower = 1.2 * (single + 0.308) - 1.17 * 1.2 * 0.934
    assert two.lines[1].baseline - two.lines[0].baseline == pytest.approx((lower + 1.2 * 0.9) * 20)
    exact = layout_text(content(("math", root), "\n次の行"), style(), fonts)
    assert exact.lines[1].baseline - exact.lines[0].baseline == pytest.approx((0.308 + 0.9) * 20)
    display = layout_text(content(("math", root)), style(line_spacing=1.2), fonts, display=True)
    assert display.lines[0].baseline == pytest.approx(1.2 * 1.0 * 20)
    # 行間 1 未満は確かめていないので、計測の行の上 × 行間のまま。
    tight = layout_text(content(("math", root)), style(line_spacing=0.9), fonts)
    assert tight.lines[0].baseline == pytest.approx(0.9 * 1.0 * 20)


def test_deep_radical_lowers_spaced_line():
    """下の深い文中の根号（\\sqrt{y_j}）の行は、行の上側を、根号の規則の上側と、下の深い数式の規則（書体の組の上側 +
    0.5 × 行間 ×（d − D × s））の大きいほうにする（校正用の資料の 20 pt、行間 1.2 で確かめた）。"""
    root = r"\sqrt{y_j}"
    fonts = FakeFonts(
        math_boxes={root: (1.5, 0.8, 0.6, 0.85, 0.6)}
    )  # インクの上下の和 1.4 em、d は 0.3 em より 0.3 em 深い
    line = layout_text(content("根号 ", ("math", root), " の値"), style(line_spacing=1.2), fonts)
    radical = max(1.2 * 0.85, 1.17 * 1.2 * 0.8)
    deep = 1.2 * 0.9 + 0.5 * 1.2 * (0.6 - FAKE_BOTTOM)
    assert deep > radical
    assert line.lines[0].baseline == pytest.approx(deep * 20)
    assert line.content_height - line.lines[0].baseline == pytest.approx(0.6 * 20)
    # 下が浅い根号は、根号の規則の上側のまま。
    shallow = FakeFonts(math_boxes={root: (1.5, 0.8, 0.25, 0.85, 0.25)})
    assert layout_text(content("根号 ", ("math", root), " の値"), style(line_spacing=1.2), shallow).lines[
        0
    ].baseline == (pytest.approx(max(1.2 * 0.9, radical) * 20))


def test_underset_rule_has_floor_of_single_spacing_top():
    """下に置くもの（\\underset・\\limits の下の添字）のある式は、インクの上下の和によらず背の高い数式の規則で a・d・h から
    置き、a = インクの上 + 0.147 em に、行の書体の組の行間 1.0 の上側（FakeFonts の組は 0.9 em）の下限を設ける
    （PowerPoint の \\underset{x \\in S}{\\arg\\min} f(x) の行間 1.0・1.15・1.5・2.0 と、\\underset{i}{\\max} x_i の行間 1.15 で
    確かめた）。"""
    expression = r"\underset{x \in S}{\arg\min} f(x)"
    fonts = FakeFonts(math_boxes={expression: (4.0, 0.7, 0.85, 0.7, 0.85)})  # a は 0.847 em より 0.9 em が大きい
    a, d = 0.9, 0.85
    for spacing, share in ((1.0, 1.0), (1.15, 1.0), (1.5, 1.0), (2.0, 0.75)):  # 行間 1.5 の e = 0.875 em は 1.17 d 以下
        line = layout_text(content(("math", expression)), style(line_spacing=spacing), fonts, display=True)
        extra = (spacing - 1) * (a + d)
        assert line.lines[0].baseline == pytest.approx((a + share * extra) * 20), spacing
        assert line.content_height == pytest.approx((a + d + extra) * 20), spacing
    # 下が浅く、インクの上下の和が 1.3 em 未満の式も同じ規則にする（1.0〜1.3 em の式の規則なら 1.15 × 0.9 em）。
    shallow = r"\underset{i}{\max} x_i"
    fonts = FakeFonts(math_boxes={shallow: (2.5, 0.6, 0.5, 0.6, 0.5)})  # 和 1.1 em。a = 0.9、h = 1.4 em
    line = layout_text(content(("math", shallow)), style(line_spacing=1.15), fonts, display=True)
    assert line.lines[0].baseline == pytest.approx((0.9 + 0.15 * 1.4) * 20)
    assert line.content_height == pytest.approx(1.15 * 1.4 * 20)
    inline = layout_text(content(("math", shallow)), style(line_spacing=1.15), fonts)
    assert inline.lines[0].baseline == pytest.approx((0.9 + 0.15 * 1.4) * 20)
    assert inline.content_height - inline.lines[0].baseline == pytest.approx(0.5 * 20)


def test_single_spacing_math_line_top_follows_ink():
    """行間がちょうど 1 の行では、規則に当たらない数式の行の上側を、数式の計測の行の上とインクの上 + 0.147 em の
    大きいほうにする（PowerPoint の \\overset{n}{x}、20 pt の上揃えで確かめた。行間 1 より大きい行は行間 × 計測の行の上）。"""
    over = r"\overset{n}{x}"
    fonts = FakeFonts(math_boxes={over: (0.6, 1.05, 0.0, 1.05, 0.0)})
    assert layout_text(content(("math", over)), style(), fonts).lines[0].baseline == pytest.approx((1.05 + 0.147) * 20)
    assert layout_text(content(("math", over)), style(line_spacing=1.000001), fonts).lines[0].baseline == pytest.approx(
        (1.05 + 0.147) * 20
    )
    for spacing in (1.02, 1.5):
        line = layout_text(content(("math", over)), style(line_spacing=spacing), fonts)
        assert line.lines[0].baseline == pytest.approx(spacing * 1.05 * 20), spacing
    # 低い式は書体の組の上側のまま（0.7 + 0.147 em < 0.9 em）。
    assert layout_text(content("a ", ("math", "x")), style(), FakeFonts()).lines[0].baseline == pytest.approx(0.9 * 20)


def test_stacked_math_rule_uses_fragment_size_and_keeps_rule_below_single_spacing():
    # 文字サイズは数式の断片の大きさで比べる（10 pt の断片で和 18 pt。要素の 20 pt は超えないが、断片の 10 pt を超える）。
    small = FakeFonts(math_boxes={r"\frac{n}{x}": (0.5, 1.7, 0.1, 1.7, 0.1)})
    smaller = layout_text(content("a ", ("math", r"\frac{n}{x}", {"size_pt": 10})), style(), small)
    assert smaller.lines[0].baseline == pytest.approx((1.7 + 0.147) * 10)  # 書体の組の上側 18 pt より下
    # 行間 1 未満は確かめていないので、上側は max(計測の行の上, 1.2 × インクの上) × 行間、下側は行の下 × 行間のまま。
    fonts = FakeFonts(math_boxes=STACK_BOXES)
    tight = layout_text(content(("math", r"\frac{a}{b}")), style(line_spacing=0.9), fonts)
    assert tight.lines[0].baseline == pytest.approx(0.9 * 1.2 * 1.2 * 20)
    lower = 1.2 * 0.9 * 20 - 0.9 * 0.9 * 20  # 文字の行の下側
    assert tight.content_height == pytest.approx(tight.lines[0].baseline + FAKE_BOTTOM * 20 + 0.9 * 0.7 * 20 - lower)


def test_valign_and_align_positions():
    fonts = FakeFonts()
    block = 18 + FAKE_BOTTOM * 20
    for valign, top in (("top", 50), ("middle", 50 + (100 - block) / 2), ("bottom", 150 - block)):
        assert layout_element(element("abc", valign=valign), fonts).block_top == pytest.approx(top)
    for align, left in (("left", 100), ("center", 185), ("right", 270)):
        assert layout_element(element("abc", align=align), fonts).line_left == pytest.approx((left,))
    justified = layout_element(element("aaaa bbbb cccc", align="justify", area=(100, 50, 100, 100)), fonts)
    assert texts(justified.text) == ["aaaa bbbb ", "cccc"]
    assert widths(justified.text) == [100, 40]  # 折り返した行は内側の幅、各明示行の最後の行は内容の幅
    assert justified.line_left == (100, 100)
    padded = layout_element(
        element("abc", padding_pt={"left": 10, "top": 5, "right": 20, "bottom": 15}, valign="bottom", align="right"),
        fonts,
    )
    assert padded.inner == Rect(110, 55, 170, 80)
    assert padded.block_top == pytest.approx(135 - block)
    assert padded.line_left == pytest.approx((250,))


def test_ink_uses_glyph_outlines(synthetic):
    layout = layout_text(content("gA"), synthetic_style(), synthetic)
    assert (layout.lines[0].ink_top, layout.lines[0].ink_bottom) == pytest.approx((14, 4))
    small = layout_text(content("ac"), synthetic_style(), synthetic)
    assert (small.lines[0].ink_top, small.lines[0].ink_bottom) == pytest.approx((10, 0))
    placed = layout_element(
        {**element(area=(10, 20, 200, 50)), "style": synthetic_style(), "content": content("gA")}, synthetic
    )
    baseline = 20 + LATIN_TOP * 20  # 欧文だけの行の上側
    assert tuple(placed.ink) == pytest.approx((10, baseline - 14, 20, 18))
    blank = layout_element({**element(), "style": synthetic_style(), "content": content("   ")}, synthetic)
    assert blank.ink is None and blank.text.lines[0].ink_top == blank.text.lines[0].ink_bottom == 0


@pytest.mark.parametrize("points", [24, 30])
def test_line_spacing_pt_model(points):
    layout = layout_text(content("a\nb\nc"), style(line_spacing_pt=points, line_spacing=1.5), FakeFonts())
    spacing = points / (1.2 * 20)  # 行間に換算する
    first = spacing * 0.9 * 20  # FakeFonts の組の上側 0.9 em は上限 0.9035 より小さいので、行間 1 以上で 行間 × 0.9 em
    assert [line.baseline for line in layout.lines] == pytest.approx([first, first + points, first + 2 * points])
    assert layout.content_height == pytest.approx(first + 2 * points + FAKE_BOTTOM * 20)


# ---- 表の自動寸法と配置の検査が同じ高さの式を使う ----


@pytest.fixture(scope="module")
def weighted(tmp_path_factory):
    """太字の書体の usWin が通常と違う合成フォント（太字が高さの式に作用することを確かめる）。"""
    directory = tmp_path_factory.mktemp("weighted")
    latin = {char: (500, 0, 700) for char in "abc"} | {" ": (250, 0, 0)}
    bold_latin = {char: (advance + 100, y_min, y_max) for char, (advance, y_min, y_max) in latin.items()}
    japanese = {char: (1000, -80, 820) for char in "漢字"}
    faces = {
        ("Test Latin", False): make_font(directory / "latin.ttf", "Test Latin", latin, ascent=900, descent=200),
        ("Test Latin", True): make_font(
            directory / "latin-bold.ttf", "Test Latin", bold_latin, ascent=950, descent=250, style="Bold"
        ),
        ("Test Japanese", False): make_font(
            directory / "japanese.ttf", "Test Japanese", japanese, ascent=1080, descent=250
        ),
        ("Test Japanese", True): make_font(
            directory / "japanese-bold.ttf", "Test Japanese", japanese, ascent=1100, descent=300, style="Bold"
        ),
    }
    return SyntheticFonts(faces, math=FakeFonts(math_boxes={"T": (1.0, 1.1, 0.35, 1.3, 0.4)}))


def test_height_combinations_use_one_formula(weighted):
    """以前は表の自動寸法と配置の検査が別の式で高さを求め、line_spacing_pt があると食い違った（2 行で 10 pt の指定に表 40 pt・検査 30 pt）。
    行間・line_spacing_pt・余白・太字・数式の組で、表と内容から決める寸法が使う layout_text の高さが PowerPoint の
    モデルの式どおりで、配置の検査が使う layout_element でも同じ文字ブロックになり、はみ出さないことを確かめる。
    文字は欧文だけなので、行の書体の組は欧文の書体（通常は usWin 0.9・0.2 em、太字は 0.95・0.25 em）だけである。"""
    latin = {False: (0.9, 0.2), True: (0.95, 0.25)}
    boxes = {None: None, "x": (0.7, 0.1), "T": (1.3, 0.4)}  # 数式の行の上下（em）。x は行を広げない
    paddings = [0, 6, {"left": 4, "top": 2, "right": 8, "bottom": 10}]
    heights = {}
    for spacing in (1.0, 1.5):
        for spacing_pt in (None, 10, 30):
            for padding in paddings:
                for bold in (False, True):
                    for expression in boxes:
                        case = (spacing, spacing_pt, json.dumps(padding), bold, expression)
                        parts = ("abc", ("math", expression), "\nab") if expression else ("abc\nab",)
                        options = {"line_spacing": spacing, "padding_pt": padding, "bold": bold}
                        if spacing_pt is not None:
                            options["line_spacing_pt"] = spacing_pt
                        measured = layout_text(content(*parts), synthetic_style(**options), weighted, width=200)
                        sp = spacing_pt / (1.2 * 20) if spacing_pt else spacing
                        top_em, bottom_em = natural(latin[bold])
                        ascent = line_top(top_em, sp) * 20
                        descent = 1.2 * sp * 20 - ascent  # 行送りに使う下側
                        first_ascent, first_descent = ascent, descent
                        if boxes[expression]:
                            first_ascent = max(ascent, sp * boxes[expression][0] * 20)
                            first_descent = max(descent, sp * boxes[expression][1] * 20)
                        second = first_ascent + first_descent + ascent
                        block = second + bottom_em * 20
                        insets = padding_insets({"padding_pt": padding})
                        assert measured.content_height == pytest.approx(block), case
                        assert measured.height == pytest.approx(
                            math.ceil((block + insets.top + insets.bottom) * 10 - 1e-6) / 10
                        ), case
                        placed = layout_element(
                            {
                                "id": "cell",
                                "kind": "box",
                                "shape": "rectangle",
                                "area": [10, 20, 200, measured.height],
                                "style": synthetic_style(**options),
                                "content": content(*parts),
                            },
                            weighted,
                        )
                        assert texts(placed.text) == texts(measured), case
                        assert placed.text.content_height == pytest.approx(measured.content_height), case
                        assert placed.inner.h >= placed.text.content_height - 1e-9, case
                        overflow = [item for item in fit_findings(placed) if item["code"] == "text_overflow"]
                        if spacing_pt == 10:
                            # 文字サイズより狭い行送りでは 1 行目の字形が枠の上へ出る（枠の高さの不足ではない）。
                            assert {item["observations"]["overflow_pt"]["side"] for item in overflow} <= {"top"}, case
                        else:
                            assert not overflow, case
                        heights[case] = measured.content_height
    # 太字の書体の値と line_spacing_pt が高さに作用している（式の分岐を通っている）。
    # （行間 1 では、書体の組が 1 つの行の高さは書体によらず 1.2 em になる。）
    assert heights[(1.5, None, "0", True, None)] != heights[(1.5, None, "0", False, None)]
    assert heights[(1.0, None, "0", True, None)] == heights[(1.0, None, "0", False, None)] == pytest.approx(48)
    # 10 pt の行送り（行間 10/24）では、行間 1 未満の 2 つの値のうち 行間 × 0.9035 のほうが大きい。
    assert line_top(LATIN_TOP, 10 / 24) == pytest.approx(10 / 24 * 0.9035)
    assert 10 / 24 * 0.9035 > LATIN_TOP - 1.2 * (1 - 10 / 24)
    assert heights[(1.0, 10, "0", False, None)] == pytest.approx(10 / 24 * 0.9035 * 20 + 10 + LATIN_BOTTOM * 20)
    assert heights[(1.5, 30, "0", False, None)] == heights[(1.0, 30, "0", False, None)]


# ---- 余白 ----


def test_normalize_padding():
    assert normalize_padding(6) == 6.0
    assert normalize_padding({"left": 12}, base=6) == {"left": 12.0, "top": 6.0, "right": 6.0, "bottom": 6.0}
    assert normalize_padding({"left": 4, "top": 4, "right": 4, "bottom": 4}) == 4.0
    assert normalize_padding({"top": 2}, base={"left": 1, "top": 1, "right": 1, "bottom": 3}) == {
        "left": 1.0,
        "top": 2.0,
        "right": 1.0,
        "bottom": 3.0,
    }
    assert normalize_padding({"bottom": 6}, base=6) == 6.0
    assert normalize_padding(None, base={"left": 2}) == {"left": 2.0, "top": 0.0, "right": 0.0, "bottom": 0.0}
    cases = [
        ((6, 6), "padding_pt は 0 以上の数値か、left・top・right・bottom を持つ辞書で指定してください"),
        ([1, 2, 3, 4], "padding_pt は 0 以上の数値か、left・top・right・bottom を持つ辞書で指定してください"),
        (True, "padding_pt は 0 以上の数値か、left・top・right・bottom を持つ辞書で指定してください"),
        (-1, "padding_pt の値は 0 以上の有限値にしてください"),
        (float("inf"), "padding_pt は 0 以上の数値か"),
        ({"left": float("nan")}, "padding_pt の値は 0 以上の有限値にしてください"),
        ({"lft": 2}, "padding_pt に未知の辺があります: lft（候補: left）"),
        ({"left": 1, "middle": 2}, "padding_pt に未知の辺があります: middle"),
    ]
    for value, message in cases:
        with pytest.raises(MeasureError, match=message.replace("(", r"\(")) as caught:
            normalize_padding(value)
        assert caught.value.code == "invalid_style" and isinstance(caught.value, InputError)
    assert padding_insets({"padding_pt": {"left": 3, "bottom": 4}}) == Insets(3, 0, 0, 4)
    assert padding_insets({}) == Insets(0, 0, 0, 0)


# ---- 文字ごとの書体 ----


def test_ambiguous_symbol_uses_japanese_font_when_latin_lacks_it(synthetic):
    runs = font_runs("A※B①", synthetic_style(), synthetic, bold=False)
    assert [(run.text, run.family, run.forced) for run in runs] == [
        ("A", "Test Latin", False),
        ("※", "Test Japanese", True),
        ("B", "Test Latin", False),
        ("①", "Test Japanese", True),
    ]
    assert layout_text(content("A※B"), synthetic_style(), synthetic).content_width == pytest.approx(40)


def test_ambiguous_symbol_in_latin_font_is_forced(synthetic):
    runs = font_runs("a→b“c”漢 x", synthetic_style(), synthetic, bold=False)
    assert [(run.text, run.family, run.forced) for run in runs] == [
        ("a", "Test Latin", False),
        ("→", "Test Latin", True),
        ("b", "Test Latin", False),
        ("“", "Test Latin", True),
        ("c", "Test Latin", False),
        ("”", "Test Latin", True),
        ("漢", "Test Japanese", False),
        (" x", "Test Latin", False),
    ]
    assert layout_text(content("→"), synthetic_style(), synthetic).content_width == pytest.approx(20)


EAST_ASIAN_FONT = "§¨°±´×÷‘’‚‛“”„⁇⁈⹀ꞵꭓ①⓿■□○●▲▼△▽◇◆◊◌◦☀☆★♪♫♯⛿✀✓➾㉈㉏"
LATIN_FONT = "→←↑↓…·′″¶†‡≠∞―—–Łé※‥⇒⇔≦≧∴αБ¡®²¼Æß─▟⟀"


def test_east_asian_font_characters_follow_powerpoint():
    """PowerPoint が lang="ja-JP" で和文の書体（a:ea）で描く、幅が W・F・H でない文字の一覧（Arial と Yu Gothic UI の組の
    2,878 文字で確かめた）。Unicode の幅や一般カテゴリーでは決めることができない（§ は和文で ¶ は欧文、± × ÷ は和文で ≠ は欧文、
    “ ” ‘ ’ は和文で ′ ″ は欧文）。※ ‥ ⇒ は欧文の側の文字（PowerPoint は欧文の書体にないので別の書体で描く）。"""
    assert all(measure.uses_east_asian_font(char) for char in EAST_ASIAN_FONT)
    assert not any(measure.uses_east_asian_font(char) for char in LATIN_FONT)
    # 幅が W・F・H の文字は今の規則（和文の書体）のままなので、範囲の中でも一覧の対象にしない（➿ ㈀ ㋿ は W）。
    assert not any(
        measure.uses_east_asian_font(char)
        for char in ("あ", "漢", "・", "〜", "Ａ", "ｱ", "➿", "㈀", "㋿", "A", "", "ab")
    )
    assert measure.EAST_ASIAN_FONT_RANGES == (
        (0x2460, 0x24FF),
        (0x25A0, 0x25FF),
        (0x2600, 0x26FF),
        (0x2700, 0x27BF),
        (0x3200, 0x32FF),
    )
    assert [measure.uses_east_asian_font(chr(code)) for code in (0x245F, 0x2500, 0x259F, 0x25FF, 0x27C0)] == [
        False,
        False,
        False,
        True,
        False,
    ]


def test_east_asian_font_characters_are_forced_to_japanese_font(synthetic):
    """一覧の文字は、和文の書体に字形があれば和文の書体で計測して forced（a:latin と a:ea の両方を和文の書体）にし、
    なければ欧文の書体で forced にする（PowerPoint の既定の見た目のまま、計測と表示を合わせる）。一覧の外の
    幅の曖昧な文字（→ …）は欧文の書体で forced のまま。幅と行の書体の組は計測した書体で決まる。"""
    fonts = FakeFonts(advance=0.5)
    runs = font_runs("a○×“b”→…", style(), fonts, bold=False)
    assert [(run.text, run.family, run.forced) for run in runs] == [
        ("a", "Latin", False),
        ("○×“", "Japanese", True),
        ("b", "Latin", False),
        ("”", "Japanese", True),
        ("→…", "Latin", True),
    ]
    lacking = FakeFonts(missing={"Japanese": "×"})
    assert [(run.text, run.family, run.forced) for run in font_runs("○×", style(), lacking, bold=False)] == [
        ("○", "Japanese", True),
        ("×", "Latin", True),
    ]
    neither = FakeFonts(missing={"Japanese": "×", "Latin": "×"})
    assert missing_glyphs("○×", style(), neither, bold=False) == ["×"]
    # 合成フォント：“ ” は和文の書体になく、欧文の書体で forced。① は和文の書体で forced（欧文の書体にもない）。
    synthetic_runs = [
        (run.text, run.family, run.forced) for run in font_runs("“①”", synthetic_style(), synthetic, bold=False)
    ]
    assert synthetic_runs == [("“", "Test Latin", True), ("①", "Test Japanese", True), ("”", "Test Latin", True)]


def test_lines_above_baseline_keep_ink_bottom(synthetic):
    """インクの上下は字形の範囲のままにする。基線より上だけにある字形（“ ”）の行は、インクの下が負（基線より上）になる。"""
    quotes = layout_text(content("“”"), synthetic_style(), synthetic)
    assert (quotes.lines[0].ink_top, quotes.lines[0].ink_bottom) == pytest.approx((14, -10))
    placed = layout_element(
        {**element(area=(10, 20, 200, 50)), "style": synthetic_style(), "content": content("“”")}, synthetic
    )
    baseline = 20 + LATIN_TOP * 20
    assert tuple(placed.ink) == pytest.approx((10, baseline - 14, 16, 4))
    mixed = layout_text(content("“g"), synthetic_style(), synthetic)
    assert (mixed.lines[0].ink_top, mixed.lines[0].ink_bottom) == pytest.approx((14, 4))


def test_missing_in_both_fonts_is_reported(synthetic):
    assert missing_glyphs("A🙂※\n✈🙂", synthetic_style(), synthetic, bold=False) == ["✈", "🙂"]
    message = "Test Latin と Test Japanese のどちらにも字形がない文字です: '🙂' (U+1F642)"
    for call in (
        lambda: layout_text(content("A🙂"), synthetic_style(), synthetic),
        lambda: font_runs("A🙂", synthetic_style(), synthetic, bold=False),
    ):
        with pytest.raises(MeasureError) as caught:
            call()
        assert str(caught.value) == message and caught.value.code == "font_glyph_missing"
        assert caught.value.details == {"characters": ["🙂"], "fonts": ["Test Latin", "Test Japanese"]}


# ---- 図形の文字領域と内容から決める寸法 ----


def test_text_area_follows_preset_shape_definitions():
    frame = Rect(0, 0, 200, 100)
    assert text_area(frame) == frame

    def inset(amount):
        return pytest.approx((amount, amount, 200 - 2 * amount, 100 - 2 * amount))

    assert tuple(text_area(frame, "rounded_rectangle")) == inset(0.29289 * 0.16667 * 100)
    assert tuple(text_area(frame, "rounded_rectangle", 10)) == inset(2.9289)  # 角の大きさの比 0.1
    assert tuple(text_area(frame, "rounded_rectangle", 80)) == inset(0.29289 * 50)  # 比は 0.5 まで
    shrink = (1 - math.sqrt(0.5)) / 2
    ellipse = (200 * shrink, 100 * shrink, 200 * math.sqrt(0.5), 100 * math.sqrt(0.5))
    assert tuple(text_area(frame, "ellipse")) == pytest.approx(ellipse)
    assert text_area(frame, "diamond") == Rect(50, 25, 100, 50)
    with pytest.raises(MeasureError, match="未対応の図形です: star") as caught:
        text_area(frame, "star")
    assert caught.value.code == "invalid_style"


@pytest.mark.parametrize(
    "shape, radius",
    [
        ("rectangle", None),
        ("rounded_rectangle", None),
        ("rounded_rectangle", 4),
        ("rounded_rectangle", 40),
        ("ellipse", None),
        ("diamond", None),
    ],
)
@pytest.mark.parametrize("size", [(120, 30), (30, 120)])
def test_frame_size_inverts_text_area(shape, radius, size):
    padding = Insets(3, 2, 5, 7)
    width, height = frame_size(shape, *size, padding, corner_radius_pt=radius)
    area = text_area(Rect(0, 0, width, height), shape, radius)
    assert (area.w - 8, area.h - 9) == pytest.approx(size)
    fixed_width, fixed_height = frame_size(shape, *size, padding, width=300, corner_radius_pt=radius)
    assert fixed_width == 300
    assert text_area(Rect(0, 0, 300, fixed_height), shape, radius).h - 9 == pytest.approx(size[1])


@pytest.mark.parametrize(
    "shape, radius",
    [("rectangle", None), ("rounded_rectangle", None), ("rounded_rectangle", 12), ("ellipse", None), ("diamond", None)],
)
def test_frames_from_measurement_hold_their_lines(shape, radius):
    fonts = FakeFonts()
    words = "あいうえおかきくけこさしすせそたちつてとなにぬねの"
    options = {"padding_pt": 6, "corner_radius_pt": radius}
    measured = layout_text(content(words), style(**options), fonts, width=150, shape=shape)
    assert len(measured.lines) > 1 and measured.height == pytest.approx(math.ceil(measured.height * 10) / 10)
    placed = layout_element(
        element(words, kind="box", shape=shape, area=(0, 0, 150, measured.height), **options), fonts
    )
    assert texts(placed.text) == texts(measured)
    assert placed.inner.h >= measured.content_height - 1e-9
    assert placed.inner.w == pytest.approx(measured.available_width)
    assert not [item for item in fit_findings(placed) if item["code"] != "text_wrap_near_limit"]
    # 幅も内容から決めた枠は、折り返す要素でも同じ行のまま置くことができ、誤差の範囲の行も出ない。
    natural = layout_text(content("あいう\nabc de"), style(**options), fonts, shape=shape)
    assert natural.available_width is None and natural.width >= natural.content_width + 0.25
    area = (0, 0, natural.width, natural.height)
    framed = layout_element(element("あいう\nabc de", kind="box", shape=shape, area=area, **options), fonts)
    assert texts(framed.text) == ["あいう", "abc de"]
    assert not fit_findings(framed)


# ---- 数式の置き方。math_mode を差し替え、その判定を計測に渡すことを確かめる ----


def test_equation_display_follows_math_mode(math_mode):
    fonts = FakeFonts()
    layout_element(element(("math", "x+1"), id="formula", kind="equation"), fonts)
    layout_element(element("値 ", ("math", "x+1"), id="label"), fonts)
    layout_element(element(("math", "x+2", {"size_pt": 30, "bold": True}), id="cell", kind="box"), fonts)
    assert math_mode == ["formula", "label", "cell"]
    assert [call["display"] for call in fonts.calls] == [True, False, False]
    assert fonts.calls[0] == {
        "expression": "x+1",
        "size": 20.0,
        "bold": False,
        "family": "Math",
        "display": True,
        "japanese": "Japanese",
        "latin": "Latin",
    }
    assert (fonts.calls[2]["size"], fonts.calls[2]["bold"]) == (30.0, True)
    assert layout_text(content(("math", "x+3")), style(), fonts, display=True) is not None
    assert fonts.calls[-1]["display"] is True


def test_text_without_math_does_not_read_math_mode(monkeypatch):
    def fail(element):
        raise AssertionError("数式のない要素で math_mode を呼んだ")

    module = types.ModuleType("slidekit.math.office_math")
    module.math_mode = fail
    monkeypatch.setitem(sys.modules, "slidekit.math.office_math", module)
    assert texts(layout_element(element("abc"), FakeFonts()).text) == ["abc"]


# ---- 断片の書式 ----


def test_fragment_size_and_bold_are_measured():
    fonts = FakeFonts(bold_advance=0.6)
    sized = layout_text(content("ab", ("text", "cd", {"size_pt": 30})), style(), fonts)
    assert sized.content_width == pytest.approx(20 + 30)
    assert sized.lines[0].baseline == pytest.approx(0.9 * 30)  # 行の上側は最も大きい文字で決まる
    assert sized.content_height == pytest.approx(0.9 * 30 + FAKE_BOTTOM * 30)
    assert layout_text(content(("emphasis", "ab")), style(), fonts).content_width == pytest.approx(24)
    assert layout_text(content(("emphasis", "ab", {"bold": False})), style(), fonts).content_width == pytest.approx(20)
    assert layout_text(content("ab"), style(bold=True), fonts).content_width == pytest.approx(24)
    assert layout_text(content(("text", "ab", {"bold": False})), style(bold=True), fonts).content_width == 20
    # 強調の数式（Math(emphasis=True)）は、指定がなければ太字で計測する（styles.fragment_style と同じ規則）。
    layout_text([{"kind": "math", "text": "y", "emphasis": True}], style(), fonts)
    layout_text([{"kind": "math", "text": "z", "emphasis": True, "style": {"bold": False}}], style(), fonts)
    assert [(call["expression"], call["bold"]) for call in fonts.calls] == [("y", True), ("z", False)]


# ---- 入力の誤り ----


def test_invalid_input_is_reported_with_codes():
    fonts = FakeFonts()
    with pytest.raises(MeasureError) as caught:
        layout_text(content("a\tb"), style(), fonts)
    assert str(caught.value) == "文章にタブなどの制御文字は使えません（改行は \\n だけを使えます）: U+0009"
    assert caught.value.code == "invalid_fragment"
    fragment_cases = [
        ([{"kind": "image", "text": "a"}], "未対応の文章の断片です: image"),
        ([{"kind": "text", "text": 3}], "文章の断片の text は文字列にしてください"),
        ("abc", "文章は断片"),
        (content(("text", "a", {"size_pt": -1})), "文章の断片の size_pt は正の有限値にしてください"),
        (content("a\r\nb"), "U\\+000D"),
    ]
    for value, message in fragment_cases:
        with pytest.raises(MeasureError, match=message) as caught:
            layout_text(value, style(), fonts)
        assert caught.value.code == "invalid_fragment"
    style_cases = [
        ({"size_pt": 0}, "size_pt は正の有限値にしてください"),
        ({"size_pt": None}, "size_pt は正の有限値にしてください"),
        ({"line_spacing": -1}, "line_spacing は正の有限値にしてください"),
        ({"line_spacing": None}, "line_spacing は正の有限値にしてください"),
        ({"word_wrap": "yes"}, "word_wrap は真偽値で指定してください"),
        ({"bold": 1}, "bold は真偽値で指定してください"),
        ({"align": "middle"}, "align は left、center、right、justify のいずれかにしてください"),
        ({"valign": "center"}, "valign は top、middle、bottom のいずれかにしてください"),
        ({"padding_pt": (1, 2)}, "padding_pt は 0 以上の数値か"),
        ({"font_latin": None}, "font_latin は書体名の文字列で指定してください"),
        ({"corner_radius_pt": -2}, "corner_radius_pt は 0 以上の有限値にしてください"),
    ]
    for overrides, message in style_cases:
        with pytest.raises(MeasureError, match=message) as caught:
            layout_text(content("a"), style(**overrides), fonts)
        assert caught.value.code == "invalid_style"
    with pytest.raises(MeasureError, match="幅は正の有限値にしてください"):
        layout_text(content("a"), style(), fonts, width=0)
    with pytest.raises(MeasureError, match="改行を含む文字列は行ごとに分けて渡してください"):
        font_runs("a\nb", style(), fonts, bold=False)


def test_shared_rules_for_control_characters_and_fragment_fonts(monkeypatch):
    """文章で受け付けない文字は fragments.control_character、断片の文字サイズと太字は styles.fragment_style で決める
    （作成 API・書き出し・検査と同じ規則を 1 か所に置く）。"""
    assert measure.control_character is fragments.control_character
    assert measure.fragment_style is styles.fragment_style
    fonts = FakeFonts()
    for char in ("\t", "\r", "\x00", "\x85", "\u2028", "\u2029", "\ud800"):
        assert fragments.control_character(f"a{char}b") == char
        with pytest.raises(MeasureError, match=f"U\\+{ord(char):04X}") as caught:
            layout_text(content(f"a{char}b"), style(), fonts)
        assert caught.value.code == "invalid_fragment"
    for char in ("\n", "\u3000", "\u200b", "\xa0"):
        assert fragments.control_character(f"a{char}b") is None
    calls = []

    def spy(element_style, fragment):
        calls.append((dict(element_style), fragment))
        return styles.fragment_style(element_style, fragment)

    monkeypatch.setattr(measure, "fragment_style", spy)
    parts = [
        {"kind": "emphasis", "text": "ab"},
        {"kind": "text", "text": "c", "style": {"size_pt": 30, "bold": False}},
        {"kind": "math", "text": "y", "emphasis": True},
    ]
    layout_text(parts, style(size_pt=16), FakeFonts(bold_advance=0.6))
    assert [fragment for _style, fragment in calls if fragment is not parts[0]] == parts[1:]
    assert all(element_style == {"size_pt": 16.0, "bold": False} for element_style, _fragment in calls)


def test_math_errors_keep_expression_and_cause():
    class Broken(FakeFonts):
        def math_box(self, expr, size, bold=False, **kwargs):
            if expr == "bad":
                raise ValueError("閉じていない括弧があります")
            return super().math_box(expr, size, bold, **kwargs)

    with pytest.raises(MeasureError) as caught:
        layout_text(content(("math", "bad")), style(), Broken())
    error = caught.value
    assert error.code == "text_measurement_failed" and error.details["expression"] == "bad"
    assert isinstance(error.__cause__, ValueError) and "閉じていない括弧" in str(error)
    with pytest.raises(MeasureError) as caught:
        layout_text(content(("math", r"x+\unknown")), style(), FakeFonts())
    assert caught.value.details["expression"] == r"x+\unknown"


def test_measure_error_is_input_error_with_diagnostic_codes():
    error = MeasureError("値が不正です", code="invalid_style", setting="size_pt")
    assert isinstance(error, InputError) and isinstance(error, ValueError)
    assert (error.code, error.details) == ("invalid_style", {"setting": "size_pt"})
    assert MeasureError("計測できません").code == "text_measurement_failed"
    assert MeasureError.CODES == {"font_glyph_missing", "invalid_fragment", "invalid_style", "text_measurement_failed"}
    with pytest.raises(ValueError, match="glyph_missing"):
        MeasureError("字形がありません", code="glyph_missing")


# ---- 場面の計測・判定・記録 ----


def test_layout_scene_measures_text_elements_and_keeps_errors(math_mode):
    class Unreadable(FakeFonts):
        def face_metrics(self, family, bold=False):
            if family == "Broken":
                raise OSError("フォントを読めません")
            return super().face_metrics(family, bold)

    fonts = Unreadable(missing={"🙂"})
    scene = {
        "slides": [
            {
                "id": "first",
                "elements": [
                    element("abc", id="title"),
                    {"id": "frame", "kind": "box", "area": [0, 0, 10, 10], "style": style(), "content": []},
                    {"id": "rule", "kind": "line", "area": [0, 0, 10, 0], "style": {}, "content": []},
                    element("a🙂", id="emoji"),
                    element("abc", id="font", font_latin="Broken"),
                ],
            },
            {
                "id": "second",
                "elements": [
                    element(("math", r"\unknown"), id="formula", kind="equation"),
                    element("  ", id="blank"),
                    element(("math", "x"), id="ok", kind="equation"),
                ],
            },
        ]
    }
    result = layout_scene(scene, fonts)
    assert set(result) == {
        ("first", "title"),
        ("first", "emoji"),
        ("first", "font"),
        ("second", "formula"),
        ("second", "ok"),
    }
    assert isinstance(result[("first", "title")], ElementLayout) and result[("first", "title")].element_id == "title"
    emoji = result[("first", "emoji")]
    assert isinstance(emoji, MeasureError) and emoji.code == "font_glyph_missing"
    assert (emoji.slide_id, emoji.element_id, emoji.details["characters"]) == ("first", "emoji", ["🙂"])
    broken = result[("first", "font")]
    assert broken.code == "text_measurement_failed" and isinstance(broken.__cause__, OSError)
    assert (broken.slide_id, broken.element_id, broken.details["reason"]) == ("first", "font", "フォントを読めません")
    formula = result[("second", "formula")]
    assert (formula.code, formula.slide_id, formula.details["expression"]) == (
        "text_measurement_failed",
        "second",
        r"\unknown",
    )
    assert [call["display"] for call in fonts.calls if call["expression"] == "x"] == [True]


def test_fit_findings_use_ink_and_frame_tolerance():
    fonts = FakeFonts()
    # 1 行目の基線は 18 pt。インクは基線の 14 pt 上から 2 pt 下まで（上端 4 pt、下端 20 pt）。
    hidden = layout_element(element("abc", area=(0, 0, 200, 19)), fonts)
    assert tuple(hidden.ink) == pytest.approx((0, 4, 30, 16)) and not hidden.visible_frame
    assert fit_findings(hidden) == []  # 塗りも線もない枠は 1.5 pt まで許す
    visible = layout_element(element("abc", area=(0, 0, 200, 19), fill="FFFFFF"), fonts)
    findings = fit_findings(visible)
    assert [item["code"] for item in findings] == ["text_overflow"] and findings[0]["severity"] == "error"
    observed = findings[0]["observations"]
    assert observed["overflow_pt"] == {"side": "bottom", "amount": 1.0}
    assert observed["ink_pt"] == [0, 4, 30, 20] and observed["text_area_pt"] == [0, 0, 200, 19]
    assert (observed["required_height_pt"], observed["available_height_pt"]) == (round(18 + FAKE_BOTTOM * 20, 2), 19)
    assert "unbreakable_width_pt" not in observed and "Rect(x, y, w, None)" in observed["suggestion"]
    assert layout_element(element("abc", area=(0, 0, 200, 19), line_color="000000"), fonts).visible_frame
    assert not layout_element(
        element("abc", area=(0, 0, 200, 19), line_color="000000", line_width_pt=0), fonts
    ).visible_frame
    padded = fit_findings(layout_element(element("abc", area=(0, 0, 200, 26), padding_pt=4), fonts))
    assert [(item["code"], item["severity"]) for item in padded] == [("text_padding_overlap", "warning")]
    assert padded[0]["observations"]["overflow_pt"] == {"side": "bottom", "amount": 2.0}
    slight = fit_findings(layout_element(element("abc", area=(0, 0, 200, 25), padding_pt=8), fonts))
    assert [item["code"] for item in slight] == ["text_overflow"]
    near = fit_findings(layout_element(element("aaaa bbbbb", area=(0, 0, 100, 100)), fonts))
    assert [(item["code"], item["severity"]) for item in near] == [("text_wrap_near_limit", "info")]
    assert (near[0]["observations"]["near_limit_lines"], near[0]["observations"]["band_pt"]) == ([1], 0.65)
    assert fit_findings(layout_element(element("   "), fonts)) == []


def test_unbreakable_part_wider_than_inner_is_overflow_without_ink():
    """内側より長い語は、計測も PowerPoint と同じく語の途中で分けるのでインクは収まるが、表示の不具合なので
    text_overflow（エラー）として報告し、分ける位置のない部分の幅と、語の途中で分けた行を示す。"""
    fonts = FakeFonts()
    placed = layout_element(element("a verylongword b", area=(0, 0, 60, 100), padding_pt=0), fonts)
    assert placed.ink.right <= 60 and placed.ink.bottom <= 100  # インクは文字領域に収まる
    findings = fit_findings(placed)
    # 内側の幅いっぱいまで詰めた行は、誤差の範囲（参考）にもなる。
    assert [item["code"] for item in findings] == ["text_overflow", "text_wrap_near_limit"]
    finding = findings[0]
    assert finding["severity"] == "error"
    assert "PowerPoint は語の途中で改行する" in finding["message"]
    observed = finding["observations"]
    assert (observed["unbreakable_width_pt"], observed["split_lines"], observed["lines"]) == (120, [2], 4)
    assert "overflow_pt" not in observed and observed["available_width_pt"] == 60
    # 内側より広い数式は途中で分けない。インクのはみ出しと、分ける位置のない部分の幅を示す。
    [wide] = fit_findings(layout_element(element(("math", "x+y+z+w+v+u+t"), area=(0, 0, 60, 100)), fonts))
    assert wide["code"] == "text_overflow" and "split_lines" not in wide["observations"]
    assert wide["observations"]["unbreakable_width_pt"] == 130
    assert wide["observations"]["overflow_pt"] == {"side": "right", "amount": 70}
    assert "文字領域からはみ出している" in wide["message"] and "数式など" in wide["message"]
    # 折り返さない要素は、分ける位置のない部分の幅では判定しない（インクだけで判定する）。
    assert fit_findings(layout_element(element("verylongword", area=(0, 0, 200, 100), word_wrap=False), fonts)) == []


def test_layout_record_has_absolute_positions():
    record = layout_record(layout_element(element("abc\nde", padding_pt={"left": 4}), FakeFonts()))
    assert json.loads(json.dumps(record)) == record
    assert record["id"] == "body" and record["word_wrap"] is True
    assert (record["frame_pt"], record["text_area_pt"], record["inner_pt"]) == (
        [100, 50, 200, 100],
        [100, 50, 200, 100],
        [104, 50, 196, 100],
    )
    assert record["block_top_pt"] == 50
    assert record["content_size_pt"] == [30, round(42 + FAKE_BOTTOM * 20, 2)]
    assert record["required_size_pt"] == [34, round(42 + FAKE_BOTTOM * 20, 2)]
    assert record["ink_pt"] == [104, 54, 134, 94] and record["unbreakable_width_pt"] == 30
    assert record["lines"] == [
        {
            "text": "abc",
            "width_pt": 30,
            "baseline_pt": 68,
            "left_pt": 104,
            "ink_top_pt": 14,
            "ink_bottom_pt": 2,
            "hard_break": True,
            "near_limit": False,
        },
        {
            "text": "de",
            "width_pt": 20,
            "baseline_pt": 92,
            "left_pt": 104,
            "ink_top_pt": 14,
            "ink_bottom_pt": 2,
            "hard_break": False,
            "near_limit": False,
        },
    ]
    assert MEASURE_VERSION == "7"


def test_results_are_kept_per_font_source():
    fonts = FakeFonts()
    first = layout_text(content("ab ", ("math", "x")), style(), fonts, width=100)
    assert layout_text(content("ab ", ("math", "x")), style(), fonts, width=100) is first
    assert len(fonts.calls) == 1  # 保持した結果を使い、数式を計測し直さない
    other = layout_text(content("ab ", ("math", "x")), style(), FakeFonts(), width=100)
    assert other is not first and other == first
    assert layout_text(content("ab ", ("math", "x")), style(size_pt=30), fonts, width=100) != first


# ---- 試験用の偽物（試験が前提にする値） ----


def test_fake_fonts_follow_documented_values():
    fonts = FakeFonts(missing={"Latin": {"※"}})
    metrics = fonts.face_metrics("Latin")
    assert fonts.face_metrics("Japanese", True) == type(metrics)("Japanese", True, 0.9, 0.3, False)
    assert (metrics.ascent, metrics.descent, metrics.has_kerning) == (0.9, 0.3, False)
    assert fonts.advance("あ", "Japanese") == fonts.advance("a", "Latin", True) == 0.5
    assert fonts.kerning("T", "A", "Latin") == 0.0
    assert fonts.glyph_bounds("a", "Latin") == (-0.1, 0.7) and fonts.glyph_bounds(" ", "Latin") == (0.0, 0.0)
    assert not fonts.has_glyph("※", "Latin") and fonts.has_glyph("※", "Japanese")
    with pytest.raises(MeasureError) as caught:
        fonts.advance("※", "Latin")
    assert caught.value.code == "font_glyph_missing"
    assert fonts.math_box("abc", 20) == MathBox(30, 14, 2, 14, 2)
    with pytest.raises(MeasureError, match=r"\\unknown"):
        fonts.math_box(r"\unknown", 20)
    assert [call["expression"] for call in fonts.calls] == ["abc", r"\unknown"]
    everywhere = FakeFonts(missing="※", kerning={("T", "A"): -0.1})
    assert not everywhere.has_glyph("※", "Japanese") and everywhere.face_metrics("Latin").has_kerning
    assert everywhere.kerning("T", "A", "Latin") == -0.1
    # 互換用の計測の関数（Project.measure_text・measure_math などが呼ぶ）
    assert fonts.measure_text("abc", 20) == pytest.approx((30, 18 + FAKE_BOTTOM * 20))
    assert fonts.measure_text("ab\nabcd", 20)[0] == pytest.approx(40)
    assert fonts.measure_math("abc", 20) == (30, 16) and fonts.measure_math_metrics("ab", 10) == (10, 8, 1)
    assert fonts.missing_glyphs("a※", latin="Latin", japanese="Latin") == ["※"]
    assert fonts.inventory()["measure_version"] == MEASURE_VERSION
    assert FakeFonts(bold_advance=0.6).advance("a", "Latin", True) == 0.6
