"""SLIDEKIT_TEST_POWERPOINT=1 を指定したときだけ、数式の計測を実際の PowerPoint の変換と比べる。

照合データ（tests/data/powerpoint_math.json）の式から 15 個を選び、照合データの道具（tests/tools/powerpoint_math_data.py）と
同じ資料と計測の方法で、PowerPoint の PDF を作り直す。計測との差が照合の基準（metrics.REFERENCE_LIMITS）に入ることを確かめる。
PowerPoint の版を変えたときと、数式の書き出しを変えたときの確認に使う。資料はフィクスチャ windows_fonts_environment で
Windows の既定フォントを使う。直列に実行する（`-p no:xdist`）。
\\text の文字（書き出しの a:rPr/@sz）が、同じ段の数式の文字と同じ大きさで描かれることも確かめる。
"""

import importlib.util
import json
import os
from pathlib import Path

import pytest

pytestmark = [
    pytest.mark.powerpoint,
    pytest.mark.skipif(
        os.environ.get("SLIDEKIT_TEST_POWERPOINT") != "1", reason="実機 PowerPoint を使う任意の検査である。"
    ),
]

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data" / "powerpoint_math.json"
# 計測の規則と書き出しの直しを確かめる式：添字の中の文章の大きさ、mod の前後の空白、大型演算子の引数と添字、: の空き、
# 根号の上の空き、伸びる括弧の大きさ、分数の斜線、文章の描画範囲、添字の土台の最後の原子、太字の数字の間のカンマ、
# 根号の後と関数の適用の後の句読点、横書きの分数の前の空き。
SELECTED = (
    r"T_{\text{max}}",
    r"a \bmod b",
    r"\int_0^1 f(x)\,dx",
    r"\sum_{i=1}^{n} a_i",
    "a:b",
    r"\sqrt{n}",
    r"\binom{n}{k}",
    "a/b",
    r"x\text{ 個}",
    r"\left(\frac{a}{b}\right)",
    r"(\frac{a}{b})^2",
    r"\mathbf{(0,0)}",
    r"\sqrt{x},y",
    r"\log x,y",
    "x a/b",
)


# \text の文字 b と、同じ段の数式の文字 2 を並べた式：添字、分数、入れ子の分数、根号の次数。
# 独立した数式の入れ子の分数は、PowerPoint が分子と分母を縮めないので、書き出しも \text を縮めない。
TEXT_SIZES = (r"x_{2\text{b}}", r"\frac{2\text{b}}{3}", r"\frac{1}{\frac{2\text{b}}{3}}", r"\sqrt[2\text{b}]{x}")


def _tool():
    spec = importlib.util.spec_from_file_location("powerpoint_math_data", ROOT / "tools" / "powerpoint_math_data.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_real_powerpoint_math_metrics(tmp_path, windows_fonts_environment):
    from slidekit.fonts import FontRegistry
    from slidekit.math import metrics

    tool = _tool()
    reference = {item["expression"] for item in json.loads(DATA.read_text(encoding="utf-8"))["items"]}
    assert set(SELECTED) <= reference and len(SELECTED) == 15
    pages = tool.layout_items(list(SELECTED))
    tool.write_deck(tmp_path, pages)
    pdf, _doctor, project = tool.render(tmp_path)
    items, failed = tool.measure_pdf(pdf, pages, math_font=project.theme["fonts"]["math"])
    assert failed == []
    with FontRegistry(project) as fonts:
        result = metrics.compare_reference(
            {"size_pt": tool.SIZE_PT, "items": items}, fonts._math_measurer(), text_fonts=fonts._text_fonts()
        )
    rows = [
        (row["mode"], row["expression"], round(row["width_em"], 4), round(row["vertical_em"], 4))
        for row in result["items"]
    ]
    print(
        "\n".join(
            f"{mode:7} 幅 {width:+.4f} em  上下 {vertical:.4f} em  {expression}"
            for mode, expression, width, vertical in rows
        )
    )
    assert result["count"] == 2 * len(SELECTED) and not result["skipped"]
    assert result["passed"], rows


def _sizes(page, area, math_key, tool):
    """枠の中の \\text の文字 b と数式の文字 2 の、PDF の文字の大きさ（pt）。文字は照合データの道具と同じ条件で選ぶ。"""
    x, y, width, height = area
    text, digit = [], []
    for char in tool._chars(page, (x - 2, y - 4, x + width + 2, y + height + 2)):
        if char["c"] == "b" and tool._font_key(char["font"]) != math_key:
            text.append(char["size"])
        elif char["c"] == "2" and tool._font_key(char["font"]) == math_key:
            digit.append(char["size"])
    return text, digit


def test_real_powerpoint_text_size_follows_math_glyphs(tmp_path, windows_fonts_environment):
    from pdf_tools import open_pdf

    tool = _tool()
    pages = tool.layout_items(list(TEXT_SIZES))
    tool.write_deck(tmp_path, pages)
    pdf, _doctor, project = tool.render(tmp_path)
    math_key = tool._font_key(project.theme["fonts"]["math"])
    rows = []
    with open_pdf(pdf) as document:
        for page, layout in zip(document, pages):
            for item in layout["items"]:
                text, digit = _sizes(page, item["area"], math_key, tool)
                rows.append((item["mode"], item["expression"], text, digit))
    print("\n".join(f"{mode:7} \\text {text}  数式 {digit}  {expression}" for mode, expression, text, digit in rows))
    assert len(rows) == 2 * len(TEXT_SIZES)
    for mode, expression, text, digit in rows:
        assert len(text) == 1 and len(digit) == 1, (mode, expression, text, digit)
        # 段の倍率の差（73 % と 60 %、20 pt で 2.6 pt）より十分に小さい差で一致する
        assert text[0] == pytest.approx(digit[0], abs=0.2), (mode, expression, text, digit)
