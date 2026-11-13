"""文字の配置の計測（measure.layout_element）を、校正用の資料を PowerPoint で変換した照合データと比べる。

照合データは tests/data/layout/powerpoint_layout.json（環境変数 SLIDEKIT_LAYOUT_FIXTURE で替えることができる）。
照合データがないとき、照合データと手元の書体のファイル（sha256）が違うときは省略する。作り方は tools/make_layout_fixture.py にある。

許容幅：行数は一致。行の幅と左端は、文字が 0.0035 em × 文字数 + 0.3 pt（PowerPoint の行の幅は計測の 0.998〜1.005 倍）、
数式を含む行はさらに数式 1 つあたり、数式の照合の基準の最大の幅の差（`slidekit.math.metrics.REFERENCE_LIMITS` の
width_max_em。1.5 pt を下限とする）を加える。基線は 0.5 pt、数式を含む行の基線は基準の上下の差（vertical_em。下限 1.5 pt）。
インクの横の端は、その端を決める行の幅の許容幅と同じにする。インクの上下は 0.6 pt、数式を含む要素は数式の基準の上下の差。
照合ケースごとに厳しく比べる項目（基線・行の幅と左端・インクの左右・インクの上下）と許容差は、タグか要素で選ぶ規則として
tests/data/layout/tolerances.toml に根拠とともに書く（例：タグ text_baseline の通常文字を目印にした数式の行は、
数式の構造によらず基線を 0.5 pt で比べる）。規則で許容差を決めた項目の違いの文言には、規則の名前を添える。
分数のように数式の字形が複数の基線に並ぶ行は、基線を比べない。数式の行の範囲には、PDF の読み取りが数式の線
（根号の記号など）を描画の範囲で含める。そこで、行の端の数式が線で描く記号（伸びる括弧・根号の記号）で始まるか終わるときは、
予測の行の範囲のその端を、数式の描画範囲の端（slidekit.math.metrics の Box.ink）にする（字形の端は送り幅のまま）。
PDF の読み取りは空白の字形を除くので、予測の行の範囲とインクの横の範囲からも、行頭と行末の空白の送り幅を除く。
書体は通常文字ごとに比べる（数式の \\text の文字を除く）。行数が違う要素は、行の位置とインクを比べない（書体は比べる）。
次の要素は違いを許す（書体は比べる）。
- 計測が誤差の範囲（text_wrap_near_limit）と示した行がある要素で、行数の差が 1 のとき：行数の違いを許す。
- タグ overflow の要素（枠より長い語や数式）で、計測が text_overflow を報告するとき：行数・行の位置・インクの違いを許す
  （資料の誤りとして報告されるため）。

この試験ファイルは、照合データを作る道具（tools/make_layout_fixture.py）と、数式の照合データを作る道具
（tests/tools/powerpoint_math_data.py）の PowerPoint を使わない部分も、合成した PDF で確かめる。
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
import functools
import importlib.util
import json
import math
import os
from pathlib import Path
import re
import shutil
import sys
import tomllib
import types

from pdf_tools import Document, text_length
import pytest

from fake_fonts import FakeFonts
from slidekit import measure
from slidekit.config import load_project
from slidekit.files import file_hash
from slidekit.fonts import FontRegistry, default_font
from slidekit.math import syntax
from slidekit.math.metrics import REFERENCE_LIMITS
from slidekit.pdf import Box, open_pdf
from slidekit.pdf_layout import _font_name, _text_in_math, read_pdf
from slidekit.pipeline import construct
from slidekit.styles import fragment_style

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "tests" / "data" / "layout"
DECK = DATA / "deck"
DEFAULT_FIXTURE = DATA / "powerpoint_layout.json"
TOLERANCES = DATA / "tolerances.toml"
FIXTURE_VARIABLE = "SLIDEKIT_LAYOUT_FIXTURE"
TOOL = ROOT / "tools" / "make_layout_fixture.py"
MATH_TOOL = ROOT / "tests" / "tools" / "powerpoint_math_data.py"

# 校正用の資料のページ（作る順）。折り返しを確かめるページは test_powerpoint_layout.py が実機で変換する。
PAGES = (
    "widths",
    "pitch",
    "wrap_ja",
    "wrap_latin",
    "wrap_mixed",
    "wrap_math",
    "explicit_wrap",
    "shapes",
    "math_lines",
    "math_stack",
    "math_stack_more",
    "math_stack_spacing",
    "math_limits",
    "math_text_lines",
    "math_text_spacing",
    "math_script_lines",
    "math_script_spacing",
    "math_script_multiline",
    "math_nested",
    "align",
    "fonts",
    "tables",
    "same_text",
)
WRAP_PAGES = ("wrap_ja", "wrap_latin", "wrap_mixed", "wrap_math", "explicit_wrap")

# 許容幅（pt。文字の幅は em あたり）。
TEXT_WIDTH_EM_PER_CHAR = 0.0035
TEXT_WIDTH_PT = 0.3
BASELINE_PT = 0.5
INK_PT = 0.6
MATH_PT = 1.5  # 数式を含む行とインクの許容幅の下限
MATH_WIDTH_EM = REFERENCE_LIMITS["width_max_em"]  # 数式 1 つあたりの幅の許容幅（数式の照合の基準と同じ）
MATH_VERTICAL_EM = REFERENCE_LIMITS["vertical_em"]  # 数式を含む行の基線とインクの上下の許容幅（同）
CHECKS = ("lines", "baselines", "extents", "ink", "fonts")
# 規則で許容差を決めることができる項目（基線、行の幅と左端、インクの左右の端、インクの上下の端）。
RULE_ITEMS = ("baseline", "width", "ink_x", "ink_y")
# 数式の字形が複数の基線に並ぶ構造（PDF から読む行の基線が、式の基線にならない）。
STACKED = re.compile(
    r"\\(?:[dtc]?frac|[dt]?binom|genfrac|over|atop|choose|overset|underset|stackrel|substack)(?![A-Za-z])"
)
_MATH_IN_LINE = re.compile(r"\$[^$]*\$")
_EDGES = ("左端", "上端", "右端", "下端")


def load_tool(path: Path = TOOL):
    """tools/ と tests/tools/ の道具をモジュールとして読む（同じプロセスでは 1 回だけ）。"""
    name = f"_slidekit_tool_{path.stem}"
    if name not in sys.modules:
        spec = importlib.util.spec_from_file_location(name, path)
        module = importlib.util.module_from_spec(spec)
        sys.modules[name] = module
        spec.loader.exec_module(module)
    return sys.modules[name]


# ---- 照合ケースごとの許容差の規則（tests/data/layout/tolerances.toml） ----


@dataclass(frozen=True)
class Tolerance:
    """規則が決める 1 項目の許容差。tolerance_pt と tolerance_em × 文字サイズの大きいほうにする。"""

    rule: str  # 規則の名前（違いの文言に添える）
    tolerance_pt: float = 0.0
    tolerance_em: float = 0.0
    stacked: bool = False  # baseline だけ：数式の字形が複数の基線に並ぶ行も比べる

    def value(self, size: float) -> float:
        return max(self.tolerance_pt, self.tolerance_em * size)


@dataclass(frozen=True)
class Rule:
    """タグか要素で選ぶ照合ケースの、項目ごとの許容差と根拠。"""

    tags: tuple[str, ...]
    elements: tuple[str, ...]
    items: Mapping[str, Tolerance]
    reason: str

    @property
    def name(self) -> str:
        return " ".join((*self.tags, *self.elements))

    def matches(self, case: Mapping) -> bool:
        tags = set(case.get("tags") or ())
        element = f"{case.get('slide_id')}/{case.get('element_id')}"
        return set(self.tags) <= tags and (not self.elements or element in self.elements)


def _positive(value, label: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value <= 0:
        raise ValueError(f"{label} は正の有限の数にする: {value!r}")
    return float(value)


def _strings(value, label: str) -> tuple[str, ...]:
    if not isinstance(value, list) or not all(isinstance(item, str) and item.strip() for item in value):
        raise ValueError(f"{label} は空でない文字列の配列にする: {value!r}")
    return tuple(value)


def parse_rules(data: Mapping, source: str = "許容差の規則") -> tuple[Rule, ...]:
    """許容差の規則（tolerances.toml の内容）を読む。形が違えば ValueError。"""
    if not isinstance(data, Mapping) or data.get("schema_version") != 1:
        raise ValueError(f"{source}: schema_version = 1 にする")
    unknown = set(data) - {"schema_version", "rules"}
    if unknown or not isinstance(data.get("rules", []), list):
        raise ValueError(f"{source}: 未対応の設定 {sorted(unknown)}（rules は表の配列にする）")
    rules = []
    for number, entry in enumerate(data.get("rules", []), 1):
        label = f"{source} の rules[{number}]"
        if not isinstance(entry, Mapping):
            raise ValueError(f"{label} は表にする")
        unknown = set(entry) - {"tags", "elements", "reason", *RULE_ITEMS}
        if unknown:
            raise ValueError(f"{label}: 未対応の項目 {sorted(unknown)}（項目は {', '.join(RULE_ITEMS)}）")
        tags = _strings(entry.get("tags", []), f"{label} の tags")
        elements = _strings(entry.get("elements", []), f"{label} の elements")
        if not tags and not elements:
            raise ValueError(f"{label}: ケースを tags か elements で選ぶ")
        if any("/" not in element for element in elements):
            raise ValueError(f"{label}: elements は「ページの識別子/要素の識別子」にする: {list(elements)}")
        reason = entry.get("reason")
        if not isinstance(reason, str) or not reason.strip():
            raise ValueError(f"{label}: reason に許容差を決めた根拠を書く")
        name = " ".join((*tags, *elements))
        items = {}
        for item in RULE_ITEMS:
            if item not in entry:
                continue
            spec = entry[item]
            keys = {"tolerance_pt", "tolerance_em", *(("stacked",) if item == "baseline" else ())}
            if not isinstance(spec, Mapping) or set(spec) - keys or not {"tolerance_pt", "tolerance_em"} & set(spec):
                raise ValueError(
                    f"{label} の {item}: tolerance_pt か tolerance_em を持つ表にする（使える鍵 {sorted(keys)}）"
                )
            stacked = spec.get("stacked", False)
            if not isinstance(stacked, bool):
                raise ValueError(f"{label} の {item}.stacked は真偽値にする: {stacked!r}")
            items[item] = Tolerance(
                rule=name,
                stacked=stacked,
                **{
                    key: _positive(spec[key], f"{label} の {item}.{key}")
                    for key in ("tolerance_pt", "tolerance_em")
                    if key in spec
                },
            )
        if not items:
            raise ValueError(f"{label}: 項目（{', '.join(RULE_ITEMS)}）を 1 つ以上書く")
        rules.append(Rule(tags=tags, elements=elements, items=items, reason=reason.strip()))
    return tuple(rules)


def load_rules(path: Path = TOLERANCES) -> tuple[Rule, ...]:
    """許容差の規則のファイルを読む。ファイルがなければ規則なし。"""
    path = Path(path)
    if not path.is_file():
        return ()
    return _load_rules(path, path.stat().st_mtime_ns)


@functools.lru_cache(maxsize=8)
def _load_rules(path: Path, _mtime: int) -> tuple[Rule, ...]:
    with path.open("rb") as stream:
        return parse_rules(tomllib.load(stream), path.name)


def case_tolerances(case: Mapping, rules) -> dict[str, Tolerance]:
    """照合ケースに合う規則の、項目ごとの許容差。複数の規則が同じ項目を決めると ValueError。"""
    result: dict[str, Tolerance] = {}
    for rule in rules:
        if not rule.matches(case):
            continue
        for item, tolerance in rule.items.items():
            if item in result:
                raise ValueError(f"許容差の規則が重なる: {item} を「{result[item].rule}」と「{rule.name}」が決める")
            result[item] = tolerance
    return result


def _allowed(tolerance: float, rule: Tolerance | None) -> str:
    """違いの文言の許容差（規則で決めたときは規則の名前を添える）。"""
    return f"許容 {tolerance:.2f}" + (f"、規則 {rule.rule}" if rule is not None else "")


# ---- 比べる処理（test_powerpoint_layout.py も使う） ----


def _size(element) -> float:
    """要素の最大の文字サイズ（断片の指定を含む）。"""
    style = element.get("style") or {}
    sizes = [float(style.get("size_pt") or 20.0)]
    for fragment in element.get("content") or []:
        own = fragment.get("style") or {}
        if own.get("size_pt"):
            sizes.append(float(own["size_pt"]))
    return max(sizes)


def _line_units(element, layout: measure.ElementLayout) -> list[list[tuple[str, dict]]]:
    """計測の行ごとの単位（文字 1 つか数式 1 つ。数式は $…$）と、その断片。

    計測の行の文字列は、断片の文字（明示改行を除く）と数式を順に分けたものなので、先頭から行の長さずつ取る。
    """
    units = []
    for fragment in element.get("content") or []:
        text = str(fragment.get("text", ""))
        if fragment.get("kind") == "math":
            units.append((f"${text}$", fragment))
        else:
            units += [(char, fragment) for char in text if char != "\n"]
    result, position = [], 0
    for line in layout.text.lines:
        start, length = position, 0
        while length < len(line.text) and position < len(units):
            length += len(units[position][0])
            position += 1
        if "".join(text for text, _fragment in units[start:position]) != line.text:
            raise ValueError(f"計測の行を断片に対応付けられない: {line.text!r}")
        result.append(units[start:position])
    return result


def _width(units, style, fonts) -> float:
    """単位の列を 1 行として計測した幅（measure.layout_text。kern を加え、行末の空白を除く）。"""
    content, sources = [], []
    for text, fragment in units:
        if fragment.get("kind") != "math" and sources and sources[-1] is fragment:
            content[-1]["text"] += text
        else:
            content.append(dict(fragment) if fragment.get("kind") == "math" else {**fragment, "text": text})
            sources.append(fragment)
    return measure.layout_text(content, style, fonts).content_width


def _edge_spaces(units, style, fonts) -> tuple[float, float]:
    """行頭の空白の送り幅（最初の字形の原点まで）と、行の幅に含まれる行末の空白の送り幅（pt）。

    PDF の読み取りは空白の字形を除くので、実際の行の範囲は最初と最後の空白でない字形で決まる。計測は行末の U+0020 と
    U+3000 を行の幅から除き、行頭の空白と、そのほかの行末の空白（U+00A0 など）を幅に含める。その送り幅を計測と同じ規則で求める。
    """
    texts = [text for text, _fragment in units]
    first = 0
    while first < len(texts) and texts[first].isspace():
        first += 1
    end = len(texts)
    while end > first and texts[end - 1] in (" ", "\u3000"):
        end -= 1
    last = end
    while last > first and texts[last - 1].isspace():
        last -= 1
    if first == len(texts) or (first == 0 and last == end):
        return 0.0, 0.0
    full = _width(units[:end], style, fonts)
    lead = full - _width(units[first:end], style, fonts) if first else 0.0
    trail = full - _width(units[:last], style, fonts) if last < end else 0.0
    return lead, trail


def _drawn_edges(expression: str) -> tuple[bool, bool]:
    """数式の左端・右端が、PowerPoint が線で描く記号（伸びる括弧、根号の記号）か。

    左端は添字の付く本体の左端を見る。分数線と根号の上の線は数式の送り幅の範囲に描かれるので、送り幅で比べてよい。
    """
    try:
        nodes = [node for node in syntax.parse(expression) if not isinstance(node, syntax.Space)]
    except ValueError:
        return False, False
    if not nodes:
        return False, False
    first, last = nodes[0], nodes[-1]
    while isinstance(first, syntax.Scripts) and first.base:
        first = first.base[0]
    # 根号の次数（\sqrt[3]{x} の 3）は字形なので、次数のない根号だけを線の端とする。
    left = (isinstance(first, syntax.Radical) and not first.degree) or (
        isinstance(first, syntax.Fenced) and first.stretchy and bool(first.open)
    )
    right = isinstance(last, syntax.Fenced) and last.stretchy and bool(last.close)
    return left, right


def _math_ink(fonts, fragment, style, display: bool) -> tuple[float, float, float] | None:
    """数式の断片の（送り幅、描画の左、描画の右）（pt。数式の原点から）。数式の計測器を持たない FontSource（偽のフォント）
    では None（行の範囲を送り幅のままにする）。計測器は FontRegistry の数式の書体ごとの計測器を使う。"""
    measurer = getattr(fonts, "_math_measurer", None)
    if measurer is None:
        return None
    values = fragment_style({"size_pt": style.get("size_pt"), "bold": bool(style.get("bold"))}, fragment)
    text_fonts = fonts._text_fonts(style.get("font_japanese"), style.get("font_latin"))
    box = measurer(style.get("font_math")).box(
        str(fragment.get("text", "")), float(values["size_pt"]), values["bold"], display=display, text_fonts=text_fonts
    )
    return box.width, box.ink[0], box.ink[1]


def _drawn_edge_offsets(members, style, fonts, display: bool) -> tuple[float, float]:
    """行の端の数式が線で描く記号で始まるか終わるときの、予測の行の範囲の左端と右端の補正（pt。右端は負の値で内側）。"""
    visible = [(text, fragment) for text, fragment in members if not text.isspace()]
    if not visible:
        return 0.0, 0.0
    start = end = 0.0
    (_first, head), (_last, tail) = visible[0], visible[-1]
    if head.get("kind") == "math" and _drawn_edges(str(head.get("text", "")))[0]:
        ink = _math_ink(fonts, head, style, display)
        start = 0.0 if ink is None else ink[1]
    if tail.get("kind") == "math" and _drawn_edges(str(tail.get("text", "")))[1]:
        ink = _math_ink(fonts, tail, style, display)
        end = 0.0 if ink is None else ink[2] - ink[0]
    return start, end


def predicted_lines(layout: measure.ElementLayout, element, fonts) -> list[dict]:
    """PDF に現れる行（空白だけの行を除く）の予測。座標は絶対の pt。

    行の範囲は、PDF の読み取りと同じく、行頭と行末の空白を除いた範囲にする。行の端の数式が線で描く記号で始まるか
    終わるときは、その端を数式の描画範囲の端にする。
    """
    units = _line_units(element, layout)
    style = element.get("style") or {}
    display = measure._display(element)
    result = []
    for line, left, members in zip(layout.text.lines, layout.line_left, units):
        if not line.text.strip():
            continue
        lead, trail = _edge_spaces(members, style, fonts)
        start, end = _drawn_edge_offsets(members, style, fonts, display)
        plain = _MATH_IN_LINE.sub("", line.text).strip()
        formulas = len(_MATH_IN_LINE.findall(line.text))
        result.append(
            {
                "text": line.text,
                "baseline": layout.block_top + line.baseline,
                "left": left + lead + start,
                "right": left + line.width - trail + end,
                "characters": len(plain),
                "math": formulas > 0,
                "formulas": formulas,
                "stacked": bool(STACKED.search(line.text)),
                "near_limit": line.near_limit,
                "inked": line.ink_top > 0 or line.ink_bottom > 0,
            }
        )
    return result


def _predicted_fonts(element, fonts) -> dict[str, set[str]]:
    """通常文字ごとに、計測が使う書体（measure.font_runs）の集合。"""
    style = element.get("style") or {}
    result: dict[str, set[str]] = {}
    for fragment in element.get("content") or []:
        kind = fragment.get("kind", "text")
        if kind == "math":
            continue
        own = fragment.get("style") or {}
        bold = own.get("bold")
        if bold is None:
            bold = bool(style.get("bold")) or kind == "emphasis"
        for part in str(fragment.get("text", "")).split("\n"):
            if part:
                for run in measure.font_runs(part, style, fonts, bold=bold):
                    for char in run.text:
                        result.setdefault(char, set()).add(run.family)
    return result


def _extent_tolerance(line, size, rule: Tolerance | None = None) -> float:
    """行の幅と左端の許容幅（pt）。文字の分に、数式を含む行は数式 1 つあたりの数式の許容幅を加える。
    規則（width）があれば、その許容差にする。"""
    if rule is not None:
        return rule.value(size)
    tolerance = TEXT_WIDTH_EM_PER_CHAR * size * line["characters"] + TEXT_WIDTH_PT
    return tolerance + line["formulas"] * max(MATH_PT, MATH_WIDTH_EM * size)


def _math_vertical(size) -> float:
    """数式を含む行の基線と、数式を含む要素のインクの上下の許容幅（pt）。"""
    return max(MATH_PT, MATH_VERTICAL_EM * size)


def _differences(element, layout, actual, fonts, checks, *, rules=None) -> tuple[list[str], list[str]]:
    """（配置の違い、書体の違い）。行数が違うときは、行の位置とインクを比べない。書体は行に依らないので比べる。

    rules は項目から規則の許容差（`Tolerance`）への対応（`case_tolerances`）で、書いた項目はその許容差で比べる。
    """
    rules = rules or {}
    if not actual.get("glyphs"):
        return ["PDF に字形がない"], []
    typefaces = _font_differences(element, actual, fonts) if "fonts" in checks else []
    lines = predicted_lines(layout, element, fonts)
    if len(lines) != actual["lines"]:
        if "lines" not in checks:
            return [], typefaces
        if any(line["near_limit"] for line in lines) and abs(len(lines) - actual["lines"]) == 1:
            return [], typefaces  # 計測が誤差の範囲と示した行がある
        widths = [round(line["right"] - line["left"], 2) for line in lines]
        return [
            f"行数: 予測 {len(lines)}、実際 {actual['lines']}（予測の行幅 {widths}、"
            f"実際の行の範囲 {actual['line_extents_pt']}、内側の幅 {round(layout.inner.w, 2)}）"
        ], typefaces
    problems = []
    size = _size(element)
    baseline_rule, width_rule = rules.get("baseline"), rules.get("width")
    for number, (line, baseline, (left, right)) in enumerate(
        zip(lines, actual["baselines_pt"], actual["line_extents_pt"]), 1
    ):
        if "baselines" in checks and (not line["stacked"] or (baseline_rule is not None and baseline_rule.stacked)):
            if baseline_rule is not None:
                tolerance = baseline_rule.value(size)
            else:
                tolerance = _math_vertical(size) if line["math"] else BASELINE_PT
            if abs(baseline - line["baseline"]) > tolerance:
                problems.append(
                    f"{number} 行目の基線: 予測 {line['baseline']:.2f}、実際 {baseline:.2f}"
                    f"（差 {baseline - line['baseline']:+.2f}、{_allowed(tolerance, baseline_rule)}）"
                )
        if "extents" in checks:
            tolerance = _extent_tolerance(line, size, width_rule)
            width, expected = right - left, line["right"] - line["left"]
            if abs(width - expected) > tolerance:
                problems.append(
                    f"{number} 行目の幅: 予測 {expected:.2f}、実際 {width:.2f}"
                    f"（差 {width - expected:+.2f}、{_allowed(tolerance, width_rule)}）"
                )
            if abs(left - line["left"]) > tolerance:
                problems.append(
                    f"{number} 行目の左端: 予測 {line['left']:.2f}、実際 {left:.2f}"
                    f"（差 {left - line['left']:+.2f}、{_allowed(tolerance, width_rule)}）"
                )
    if "ink" in checks:
        problems += _ink_differences(element, layout, actual, lines, rules)
    return problems, typefaces


def _ink_differences(element, layout, actual, lines, rules=None) -> list[str]:
    rules = rules or {}
    ink = layout.ink
    if ink is None or actual.get("ink_pt") is None:
        if (ink is None) != (actual.get("ink_pt") is None):
            return [f"インク: 予測 {None if ink is None else 'あり'}、実際 {actual.get('ink_pt')}"]
        return []
    has_math = any(fragment.get("kind") == "math" for fragment in element.get("content") or [])
    size = _size(element)
    default = _math_vertical(size) if has_math else INK_PT
    ink_x, ink_y, width_rule = rules.get("ink_x"), rules.get("ink_y"), rules.get("width")
    vertical = default if ink_y is None else ink_y.value(size)
    # 計測のインクの横の範囲はインクのある行の範囲である。行の範囲と同じく、行頭と行末の空白を除く。
    # 横の端の許容幅は、その端を決める行の幅の許容幅と同じにする（PowerPoint の行の幅の差がそのまま端に出る）。
    inked = [line for line in lines if line["inked"]]
    if inked:
        first = min(inked, key=lambda line: line["left"])
        last = max(inked, key=lambda line: line["right"])
        left, right = first["left"], last["right"]
        horizontal = (_extent_tolerance(first, size, width_rule), _extent_tolerance(last, size, width_rule))
        horizontal_rule = width_rule
    else:
        left, right = ink.x, ink.x + ink.w
        horizontal, horizontal_rule = (default, default), None
    if ink_x is not None:
        horizontal, horizontal_rule = (ink_x.value(size),) * 2, ink_x
    expected = (left, ink.y, right, ink.y + ink.h)
    tolerances = (horizontal[0], vertical, horizontal[1], vertical)
    sources = (horizontal_rule, ink_y, horizontal_rule, ink_y)
    return [
        f"インクの{edge}: 予測 {a:.2f}、実際 {b:.2f}（差 {b - a:+.2f}、{_allowed(tolerance, rule)}）"
        for edge, a, b, tolerance, rule in zip(_EDGES, expected, actual["ink_pt"], tolerances, sources)
        if abs(b - a) > tolerance
    ]


def _font_differences(element, actual, fonts) -> list[str]:
    predicted = _predicted_fonts(element, fonts)
    # 数式の \text{…} の文字は数式の書き出しが書体を決めるので比べない。PDF の書体は文字ごとに最初に描いた字形のものなので、
    # 同じ文字が通常文字にあっても比べない。
    in_text = "".join(
        _text_in_math(str(fragment.get("text", "")))
        for fragment in element.get("content") or []
        if fragment.get("kind") == "math"
    )
    problems = []
    for char, pdf_font in sorted((actual.get("fonts") or {}).items()):
        families = predicted.get(char)
        if not families or char in in_text or char.isspace():
            continue
        if _font_name(pdf_font) not in {_font_name(family) for family in families}:
            problems.append(f"書体: {char!r} は予測 {'・'.join(sorted(families))}、PDF {pdf_font}")
    return problems


def compare_case(case, fonts, *, checks=CHECKS, rules=None) -> list[str]:
    """照合データの 1 要素を計測と比べ、違いの文言（要素の識別子・タグ・予測と実際の値を含む）を返す。違いがなければ空。

    rules は許容差の規則（`Rule` の列）で、省略すると tests/data/layout/tolerances.toml を使う。
    """
    tags = list(case.get("tags") or ())
    label = f"{case.get('slide_id')}/{case.get('element_id')} [{' '.join(tags)}]"
    try:
        layout = measure.layout_element(case["element"], fonts)
    except measure.MeasureError as error:
        return [f"{label}: 計測できない: {error}"]
    try:
        tolerances = case_tolerances(case, load_rules() if rules is None else rules)
        placement, typefaces = _differences(case["element"], layout, case["actual"], fonts, checks, rules=tolerances)
    except ValueError as error:
        return [f"{label}: {error}"]
    if placement and "overflow" in tags:
        if any(finding["code"] == "text_overflow" for finding in measure.fit_findings(layout)):
            placement = []  # 枠より長い内容は資料の誤りとして報告されるので、配置の一致は求めない
        else:
            placement.insert(0, "枠より長い内容の要素（overflow）だが、計測が text_overflow を報告しない")
    return [f"{label}: {problem}" for problem in placement + typefaces]


def font_mismatches(data, environment) -> list[str]:
    """照合データに記録した書体と、手元で解決した書体のファイル（sha256 と書体番号）の違い。"""
    problems = []
    for family, record in sorted(((data.get("environment") or {}).get("fonts") or {}).items()):
        specification = (environment.get("fonts") or {}).get(family) or default_font(family, environment)
        for prefix, label in (("", family), ("bold_", f"{family}（太字）")):
            expected = record.get(prefix + "sha256")
            if not expected:
                continue
            path = (specification or {}).get(prefix + "path")
            if not path or not Path(path).is_file():
                problems.append(f"{label} が手元にない")
            elif file_hash(path) != expected or int((specification or {}).get(prefix + "index", 0)) != int(
                record.get(prefix + "index", 0)
            ):
                problems.append(f"{label} のファイルが違う")
    return problems


def fixture_path() -> Path:
    value = os.environ.get(FIXTURE_VARIABLE)
    return Path(value) if value else DEFAULT_FIXTURE


def test_measurement_matches_powerpoint_layout(request):
    path = fixture_path()
    if not path.is_file():
        pytest.skip(f"照合データがない（tools/make_layout_fixture.py で作る）: {path.name}")
    request.getfixturevalue("windows_fonts_environment")
    data = json.loads(path.read_text(encoding="utf-8"))
    assert data.get("schema_version") == 1, f"照合データの形式が違う: {data.get('schema_version')}"
    project = load_project(DECK / "deck.toml")
    mismatches = font_mismatches(data, project.environment)
    if mismatches:
        pytest.skip(
            "試験データと異なるフォントのため省略。tools/make_layout_fixture.py で作り直す（"
            + "、".join(mismatches)
            + "）"
        )
    with FontRegistry(project) as registry:
        results = [compare_case(case, registry) for case in data["cases"]]
    failed = [messages for messages in results if messages]
    if failed:
        pytest.fail(
            f"PowerPoint の配置と計測が違う要素が {len(failed)} 個ある（照合データ {path.name}、要素 {len(results)} 個）:\n"
            + "\n".join(message for messages in failed for message in messages),
            pytrace=False,
        )


# ---- 比べる処理の試験（偽のフォントで、予測から作った実際の値を使う） ----


def element(content, *, id="sample", kind="text", area=(10, 20, 200, 100), **style):
    base = {
        "size_pt": 20.0,
        "line_spacing": 1.0,
        "font_latin": "Arial",
        "font_japanese": "Yu Gothic UI",
        "font_math": "Cambria Math",
        "word_wrap": True,
        "align": "left",
        "valign": "top",
        "padding_pt": 0.0,
        "bold": False,
        "color": "17212D",
        "fill": None,
        "line_color": None,
    }
    base.update(style)
    return {"id": id, "kind": kind, "area": list(area), "role": "body", "style": base, "content": content}


def text(value, **options):
    return {"kind": "text", "text": value, **options}


def math_fragment(value):
    return {"kind": "math", "text": value}


def actual_from_prediction(value, fonts, *, pdf_fonts=None):
    """予測と同じ値の actual（PowerPoint が計測どおりに描いた場合）。"""
    layout = measure.layout_element(value, fonts)
    lines = predicted_lines(layout, value, fonts)
    ink = layout.ink
    chars = {
        char
        for fragment in value["content"]
        if fragment["kind"] != "math"
        for char in fragment["text"]
        if not char.isspace()
    }
    return {
        "lines": len(lines),
        "baselines_pt": [line["baseline"] for line in lines],
        "line_extents_pt": [[line["left"], line["right"]] for line in lines],
        "ink_pt": None if ink is None else [ink.x, ink.y, ink.x + ink.w, ink.y + ink.h],
        "sizes_pt": [20.0],
        "fonts": pdf_fonts if pdf_fonts is not None else {char: "ArialMT" for char in sorted(chars)},
        "glyphs": max(1, len(chars)),
        "drawings": 0,
    }


def case(value, actual, *tags):
    return {"slide_id": "page", "element_id": value["id"], "tags": ["page", *tags], "element": value, "actual": actual}


def test_comparison_accepts_layout_that_matches_measurement():
    fonts = FakeFonts()
    for value in (
        element([text("Two short\nlines of text")]),
        element([text("Wrapped words in a narrow frame")], area=(10, 20, 120, 120), align="center"),
        element([text("Sum "), math_fragment("x+1")], area=(10, 20, 300, 60), valign="middle"),
    ):
        assert compare_case(case(value, actual_from_prediction(value, fonts)), fonts) == []


@pytest.mark.parametrize("expression", ["m^{1/4}", r"x^{\frac{a}{b}}"])
def test_text_markers_check_math_baseline_with_text_tolerance(expression):
    fonts = FakeFonts()
    value = element([text("["), math_fragment(expression), text("]")])
    actual = actual_from_prediction(value, fonts)
    actual["baselines_pt"][0] += 0.8
    assert compare_case(case(value, actual), fonts, checks=("baselines",)) == []
    differences = compare_case(case(value, actual, "text_baseline"), fonts, checks=("baselines",))
    assert len(differences) == 1 and "許容 0.50、規則 text_baseline" in differences[0]


def rules(*entries):
    return parse_rules({"schema_version": 1, "rules": [{"reason": "試験", **entry} for entry in entries]})


def test_tolerance_rules_override_selected_items():
    """規則は書いた項目だけの許容差を決める。インクの左右は、ink_x がなければ行の幅（width）の許容差に従う。"""
    fonts = FakeFonts()
    value = element([text("a" * 40)], area=(10, 20, 500, 60))  # 通常の幅の許容差は 3.1 pt
    actual = actual_from_prediction(value, fonts)
    actual["line_extents_pt"][0][1] += 1.0
    actual["ink_pt"][2] += 1.0
    actual["ink_pt"][3] += 0.5
    strict = case(value, actual, "strict")
    assert compare_case(strict, fonts, rules=()) == []
    width, ink = compare_case(strict, fonts, rules=rules({"tags": ["strict"], "width": {"tolerance_pt": 0.5}}))
    assert "1 行目の幅" in width and "（差 +1.00、許容 0.50、規則 strict）" in width
    assert "インクの右端" in ink and "許容 0.50、規則 strict" in ink
    # ink_x と ink_y は、インクの左右と上下の許容差を別に決める。tolerance_em は要素の最大の文字サイズの em。
    both = rules(
        {"tags": ["strict"], "width": {"tolerance_pt": 0.5}, "ink_x": {"tolerance_pt": 2.0}},
        {"tags": ["strict"], "ink_y": {"tolerance_em": 0.02, "tolerance_pt": 0.1}},
    )
    width, bottom = compare_case(strict, fonts, rules=both)
    assert "1 行目の幅" in width and "インクの下端" in bottom and "（差 +0.50、許容 0.40、規則 strict）" in bottom
    # 要素で選ぶ規則。tags と elements の両方を書くと、両方に合うケースだけを選ぶ。
    narrow = {"width": {"tolerance_pt": 0.5}}
    assert compare_case(strict, fonts, rules=rules({"elements": ["page/other"], **narrow})) == []
    assert len(compare_case(strict, fonts, rules=rules({"elements": ["page/sample"], **narrow}))) == 2
    assert compare_case(strict, fonts, rules=rules({"tags": ["other"], "elements": ["page/sample"], **narrow})) == []


def test_tolerance_rule_compares_stacked_baselines_when_asked():
    fonts = FakeFonts()
    value = element([math_fragment(r"\frac{a}{b}")], kind="equation", area=(10, 20, 200, 80))
    actual = actual_from_prediction(value, fonts)
    actual["baselines_pt"][0] += 0.8
    stacked = case(value, actual, "nested")
    plain = rules({"tags": ["nested"], "baseline": {"tolerance_pt": 0.5}})
    assert compare_case(stacked, fonts, checks=("baselines",), rules=plain) == []  # 分数の行の基線は比べない
    strict = rules({"tags": ["nested"], "baseline": {"tolerance_pt": 0.5, "stacked": True}})
    (message,) = compare_case(stacked, fonts, checks=("baselines",), rules=strict)
    assert "1 行目の基線" in message and "許容 0.50、規則 nested" in message


def test_overlapping_tolerance_rules_are_reported():
    fonts = FakeFonts()
    value = element([text("Hello")])
    overlapping = rules(
        {"tags": ["a"], "baseline": {"tolerance_pt": 0.5}},
        {"elements": ["page/sample"], "baseline": {"tolerance_pt": 1}},
    )
    (message,) = compare_case(case(value, actual_from_prediction(value, fonts), "a"), fonts, rules=overlapping)
    assert "許容差の規則が重なる: baseline を「a」と「page/sample」が決める" in message


@pytest.mark.parametrize(
    "data",
    [
        {"rules": []},
        {"schema_version": 1, "rules": [{"tags": ["a"], "baseline": {"tolerance_pt": 0.5}}]},  # reason がない
        {"schema_version": 1, "rules": [{"reason": "r", "baseline": {"tolerance_pt": 0.5}}]},  # 選ぶ条件がない
        {"schema_version": 1, "rules": [{"reason": "r", "tags": ["a"]}]},  # 項目がない
        {"schema_version": 1, "rules": [{"reason": "r", "tags": ["a"], "height": {"tolerance_pt": 1}}]},
        {"schema_version": 1, "rules": [{"reason": "r", "tags": ["a"], "baseline": {"tolerance_pt": 0}}]},
        {"schema_version": 1, "rules": [{"reason": "r", "tags": ["a"], "baseline": {"stacked": True}}]},
        {"schema_version": 1, "rules": [{"reason": "r", "tags": ["a"], "width": {"tolerance_pt": 1, "stacked": True}}]},
        {"schema_version": 1, "rules": [{"reason": "r", "elements": ["sample"], "width": {"tolerance_pt": 1}}]},
        {"schema_version": 1, "rules": [{"reason": "r", "tags": "a", "width": {"tolerance_pt": 1}}]},
        {"schema_version": 1, "cases": []},
    ],
)
def test_tolerance_rules_reject_malformed_data(data):
    with pytest.raises(ValueError):
        parse_rules(data)


def test_stored_tolerance_rules_select_calibration_cases(tmp_path):
    """保存した規則は根拠を持ち、どれも校正用の資料の要素を選ぶ。text_baseline の規則は、以前の試験の特例
    （数式の構造によらず基線を 0.5 pt で比べる）と同じ。"""
    stored = load_rules()
    assert stored and all(rule.reason for rule in stored)
    tool = load_tool()
    _project, scene = calibration_scene(tmp_path)
    cases = [
        {"slide_id": slide["id"], "element_id": value["id"], "tags": tool.case_tags(slide, value)}
        for slide in scene["slides"]
        for value in tool.text_elements(slide)
    ]
    for rule in stored:
        assert any(rule.matches(item) for item in cases), rule.name
    for item in cases:
        case_tolerances(item, stored)  # 重なる規則がない
    (text_baseline,) = [rule for rule in stored if rule.tags == ("text_baseline",)]
    assert text_baseline.items == {"baseline": Tolerance(rule="text_baseline", tolerance_pt=0.5, stacked=True)}


def test_comparison_reports_identifiers_tags_and_values():
    fonts = FakeFonts()
    value = element([text("First line\nSecond line")], id="lead")
    actual = actual_from_prediction(value, fonts)
    actual["baselines_pt"][1] += 0.8
    actual["line_extents_pt"][0][1] += 2.0
    actual["ink_pt"][3] -= 1.0
    width, baseline, ink = compare_case(case(value, actual, "pitch", "valign:top"), fonts)
    for message in (width, baseline, ink):
        assert message.startswith("page/lead [page pitch valign:top]: ")
    assert "1 行目の幅: 予測 100.00、実際 102.00（差 +2.00、許容 1.00）" in width
    assert "2 行目の基線" in baseline and "差 +0.80、許容 0.50" in baseline
    assert "インクの下端" in ink and "差 -1.00、許容 0.60" in ink
    actual["lines"] = 3
    (message,) = compare_case(case(value, actual), fonts)
    assert "行数: 予測 2、実際 3" in message and "内側の幅 200.0" in message


def test_text_width_tolerance_grows_with_characters():
    fonts = FakeFonts()
    value = element([text("a" * 40)], area=(10, 20, 500, 60))  # 40 文字 × 20 pt で 0.0035 em × 40 + 0.3 pt = 3.1 pt
    actual = actual_from_prediction(value, fonts)
    actual["line_extents_pt"][0][1] += 3.0
    actual["ink_pt"][2] += 3.0  # インクの横の端も、その端を決める行の幅の許容幅で比べる
    assert compare_case(case(value, actual), fonts) == []
    actual["line_extents_pt"][0][1] += 0.2
    actual["ink_pt"][2] += 0.2
    width, ink = compare_case(case(value, actual), fonts)
    assert "1 行目の幅" in width and "許容 3.10" in width
    assert "インクの右端" in ink and "許容 3.10" in ink
    # インクの上下は文字数によらない。
    actual = actual_from_prediction(value, fonts)
    actual["ink_pt"][3] += 0.7
    (message,) = compare_case(case(value, actual), fonts)
    assert "インクの下端" in message and "許容 0.60" in message


def test_ink_edge_tolerance_follows_the_line_at_that_edge():
    """インクの横の端の許容幅は、その端を決める行（右端なら最も右に届く行）の文字数で決まる。"""
    fonts = FakeFonts()
    value = element(
        [text("a" * 30 + "\nbb")], area=(10, 20, 500, 80)
    )  # 1 行目 30 文字で 2.4 pt、2 行目 2 文字で 0.44 pt
    actual = actual_from_prediction(value, fonts)
    actual["ink_pt"][2] += 2.3
    actual["line_extents_pt"][0][1] += 2.3
    assert compare_case(case(value, actual), fonts) == []
    right = element([text("b\n" + "a" * 4)], area=(10, 20, 500, 80), align="right")
    actual = actual_from_prediction(right, fonts)
    actual["ink_pt"][0] -= 0.5  # 左端は 4 文字の 2 行目で決まる（0.0035 × 20 × 4 + 0.3 = 0.58 pt）
    assert compare_case(case(right, actual), fonts) == []
    actual["ink_pt"][0] -= 0.2
    (message,) = compare_case(case(right, actual), fonts)
    assert "インクの左端" in message and "許容 0.58" in message


def test_math_lines_use_wider_tolerance_and_skip_stacked_baselines():
    fonts = FakeFonts()
    value = element([math_fragment(r"\frac{a}{b}")], kind="equation", area=(10, 20, 200, 80))
    actual = actual_from_prediction(value, fonts)
    actual["baselines_pt"][0] += 9.0  # 分数の分母の基線を読んだ場合
    # 数式の許容幅は数式の照合の基準と同じ：幅は 1 式あたり 0.12 em（20 pt で 2.4 pt）、上下は 0.1 em（2 pt）。
    assert (MATH_WIDTH_EM, MATH_VERTICAL_EM) == (0.12, 0.1)
    actual["line_extents_pt"][0][1] += 2.6
    actual["ink_pt"][1] -= 1.9
    assert compare_case(case(value, actual), fonts) == []
    actual["ink_pt"][1] -= 0.2
    (message,) = compare_case(case(value, actual), fonts)
    assert "インクの上端" in message and "許容 2.00" in message
    # 小さい文字でも 1.5 pt を下限とする。文字と数式が混じる行は、文字の分と数式の分を加える。
    small = element([text("ab "), math_fragment("x"), text(" cd")], area=(10, 20, 300, 60), size_pt=10.0)
    lines = predicted_lines(measure.layout_element(small, fonts), small, fonts)
    assert lines[0]["formulas"] == 1 and lines[0]["characters"] == 6  # 数式を除いた「ab  cd」
    actual = actual_from_prediction(small, fonts)
    actual["line_extents_pt"][0][1] += 0.0035 * 10 * 6 + 0.3 + 1.5 - 0.01
    actual["baselines_pt"][0] += 1.49
    assert compare_case(case(small, actual), fonts) == []


class InkFonts(FakeFonts):
    """数式の描画範囲の左右を返す偽のフォント（FontRegistry の数式の計測器の代わり）。inks は式から（左, 右）の em。"""

    def __init__(self, inks, **options):
        super().__init__(**options)
        self.inks = inks

    def _text_fonts(self, japanese=None, latin=None):
        return (japanese, latin)

    def _math_measurer(self, family=None):
        fonts = self

        class Measurer:
            def box(self, expr, size, bold=False, *, display=False, text_fonts=None):
                box = fonts.math_box(expr, size, bold)
                left, right = fonts.inks.get(expr, (0.0, box.width / size))
                return types.SimpleNamespace(width=box.width, ink=(left * size, right * size, box.ascent, box.descent))

        return Measurer()


def test_drawn_math_edges_use_ink_range():
    """行の端の数式が伸びる括弧か根号の記号で始まるか終わるとき（PowerPoint は線で描き、PDF の読み取りは描画の範囲で
    行の範囲に含める）、予測の行の範囲のその端を数式の描画範囲の端にする。字形の端は送り幅のまま。"""
    assert _drawn_edges(r"\left(\frac{1}{2}\right)") == (True, True)
    assert _drawn_edges(r"\left(x\right)^2") == (True, False)
    assert _drawn_edges(r"\left.x\right|") == (False, True)
    assert _drawn_edges(r"\sqrt{n}") == (True, False)
    assert _drawn_edges(r"\,\sqrt{n}") == (True, False)
    assert _drawn_edges(r"\sqrt[3]{n}") == (False, False)
    assert _drawn_edges(r"x^2") == _drawn_edges(r"(x)") == _drawn_edges(r"\frac{a}{b}") == (False, False)
    paren, root = r"\left(\frac{1}{2}\right)", r"\sqrt{n}"
    fonts = InkFonts(
        {paren: (0.1, 11.9), root: (0.05, 4.0)}
    )  # 送り幅は記法の 1 文字 0.5 em（paren は 12 em、root は 4 em）
    alone = element([math_fragment(paren)], area=(10, 20, 300, 60))
    (line,) = predicted_lines(measure.layout_element(alone, fonts), alone, fonts)
    assert (line["left"], line["right"]) == pytest.approx((10 + 2, 10 + 238))
    mixed = element([math_fragment(root), text(" ab")], area=(10, 20, 300, 60))
    (line,) = predicted_lines(measure.layout_element(mixed, fonts), mixed, fonts)
    assert (line["left"], line["right"]) == pytest.approx((10 + 1, 10 + 80 + 10 + 20))  # 右端は文字の送り幅
    ending = element([text("ab "), math_fragment(root)], area=(10, 20, 300, 60))
    (line,) = predicted_lines(measure.layout_element(ending, fonts), ending, fonts)
    assert (line["left"], line["right"]) == pytest.approx((10, 10 + 30 + 80))  # 根号の右端は送り幅
    # 偽のフォント（数式の計測器がない）では送り幅のまま。
    plain = FakeFonts()
    (line,) = predicted_lines(measure.layout_element(alone, plain), alone, plain)
    assert (line["left"], line["right"]) == pytest.approx((10, 10 + 240))


def test_empty_lines_are_not_counted():
    fonts = FakeFonts()
    value = element([text("First\n\nThird")])
    actual = actual_from_prediction(value, fonts)
    assert actual["lines"] == 2 and len(measure.layout_element(value, fonts).text.lines) == 3
    assert compare_case(case(value, actual), fonts) == []


def test_overflow_case_needs_text_overflow_from_measurement():
    fonts = FakeFonts()
    long_word = element([text("Internationalization")], area=(10, 20, 125, 100))  # 20 文字 × 10 pt は内側より広い
    # 計測も PowerPoint と同じく語の途中で分け、分ける位置のない部分が内側より長いので text_overflow を報告する。
    layout = measure.layout_element(long_word, fonts)
    assert [line.text for line in layout.text.lines] == ["Internationa", "lization"]
    assert any(finding["code"] == "text_overflow" for finding in measure.fit_findings(layout))
    assert compare_case(case(long_word, actual_from_prediction(long_word, fonts)), fonts) == []
    # 分ける位置が違っても、overflow のタグの要素は配置の違いを許す（資料の誤りとして報告されるため）。
    actual = actual_from_prediction(long_word, fonts)
    actual.update(
        lines=3,
        baselines_pt=[actual["baselines_pt"][0] + 24.0 * number for number in range(3)],
        line_extents_pt=[[10.0, 110.0], [10.0, 110.0], [10.0, 10.0]],
    )
    assert compare_case(case(long_word, actual, "overflow"), fonts) == []
    (message,) = compare_case(case(long_word, actual), fonts)
    assert "行数: 予測 2、実際 3" in message
    # 配置の違いを許しても、書体は比べる。
    (message,) = compare_case(case(long_word, {**actual, "fonts": {"I": "MS-Gothic"}}, "overflow"), fonts)
    assert "書体: 'I' は予測 Arial、PDF MS-Gothic" in message
    fits = element([text("Short")], area=(10, 20, 120, 100))
    wrong = actual_from_prediction(fits, fonts)
    wrong["baselines_pt"][0] += 2.0
    first, second = compare_case(case(fits, wrong, "overflow"), fonts)
    assert "text_overflow を報告しない" in first and "1 行目の基線" in second


def test_near_limit_line_count_difference_is_accepted():
    fonts = FakeFonts()
    # 内側の幅 100.1 pt に 10 文字（100 pt）：計測は 1 行で、誤差の範囲（near_limit）と示す。
    value = element([text("abcde abcd")], area=(10, 20, 100.1, 100))
    layout = measure.layout_element(value, fonts)
    assert len(layout.text.lines) == 1 and layout.text.lines[0].near_limit
    actual = actual_from_prediction(value, fonts)
    actual.update(lines=2, baselines_pt=actual["baselines_pt"] * 2, line_extents_pt=[[10.0, 70.0], [10.0, 50.0]])
    assert compare_case(case(value, actual), fonts) == []
    # 行数の違いを許しても、書体は比べる。
    (message,) = compare_case(case(value, {**actual, "fonts": {"a": "MS-Gothic"}}), fonts)
    assert "書体: 'a' は予測 Arial、PDF MS-Gothic" in message
    actual.update(lines=3, baselines_pt=actual["baselines_pt"] + [0.0], line_extents_pt=actual["line_extents_pt"] * 2)
    (message,) = compare_case(case(value, actual), fonts)
    assert "行数: 予測 1、実際 3" in message
    # 行数が違う要素も、書体は比べる（行の位置とインクは比べない）。
    line_count, typeface = compare_case(case(value, {**actual, "fonts": {"a": "MS-Gothic"}}), fonts)
    assert "行数: 予測 1、実際 3" in line_count and "書体: 'a'" in typeface


def test_font_of_each_character_is_compared():
    fonts = FakeFonts(missing={"Arial": {"※"}})
    value = element([text("Note ※ here")], area=(10, 20, 300, 60))
    good = actual_from_prediction(value, fonts, pdf_fonts={"N": "ArialMT", "※": "ABCDEF+YuGothicUI-Regular"})
    assert compare_case(case(value, good), fonts) == []
    bad = actual_from_prediction(value, fonts, pdf_fonts={"N": "ArialMT", "※": "MS-Gothic"})
    (message,) = compare_case(case(value, bad), fonts)
    assert "書体: '※' は予測 Yu Gothic UI、PDF MS-Gothic" in message
    # 数式の \text の文字は比べない。
    mixed = element([text("abc "), math_fragment(r"x\text{abc}")], area=(10, 20, 300, 60))
    other = actual_from_prediction(mixed, fonts, pdf_fonts={"a": "CambriaMath", "b": "ArialMT"})
    assert compare_case(case(mixed, other), fonts) == []
    # \text の外の数式の文字（ここでは T）と同じ通常文字は比べる。
    formula = element([text("The forecast "), math_fragment("T")], area=(10, 20, 300, 60))
    wrong = actual_from_prediction(formula, fonts, pdf_fonts={"T": "MS-Gothic", "h": "MS-Gothic", "e": "ArialMT"})
    assert [message.split("]: ", 1)[1] for message in compare_case(case(formula, wrong), fonts)] == [
        "書体: 'T' は予測 Arial、PDF MS-Gothic",
        "書体: 'h' は予測 Arial、PDF MS-Gothic",
    ]


def test_missing_glyphs_are_reported():
    fonts = FakeFonts()
    value = element([text("Hello")])
    actual = actual_from_prediction(value, fonts)
    actual.update(glyphs=0, lines=0, baselines_pt=[], line_extents_pt=[], ink_pt=None)
    (message,) = compare_case(case(value, actual), fonts)
    assert message.endswith("PDF に字形がない")


def test_fixture_fonts_must_match_local_files(tmp_path):
    regular, bold = tmp_path / "sample.ttf", tmp_path / "sample-bold.ttf"
    regular.write_bytes(b"regular")
    bold.write_bytes(b"bold")
    environment = {"fonts": {"Sample": {"path": str(regular), "index": 0, "bold_path": str(bold), "bold_index": 0}}}
    record = {"sha256": file_hash(regular), "index": 0, "bold_sha256": file_hash(bold), "bold_index": 0}
    data = {"environment": {"fonts": {"Sample": record}}}
    assert font_mismatches(data, environment) == []
    bold.write_bytes(b"changed")
    assert font_mismatches(data, environment) == ["Sample（太字） のファイルが違う"]
    data["environment"]["fonts"]["Missing"] = {"sha256": "0" * 64, "index": 0}
    assert font_mismatches(data, environment)[0] == "Missing が手元にない"


def test_fixture_is_skipped_when_absent(tmp_path, monkeypatch):
    monkeypatch.setenv(FIXTURE_VARIABLE, str(tmp_path / "none.json"))
    assert fixture_path() == tmp_path / "none.json" and not fixture_path().exists()
    monkeypatch.delenv(FIXTURE_VARIABLE)
    assert fixture_path() == DEFAULT_FIXTURE


# ---- 照合データを作る道具（合成した実行と PDF で確かめる） ----

COLOR = "17212D"
RGB = tuple(int(COLOR[i : i + 2], 16) / 255 for i in (0, 2, 4))
# ローカルのパスに見える例（ソースに書かないよう、実行時に組み立てる）。
UNIX_HOME = "/".join(("", "home", "someone", "notes.txt"))
WINDOWS_USER = "\\".join(("C:", "Users", "someone", "AppData"))
WINDOWS_PROGRAM = "\\".join(("C:", "Office", "POWERPNT.EXE"))


def scene_element(id, area, content, *, kind="text", **extra):
    style = {
        "size_pt": 20.0,
        "line_spacing": 1.0,
        "color": COLOR,
        "padding_pt": 0.0,
        "word_wrap": True,
        "align": "left",
        "valign": "top",
        "font_latin": "Helvetica",
        "font_japanese": "Helvetica",
        "font_math": "Times-Roman",
    }
    return {"id": id, "kind": kind, "area": list(area), "role": "body", "style": style, "content": content, **extra}


def synthetic_run(root, *, first_text="Hello world", fonts=("Helvetica", "Times-Roman")):
    """2 ページの場面・生成記録・標準表示の PDF を持つ実行ディレクトリを作る。"""
    run = root / "runs" / "0001"
    standard = run / "renders" / "powerpoint" / "standard"
    standard.mkdir(parents=True)
    background = {
        "id": "_theme/background",
        "kind": "box",
        "area": [0, 0, 720, 540],
        "role": "body",
        "style": {},
        "content": [],
        "semantic": "background",
    }
    scene = {
        "schema_version": 1,
        "size": [720, 540],
        "slides": [
            {
                "id": "first",
                "number": 1,
                "elements": [
                    background,
                    scene_element("lead", (40, 40, 300, 40), [text(first_text)], description="tags: c1 sample"),
                    scene_element("note", (40, 120, 300, 60), [text("Two\nlines")]),
                    scene_element("blank", (400, 120, 100, 40), [text("   ")]),
                ],
            },
            {
                "id": "second",
                "number": 2,
                "elements": [
                    scene_element("formula", (40, 40, 200, 60), [math_fragment("x+1")], kind="equation"),
                ],
            },
        ],
    }
    (run / "deck.json").write_text(json.dumps(scene, ensure_ascii=False), encoding="utf-8")
    inventory = {
        family: {
            "path": str(root / f"{family}.ttf"),
            "index": 0,
            "sha256": f"{number:064d}",
            "version": "Version 1.0",
            "font_names": [family],
            "roles": [],
        }
        for number, family in enumerate(fonts, 1)
    }
    inventory.get("Helvetica", {})["bold"] = {
        "path": str(root / "bold.ttf"),
        "index": 0,
        "sha256": "f" * 64,
        "version": "Version 1.0",
    }
    manifest = {
        "schema_version": 2,
        "environment": {
            "fonts": {"fonts": inventory, "roles": {}},
            "packages": {"pypdfium2": "5.14.0", "pypdf": "6.21.1", "fonttools": "4.66.1", "lxml": "6"},
            "renderers": {
                "powerpoint": {"version": "16.0.1.2", "executable": WINDOWS_PROGRAM, "state_dir": WINDOWS_USER}
            },
        },
    }
    (run / "manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
    with Document() as document:
        page = document.new_page(width=720, height=540)
        for x, y, value in ((44, 64, first_text), (44, 144, "Two"), (44, 168, "lines")):
            page.insert_text((x, y), value, fontname="helv", fontsize=20, color=RGB)
        page = document.new_page(width=720, height=540)
        page.insert_text((44, 80), "x+1", fontname="tiro", fontsize=20, color=RGB)
        document.save(standard / "deck.pdf")
    return run


def test_make_layout_fixture_writes_cases_without_paths(tmp_path):
    tool = load_tool()
    run = synthetic_run(tmp_path)
    output = tmp_path / "fixture.json"
    assert tool.main(["--run", str(run), "--output", str(output)]) == 0
    raw = output.read_text(encoding="utf-8")
    assert str(tmp_path) not in raw and "Users" not in raw and "POWERPNT" not in raw
    data = json.loads(raw)
    assert (data["schema_version"], data["measure_version"]) == (1, measure.MEASURE_VERSION)
    assert data["environment"]["powerpoint"] == "16.0.1.2"
    assert data["environment"]["packages"] == {"pypdfium2": "5.14.0", "pypdf": "6.21.1", "fonttools": "4.66.1"}
    assert data["environment"]["fonts"] == {
        "Helvetica": {
            "version": "Version 1.0",
            "index": 0,
            "sha256": f"{1:064d}",
            "bold_version": "Version 1.0",
            "bold_index": 0,
            "bold_sha256": "f" * 64,
        },
        "Times-Roman": {"version": "Version 1.0", "index": 0, "sha256": f"{2:064d}"},
    }
    assert [slide["id"] for slide in data["slides"]] == ["first", "second"] and data["unread"] == []
    cases = {case["element_id"]: case for case in data["cases"]}
    assert list(cases) == ["lead", "note", "formula"]
    lead = cases["lead"]
    assert lead["tags"][:3] == ["first", "c1", "sample"]
    assert {"kind:text", "wrap", "align:left", "valign:top", "size:20", "spacing:1", "latin:Helvetica"} <= set(
        lead["tags"]
    )
    assert lead["element"]["description"] == "tags: c1 sample"
    width = text_length("Hello world", fontname="helv", fontsize=20)
    actual = lead["actual"]
    assert (actual["lines"], actual["baselines_pt"], actual["glyphs"]) == (1, [64.0], 10)
    assert actual["line_extents_pt"] == [[44.0, pytest.approx(44 + width, abs=0.001)]]
    assert actual["sizes_pt"] == [20.0] and actual["fonts"]["H"] == "Helvetica"
    assert actual["ink_pt"][0] == 44.0 and actual["ink_pt"][1] < 64 < actual["ink_pt"][3]
    assert cases["note"]["actual"]["lines"] == 2 and "newline" in cases["note"]["tags"]
    formula = cases["formula"]
    assert {"math", "display", "kind:equation"} <= set(formula["tags"])
    assert (formula["actual"]["lines"], formula["actual"]["glyphs"], formula["actual"]["fonts"]) == (1, 3, {})


def test_make_layout_fixture_refuses_local_paths(tmp_path, capsys):
    tool = load_tool()
    run = synthetic_run(tmp_path, first_text=f"See {UNIX_HOME}")
    output = tmp_path / "fixture.json"
    assert tool.main(["--run", str(run), "--output", str(output)]) == 2
    assert "ローカルのパスに見える文字列" in capsys.readouterr().err and not output.exists()
    for value in (WINDOWS_USER, f"x {UNIX_HOME}", "~/notes", "/".join(("", "mnt", "c", "Windows"))):
        with pytest.raises(tool.FixtureError):
            tool.check_public({"text": value})
    tool.check_public({"text": "read/write 1:2 a/b https://example.com"})


def test_make_layout_fixture_needs_font_records_and_pdf(tmp_path, capsys):
    tool = load_tool()
    run = synthetic_run(tmp_path, fonts=("Helvetica",))
    assert tool.main(["--run", str(run), "--output", str(tmp_path / "a.json")]) == 2
    assert "生成記録に書体の記録がありません: Times-Roman" in capsys.readouterr().err
    (run / "renders" / "powerpoint" / "standard" / "deck.pdf").unlink()
    assert tool.main(["--run", str(run), "--output", str(tmp_path / "b.json")]) == 2
    assert "標準表示の PDF がありません" in capsys.readouterr().err


def test_make_layout_fixture_help(capsys):
    with pytest.raises(SystemExit) as raised:
        load_tool().main(["--help"])
    assert raised.value.code == 0
    output = capsys.readouterr().out
    assert "--run" in output and "--output" in output and "--pdf" in output


def test_fixture_cases_compare_with_their_own_pdf_reading(tmp_path):
    """道具が書いた actual を、同じ PDF の読み取り（read_pdf）と比べる（道具と試験の値の形が合っていること）。"""
    tool = load_tool()
    run = synthetic_run(tmp_path)
    data = tool.make_fixture(run, generated_at="2026-10-12T00:00:00+00:00")
    scene = json.loads((run / "deck.json").read_text(encoding="utf-8"))
    readings = read_pdf(run / "renders" / "powerpoint" / "standard" / "deck.pdf", scene)
    for item in data["cases"]:
        page = [slide["id"] for slide in scene["slides"]].index(item["slide_id"])
        assert item["actual"] == tool.actual_record(readings[page].elements[item["element_id"]])


def test_edge_spaces_are_excluded_like_pdf_reading(tmp_path):
    """行頭・行末に空白がある行を計測どおりの位置に描いた PDF の actual が、予測と一致する。

    PDF の読み取りは空白の字形を除くので、予測の行の範囲とインクの横の範囲からも空白の送り幅（kern を含む）を除く。
    """
    tool = load_tool()
    fonts = FakeFonts(advance=0.6, kerning={(" ", "a"): -0.1})  # 送り幅は Courier と同じ 0.6 em（20 pt で 12 pt）
    values = [
        element([text("abcd\n\u3000ef")], id="lead", area=(40, 40, 200, 60)),
        element([text("abcd\u3000\nef")], id="trail", area=(40, 160, 200, 60)),
        element([text(" ab")], id="space", area=(300, 40, 200, 30)),
        element([text("\u3000ef\u3000")], id="center", area=(300, 160, 100, 30), align="center"),
    ]
    for value in values:
        value["style"].update(font_latin="Courier", font_japanese="Courier")
    # 空白でない字形の原点の x（pt）。空白は描かない（描いても読み取りが除く）。
    # space は U+0020（12 pt）と、空白と a の kern（-2 pt）の後。center は行末の U+3000 を除いた U+3000・e・f の幅 36 pt を
    # 中央に置き、U+3000 の後（行末の U+3000 は U+0020 と同じく幅に含めない）。
    origins = {
        "lead": [[40, 52, 64, 76], [52, 64]],
        "trail": [[40, 52, 64, 76], [40, 52]],
        "space": [[310, 322]],
        "center": [[344, 356]],
    }
    path = tmp_path / "spaces.pdf"
    with Document() as document:
        page = document.new_page(width=720, height=540)
        for value in values:
            layout = measure.layout_element(value, fonts)
            lines = [line for line in layout.text.lines if line.text.strip()]
            for line, positions in zip(lines, origins[value["id"]], strict=True):
                visible = [char for char in line.text if not char.isspace()]
                for char, x in zip(visible, positions, strict=True):
                    page.insert_text(
                        (x, layout.block_top + line.baseline), char, fontname="cour", fontsize=20, color=RGB
                    )
        document.save(path)
    (reading,) = read_pdf(path, {"slides": [{"id": "page", "elements": values}]})
    extents = {}
    for value in values:
        actual = tool.actual_record(reading.elements[value["id"]])
        extents[value["id"]] = actual["line_extents_pt"]
        # 縦のインクは Courier の輪郭で決まり偽のフォントの値と違うので、予測の値にして横の範囲だけを比べる。
        ink = measure.layout_element(value, fonts).ink
        actual["ink_pt"][1], actual["ink_pt"][3] = ink.y, ink.y + ink.h
        assert compare_case(case(value, actual), fonts) == [], value["id"]
    assert extents == {
        "lead": [[40, 88], [52, 76]],
        "trail": [[40, 88], [40, 64]],
        "space": [[310, 334]],
        "center": [[344, 368]],
    }


# ---- 校正用の資料 ----


def calibration_scene(tmp_path):
    """校正用の資料を偽のフォントで構築する（Windows のフォントがなくても資料の形を確かめる）。"""
    deck = shutil.copytree(DECK, tmp_path / "deck")
    project = load_project(deck / "deck.toml")
    project.fonts = FakeFonts()
    return project, construct(project)


def test_calibration_deck_has_generic_cases_on_each_page(tmp_path):
    tool = load_tool()
    project, scene = calibration_scene(tmp_path)
    assert project.config["metadata"]["author"] == "slidekit example"
    assert tuple(slide["id"] for slide in scene["slides"]) == PAGES
    tool.check_public(scene)
    cases = [(slide, value) for slide in scene["slides"] for value in tool.text_elements(slide)]
    assert len(cases) >= 150
    assert len({(slide["id"], value["id"]) for slide, value in cases}) == len(cases)
    for slide in scene["slides"]:
        assert tool.text_elements(slide), slide["id"]
    fonts = FakeFonts()
    for slide, value in cases:
        tags = tool.case_tags(slide, value)
        layout = measure.layout_element(value, fonts)
        if "overflow" in tags:
            # 枠より長い語や数式の要素は、計測が text_overflow を報告する。
            assert any(finding["code"] == "text_overflow" for finding in measure.fit_findings(layout)), value["id"]
        if slide["id"] in WRAP_PAGES and value["kind"] == "text":
            assert value["style"]["word_wrap"] is (value["id"] != "nowrap_explicit"), value["id"]
    families = {value["style"][key] for _slide, value in cases for key in ("font_latin", "font_japanese")}
    assert families == {"Arial", "Segoe UI", "Calibri", "Yu Gothic UI", "Meiryo", "BIZ UDGothic"}


def test_kinsoku_cases_put_half_of_checked_character_in_first_line(tmp_path):
    """行頭禁則の例の枠は、前の文字列と確かめる文字の半分が 1 行目に入る幅にする。全角空白も文の途中の幅で計測する
    （行末の全角空白は幅に含めないので、行末で計測すると 0 になり、全角空白が行末にはみ出す形を確かめることができない）。"""
    slides = load_tool(DECK / "slides.py")
    _project, scene = calibration_scene(tmp_path)
    (page,) = [slide for slide in scene["slides"] if slide["id"] == "wrap_ja"]
    areas = {value["id"]: value["area"] for value in page["elements"]}
    for id, before, target, _after, _tag in slides.KINSOKU_START:
        assert areas[f"start_{id}"][2] == pytest.approx(
            10 * len(before) + 5, abs=0.01
        ), id  # 偽のフォントは 1 文字 10 pt


class FamilyFonts(FakeFonts):
    """資料の書体（Arial・Yu Gothic UI）とほかの書体で送り幅が違う偽のフォント。"""

    def advance(self, char, family, bold=False):
        return super().advance(char, family, bold) * (1.0 if family in ("Arial", "Yu Gothic UI") else 1.4)


THEME_OVERRIDES = """
[fonts]
latin = "Calibri"
japanese = "Meiryo"

[defaults]
line_spacing = 1.5
align = "right"
valign = "bottom"
padding_pt = 3

[defaults.box]
padding_pt = 1

[defaults.equation]
word_wrap = true

[roles.body]
size_pt = 30
"""
LAYOUT_KEYS = (
    "size_pt",
    "bold",
    "line_spacing",
    "line_spacing_pt",
    "font_latin",
    "font_japanese",
    "font_math",
    "word_wrap",
    "align",
    "valign",
    "padding_pt",
)


def test_calibration_deck_does_not_depend_on_theme_defaults(tmp_path):
    """要素の書式を明示し、枠の寸法も同じ書体と値で求めるので、テーマの既定を変えても要素の書式と枠は変わらない。"""

    def build(directory, overrides=""):
        deck = shutil.copytree(DECK, directory)
        theme = deck / "theme.toml"
        theme.write_text(theme.read_text(encoding="utf-8") + overrides, encoding="utf-8")
        project = load_project(deck / "deck.toml")
        project.fonts = FamilyFonts(roles=project.theme["fonts"])
        scene = construct(project)
        return project, [
            (
                slide["id"],
                value["id"],
                value["kind"],
                value["area"],
                {key: value["style"].get(key) for key in LAYOUT_KEYS},
            )
            for slide in scene["slides"]
            for value in slide["elements"]
            if value["kind"] in ("text", "box", "equation") and not value["id"].startswith("_theme/")
        ]

    _project, expected = build(tmp_path / "deck")
    project, changed = build(tmp_path / "changed", THEME_OVERRIDES)
    assert project.theme["fonts"]["latin"] == "Calibri" and project.theme["defaults"]["padding_pt"] == 3
    assert changed == expected


def test_calibration_deck_text_is_public():
    tool = load_tool()
    for path in sorted(DECK.iterdir()):
        if path.is_file():
            tool.check_public({"source": path.read_text(encoding="utf-8")})


# ---- 数式の照合データを作る道具（PowerPoint を使わない部分） ----


def test_math_data_tool_measures_marked_inline_and_display_math(tmp_path):
    tool = load_tool(MATH_TOOL)
    pdf = tmp_path / "math.pdf"
    with Document() as document:
        page = document.new_page(width=720, height=540)
        page.insert_text((100, 100), "[", fontname="helv", fontsize=20, color=RGB)
        page.insert_text((110, 100), "x+[", fontname="tiro", fontsize=20, color=RGB)  # 式の中の [ は目印にしない
        page.insert_text((150, 100), "]", fontname="helv", fontsize=20, color=RGB)
        page.draw_rect((100, 200, 160, 230), color=None, fill=RGB)
        document.save(pdf)
    inline = [{"id": "inline_000", "mode": "inline", "expression": "x+[", "area": [96.0, 76.0, 200.0, 40.0]}]
    display = [
        {"id": "display_000", "mode": "display", "expression": r"\frac{a}{b}", "area": [96.0, 190.0, 200.0, 50.0]}
    ]
    items, failed = tool.measure_pdf(pdf, [{"id": "page", "items": inline + display}], math_font="Times-Roman")
    assert failed == []
    first, second = items
    opening = 100 + text_length("[", fontname="helv", fontsize=20)
    assert first["mode"] == "inline" and first["advance_pt"] == pytest.approx(150 - opening, abs=0.001)
    left, right, above, below = first["ink_pt"]
    assert 0 <= left < 10 + 1 and right <= first["advance_pt"] and 0 < above < 20 and below >= 0
    assert second == {
        "expression": r"\frac{a}{b}",
        "mode": "display",
        "ink_width_pt": pytest.approx(60, abs=0.25),
        "ink_height_pt": pytest.approx(30, abs=0.25),
    }


def test_math_data_ink_box_keeps_position_with_fractional_clip(tmp_path):
    """切り抜きの端が画素の途中にあっても、描画範囲の位置は変わらない（画像の原点から pt に戻す）。"""
    tool = load_tool(MATH_TOOL)
    path = tmp_path / "box.pdf"
    with Document() as document:
        document.new_page(width=720, height=540).draw_rect((120, 200, 160, 230), color=None, fill=RGB)
        document.save(path)
    with open_pdf(path) as document:
        for offset in (0.0, 0.06, 0.12, 0.5, 0.99):
            rect = Box(100 + offset, 190 + offset, 200 + offset, 250 + offset)
            assert tool.ink_box(document[0], rect) == pytest.approx((120, 200, 160, 230), abs=1e-6), offset


@pytest.mark.parametrize("mode", ["inline", "display"])
@pytest.mark.parametrize("edge", ["top", "bottom"])
def test_math_data_tool_rejects_vertically_clipped_ink(tmp_path, mode, edge):
    """測定枠より高い式は、切り取った寸法を照合データに書かず、測定不能として返す。"""
    tool = load_tool(MATH_TOOL)
    path = tmp_path / "clipped.pdf"
    area = [96.0, 76.0, 200.0, 40.0]
    top, bottom = (60, 105) if edge == "top" else (90, 130)
    with Document() as document:
        page = document.new_page(width=720, height=540)
        if mode == "inline":
            page.insert_text((100, 100), "[", fontname="helv", fontsize=20, color=RGB)
            page.insert_text((150, 100), "]", fontname="helv", fontsize=20, color=RGB)
        page.draw_rect((115, top, 140, bottom), color=None, fill=RGB)
        document.save(path)
    item = {"id": "clipped", "mode": mode, "expression": "tall", "area": area}
    items, failed = tool.measure_pdf(path, [{"id": "page", "items": [item]}], math_font="Times-Roman")
    assert items == [] and failed == [f"{mode}: tall"]


@pytest.mark.parametrize("edge", ["top", "bottom"])
def test_math_data_tool_rejects_ink_clipped_at_page_edge(tmp_path, edge):
    tool = load_tool(MATH_TOOL)
    path = tmp_path / "page-edge.pdf"
    top, bottom = (-5, 20) if edge == "top" else (80, 105)
    with Document() as document:
        page = document.new_page(width=200, height=100)
        page.draw_rect((50, top, 75, bottom), color=None, fill=RGB)
        document.save(path)
    item = {"id": "clipped", "mode": "display", "expression": "tall", "area": [40, 0, 60, 100]}
    items, failed = tool.measure_pdf(path, [{"id": "page", "items": [item]}])
    assert items == [] and failed == ["display: tall"]


def test_math_data_tool_ignores_disconnected_ink_from_previous_row(tmp_path):
    """計測画像の上端に前の行の描画があっても、空白の行で離れていれば対象の式のクリップとは扱わない。"""
    tool = load_tool(MATH_TOOL)
    path = tmp_path / "previous-row.pdf"
    with Document() as document:
        page = document.new_page(width=720, height=540)
        page.insert_text((100, 100), "[", fontname="helv", fontsize=20, color=RGB)
        page.insert_text((150, 100), "]", fontname="helv", fontsize=20, color=RGB)
        page.draw_rect((115, 68, 140, 74), color=None, fill=RGB)
        page.draw_rect((115, 88, 140, 105), color=None, fill=RGB)
        document.save(path)
    item = {"id": "inline", "mode": "inline", "expression": "x", "area": [96, 76, 200, 40]}
    items, failed = tool.measure_pdf(path, [{"id": "page", "items": [item]}], math_font="Times-Roman")
    assert failed == [] and len(items) == 1
    assert items[0]["ink_pt"][2:] == [12, 5]


@pytest.mark.parametrize("opening,closing", [("]", "]"), ("[", "["), ("]", "[")])
def test_math_data_tool_finds_marker_sides_by_position(tmp_path, monkeypatch, opening, closing):
    """文字コードの対応が逆でも、本文中の角括弧を飛ばし、位置が外側の 2 つを目印にする。"""
    tool = load_tool(MATH_TOOL)
    path = tmp_path / "reversed-marks.pdf"
    with Document() as document:
        page = document.new_page(width=720, height=540)
        for x, text in ((100, "["), (120, "[]"), (150, "]")):
            page.insert_text((x, 100), text, fontname="helv", fontsize=20, color=RGB)
        document.save(path)
    read_chars = tool._chars

    def swapped_chars(page, rect):
        chars = read_chars(page, rect)
        for char in chars:
            if char["origin"][0] == 100:
                char["c"] = opening
            elif char["origin"][0] == 150:
                char["c"] = closing
        return chars[::-1]  # 抽出順にも依存しない。

    monkeypatch.setattr(tool, "_chars", swapped_chars)
    item = {"id": "inline", "mode": "inline", "expression": r"\text{[]}", "area": [96, 76, 200, 40]}
    items, failed = tool.measure_pdf(path, [{"id": "page", "items": [item]}], math_font="Times-Roman")
    assert failed == [] and len(items) == 1
    assert items[0]["advance_pt"] == pytest.approx(50 - text_length("[", fontname="helv", fontsize=20), abs=0.001)


def test_math_data_tool_lays_out_generic_deck(tmp_path):
    tool = load_tool(MATH_TOOL)
    accepted, rejected = tool.supported(("x^2", r"\unknownmacro{x}"))
    assert accepted == ["x^2"] and list(rejected) == [r"\unknownmacro{x}"]
    pages = tool.layout_items([f"x_{number}" for number in range(30)])
    inline = [item for page in pages if page["id"].startswith("inline") for item in page["items"]]
    display = [item for page in pages if page["id"].startswith("display") for item in page["items"]]
    assert len(inline) == len(display) == 30
    for page in pages:
        boxes = [item["area"] for item in page["items"]]
        assert all(x + w <= 960 and y + h <= 540 for x, y, w, h in boxes)
        assert not any(
            a[0] < b[0] + b[2] and b[0] < a[0] + a[2] and a[1] < b[1] + b[3] and b[1] < a[1] + a[3]
            for i, a in enumerate(boxes)
            for b in boxes[i + 1 :]
        )
    deck = tool.write_deck(tmp_path / "work", pages)
    assert 'author = "slidekit example"' in deck.read_text(encoding="utf-8")
    project = load_project(deck)
    project.fonts = FakeFonts()
    scene = construct(project)
    assert len(scene["slides"]) == len(pages)
    assert math.isclose(scene["slides"][0]["elements"][1]["style"]["size_pt"], tool.SIZE_PT)


def test_math_data_tool_help(capsys):
    with pytest.raises(SystemExit) as raised:
        load_tool(MATH_TOOL).main(["--help"])
    assert raised.value.code == 0
    assert "--output" in capsys.readouterr().out
