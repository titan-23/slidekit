"""表示ソフトの PDF から、要素ごとの実際の行と寸法を読み取る。

PDF の検査と、計測の照合データの作成が使う。PDF の読み取りには `slidekit.pdf` を使い、単位は pt とする。
字形は描画順に読み、通常文字は場面の文字列との対応付けで、数式の字形と数式の線は描画順と位置で要素に割り当てる。
描画順と数式の線の位置は、PowerPoint 16.0（build 20430）の PDF で確かめた規則による（要素の順に描き、数式の線は
その構造の最初の字形の 1〜2 つ前に描く）。
回転・反転した要素と、数式を画像にした表示での数式を含む要素は読まない。場面データは変更しない。
書体名の正規化（`_font_name`）、数式の `\\text{...}` の取り出し（`_text_in_math`）、要素の枠（`_area`）、位置で探すときの
1 em（`_em`）は、PDF の検査（`pdf_validation`）も読み込んで使う。照合用の文字の正規化（`_normalize`）は、PPTX と PDF の検査の
`pptx_validation._normalize_text` と同じ規則にする。
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field
import difflib
import math
from pathlib import Path
import re
import statistics
from typing import Any
import unicodedata

# 行とまとまりの判定（行送りに対する比）。
_LINE_RATIO = 0.6  # 基線がこれより離れたら別の行
_MATH_SPLIT_RATIO = 0.9  # 数式の字形の基線がこれ以上変わったら別のまとまり
_NEAR_SIZE_RATIO = 0.12  # 要素の文字サイズに近い字形とみなす差（文字サイズに対する比）
# 書体名の末尾から除く、太字などの語。
_STYLE_WORDS = re.compile(r"(?:[-, ]?(?:regular|bold|italic|oblique|light|semibold|semilight|medium|psmt|mt))+$", re.I)
# 数式の線の割り当て。
_DRAWINGS_BEFORE = 2  # 要素の最初の字形の前に含める描画の数（数式の線は構造の最初の字形の 1〜2 つ前に描かれる。
# それより前は、要素の線の条件に合う描画が続く間だけ含める）
_DRAWINGS_AFTER = 2  # 要素の字形の直後に含める描画の数
_SAME_FRAME_PT = 1.0  # 要素の枠と同じ矩形とみなす差
_COLOR_TOLERANCE = 0.02  # 文字色と同じとみなす差（各成分 0〜1）
_DEFAULT_COLOR = "17212D"  # writer と同じ既定の文字色
_TEXT_IN_MATH = re.compile(r"\\text\s*\{([^{}]*)\}")
# 分数線とみなす描画（横長で細い）：高さが文字サイズのこの比以下で、幅が高さの 2 倍以上。
_BAR_HEIGHT_RATIO = 0.15
# 数式の線を行に含めるとき、線の縦の中心と比べる位置（基線から上へ、文字サイズに対する比。数式の軸の高さ）。
_AXIS_RATIO = 0.25


@dataclass(frozen=True)
class RenderedLine:
    """PDF に描かれた 1 行。"""

    baseline: float  # 行の基線の y
    left: float  # 行の最初の字形の原点の x（行に数式の線があれば、その左端との小さい方）
    right: float  # 行の最後の字形の送り幅の右端（行に数式の線があれば、その右端との大きい方）


@dataclass(frozen=True)
class RenderedText:
    """1 つの要素に割り当てた字形と数式の線から求めた、実際の行と寸法。"""

    element_id: str
    glyphs: int  # 割り当てた字形の数（空白を除く）
    lines: tuple[RenderedLine, ...]  # 上から順
    ink: tuple[float, float, float, float] | None  # 字形の輪郭と数式の線の外接矩形 (x0, y0, x1, y1)。なければ None
    drawings: int  # 含めた数式の線の数
    sizes: tuple[float, ...]  # PDF に書かれた文字サイズ（重複を除き昇順。0.001 pt に丸める）
    fonts: dict[str, str] = field(
        default_factory=dict
    )  # 数式の書体でない字形の、文字ごとの PDF の書体名（最初に描かれたもの）
    math_glyphs: int = 0  # glyphs のうち数式の一部として割り当てた字形の数（数式の書体の字形と \text{...} の字形）


@dataclass(frozen=True)
class SlideReading:
    """1 ページの読み取りの結果。"""

    elements: dict[str, RenderedText]  # 読んだ要素（場面の順）
    unassigned: int  # どの要素にも割り当てなかった字形の数


@dataclass
class _Glyph:
    char: str
    key: str  # 照合に使う文字列（_normalize の結果）
    font: str  # PDF の書体名（部分集合の接頭辞を除く）
    size: float
    x: float  # 原点の x
    baseline: float  # 原点の y
    bbox: tuple[float, float, float, float]  # 横は送り幅、縦は字形の輪郭の範囲
    seqno: int
    math_font: bool  # 数式の書体で描かれた字形か


def read_pdf(path: str | Path, scene: Mapping, *, rasterized_math: bool = False) -> list[SlideReading]:
    """PDF の各ページを場面の同じ順のページとして読む。ページ数が違えば ValueError にする。"""
    from .pdf import open_pdf

    slides = list(scene.get("slides", []))
    with open_pdf(path) as document:
        if len(document) != len(slides):
            raise ValueError(f"PDF と場面のページ数が一致しない: PDF={len(document)}, 場面={len(slides)}")
        return [read_slide(page, slide, rasterized_math=rasterized_math) for page, slide in zip(document, slides)]


def read_slide(
    page: Any, slide: Mapping, *, math_fonts: Iterable[str] | None = None, rasterized_math: bool = False
) -> SlideReading:
    """PDF の 1 ページ（`slidekit.pdf` のページ）を、場面の 1 ページの要素へ割り当てて読む。

    math_fonts は数式の書体名である。省略すると、数式を含む要素の font_math を使う。
    rasterized_math が真なら、数式を含む要素は画像で描かれたものとみなし、読まず、期待する文字にも含めない。
    """
    elements = [element for element in slide.get("elements", []) if isinstance(element, Mapping)]
    contents = [_content(element) for element in elements]
    if math_fonts is None:
        math_fonts = [
            (element.get("style") or {}).get("font_math")
            for element, content in zip(elements, contents)
            if _has_math(content)
        ]
    glyphs = _page_glyphs(page, {_font_name(name) for name in math_fonts if isinstance(name, str) and name})
    present = [not (rasterized_math and _has_math(content)) for content in contents]
    owners, in_math = _assign(elements, contents, present, glyphs)

    lookup = {element.get("id"): element for element in elements}
    readable = [
        present[index]
        and not _transformed(element, lookup)
        and (bool(_normalize(_normal_text(contents[index]))) or _has_math(contents[index]))
        for index, element in enumerate(elements)
    ]
    assigned: dict[int, list[int]] = {}
    for number, owner in enumerate(owners):
        if owner is not None:
            assigned.setdefault(owner, []).append(number)
    spans = sorted(
        (min(glyphs[n].seqno for n in numbers), max(glyphs[n].seqno for n in numbers), index)
        for index, numbers in assigned.items()
    )
    ranges = [
        (low, high, index)
        for low, high, index in spans
        if readable[index] and _area(elements[index]) and any(in_math[n] for n in assigned[index])
    ]
    drawings = _attach_drawings(page, elements, contents, ranges, spans) if ranges else {}

    readings = {}
    for index, element in enumerate(elements):
        if readable[index]:
            numbers = assigned.get(index, [])
            readings[str(element.get("id"))] = _reading(
                element, [glyphs[n] for n in numbers], [in_math[n] for n in numbers], drawings.get(index, [])
            )
    return SlideReading(elements=readings, unassigned=sum(owner is None for owner in owners))


def _normalize(value: str) -> str:
    """照合用に NFKC で正規化し、空白と書式文字（U+200B など）を除く。"""
    return "".join(
        char
        for char in unicodedata.normalize("NFKC", value)
        if not char.isspace() and unicodedata.category(char) != "Cf"
    )


def _font_name(name: str) -> str:
    """PDF と場面の書体名を、部分集合の接頭辞・太字などの語・記号を除いた形にそろえる。"""
    name = re.sub(r"^[A-Z]{6}\+", "", name)
    name = _STYLE_WORDS.sub("", name)
    return re.sub(r"[^\w]", "", unicodedata.normalize("NFKC", name)).lower()


def _content(element: Mapping) -> list[Mapping]:
    value = element.get("content", [])
    if isinstance(value, str):
        return [{"kind": "text", "text": value}]
    return [fragment for fragment in value if isinstance(fragment, Mapping)] if isinstance(value, (list, tuple)) else []


def _has_math(content: list[Mapping]) -> bool:
    return any(fragment.get("kind") == "math" for fragment in content)


def _normal_text(content: list[Mapping]) -> str:
    return "".join(str(fragment.get("text", "")) for fragment in content if fragment.get("kind") != "math")


def _text_in_math(expression: str) -> str:
    """数式のうち、文字用の書体で描かれる \\text{...} の中身を出現順につなげる（- は書き出しと同じくハイフン U+2010 にする）。"""
    return "".join(_TEXT_IN_MATH.findall(expression)).replace("-", "\u2010")


def _area(element: Mapping) -> tuple[float, float, float, float] | None:
    value = element.get("area")
    if (
        isinstance(value, (list, tuple))
        and len(value) == 4
        and all(isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v) for v in value)
    ):
        x, y, w, h = map(float, value)
        return x, y, w, h
    return None


def _transformed(element: Mapping, lookup: Mapping) -> bool:
    """要素か親のまとまりが回転・反転しているか。"""
    seen = set()
    current = element
    while isinstance(current, Mapping) and id(current) not in seen:
        seen.add(id(current))
        style = current.get("style") or {}
        rotation = current.get("rotation", style.get("rotation", 0))
        if isinstance(rotation, (int, float)) and not isinstance(rotation, bool) and rotation % 360:
            return True
        if current.get("flip_h", style.get("flip_h")) or current.get("flip_v", style.get("flip_v")):
            return True
        current = lookup.get(current.get("parent_id"))
    return False


def _number(value: Any, default: float) -> float:
    """正の有限の数ならその値、そうでなければ既定値。"""
    if isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value) and value > 0:
        return float(value)
    return default


def _size(element: Mapping) -> float:
    return _number((element.get("style") or {}).get("size_pt"), 20.0)


def _em(element: Mapping, content: list[Mapping]) -> float:
    """位置で探すときの 1 em。数式の断片の文字サイズが大きければそれを使う。"""
    sizes = [
        _number(fragment["style"].get("size_pt"), 0.0)
        for fragment in content
        if fragment.get("kind") == "math" and isinstance(fragment.get("style"), Mapping)
    ]
    return max([_size(element), *sizes])


def _page_glyphs(page: Any, math_fonts: set[str]) -> list[_Glyph]:
    """描画順の字形。空白・書式文字と、見えない文字を除く。塗りと輪郭で 2 回描いた字形は 1 つにする。

    縦の範囲は字形の輪郭から求めた範囲（`Char.ink`）を使う。ページの外の文字も読む。
    """
    glyphs = []
    previous = None
    for span in page.spans():
        if span.kind == 3 or span.opacity <= 0:
            continue
        signature = [(char.unicode, char.origin) for char in span.chars]
        if span.kind == 1 and previous == (0, span.seqno - 1, signature):
            continue
        previous = (span.kind, span.seqno, signature)
        math_font = _font_name(span.font) in math_fonts
        for char in span.chars:
            if not 0 <= char.unicode <= 0x10FFFF:
                continue
            text = chr(char.unicode)
            key = _normalize(text)
            if not key:
                continue
            x, y = char.origin
            glyphs.append(_Glyph(text, key, span.font, span.size, x, y, tuple(char.ink), span.seqno, math_font))
    return glyphs


def _assign(
    elements: list[Mapping], contents: list[list[Mapping]], present: list[bool], glyphs: list[_Glyph]
) -> tuple[list[int | None], list[bool]]:
    """字形ごとに要素の番号（割り当てなければ None）と、数式の一部かを返す。"""
    owners: list[int | None] = [None] * len(glyphs)
    in_math = [glyph.math_font for glyph in glyphs]
    if not elements:
        return owners, in_math
    # 通常文字：要素の順に連結した期待する文字列と、描画順の字形の文字列を対応付ける。
    # 数式の \text{...} は文字用の書体で描かれるので、期待する文字列に含め、数式の一部として扱う。
    expected, sources = [], []
    for index, content in enumerate(contents):
        if not present[index]:
            continue
        for fragment in content:
            from_math = fragment.get("kind") == "math"
            text = str(fragment.get("text", ""))
            for char in _normalize(_text_in_math(text) if from_math else text):
                expected.append(char)
                sources.append((index, from_math))
    actual, positions = [], []
    for number, glyph in enumerate(glyphs):
        if not glyph.math_font:
            actual.extend(glyph.key)
            positions.extend([number] * len(glyph.key))
    matcher = difflib.SequenceMatcher(None, expected, actual, autojunk=False)
    for block in matcher.get_matching_blocks():
        for offset in range(block.size):
            number = positions[block.b + offset]
            if owners[number] is None:
                owners[number], from_math = sources[block.a + offset]
                in_math[number] = in_math[number] or from_math
    # 数式の書体の字形：前後で対応した字形の要素の範囲のうち、数式を含む要素の枠で探す。
    previous, following = [0] * len(glyphs), [len(elements) - 1] * len(glyphs)
    last = 0
    for number, owner in enumerate(owners):
        last = owner if owner is not None else last
        previous[number] = last
    last = len(elements) - 1
    for number in range(len(glyphs) - 1, -1, -1):
        last = owners[number] if owners[number] is not None else last
        following[number] = last
    searchable = [
        present[index] and _has_math(content) and _area(elements[index]) is not None
        for index, content in enumerate(contents)
    ]
    for number, glyph in enumerate(glyphs):
        if owners[number] is None and glyph.math_font:
            low, high = sorted((previous[number], following[number]))
            candidates = [index for index in range(low, high + 1) if searchable[index]]
            owners[number] = _locate(glyph, candidates, elements, contents)
    return owners, in_math


def _locate(glyph: _Glyph, candidates: list[int], elements: list[Mapping], contents: list[list[Mapping]]) -> int | None:
    """字形の中心（横は外接矩形の中央、縦は原点）を枠に含む最初の要素。なければ枠を 1 em 広げて探し、広げた枠に中心を
    含む要素のうち、中心から元の枠までの距離（横と縦の和）が最も小さい要素にする（同じなら候補の順で先の要素）。

    下揃えの要素の分母や下の添字は、枠の下端をわずかに出て描かれることがある（PowerPoint 16.0.20430.20182 で 0.03〜0.14 pt）。
    広げた枠は隣の要素の枠と重なるので、候補の順で最初の要素にすると、隣の要素に割り当ててしまう。
    """
    center = (glyph.bbox[0] + glyph.bbox[2]) / 2
    for index in candidates:
        x, y, w, h = _area(elements[index])
        if x <= center <= x + w and y <= glyph.baseline <= y + h:
            return index
    best, nearest = None, math.inf
    for index in candidates:
        x, y, w, h = _area(elements[index])
        margin = _em(elements[index], contents[index])
        if not (x - margin <= center <= x + w + margin and y - margin <= glyph.baseline <= y + h + margin):
            continue
        distance = max(x - center, center - x - w, 0.0) + max(y - glyph.baseline, glyph.baseline - y - h, 0.0)
        if distance < nearest:
            best, nearest = index, distance
    return best


def _attach_drawings(
    page: Any,
    elements: list[Mapping],
    contents: list[list[Mapping]],
    ranges: list[tuple[int, int, int]],
    spans: list[tuple[int, int, int]] = (),
) -> dict[int, list[tuple[float, float, float, float]]]:
    """数式の線（根号・分数線・伸びた括弧）を、描画順と位置と色で要素へ割り当てる。

    ranges は数式の字形を持つ要素ごとの（最初の seqno、最後の seqno、要素の番号）、spans は字形を持つすべての要素の
    同じ値で、どちらも seqno の順に並べる。描画の seqno が、要素の字形の範囲の中、最初の字形の 2 つ前から
    （前の要素の最後の字形より後に限る）、最後の字形の 2 つ後まで（次の数式の要素の最初の字形より前に限る）のどれかに
    あれば、その要素の線の候補にする。候補が複数あれば、範囲の中、前、後の順に選ぶ（PowerPoint は数式の線を、その構造の
    最初の字形の 1〜2 つ前に描き、後には描かない）。入れ子の構造では外側の線がさらに前に描かれる（`\\left(\\frac{\\sqrt{x}}{y}\\right)`
    の括弧は 3 つ前）ので、2 つ前より前も、その要素の線の条件に合う描画が続く間は、前の要素の最後の字形の後までさかのぼって
    候補にする。不透明度が 0 の描画（説明のある要素に PowerPoint が加える見えない矩形など）は除く。
    """
    frames = [area for area in map(_area, elements) if area is not None]
    painted = [area for area in map(_painted_range, elements) if area is not None]
    colors = {index: _colors(elements[index], contents[index]) for _, _, index in ranges}
    # 描画ごとの（seqno、外接矩形、線幅で広げる幅）。数式の線になりえないもの（図形の枠、塗った範囲の中、色のない描画）は
    # None にする。不透明度が 0 の描画は読まない。
    lines: dict[int, tuple[tuple[float, float, float, float], tuple[float, ...], float] | None] = {}
    for drawing in page.drawings():
        seqno = drawing.seqno
        kind = drawing.kind
        opacity = drawing.fill_opacity if kind in ("f", "fs") else drawing.stroke_opacity
        if opacity <= 0:
            continue
        box = tuple(drawing.rect)
        rgb = _rgb(drawing.fill if kind in ("f", "fs") else drawing.color)
        # 図形の枠や線・矢印は、数式の線として扱わない。
        if (
            rgb is None
            or any(
                all(abs(a - b) <= _SAME_FRAME_PT for a, b in zip(box, (fx, fy, fx + fw, fy + fh)))
                for fx, fy, fw, fh in frames
            )
            or any(a <= box[0] and b <= box[1] and box[2] <= c and box[3] <= d for a, b, c, d in painted)
        ):
            lines[seqno] = None
            continue
        # 線で描いた分数線などは、線幅の半分までを塗る範囲に含める。
        width = drawing.width if kind in ("s", "fs") else None
        lines[seqno] = (box, rgb, float(width) / 2 if isinstance(width, (int, float)) and width > 0 else 0.0)

    def fits(line, index: int, order: int) -> bool:
        box, rgb, _grow = line
        x, y, w, h = _area(elements[index])
        margin = _em(elements[index], contents[index]) * (0.5 if order == 2 else 1.0)
        return (
            x - margin <= box[0]
            and box[2] <= x + w + margin
            and y - margin <= box[1]
            and box[3] <= y + h + margin
            and any(max(abs(a - b) for a, b in zip(rgb, color)) <= _COLOR_TOLERANCE for color in colors[index])
        )

    windows = []
    for position, (low, high, index) in enumerate(ranges):
        upper = ranges[position + 1][0] if position + 1 < len(ranges) else math.inf
        previous = max((end for start, end, other in spans if other != index and end < low), default=-math.inf)
        first = max(low - _DRAWINGS_BEFORE, previous + 1)
        # 2 つ前より前は、その要素の線の条件に合う描画が続く間だけ広げる（描画でない番号は飛ばす）。
        number = first - 1
        while number > previous and number >= min(lines, default=0):
            if number in lines:
                if lines[number] is None or not fits(lines[number], index, 1):
                    break
                first = number
            number -= 1
        windows.append((low, high, index, first, min(high + _DRAWINGS_AFTER, upper - 1)))
    result: dict[int, list[tuple[float, float, float, float]]] = {}
    for seqno, line in lines.items():
        if line is None:
            continue
        candidates = []
        for position, (low, high, index, first, last) in enumerate(windows):
            if low <= seqno <= high:
                candidates.append((0, position, index))
            elif first <= seqno < low:
                candidates.append((1, position, index))
            elif high < seqno <= last:
                candidates.append((2, position, index))
        for order, _position, index in sorted(candidates):
            if fits(line, index, order):
                box, _rgb_value, grow = line
                result.setdefault(index, []).append((box[0] - grow, box[1] - grow, box[2] + grow, box[3] + grow))
                break
    return result


def _painted_range(element: Mapping) -> tuple[float, float, float, float] | None:
    """線・矢印の描画範囲（端点の外接矩形を線幅と矢じりの分だけ広げた範囲）。"""
    if element.get("kind") not in ("line", "arrow"):
        return None
    points = [element.get("start"), element.get("end")]
    if not all(
        isinstance(point, (list, tuple))
        and len(point) == 2
        and all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in point)
        for point in points
    ):
        area = _area(element)
        if area is None:
            return None
        points = [(area[0], area[1]), (area[0] + area[2], area[1] + area[3])]
    width = _number((element.get("style") or {}).get("line_width_pt"), 1.0)
    margin = max(width / 2, 2 * width if element.get("kind") == "arrow" else 0.0) + 1.0
    (x0, y0), (x1, y1) = points
    return min(x0, x1) - margin, min(y0, y1) - margin, max(x0, x1) + margin, max(y0, y1) + margin


def _colors(element: Mapping, content: list[Mapping]) -> list[tuple[float, float, float]]:
    """数式の線に使われうる色（要素の文字色と、数式の断片の色）。"""
    values = [(element.get("style") or {}).get("color") or _DEFAULT_COLOR]
    values += [
        fragment["style"].get("color")
        for fragment in content
        if fragment.get("kind") == "math"
        and isinstance(fragment.get("style"), Mapping)
        and fragment["style"].get("color")
    ]
    colors = []
    for value in values:
        text = str(value).removeprefix("#")
        if re.fullmatch(r"[0-9a-fA-F]{6}", text):
            colors.append(tuple(int(text[i : i + 2], 16) / 255 for i in (0, 2, 4)))
    return colors


def _rgb(value: Any) -> tuple[float, ...] | None:
    """色（灰色・RGB・CMYK の 0〜1 の値の列）を RGB にする。"""
    if not isinstance(value, (list, tuple)) or not all(isinstance(v, (int, float)) for v in value):
        return None
    if len(value) == 1:
        return (float(value[0]),) * 3
    if len(value) == 3:
        return tuple(map(float, value))
    if len(value) == 4:
        c, m, y, k = map(float, value)
        return tuple(1 - min(1.0, v + k) for v in (c, m, y))
    return None


def _reading(
    element: Mapping, glyphs: list[_Glyph], in_math: list[bool], drawings: list[tuple[float, float, float, float]]
) -> RenderedText:
    boxes = [glyph.bbox for glyph in glyphs] + drawings
    ink = (
        (
            min(box[0] for box in boxes),
            min(box[1] for box in boxes),
            max(box[2] for box in boxes),
            max(box[3] for box in boxes),
        )
        if boxes
        else None
    )
    fonts: dict[str, str] = {}
    for glyph in glyphs:
        if not glyph.math_font:
            fonts.setdefault(glyph.char, glyph.font)
    sizes = tuple(sorted({round(glyph.size, 3) for glyph in glyphs}))
    return RenderedText(
        element_id=str(element.get("id")),
        glyphs=len(glyphs),
        lines=_lines(element, glyphs, in_math, drawings),
        ink=ink,
        drawings=len(drawings),
        sizes=sizes,
        fonts=fonts,
        math_glyphs=sum(in_math),
    )


def _lines(
    element: Mapping, glyphs: list[_Glyph], in_math: list[bool], drawings: list[tuple[float, float, float, float]] = ()
) -> tuple[RenderedLine, ...]:
    """字形を行にまとめる。分数の分子・分母や添字は、近い行に含める。数式の線は近い行の範囲に含める。"""
    style = element.get("style") or {}
    size = _size(element)
    pitch = _number(style.get("line_spacing_pt"), 1.2 * _number(style.get("line_spacing"), 1.15) * size)
    # 通常文字は、基線が行送りの 0.6 倍より離れたら別の行としてまとめる。
    groups: list[list[_Glyph]] = []
    for glyph in sorted((g for g, flag in zip(glyphs, in_math) if not flag), key=lambda g: g.baseline):
        if groups and glyph.baseline - groups[-1][0].baseline <= _LINE_RATIO * pitch:
            groups[-1].append(glyph)
        else:
            groups.append([glyph])
    lines = [(statistics.median(g.baseline for g in group), group) for group in groups]
    # 数式の字形は、描画順に続く字形をまとまりにし、基線が行送りの 0.9 倍以上変わったら分ける。
    # ただし、間に分数線がある分子と分母と、伸びた括弧が縦に覆って横に挟む 2 段（独立した数式の \binom など。分数線がない）は
    # 分けない（独立した数式では、2 段の基線の差が行送りに近いか超える。PowerPoint が改行した数式の次の行とは、
    # 分数線か挟む括弧の有無で区別する）。
    zones: list[list[_Glyph]] = []
    current: list[_Glyph] = []
    for glyph, flag in zip(glyphs, in_math):
        if not flag:
            if current:
                zones.append(current)
            current = []
            continue
        if (
            current
            and abs(glyph.baseline - _main_baseline(current, size)) >= _MATH_SPLIT_RATIO * pitch
            and not _fraction_bar(drawings, current, glyph, size)
            and not _enclosing_delimiter(drawings, current, glyph, size)
        ):
            zones.append(current)
            current = []
        current.append(glyph)
    if current:
        zones.append(current)
    # まとまりの基線がどの行からも 0.6 倍より離れていれば、別の行として扱う。
    for zone in zones:
        main = _main_baseline(zone, size)
        nearest = min(range(len(lines)), key=lambda i: abs(lines[i][0] - main), default=None)
        if nearest is not None and abs(lines[nearest][0] - main) <= _LINE_RATIO * pitch:
            lines[nearest][1].extend(zone)
        else:
            lines.append((main, list(zone)))
    rows = [
        [baseline, min(g.x for g in members), max(g.bbox[2] for g in members)]
        for baseline, members in sorted(lines, key=lambda line: line[0])
    ]
    # 数式の線（根号の記号、分数線、伸びた括弧）は字形でないので、縦の中心が数式の軸に最も近い行の範囲に含める。
    for x0, y0, x1, y1 in drawings:
        if not rows:
            break
        center = (y0 + y1) / 2
        row = min(rows, key=lambda item: abs(item[0] - _AXIS_RATIO * size - center))
        if abs(row[0] - _AXIS_RATIO * size - center) <= pitch:
            row[1], row[2] = min(row[1], x0), max(row[2], x1)
    return tuple(RenderedLine(baseline=baseline, left=left, right=right) for baseline, left, right in rows)


def _enclosing_delimiter(drawings, zone: list[_Glyph], glyph: _Glyph, size: float) -> bool:
    """まとまりと字形を、上の基線より上から下の基線より下まで縦に覆い、両方を横に挟む描画（\\binom などの伸びた括弧）があるか。"""
    upper = _main_baseline(zone, size)
    low, high = sorted((upper, glyph.baseline))
    for x0, y0, x1, y1 in drawings:
        if (
            y0 < low
            and y1 > high
            and x0 <= glyph.bbox[0]
            and glyph.bbox[2] <= x1
            and any(x0 <= g.bbox[0] and g.bbox[2] <= x1 for g in zone)
        ):
            return True
    return False


def _fraction_bar(drawings, zone: list[_Glyph], glyph: _Glyph, size: float) -> bool:
    """まとまり（分子）と字形（分母）の間に、両方と横の範囲が重なる横長の細い描画（分数線）があるか。"""
    upper = _main_baseline(zone, size)
    low, high = sorted((upper, glyph.baseline))
    for x0, y0, x1, y1 in drawings:
        height = y1 - y0
        if height > _BAR_HEIGHT_RATIO * size or x1 - x0 < 2 * height or not low < (y0 + y1) / 2 < high:
            continue
        if x0 <= glyph.bbox[2] and glyph.bbox[0] <= x1 and any(x0 <= g.bbox[2] and g.bbox[0] <= x1 for g in zone):
            return True
    return False


def _main_baseline(zone: list[_Glyph], size: float) -> float:
    """まとまりの基線：要素の文字サイズに近い字形の基線の最頻値（同数なら中央の値）。"""
    values = [round(g.baseline, 1) for g in zone if abs(g.size - size) < _NEAR_SIZE_RATIO * size] or [
        round(g.baseline, 1) for g in zone
    ]
    counts = Counter(values).most_common()
    best = sorted(value for value, count in counts if count == counts[0][1])
    return best[len(best) // 2]
