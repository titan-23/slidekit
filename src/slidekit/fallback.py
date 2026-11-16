"""数式の互換画像の埋め込みと、互換表示の確認用複製の作成。

互換画像を付ける生成の流れは次のとおりである。

1. `prepare_final_form`：中間 PPTX の数式の図形を `mc:AlternateContent` に置き換え、互換画像を仮の画像にした最終形の PPTX を作る。
2. 最終形の PPTX を PowerPoint で 1 回だけ変換する。PowerPoint は `mc:Choice` の編集可能な数式を描くので、この PDF を標準表示の確認に使う。
3. `finalize_math`：その PDF から互換画像を切り出し、仮の画像の中身だけを差し替えて最終 PPTX を書く。XML は変えない。
4. `compare_package_structure`：最終形の PPTX と最終 PPTX が、互換画像の中身を除いて同一であることを確かめる。
5. `prepare_fallback`：互換表示の確認のために、最終 PPTX の互換画像を通常の図形へ展開した確認用複製を作る。

互換画像の図形は数式の図形の矩形を `MAX_MARGIN_PT` 広げた範囲（ページ内）に置く。切り出す範囲は変換した PDF から決まるので、
画像のうち切り出す範囲の外側は、ほぼ透明（不透明度 `ALPHA_FLOOR`）にする。XML を変換の前に決めるためである。
"""

from __future__ import annotations

from collections import Counter
from copy import deepcopy
from functools import lru_cache
import hashlib
import io
import math
from pathlib import Path
import posixpath
import re
import struct
from typing import Any, Iterable
from urllib.parse import unquote
import zlib

from lxml import etree

from .files import file_hash
from .ooxml import EMU_PER_PT, NS, read_package, slide_order, tag, write_package, xml_bytes

# 切り出す範囲を数式の図形の矩形から広げる上限（pt）。互換画像の図形もこの範囲に置く。
MAX_MARGIN_PT = 16.0
# 自身の文字と塗りを判定するときに、数式の図形の矩形を広げる幅（pt）。
TOLERANCE_PT = 0.5
# 仮の画像の色。標準表示で間違って互換画像が描かれたときに目立つ色にする。
PLACEHOLDER_COLOR = (255, 0, 255)
# 切り出す範囲の外側で、透明度を段階的に下げる画素の数。
FEATHER_PIXELS = 2
# 切り出す範囲の外側の画素の不透明度（0〜255）。PowerPoint は PDF へ書き出すときに完全に透明な画素（不透明度 0）の色を黒にする
# ので、0 にすると、PDF を縮小して描いたときに不透明な画素と黒が混ざり、境界が灰色の線として見える（白い背景の上の塗りのない
# 文字枠や楕円の外接矩形の境界で、2 倍の確認画像の明度 235）。1/255 の不透明度は見た目を変えず、画素の色を残す。
ALPHA_FLOOR = 1
_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
_EPSILON = 1e-6
# 数式の線の塗りの色を自身の文字の色と同じとみなす差（各成分 0〜1）。
_COLOR_TOLERANCE = 0.02
# ほかの要素の描画とみなすときに、その要素の範囲を広げる幅（pt）。枠線は線幅の半分を加える。
_REGION_MARGIN_PT = 1.0
# 線・矢印の範囲は、さらに線幅のこの倍数だけ広げる。矢印の先端の幅は線幅の 5 倍までで、線の両側へ半分ずつ出る。
_LINE_END_FACTOR = 3.0
# 線幅が分からない線・矢印（雛形の書式の線）に使う線幅（pt）。
_DEFAULT_LINE_PT = 1.0
# 始点と終点を結ぶ線として扱う図形の形（`a:prstGeom` の `prst`）。
_LINE_GEOMETRIES = frozenset({"line", "straightConnector1"})
# 吹き出しの本体の位置の式（writer._custom_geometry の a:gd の fmla。`*/ w 定数 W`）。
_BODY_GUIDE = re.compile(r"\*/ ([wh]) (-?\d+) (\d+)")
# 範囲を求める図形の要素。
_SHAPE_TAGS = frozenset(tag("p", name) for name in ("sp", "cxnSp", "pic", "graphicFrame"))
# 関係の識別子を指す属性（`r:` の名前空間）の値。ページ全体と、`mc:Fallback` の中だけ。
_RELATIONSHIP_VALUES = etree.XPath("//@*[namespace-uri()=$namespace]")
_FALLBACK_RELATIONSHIP_VALUES = etree.XPath(
    "//mc:Fallback//@*[namespace-uri()=$namespace]", namespaces={"mc": NS["mc"]}
)


def fallback_required(scene: dict, project: Any) -> bool:
    """場面に数式があり、`math.fallback` が `required` のときに真を返す。"""
    from .validators import _has_math

    return project.config.get("math", {}).get("fallback", "off") == "required" and any(
        _has_math(element) for slide in scene.get("slides", []) for element in slide.get("elements", [])
    )


def outline_width_pt(shape) -> float:
    """図形の枠線の幅（pt）。`p:spPr/a:ln` があり `a:noFill` でなければ `w`（既定 12700 EMU）、なければ 0。"""
    outline = shape.find("p:spPr/a:ln", NS)
    if outline is None or outline.find("a:noFill", NS) is not None:
        return 0.0
    return int(outline.get("w", str(EMU_PER_PT))) / EMU_PER_PT


def _math_rectangle(shape):
    from .pdf import Box

    transform = shape.find("p:spPr/a:xfrm", NS)
    if transform is None:
        raise ValueError("数式の図形に位置がない")
    if int(transform.get("rot", "0")) % 21600000 or any(
        transform.get(key) in {"1", "true"} for key in ("flipH", "flipV")
    ):
        raise ValueError("数式を含む図形は回転・反転できない")
    offset, extent = transform.find("a:off", NS), transform.find("a:ext", NS)
    if offset is None or extent is None:
        raise ValueError("数式の図形に位置または寸法がない")
    x, y = int(offset.get("x")) / EMU_PER_PT, int(offset.get("y")) / EMU_PER_PT
    width, height = int(extent.get("cx")) / EMU_PER_PT, int(extent.get("cy")) / EMU_PER_PT
    if min(width, height) <= 0:
        raise ValueError("数式の図形の幅・高さは正にする")
    return Box(x, y, x + width, y + height)


def _body_rectangle(shape, rectangle):
    """吹き出しの本体の矩形。a:custGeom の式 bL・bT・bR・bB（`*/ w 定数 W` の形）と a:xfrm の矩形 rectangle から求める。

    式のない図形（既定図形と、slidekit が書かない自由図形）は rectangle をそのまま返す。
    """
    from .pdf import Box

    guides = {
        node.get("name"): (node.get("fmla") or "").strip()
        for node in shape.findall("p:spPr/a:custGeom/a:gdLst/a:gd", NS)
    }
    values = {}
    for name in ("bL", "bT", "bR", "bB"):
        match = _BODY_GUIDE.fullmatch(guides.get(name, ""))
        if match is None or int(match.group(3)) <= 0:
            return rectangle
        axis, value, base = match.group(1), int(match.group(2)), int(match.group(3))
        values[name] = value / base * (rectangle.width if axis == "w" else rectangle.height)
    body = Box(
        rectangle.x0 + values["bL"],
        rectangle.y0 + values["bT"],
        rectangle.x0 + values["bR"],
        rectangle.y0 + values["bB"],
    )
    return body if not body.is_empty and _contains(rectangle, body, 0.01) else rectangle


def _contains(outer, inner, tolerance: float = _EPSILON) -> bool:
    return (
        inner.x0 >= outer.x0 - tolerance
        and inner.y0 >= outer.y0 - tolerance
        and inner.x1 <= outer.x1 + tolerance
        and inner.y1 <= outer.y1 + tolerance
    )


def _expanded(rectangle, amount: float):
    from .pdf import Box

    return Box.of(rectangle).expanded(amount)


def _snap_outward(rectangle, scale: float):
    """1 / scale pt の格子へ外向きにそろえる。"""
    from .pdf import Box

    low = [math.floor(value * scale + _EPSILON) / scale for value in (rectangle.x0, rectangle.y0)]
    high = [math.ceil(value * scale - _EPSILON) / scale for value in (rectangle.x1, rectangle.y1)]
    return Box(*low, *high)


def _frame(rectangle, page_rect, scale: float):
    """互換画像の図形の範囲。数式の図形の矩形を MAX_MARGIN_PT 広げ、ページ内で格子にそろえる。"""
    return _snap_outward(_expanded(rectangle, MAX_MARGIN_PT) & page_rect, scale) & page_rect


def _points(rectangle) -> list[float]:
    return [round(value, 4) for value in rectangle]


def _margins(rectangle, crop) -> list[float]:
    """切り出す範囲が矩形から広がった幅（pt）。[左, 上, 右, 下]。"""
    return [
        round(value, 4)
        for value in (rectangle.x0 - crop.x0, rectangle.y0 - crop.y0, crop.x1 - rectangle.x1, crop.y1 - rectangle.y1)
    ]


def _same_color(a, b) -> bool:
    return (
        isinstance(a, (tuple, list))
        and isinstance(b, (tuple, list))
        and len(a) == len(b)
        and all(abs(x - y) <= _COLOR_TOLERANCE for x, y in zip(a, b))
    )


def _segment_distance(x: float, y: float, segment) -> float:
    """点 (x, y) から線分（x0, y0, x1, y1）までの距離。"""
    x0, y0, x1, y1 = segment
    dx, dy = x1 - x0, y1 - y0
    length = dx * dx + dy * dy
    t = 0.0 if length <= 0 else min(1.0, max(0.0, ((x - x0) * dx + (y - y0) * dy) / length))
    return math.hypot(x - x0 - t * dx, y - y0 - t * dy)


class _Region:
    """ほかの要素の範囲。`neighbors` の 1 項目（`{"id", "box_pt", "tolerance_pt", "segment_pt"}`）から作る。"""

    def __init__(self, item):
        from .pdf import Box

        self.id = item["id"]
        self.box = Box.of(item["box_pt"])
        self.tolerance = float(item["tolerance_pt"])
        self.segment = tuple(item["segment_pt"]) if item.get("segment_pt") else None
        self.reach = self.box.expanded(self.tolerance)

    def drew(self, box, probe) -> bool:
        """外接矩形 box の描画を、この要素の描画とみなせるか。

        次のどちらかなら真とする。
        - box と要素の範囲が、互いに相手を `tolerance` 広げた範囲に収まる（1 つの塗りで描いた線・矢印の全体や、箱の塗り）。
        - 線の形の要素で、box が数式の矩形を広げた範囲 probe に収まらず、box の 4 隅が線分から `tolerance` 以内にある
          （別に描いた矢印の先端や破線の一部）。probe に収まる描画は数式の括弧などと区別できないので、この規則を使わない。
        """
        if not _contains(self.reach, box):
            return False
        if _contains(box.expanded(self.tolerance), self.box):
            return True
        return (
            self.segment is not None
            and not _contains(probe, box)
            and all(
                _segment_distance(x, y, self.segment) <= self.tolerance
                for x in (box.x0, box.x1)
                for y in (box.y0, box.y1)
            )
        )

    def contains(self, rectangle) -> bool:
        """線の形でない要素が rectangle を含むか（数式を載せた箱など）。"""
        return self.segment is None and _contains(self.reach, rectangle)


def _outside(box, bound) -> list:
    """box のうち bound の外の部分。左・右・上・下の帯（重なってもよい）の列。"""
    from .pdf import Box

    parts = [
        Box(box.x0, box.y0, min(box.x1, bound.x0), box.y1),
        Box(max(box.x0, bound.x1), box.y0, box.x1, box.y1),
        Box(box.x0, box.y0, box.x1, min(box.y1, bound.y0)),
        Box(box.x0, max(box.y0, bound.y1), box.x1, box.y1),
    ]
    return [part for part in parts if not part.is_empty]


def crop_rectangle(
    page,
    rectangle,
    *,
    outline_pt: float,
    traces: list[dict],
    drawings: list[dict],
    scale: float,
    max_margin_pt: float = MAX_MARGIN_PT,
    tolerance_pt: float = TOLERANCE_PT,
    math_fonts: Iterable[str] = (),
    others: Iterable = (),
    neighbors: Iterable[dict] = (),
    shape_rectangle=None,
) -> tuple[Any, dict]:
    """数式の図形自身の描画と枠線を含む切り出し範囲を、変換した PDF のページから求める。

    `rectangle` は自身の文字を判定する矩形で、`shape_rectangle` は図形の a:xfrm の矩形である（省くと rectangle）。
    吹き出しでは rectangle を本体にし、shape_rectangle をしっぽまでの範囲にする。切り出しの範囲・上限・枠線・図形の枠の判定は
    shape_rectangle で、1・2 の自身の文字の判定は rectangle で行う（しっぽの側の帯にある隣の要素の文字を含めないため）。

    `traces` は `page.spans()`、`drawings` は `page.drawings()` の結果（`slidekit.pdf` のページ。ページごとに 1 回求めて使い回す）。
    PowerPoint は図形ごとに続けて描き、数式の線（伸びる括弧・根号・分数線）はその構造の最初の字形より前に描く。
    `neighbors` は同じページのほかの要素の範囲（`{"id", "box_pt", "tolerance_pt", "segment_pt"}` の列。`_shape_regions` の形）で、
    その要素の描画とみなせる塗り（`_Region.drew`。1 つの塗りで描いた矢印の全体など）は 3・4 で含めない。
    1. 自身の文字：基点が矩形の中にある文字と、基点が矩形を `tolerance_pt` 広げた範囲にあり送り幅の横の中央もその範囲に
       ある文字。外接矩形を含める（枠の辺に接して始まる隣の要素の最初の文字は、中央が外にあるので含めない）。
    2. `math_fonts`（数式の書体名。`pdf_layout._font_name` の形）を渡すと、自身の文字の後に続く数式の書体の文字列も、
       外接矩形が矩形を `max_margin_pt` 広げた範囲（上限）に掛かり、基点がほかの数式の図形の矩形（`others`）にない間は
       自身の文字とする。枠の外へはみ出した分母などが当たる。
    3. 自身の文字を含む文字列の描画順の最小 low と最大 high の間に描いた塗り（`type` が `"f"`）のうち、上限に掛かり
       数式の図形の枠でないものを含める。high の後、次の他の文字列より前に描いた塗りは、中心が範囲内で面積が矩形以下のもの
       だけを含める。
    4. low の直前から描画順にさかのぼり、数式の線の条件に合う塗り（自身の文字の色で、上限に掛かり、数式の図形の枠でも
       ほかの要素の描画でもないもの）が続く間は、矩形に掛かるものを含める（括弧・根号）。他の文字列か、数式の線の条件に合わない
       描画で止める（図形の塗りと枠線は線より前に描かれ、ほかの要素の描画より前は別の図形である）。矩形に掛からない塗りは、
       直前に描いた別の図形（文字のない飾りの線など）と区別できないので含めない。
    5. 枠線があれば、矩形を枠線の幅の半分と 1 / scale pt だけ広げる（外側の半分と縁のぼかし）。
    6. 1〜5 を上限とページで切り詰め、1 / scale pt の格子へ外向きにそろえる。上限を超えた描画があれば `clipped` を真にする。
       そのときは、上限の外の部分に範囲が掛かるほかの要素（数式を載せた箱など、数式の矩形を含む要素を除く）の識別子を
       `nearby_elements` に入れる。上限を超えた描画が、その要素のものかもしれない。
    不透明度が 0 の文字と描画（説明のある要素に PowerPoint が加える読み上げ用の文字など）は読まない。
    戻り値は範囲と `{"margin_pt": [左, 上, 右, 下], "clipped": 切り詰めたか, "outline_pt": 枠線の幅,
    "nearby_elements": [識別子]}`。
    """
    from .pdf import Box
    from .pdf_layout import _font_name

    rectangle = Box.of(rectangle)
    outer = rectangle if shape_rectangle is None else Box.of(shape_rectangle) | rectangle
    probe = _expanded(rectangle, tolerance_pt)
    limit = _expanded(outer, max_margin_pt) & page.rect
    blocked = [_expanded(Box.of(other), tolerance_pt) for other in others]
    regions = [_Region(item) for item in neighbors]
    math_fonts = set(math_fonts)

    def inside(area, x: float, y: float) -> bool:
        return area.x0 <= x <= area.x1 and area.y0 <= y <= area.y1

    def own_character(character) -> bool:
        (x, y), box = character.origin, character.bbox
        return inside(rectangle, x, y) or (inside(probe, x, y) and probe.x0 <= (box[0] + box[2]) / 2 <= probe.x1)

    def frame_like(box) -> bool:
        return all(abs(a - b) <= 1.0 for a, b in zip(box, outer))

    def foreign(box) -> bool:
        return any(region.drew(box, probe) for region in regions)

    spans = [span for span in traces if span.kind != 3 and span.opacity > 0]
    fills = sorted(
        (drawing for drawing in drawings if drawing.kind == "f" and drawing.fill_opacity > 0),
        key=lambda drawing: drawing.seqno,
    )
    wanted = outer
    # 含めた描画と文字の範囲。上限を超えたときに、上限の外の部分に掛かる要素を探す。
    included = []

    def include(box) -> None:
        nonlocal wanted
        box = Box.of(box)
        wanted |= box
        included.append(box)

    own = {span.seqno: span for span in spans if any(own_character(character) for character in span.chars)}
    if own:
        for span in own.values():
            for character in span.chars:
                if own_character(character):
                    include(character.bbox)
        low, high = min(own), max(own)
        stop = math.inf
        for span in sorted((span for span in spans if span.seqno > high), key=lambda span: span.seqno):
            continued = (
                math_fonts
                and _font_name(span.font) in math_fonts
                and any(Box.of(character.bbox).intersects(limit) for character in span.chars)
                and not any(inside(area, *character.origin) for area in blocked for character in span.chars)
            )
            if not continued:
                stop = span.seqno
                break
            own[span.seqno] = span
            high = span.seqno
            for character in span.chars:
                include(character.bbox)
        colors = [span.color for span in own.values()]
        area = rectangle.width * rectangle.height
        for drawing in fills:
            box = drawing.rect
            if low <= drawing.seqno <= high:
                if box.intersects(limit) and not frame_like(box) and not foreign(box):
                    include(box)
            elif high < drawing.seqno < stop:
                if (
                    inside(probe, (box.x0 + box.x1) / 2, (box.y0 + box.y1) / 2)
                    and box.width * box.height <= area
                    and not foreign(box)
                ):
                    include(box)
        # low の直前に続く数式の線。矩形に掛かるものを含め、矩形に掛からず上限に掛かるもの（分数線か、すぐ前に描いた
        # 別の図形の飾りの線）は含めずにさかのぼる。ほかの要素の描画（矢印など）で止める。
        items = {span.seqno: None for span in spans if span.seqno < low}
        items.update({drawing.seqno: drawing for drawing in drawings if drawing.seqno < low})
        for number in sorted(items, reverse=True):
            drawing = items[number]
            if drawing is None:
                break
            opacity = drawing.fill_opacity if drawing.kind == "f" else drawing.stroke_opacity
            if opacity <= 0:
                continue
            box = drawing.rect
            if (
                drawing.kind != "f"
                or frame_like(box)
                or not box.intersects(limit)
                or not any(_same_color(drawing.fill, color) for color in colors)
                or foreign(box)
            ):
                break
            if box.intersects(probe):
                include(box)
    if outline_pt > 0:
        include(_expanded(outer, outline_pt / 2 + 1 / scale))
    clipped = not _contains(limit, wanted)
    nearby = []
    if clipped:
        beyond = [part for box in included for part in _outside(box, limit)]
        nearby = [
            region.id
            for region in regions
            if not region.contains(outer) and any(region.reach.intersects(part) for part in beyond)
        ]
    crop = _snap_outward(wanted & limit, scale) & page.rect
    return crop, {
        "margin_pt": _margins(outer, crop),
        "clipped": clipped,
        "outline_pt": outline_pt,
        "nearby_elements": list(dict.fromkeys(nearby)),
    }


def _crop_png(image, page_rect, scale: float, frame, crop) -> bytes:
    """frame の範囲を描き、crop の外側をほぼ透明（不透明度 `ALPHA_FLOOR`）にした PNG を返す。frame と crop が同じなら
    不透明の PNG にする。image はページ全体を scale 倍で描いた画像（`slidekit.pdf.Page.render`）である。

    crop の外側の画素の色は、crop の端の画素の色を外へ延ばしたものにする。色を黒にすると、表示ソフトが画像を
    拡大・縮小するときに crop の境界が灰色の線として見える。
    不透明度を 0 にしないのは、PowerPoint が PDF へ書き出すときに不透明度 0 の画素の色を黒にするためである。
    crop のすぐ外側の `FEATHER_PIXELS` 画素は、透明度を段階的に下げる。
    """
    from PIL import Image

    region = _pixel_intersection(frame.pixels(scale), page_rect.pixels(scale))
    base = image.crop(region)
    inner = _pixel_intersection(crop.pixels(scale), region)
    if inner == region:
        return _png(base.width, base.height, base.tobytes(), alpha=False)
    width, height = base.size
    left, top, right, bottom = inner[0] - region[0], inner[1] - region[1], inner[2] - region[0], inner[3] - region[1]
    image = base.copy()
    alpha = Image.new("L", (width, height), ALPHA_FLOOR)
    if right > left and bottom > top:
        nearest = Image.Resampling.NEAREST
        if left > 0:
            image.paste(image.crop((left, top, left + 1, bottom)).resize((left, bottom - top), nearest), (0, top))
        if right < width:
            image.paste(
                image.crop((right - 1, top, right, bottom)).resize((width - right, bottom - top), nearest), (right, top)
            )
        if top > 0:
            image.paste(image.crop((0, top, width, top + 1)).resize((width, top), nearest), (0, 0))
        if bottom < height:
            image.paste(
                image.crop((0, bottom - 1, width, bottom)).resize((width, height - bottom), nearest), (0, bottom)
            )
        # 外側の段から順に塗り、crop の中を不透明にする。
        for step in range(FEATHER_PIXELS, -1, -1):
            level = 255 * (FEATHER_PIXELS + 1 - step) // (FEATHER_PIXELS + 1)
            alpha.paste(
                level, (max(0, left - step), max(0, top - step), min(width, right + step), min(height, bottom + step))
            )
    image.putalpha(alpha)
    return _png(width, height, image.tobytes())


def _pixel_intersection(a, b) -> tuple[int, int, int, int]:
    """画素の範囲（左、上、右、下）の共通部分。"""
    left, top = max(a[0], b[0]), max(a[1], b[1])
    return left, top, max(left, min(a[2], b[2])), max(top, min(a[3], b[3]))


def _png(width: int, height: int, pixels: bytes, *, alpha: bool = True) -> bytes:
    """画素（alpha が真なら RGBA、偽なら RGB）から PNG を作る。各行はフィルターなしで zlib に通す（PIL の PNG より小さい）。"""
    channels = 4 if alpha else 3
    stride = width * channels
    rows = b"".join(b"\x00" + pixels[start : start + stride] for start in range(0, stride * height, stride))

    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

    header = struct.pack(">IIBBBBB", width, height, 8, 6 if alpha else 2, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b"")


def _page_crops(page, shapes: list[dict], scale: float, max_margin_pt: float) -> list[dict]:
    from .pdf import Box

    traces, drawings = page.spans(), page.drawings()
    # ページは 1 回だけ画像にし、数式ごとの切り出しはその画像から作る。
    image = page.render(scale)
    results = []
    for shape in shapes:
        rectangle = Box.of(shape["rectangle_pt"])
        # 吹き出しは本体（body_pt）で自身の文字を判定し、しっぽまでの範囲（rectangle_pt）を切り出す。
        body = Box.of(shape["body_pt"]) if shape.get("body_pt") else rectangle
        others = [other.get("body_pt") or other["rectangle_pt"] for other in shapes if other is not shape]
        crop, details = crop_rectangle(
            page,
            body,
            shape_rectangle=rectangle,
            outline_pt=shape["outline_pt"],
            traces=traces,
            drawings=drawings,
            scale=scale,
            max_margin_pt=max_margin_pt,
            math_fonts=shape.get("math_fonts", ()),
            others=others,
            neighbors=shape.get("neighbors", ()),
        )
        frame = Box.of(shape["frame_pt"]) if shape.get("frame_pt") else crop
        # 互換画像の図形の外は描くことができないので、切り出す範囲を図形の範囲に収める。
        if not _contains(frame, crop, 1 / scale):
            details["clipped"] = True
        crop &= frame
        details["margin_pt"] = _margins(rectangle, crop)
        results.append(
            {
                "key": shape["key"],
                "crop_pt": _points(crop),
                **details,
                "png": _crop_png(image, page.rect, scale, frame, crop),
            }
        )
    return results


def render_page_crops(pdf: str, index: int, shapes: list[dict], scale: float, max_margin_pt: float) -> list[dict]:
    """PDF の 1 ページから、数式ごとの互換画像を切り出す。子プロセスでも動かすことができる。

    `shapes` は `[{"key", "rectangle_pt", "body_pt", "outline_pt", "frame_pt", "math_fonts", "neighbors"}]`。`body_pt` は吹き出しの
    本体の矩形（`_body_rectangle`）で、省くと rectangle_pt で自身の文字を判定する。`frame_pt`（互換画像の図形の範囲）を
    省くと、切り出す範囲そのものを画像にする。`math_fonts` は数式の書体名で、省くと枠の外へはみ出した数式の文字を探さない。
    `neighbors` は同じページのほかの要素の範囲（`crop_rectangle` の引数）で、省くとほかの要素の描画を区別しない。
    戻り値は `[{"key", "crop_pt", "margin_pt", "clipped", "outline_pt", "nearby_elements", "png"}]`。
    """
    from .pdf import open_pdf

    with open_pdf(pdf) as document:
        return _page_crops(document[index], shapes, scale, max_margin_pt)


@lru_cache(maxsize=256)
def _placeholder_png(width: int, height: int) -> bytes:
    """寸法を決めた 1 色の PNG（1 ビットの色表）。"""
    from PIL import Image

    image = Image.new("P", (width, height), 0)
    image.putpalette(PLACEHOLDER_COLOR)
    buffer = io.BytesIO()
    image.save(buffer, "PNG", bits=1)
    return buffer.getvalue()


def _pixel_size(frame, scale: float) -> tuple[int, int]:
    left, top, right, bottom = frame.pixels(scale)
    return right - left, bottom - top


def _fallback_shape(original, relationship_id: str, frame):
    """互換画像の図形。位置と寸法は frame（pt）、線はなし（枠線は画像に含める）。"""
    shape = etree.Element(tag("p", "sp"))
    shape.append(deepcopy(original.find("p:nvSpPr", NS)))
    properties = etree.SubElement(shape, tag("p", "spPr"))
    transform = etree.SubElement(properties, tag("a", "xfrm"))
    x0, y0, x1, y1 = (round(value * EMU_PER_PT) for value in frame)
    etree.SubElement(transform, tag("a", "off"), x=str(x0), y=str(y0))
    etree.SubElement(transform, tag("a", "ext"), cx=str(x1 - x0), cy=str(y1 - y0))
    geometry = etree.SubElement(properties, tag("a", "prstGeom"), prst="rect")
    etree.SubElement(geometry, tag("a", "avLst"))
    fill = etree.SubElement(properties, tag("a", "blipFill"), dpi="0", rotWithShape="1")
    blip = etree.SubElement(fill, tag("a", "blip"))
    blip.set(tag("r", "embed"), relationship_id)
    stretch = etree.SubElement(fill, tag("a", "stretch"))
    etree.SubElement(stretch, tag("a", "fillRect"))
    outline = etree.SubElement(properties, tag("a", "ln"))
    etree.SubElement(outline, tag("a", "noFill"))
    etree.SubElement(properties, tag("a", "effectLst"))
    return shape


def _relationships_name(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def _relationships_source(name: str) -> str:
    """関係の部品名から、関係を持つ部品の名前を返す（`_rels/.rels` はパッケージ自身で ""）。"""
    directory, base = posixpath.dirname(posixpath.dirname(name)), posixpath.basename(name)
    return directory if base == ".rels" else posixpath.join(directory, base[: -len(".rels")])


def _part_target(source: str, target: str) -> str:
    target = unquote(target).split("#", 1)[0]
    if target.startswith("/"):
        return posixpath.normpath(target).lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), target))


def _placement(transform) -> tuple | None:
    """`a:xfrm`（`p:xfrm`）の（x, y, 幅, 高さ（pt）, 回転（ラジアン）, 左右反転, 上下反転）。位置か寸法がなければ None。"""
    if transform is None:
        return None
    offset, extent = transform.find("a:off", NS), transform.find("a:ext", NS)
    if offset is None or extent is None:
        return None
    return (
        int(offset.get("x", "0")) / EMU_PER_PT,
        int(offset.get("y", "0")) / EMU_PER_PT,
        int(extent.get("cx", "0")) / EMU_PER_PT,
        int(extent.get("cy", "0")) / EMU_PER_PT,
        math.radians(int(transform.get("rot", "0")) / 60000),
        transform.get("flipH") in {"1", "true"},
        transform.get("flipV") in {"1", "true"},
    )


def _orient(point, placement) -> tuple[float, float]:
    """図形の座標の点に、図形の中心を軸にした反転と回転（時計回り）を当てる。"""
    x, y, width, height, angle, flip_h, flip_v = placement
    center_x, center_y = x + width / 2, y + height / 2
    px, py = point
    if flip_h:
        px = 2 * center_x - px
    if flip_v:
        py = 2 * center_y - py
    if angle:
        dx, dy = px - center_x, py - center_y
        px = center_x + dx * math.cos(angle) - dy * math.sin(angle)
        py = center_y + dx * math.sin(angle) + dy * math.cos(angle)
    return px, py


def _group_mapping(group, outer):
    """グループの子の座標をページの座標にする関数。outer は外側のグループの関数。"""
    transform = group.find("p:grpSpPr/a:xfrm", NS)
    placement = _placement(transform)
    if placement is None:
        return outer
    x, y, width, height = placement[:4]
    child_offset, child_extent = transform.find("a:chOff", NS), transform.find("a:chExt", NS)
    child_x, child_y = (
        (int(child_offset.get("x", "0")) / EMU_PER_PT, int(child_offset.get("y", "0")) / EMU_PER_PT)
        if child_offset is not None
        else (x, y)
    )
    child_width, child_height = (
        (int(child_extent.get("cx", "0")) / EMU_PER_PT, int(child_extent.get("cy", "0")) / EMU_PER_PT)
        if child_extent is not None
        else (width, height)
    )
    scale_x = width / child_width if child_width > 0 else 1.0
    scale_y = height / child_height if child_height > 0 else 1.0

    def mapping(point):
        local = (x + (point[0] - child_x) * scale_x, y + (point[1] - child_y) * scale_y)
        return outer(_orient(local, placement))

    return mapping


def _shape_region(node, mapping) -> dict | None:
    """図形の範囲（`crop_rectangle` の `neighbors` の 1 項目）。名前か位置がなければ None。

    `box_pt` は回転と反転を当てた外接矩形。線の形の図形は始点と終点（`segment_pt`）も持つ。
    `tolerance_pt` は描画の広がりの余裕で、線・矢印は矢印の先端の幅、ほかは枠線の外側の半分を含める。
    """
    from .pdf import Box

    name = node.find(".//p:cNvPr", NS)
    transform = node.find("p:xfrm" if node.tag == tag("p", "graphicFrame") else "p:spPr/a:xfrm", NS)
    placement = _placement(transform)
    if name is None or name.get("name") is None or placement is None:
        return None
    x, y, width, height = placement[:4]
    corners = [
        mapping(_orient(point, placement))
        for point in ((x, y), (x + width, y), (x + width, y + height), (x, y + height))
    ]
    xs, ys = [point[0] for point in corners], [point[1] for point in corners]
    region = {"id": name.get("name"), "box_pt": _points(Box(min(xs), min(ys), max(xs), max(ys)))}
    line = outline_width_pt(node)
    geometry = node.find("p:spPr/a:prstGeom", NS)
    if geometry is not None and geometry.get("prst") in _LINE_GEOMETRIES:
        start, end = (mapping(_orient(point, placement)) for point in ((x, y), (x + width, y + height)))
        region["segment_pt"] = [round(value, 4) for value in (*start, *end)]
        region["tolerance_pt"] = round(_REGION_MARGIN_PT + _LINE_END_FACTOR * (line or _DEFAULT_LINE_PT), 4)
    else:
        region["tolerance_pt"] = round(_REGION_MARGIN_PT + line / 2, 4)
    return region


def _shape_regions(tree) -> list[tuple[Any, dict]]:
    """スライドの図形ごとの範囲。（図形, 範囲）の重なり順の列で、範囲は `_shape_region` の形。

    グループの中の図形はグループの座標変換を当て、`mc:AlternateContent` は `mc:Choice` の図形を読む。
    位置のない図形（雛形の位置を使うプレースホルダー）とグループ自身は含めない。
    """
    regions = []

    def visit(container, mapping) -> None:
        for child in container:
            if child.tag == tag("mc", "AlternateContent"):
                choice = child.find("mc:Choice", NS)
                if choice is not None:
                    visit(choice, mapping)
            elif child.tag == tag("p", "grpSp"):
                visit(child, _group_mapping(child, mapping))
            elif child.tag in _SHAPE_TAGS:
                region = _shape_region(child, mapping)
                if region is not None:
                    regions.append((child, region))

    visit(tree, lambda point: point)
    return regions


def _neighbors(shapes: list[tuple[Any, dict]], tree) -> None:
    """同じスライドの数式の図形の項目（（図形, 項目）の列）に、ほかの図形の範囲（`neighbors`）を加える。"""
    regions = _shape_regions(tree)
    for shape, item in shapes:
        item["neighbors"] = [region for node, region in regions if node is not shape]


def _math_fonts(shape) -> list[str]:
    """数式の図形の数式の文字（`\\text` の文字を除く）の書体名。`pdf_layout._font_name` の形で、切り出しが PDF の書体名と比べる。"""
    from .pdf_layout import _font_name

    names = set()
    for run in shape.iter(tag("m", "r")):
        latin = run.find("a:rPr/a:latin", NS)
        if run.find("m:rPr/m:nor", NS) is None and latin is not None and latin.get("typeface"):
            names.add(_font_name(latin.get("typeface")))
    return sorted(names)


def _math_nodes(root) -> tuple[list, list]:
    """スライドの数式を通常の図形（中間 PPTX）と `mc:AlternateContent`（最終形）に分けて返す。"""
    wrapped = [
        node for node in root.iter(tag("mc", "AlternateContent")) if node.find("mc:Choice//m:oMath", NS) is not None
    ]
    plain = [
        shape
        for shape in root.iter(tag("p", "sp"))
        if shape.find(".//m:oMath", NS) is not None
        and next(shape.iterancestors(tag("mc", "AlternateContent")), None) is None
    ]
    return plain, wrapped


def _package_math(parts: dict[str, bytes], ordered: list[str]) -> tuple[bool, bool]:
    plain = wrapped = False
    for name in ordered:
        found = _math_nodes(etree.fromstring(parts[name]))
        plain, wrapped = plain or bool(found[0]), wrapped or bool(found[1])
    return plain, wrapped


def _wrap_math(
    parts: dict[str, bytes], ordered: list[str], slide_size: tuple[float, float], scale: float
) -> list[dict]:
    """中間 PPTX の部品の数式の図形を `mc:AlternateContent` に置き換え、仮の画像を加える。parts を書き換え、図形の一覧を返す。"""
    from .pdf import Box

    page = Box(0.0, 0.0, *map(float, slide_size))
    shapes = []
    for number, name in enumerate(ordered, 1):
        root = etree.fromstring(parts[name])
        plain, wrapped = _math_nodes(root)
        if wrapped:
            raise ValueError("最終形の PPTX を作る入力は、数式を通常の図形として持つ中間 PPTX にする")
        if not plain:
            continue
        tree = root.find("p:cSld/p:spTree", NS)
        relation_name = _relationships_name(name)
        relationships = (
            etree.fromstring(parts[relation_name])
            if relation_name in parts
            else etree.Element(tag("pr", "Relationships"), nsmap={None: NS["pr"]})
        )
        next_id = (
            max(
                [int(node.get("Id")[3:]) for node in relationships if re.fullmatch(r"rId\d+", node.get("Id", ""))] + [0]
            )
            + 1
        )
        for shape in plain:
            if shape.getparent() is not tree:
                raise ValueError("数式を含む図形はスライドの直下に置く")
            rectangle = _math_rectangle(shape)
            if not _contains(page, rectangle):
                raise ValueError(f"{number} ページ目の数式の枠がページの外にある")
            frame = _frame(rectangle, page, scale)
            media_name = f"ppt/media/slidekit_math_{number:04d}_{len(shapes) + 1:04d}.png"
            if media_name in parts:
                raise ValueError(f"互換画像の名前が重複する: {media_name}")
            parts[media_name] = _placeholder_png(*_pixel_size(frame, scale))
            relation_id = f"rId{next_id}"
            next_id += 1
            etree.SubElement(
                relationships,
                tag("pr", "Relationship"),
                Id=relation_id,
                Type=NS["r"] + "/image",
                Target=posixpath.relpath(media_name, posixpath.dirname(name)),
            )
            alternative = etree.Element(tag("mc", "AlternateContent"), nsmap={"mc": NS["mc"], "a14": NS["a14"]})
            choice = etree.SubElement(alternative, tag("mc", "Choice"), Requires="a14")
            fallback = etree.SubElement(alternative, tag("mc", "Fallback"))
            fallback.append(_fallback_shape(shape, relation_id, frame))
            position = tree.index(shape)
            tree.remove(shape)
            choice.append(shape)
            tree.insert(position, alternative)
            shapes.append(
                {
                    "slide_number": number,
                    "slide_id": root.find("p:cSld", NS).get("name"),
                    "element_id": shape.find("p:nvSpPr/p:cNvPr", NS).get("name"),
                    "rectangle": rectangle,
                    "body": _body_rectangle(shape, rectangle),
                    "frame": frame,
                    "outline_pt": outline_width_pt(shape),
                    "image": media_name,
                    "native_math": len(shape.findall(".//m:oMath", NS)),
                    "math_fonts": _math_fonts(shape),
                }
            )
        _neighbors(list(zip(plain, shapes[-len(plain) :])), tree)
        parts[name] = xml_bytes(root)
        parts[relation_name] = xml_bytes(relationships)
    if shapes:
        content_types = etree.fromstring(parts["[Content_Types].xml"])
        if not any(node.get("Extension", "").lower() == "png" for node in content_types):
            etree.SubElement(content_types, tag("ct", "Default"), Extension="png", ContentType="image/png")
            parts["[Content_Types].xml"] = xml_bytes(content_types)
    return shapes


def _read_final_form(parts: dict[str, bytes], ordered: list[str]) -> list[dict]:
    """最終形の PPTX の数式ごとに、矩形・互換画像の図形の範囲・枠線・画像の部品名を読む。"""
    shapes, images = [], set()
    for number, name in enumerate(ordered, 1):
        root = etree.fromstring(parts[name])
        _, wrapped = _math_nodes(root)
        if not wrapped:
            continue
        tree = root.find("p:cSld/p:spTree", NS)
        relation_name = _relationships_name(name)
        relationships = etree.fromstring(parts[relation_name]) if relation_name in parts else []
        targets = {
            node.get("Id"): _part_target(name, node.get("Target", ""))
            for node in relationships
            if node.get("TargetMode") != "External"
        }
        placed = []
        for alternative in wrapped:
            if alternative.getparent() is not tree:
                raise ValueError("数式を含む図形はスライドの直下に置く")
            original = alternative.find("mc:Choice/p:sp", NS)
            fallback = alternative.find("mc:Fallback", NS)
            blip = None if fallback is None or len(fallback) != 1 else fallback[0].find("p:spPr/a:blipFill/a:blip", NS)
            if original is None or blip is None:
                raise ValueError(f"{name}: 数式の互換画像の形が不正である")
            image = targets.get(blip.get(tag("r", "embed")))
            if image not in parts:
                raise ValueError(f"{name}: 数式の互換画像の部品がない")
            if image in images:
                raise ValueError(f"{name}: 互換画像を複数の数式で共有している: {image}")
            images.add(image)
            rectangle, frame = _math_rectangle(original), _math_rectangle(fallback[0])
            if not _contains(frame, rectangle, 0.01):
                raise ValueError(f"{name}: 互換画像の図形が数式の図形を覆っていない")
            shapes.append(
                {
                    "slide_number": number,
                    "slide_id": root.find("p:cSld", NS).get("name"),
                    "element_id": original.find("p:nvSpPr/p:cNvPr", NS).get("name"),
                    "rectangle": rectangle,
                    "body": _body_rectangle(original, rectangle),
                    "frame": frame,
                    "outline_pt": outline_width_pt(original),
                    "image": image,
                    "native_math": len(original.findall(".//m:oMath", NS)),
                    "math_fonts": _math_fonts(original),
                }
            )
            placed.append((original, shapes[-1]))
        _neighbors(placed, tree)
    return shapes


def prepare_final_form(native_pptx: Path, output: Path, *, scale: float = 4, source_sha256: str | None = None) -> dict:
    """互換画像を仮の画像にした最終形の PPTX を作る。

    スライドの XML と関係は、`finalize_math` が書く最終 PPTX と同一になる（互換画像の中身だけが違う）。
    仮の画像は互換画像の図形の範囲を scale 倍した寸法の 1 色の PNG で、数式ごとに別の名前にする。
    scale は `finalize_math` に渡す値と同じにする。source_sha256 は中間 PPTX のハッシュ値で、分かっていれば渡す（求め直さない）。
    """
    if not math.isfinite(scale) or scale <= 0:
        raise ValueError("画像の倍率は正の有限値にする")
    if Path(native_pptx).resolve() == Path(output).resolve():
        raise ValueError("中間 PPTX と最終形の PPTX は別のパスにする")
    parts = read_package(native_pptx)
    ordered, slide_size = slide_order(parts)
    shapes = _wrap_math(parts, ordered, slide_size, scale)
    write_package(parts, output)
    return {
        "output": str(output),
        "sha256": file_hash(output),
        "source_sha256": source_sha256 or file_hash(native_pptx),
        "slides": len(ordered),
        "fallback_shapes": len(shapes),
        "native_math": sum(shape["native_math"] for shape in shapes),
        "image_scale": scale,
        "shapes": [
            {
                "slide_number": shape["slide_number"],
                "slide_id": shape["slide_id"],
                "element_id": shape["element_id"],
                "rectangle_pt": _points(shape["rectangle"]),
                "frame_pt": _points(shape["frame"]),
                "outline_pt": shape["outline_pt"],
                "image": shape["image"],
                "native_math": shape["native_math"],
            }
            for shape in shapes
        ],
    }


def finalize_math(
    final_form_pptx: Path,
    final_form_pdf: Path,
    output: Path,
    *,
    scale: float = 4,
    workers=None,
    source_pptx_sha256: str | None = None,
    source_pdf_sha256: str | None = None,
) -> dict:
    """最終形の PPTX を変換した PDF から互換画像を切り出し、仮の画像の中身だけを差し替えて最終 PPTX を書く。XML は変えない。

    数式を通常の図形として持つ中間 PPTX とその PDF も受け付ける。そのときは最終形の形に置き換えてから切り出す。
    `workers` は `map(function, arguments)` を持つもの（`parallel.Workers` など）で、ページごとの切り出しを任せる。
    `None` なら同じプロセスで順に切り出す。どちらでも結果は同じである。scale は `prepare_final_form` に渡した値と同じにする。
    source_pptx_sha256・source_pdf_sha256 は入力の PPTX と PDF のハッシュ値で、分かっていれば渡す（求め直さない）。
    切り出しでは同じスライドのほかの図形の範囲（`_shape_regions`）を渡し、その図形の描画を含めない。戻り値の `crops` の
    `nearby_elements` は、切り詰めたときに上限の外の部分に範囲が掛かるほかの図形の名前（要素の識別子）である。
    """
    from .pdf import open_pdf

    if not math.isfinite(scale) or scale <= 0:
        raise ValueError("画像の倍率は正の有限値にする")
    if Path(final_form_pptx).resolve() == Path(output).resolve():
        raise ValueError("入力の PPTX と最終 PPTX は別のパスにする")
    parts = read_package(final_form_pptx)
    ordered, slide_size = slide_order(parts)
    plain, wrapped = _package_math(parts, ordered)
    if plain and wrapped:
        raise ValueError("中間 PPTX の数式と最終形の数式が混在している")
    shapes = _wrap_math(parts, ordered, slide_size, scale) if plain else _read_final_form(parts, ordered)
    tasks: dict[int, list[dict]] = {}
    for key, shape in enumerate(shapes):
        tasks.setdefault(shape["slide_number"] - 1, []).append(
            {
                "key": key,
                "rectangle_pt": list(shape["rectangle"]),
                **({"body_pt": list(shape["body"])} if list(shape["body"]) != list(shape["rectangle"]) else {}),
                "outline_pt": shape["outline_pt"],
                "frame_pt": list(shape["frame"]),
                "math_fonts": shape["math_fonts"],
                "neighbors": shape["neighbors"],
            }
        )
    with open_pdf(final_form_pdf) as document:
        if len(document) != len(ordered):
            raise ValueError(f"ページ数が一致しない: PPTX={len(ordered)}, PDF={len(document)}")
        for number, page in enumerate(document, 1):
            if abs(page.width - slide_size[0]) > 0.1 or abs(page.height - slide_size[1]) > 0.1:
                raise ValueError(f"{number} ページ目の PDF と PPTX の寸法が異なる")
        for shape in shapes:
            if not _contains(document[shape["slide_number"] - 1].rect, shape["rectangle"]):
                raise ValueError(f"{shape['slide_number']} ページ目の数式の枠が PDF の外にある")
        pages = (
            None
            if workers is not None
            else [_page_crops(document[index], items, scale, MAX_MARGIN_PT) for index, items in tasks.items()]
        )
    if pages is None:
        pages = workers.map(
            render_page_crops,
            [(str(final_form_pdf), index, items, scale, MAX_MARGIN_PT) for index, items in tasks.items()],
        )
    results = {result["key"]: result for page in pages for result in page}
    crops = []
    for key, shape in enumerate(shapes):
        result = results[key]
        parts[shape["image"]] = result["png"]
        crops.append(
            {
                "slide_number": shape["slide_number"],
                "slide_id": shape["slide_id"],
                "element_id": shape["element_id"],
                "rectangle_pt": _points(shape["rectangle"]),
                "frame_pt": _points(shape["frame"]),
                "crop_pt": result["crop_pt"],
                "margin_pt": result["margin_pt"],
                "outline_pt": result["outline_pt"],
                "clipped": result["clipped"],
                "nearby_elements": result["nearby_elements"],
                "image": shape["image"],
                "image_sha256": hashlib.sha256(result["png"]).hexdigest(),
                "native_math": shape["native_math"],
            }
        )
    write_package(parts, output)
    return {
        "output": str(output),
        "sha256": file_hash(output),
        "input_form": "native" if plain else "final_form",
        "slides": len(ordered),
        "fallback_shapes": len(crops),
        "native_math": sum(crop["native_math"] for crop in crops),
        "image_scale": scale,
        "source_pptx_sha256": source_pptx_sha256 or file_hash(final_form_pptx),
        "source_pdf_sha256": source_pdf_sha256 or file_hash(final_form_pdf),
        "crops": crops,
    }


def _fallback_images(parts: dict[str, bytes]) -> set[str]:
    """スライドの `mc:Fallback` の中だけから参照される部品（互換画像）の名前。"""
    slides = set(slide_order(parts)[0])
    inside_only, others = set(), set()
    for name in parts:
        if not name.endswith(".rels"):
            continue
        source = _relationships_source(name)
        inside: Counter[str] = Counter()
        everywhere: Counter[str] = Counter()
        if source in slides and source in parts:
            root = etree.fromstring(parts[source])
            everywhere.update(map(str, _RELATIONSHIP_VALUES(root, namespace=NS["r"])))
            inside.update(map(str, _FALLBACK_RELATIONSHIP_VALUES(root, namespace=NS["r"])))
        for relationship in etree.fromstring(parts[name]):
            if relationship.get("TargetMode") == "External":
                continue
            target = _part_target(source, relationship.get("Target", ""))
            identifier = relationship.get("Id")
            (inside_only if inside[identifier] and inside[identifier] == everywhere[identifier] else others).add(target)
    return inside_only - others


def compare_package_structure(a_pptx: Path, b_pptx: Path) -> dict:
    """2 つの PPTX の部品を比べる。互換画像（`mc:Fallback` だけが参照する画像）は名前だけを比べ、ほかの部品は中身も比べる。

    最終形の PPTX と最終 PPTX が、スライドの XML・関係・ほかのメディアまで同一で、互換画像の中身だけが違うことを確かめる。
    戻り値の `differences` は `{"part", "reason"}` の列で、`reason` は `only_in_first`・`only_in_second`・`content` のどれか。
    """
    first, second = read_package(a_pptx), read_package(b_pptx)
    images = _fallback_images(first) & _fallback_images(second)
    differences = []
    for name in sorted(set(first) | set(second)):
        if name not in second:
            differences.append({"part": name, "reason": "only_in_first"})
        elif name not in first:
            differences.append({"part": name, "reason": "only_in_second"})
        elif name not in images and first[name] != second[name]:
            differences.append({"part": name, "reason": "content"})
    common = set(first) & set(second)
    return {
        "matched": not differences,
        "parts": len(common),
        "fallback_images": len(images & common),
        "differences": differences,
    }


def _reachable_parts(parts: dict[str, bytes]) -> set[str]:
    """パッケージの関係（`_rels/.rels`）を起点に、関係を順にたどって行き着く部品の名前。"""
    reached: set[str] = set()
    pending = [""]
    while pending:
        source = pending.pop()
        relation_name = _relationships_name(source)
        if relation_name not in parts:
            continue
        for relationship in etree.fromstring(parts[relation_name]):
            if relationship.get("TargetMode") == "External":
                continue
            target = _part_target(source, relationship.get("Target", ""))
            if target in parts and target not in reached:
                reached.add(target)
                pending.append(target)
    return reached


def keep_slides(parts: dict[str, bytes], names: Iterable[str]) -> dict[str, bytes]:
    """names のスライドだけを残した部品の辞書を返す（parts は変えない）。

    外すスライドを `p:sldIdLst` と資料本体の関係、章（`p14:sectionLst`）の `p14:sldId`、目的別スライド ショー（`p:custShowLst`）の参照から外し、
    パッケージの関係からたどることができなくなった部品と、その `[Content_Types].xml` の `Override` を消す。
    外すスライドが残る部品から参照されているときは ValueError にする。
    """
    keep = set(names)
    ordered, _ = slide_order(parts)
    unknown = keep - set(ordered)
    if unknown:
        raise ValueError(f"残すスライドが PPTX にない: {', '.join(sorted(unknown))}")
    result = dict(parts)
    presentation = etree.fromstring(parts["ppt/presentation.xml"])
    relationships = etree.fromstring(parts["ppt/_rels/presentation.xml.rels"])
    targets = {
        node.get("Id"): _part_target("ppt/presentation.xml", node.get("Target", ""))
        for node in relationships
        if node.get("TargetMode") != "External"
    }
    removed_ids, removed_relationships = set(), set()
    for node in presentation.findall("p:sldIdLst/p:sldId", NS):
        identifier = node.get(tag("r", "id"))
        if targets.get(identifier) not in keep:
            removed_ids.add(node.get("id"))
            removed_relationships.add(identifier)
            node.getparent().remove(node)
    if not removed_relationships:
        return result
    for node in list(presentation.iter(f"{{{_P14}}}sldId")):
        if node.get("id") in removed_ids:
            node.getparent().remove(node)
    for node in presentation.findall("p:custShowLst/p:custShow/p:sldLst/p:sld", NS):
        if node.get(tag("r", "id")) in removed_relationships:
            node.getparent().remove(node)
    for node in list(relationships):
        if node.get("Id") in removed_relationships:
            relationships.remove(node)
    result["ppt/presentation.xml"] = xml_bytes(presentation)
    result["ppt/_rels/presentation.xml.rels"] = xml_bytes(relationships)
    reachable = _reachable_parts(result)
    stranded = (set(ordered) - keep) & reachable
    if stranded:
        raise ValueError(f"外すスライドが残る部品から参照されている: {', '.join(sorted(stranded))}")
    for name in list(result):
        if name in {"[Content_Types].xml", "_rels/.rels"}:
            continue
        if (_relationships_source(name) if name.endswith(".rels") else name) not in reachable:
            del result[name]
    content_types = etree.fromstring(result["[Content_Types].xml"])
    for node in list(content_types):
        if node.tag == tag("ct", "Override") and unquote(node.get("PartName", "")).lstrip("/") not in result:
            content_types.remove(node)
    result["[Content_Types].xml"] = xml_bytes(content_types)
    return result


def prepare_fallback(
    final_pptx: Path, output: Path, *, math_only: bool = False, source_sha256: str | None = None
) -> dict:
    """最終 PPTX の互換画像を通常の図形へ展開した確認用複製を作る。

    math_only が真なら、数式のページだけを残す（`keep_slides`）。戻り値の `slides` は数式のページの元の番号、
    `kept_slides` は複製に残したページの、複製の中の番号（1 から）と識別子である。
    source_sha256 は最終 PPTX のハッシュ値で、分かっていれば渡す（求め直さない）。
    """
    if Path(final_pptx).resolve() == Path(output).resolve():
        raise ValueError("最終 PPTX と確認用複製は別のパスにする")
    parts = read_package(final_pptx)
    ordered, _ = slide_order(parts)
    replaced = 0
    slides, identifiers, math_names = [], [], []
    for number, name in enumerate(ordered, 1):
        root = etree.fromstring(parts[name])
        identifiers.append(root.find("p:cSld", NS).get("name"))
        page_count = 0
        for alternative in reversed(root.findall(".//mc:AlternateContent", NS)):
            choice = alternative.find("mc:Choice", NS)
            if choice is None or not choice.findall(".//m:oMath", NS):
                continue
            fallback = alternative.find("mc:Fallback", NS)
            if fallback is None or len(fallback) != 1 or not fallback.findall(".//a:blip", NS):
                raise ValueError(f"{name}: 数式の互換画像がない")
            parent = alternative.getparent()
            position = parent.index(alternative)
            parent.insert(position, deepcopy(fallback[0]))
            parent.remove(alternative)
            replaced += 1
            page_count += 1
        if root.findall(".//m:oMath", NS):
            raise ValueError(f"{name}: 互換画像のない数式が残っている")
        if page_count:
            slides.append({"number": number, "id": identifiers[-1]})
            math_names.append(name)
            parts[name] = xml_bytes(root)
    if math_only:
        if not math_names:
            raise ValueError("確認用複製に残す数式のページがない")
        parts = keep_slides(parts, math_names)
        kept = [{"number": index, "id": slide["id"]} for index, slide in enumerate(slides, 1)]
    else:
        kept = [{"number": number, "id": identifier} for number, identifier in enumerate(identifiers, 1)]
    write_package(parts, output)
    return {
        "output": str(output),
        "sha256": file_hash(output),
        "source_sha256": source_sha256 or file_hash(final_pptx),
        "alternatives_replaced": replaced,
        "slides": slides,
        "kept_slides": kept,
        "math_only": math_only,
    }
