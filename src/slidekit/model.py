"""編集可能なスライドの内容と配置を表す。ファイルには出力しない。

書式は `styles` の順に重ねて解決する（ライブラリ < テーマの `[defaults]` < `[defaults.<種類>]` < 役割 < 要素の
`style=` < キーワード引数）。解決後の場面の書式では色を 16 進数にし、塗らない・線なしを `None` で表す。

入力の誤りはページと要素の識別子を持つ `errors.InputError`（受け付けない引数名は `ArgumentNameError`）にする。
識別子・領域・接続などの構造の誤りは、診断と同じ識別子（`code`）の例外にする（場面の検査では確かめない）。
資料の `collect_errors` が真のときは、続行できる誤り（書式の名前と値、役割、断片の書式とリンク、出典など）を
`Deck.issues` に集めて続け、`Deck.to_dict()` の結果の `issues` に診断として入れる。`stack.take` で受け取った領域が
積み上げる領域を超えた分も、集めるかどうかによらず、同じ `issues` に配置の診断（`layout_bounds_overflow`）として入れる。

領域の幅・高さの `None` は「内容から決める寸法」で、要素の作成で数値にする（文字・図形・数式は `measure.layout_text`、
画像は縦横比）。`slide.measure()`・`element.measure()` は同じ計算の結果（`measure.TextLayout`）を返し、
`slide.stack()` は要素を上から順に間隔を空けて置く。

まとまりは条件（`_pptx_groups`）を満たすと、場面で `pptx_group` が真になり、PowerPoint のグループとして書き出される。
ページ種別の飾り（テーマの `page_kinds.<種別>.decorations`）は `_theme/decoration/<識別子>` の要素になる。
`slide.references()` は文献一覧を、出典の文献キーと `project.bibliography` から作る。
"""

from __future__ import annotations

from collections.abc import Mapping
import copy
from dataclasses import dataclass
import math
import os
from pathlib import Path
import re
import sys

from . import styles
from .errors import ArgumentNameError, InputError

# 断片の型は fragments にある。slidekit の公開 API が model から読み込むため、ここで読み込む。
from .fragments import Emphasis, Math, Text, content_fragments, valid_link
from .layout import Rect, anchor_options, as_rect, finite_number
from .measure import layout_element, layout_text, normalize_padding, phrase_wrapped_content
from .schema import PAGE_KINDS, role_usage_problems

_KINDS = frozenset({"text", "box", "equation", "image", "line", "arrow", "group"})
# 内容から決める寸法に対応する種類（文字の計測で決める種類と、縦横比で決める画像）。
_MEASURED_KINDS = frozenset({"text", "box", "equation"})
_SIZED_KINDS = _MEASURED_KINDS | {"image"}
# 種類ごとの専用の引数と、すべての要素に共通の引数。受け付けない名前の候補にも使う。
_SPECIFIC = {
    "text": ("text",),
    "box": ("text", "shape"),
    "equation": ("expression",),
    "image": ("path", "fit", "page", "raster_scale", "vector"),
    "line": ("start", "end"),
    "arrow": ("start", "end"),
    "group": ("pptx_group",),
}
_COMMON = ("id", "role", "sources", "style")
_TRANSFORMS = ("rotation", "flip_h", "flip_v")
# 場面の書式で float にそろえる数値の項目（padding_pt は normalize_padding が扱う）。
_NUMBER_KEYS = ("size_pt", "min_size_pt", "line_spacing", "line_width_pt", "line_spacing_pt", "corner_radius_pt")
_HEX = re.compile("[0-9a-fA-F]{6}")
_RASTER_FORMATS = (".png", ".jpg", ".jpeg")
_IMAGE_FORMATS = (*_RASTER_FORMATS, ".svg", ".pdf")
# 積み上げの横位置。
_PLACES = ("left", "center", "right")
# SVG の長さの単位（CSS の px に対する比。縦横比を求めるだけなので、相対の単位は扱わない）。
_SVG_UNITS = {"": 1.0, "px": 1.0, "pt": 4 / 3, "pc": 16.0, "mm": 96 / 25.4, "cm": 96 / 2.54, "in": 96.0}
_SVG_LENGTH = re.compile(r"\s*([+]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)\s*(px|pt|pc|mm|cm|in)?\s*")
_EPSILON = 1e-9


def _shown(value) -> str:
    """誤りの文言に入れる値（repr を 40 文字まで）。"""
    text = repr(value)
    return text if len(text) <= 40 else text[:39] + "…"


def _option_number(value) -> float | None:
    """有限の数（真偽値を除く）なら float、そうでなければ None。"""
    try:
        return finite_number(value, "値")
    except ValueError:
        return None


def _quarter_turn(degrees: float) -> tuple[float, float]:
    """時計回りの回転角の (cos, sin)。90 度の倍数は誤差のない値にする（場面の座標を安定させる）。"""
    quarter = degrees / 90
    if quarter == int(quarter):
        return ((1.0, 0.0), (0.0, 1.0), (-1.0, 0.0), (0.0, -1.0))[int(quarter) % 4]
    radians = math.radians(degrees)
    return math.cos(radians), math.sin(radians)


def _transformed(point, center, *, rotation=0.0, flip_h=False, flip_v=False):
    """図形の変換を点に施す。中心について左右反転・上下反転し、その後に rotation 度だけ時計回りに回す
    （PowerPoint の図形の変換と同じ順。y は下向き）。"""
    x, y = point
    cx, cy = center
    if flip_h:
        x = 2 * cx - x
    if flip_v:
        y = 2 * cy - y
    if rotation:
        cos, sin = _quarter_turn(rotation)
        dx, dy = x - cx, y - cy
        x, y = cx + dx * cos - dy * sin, cy + dx * sin + dy * cos
    return (x, y)


def _svg_length(value) -> float | None:
    """SVG の width・height（px に換算）。% や em などの相対の長さと、0 以下の値は None。"""
    match = _SVG_LENGTH.fullmatch(value) if isinstance(value, str) else None
    if match is None:
        return None
    length = float(match.group(1)) * _SVG_UNITS[match.group(2) or ""]
    return length if math.isfinite(length) and length > 0 else None


def _svg_size(data: bytes) -> tuple[float, float] | None:
    """SVG の縦横比に使う寸法。width・height がともに絶対の長さならその値、なければ viewBox の幅と高さ。"""
    from lxml import etree

    try:
        root = etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True))
    except etree.XMLSyntaxError:
        return None
    if not isinstance(root.tag, str) or etree.QName(root).localname != "svg":
        return None
    width, height = _svg_length(root.get("width")), _svg_length(root.get("height"))
    if width is not None and height is not None:
        return (width, height)
    try:
        numbers = [float(value) for value in re.split(r"[\s,]+", (root.get("viewBox") or "").strip())]
    except ValueError:
        return None
    if len(numbers) == 4 and all(map(math.isfinite, numbers)) and numbers[2] > 0 and numbers[3] > 0:
        return (numbers[2], numbers[3])
    return None


def _intrinsic_size(path: Path, page: int | None) -> tuple[float, float] | None:
    """画像の縦横比に使う寸法。PNG・JPEG は画素数、PDF は指定ページの寸法（pt）、SVG は `_svg_size`。

    寸法が分からない SVG は None。PDF の指定ページがなければ ValueError、ファイルの読み込みに失敗したら読み込みの例外にする。
    """
    suffix = path.suffix.lower()
    if suffix in _RASTER_FORMATS:
        from PIL import Image

        with Image.open(path) as source:
            size = source.size
            source.verify()
        return size
    if suffix == ".pdf":
        from .pdf import page_size

        return page_size(path, page)
    return _svg_size(path.read_bytes())


def _identifier(value, kind="識別子", *, code="invalid_element_id"):
    if (
        not isinstance(value, str)
        or not value.strip()
        or "/" in value
        or value.startswith("_theme")
        or any(ord(c) < 32 for c in value)
    ):
        raise InputError(
            f"{kind} は空でなく、/ と予約接頭辞 _theme を含まない文字列にしてください: {_shown(value)}",
            code=code,
            value=_shown(value),
        )
    return value


def key_id(*keys: str | int) -> str:
    """キーを _ でつないで要素の識別子にする（例：key_id("step", "prepare") は "step_prepare"）。

    繰り返し作る要素は、番号ではなく内容のキーから識別子を作る（並べ替えても同じ要素を指す）。
    各キーは空でない文字列か整数（真偽値は不可）で、/ と制御文字を含めない。結果は要素の識別子の規則で確かめる。
    """
    if not keys:
        raise InputError("key_id にはキーを 1 つ以上指定してください", code="invalid_element_id")
    parts = []
    for key in keys:
        if isinstance(key, bool) or not isinstance(key, (str, int)) or not str(key).strip():
            raise InputError(
                f"key_id のキーは空でない文字列か整数で指定してください: {_shown(key)}",
                code="invalid_element_id",
                value=_shown(key),
            )
        if "/" in str(key) or any(ord(c) < 32 for c in str(key)):
            raise InputError(
                f"key_id のキーに / と制御文字は使えません: {_shown(key)}", code="invalid_element_id", value=_shown(key)
            )
        parts.append(str(key))
    return _identifier("_".join(parts), "key_id で作る識別子")


def _plain_text(fragments) -> str:
    """断片の列を通常の文字列にする。数式の断片は $…$ で囲む（場面のページの題名に使う）。"""
    return "".join(
        f"${fragment['text']}$" if fragment["kind"] == "math" else fragment["text"] for fragment in fragments
    )


def _pptx_groups(elements: list[dict]) -> set[str]:
    """ページの場面の要素（作成順）のうち、PowerPoint のグループとして書き出すまとまりの識別子。

    pptx_group が偽でなく、次をすべて満たすまとまりを書き出す。入れ子のまとまりも外側によらず同じ条件で判定する。
    1. 描く子孫（まとまり以外）が 1 つ以上ある。
    2. 子孫が作成順でまとまりの直後に続く。間にまとまりの外の要素があると、グループにしたときに重なり順が変わるため。
    3. 子孫に数式の断片がない。数式の互換画像はスライド直下の図形だけに付ける（数式を含むまとまりは展開して書き出す）。
    """
    position = {element["id"]: index for index, element in enumerate(elements)}
    children: dict[str, list[str]] = {}
    for element in elements:
        if element.get("parent_id") is not None:
            children.setdefault(element["parent_id"], []).append(element["id"])
    result = set()
    for index, element in enumerate(elements):
        if element["kind"] != "group" or element.get("pptx_group") is False:
            continue
        descendants, pending = set(), list(children.get(element["id"], ()))
        while pending:
            identifier = pending.pop()
            if identifier not in descendants and identifier != element["id"]:
                descendants.add(identifier)
                pending.extend(children.get(identifier, ()))
        members = [elements[position[identifier]] for identifier in descendants]
        if not any(member["kind"] != "group" for member in members):
            continue
        if sorted(position[identifier] for identifier in descendants) != list(
            range(index + 1, index + 1 + len(descendants))
        ):
            continue
        if any(fragment.get("kind") == "math" for member in members for fragment in member.get("content") or ()):
            continue
        result.add(element["id"])
    return result


def _page_size(project):
    size = tuple(project.size)
    if len(size) != 2 or any(
        isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value <= 0
        for value in size
    ):
        raise InputError(
            f"ページ寸法は正の有限の数 2 つで指定してください: {_shown(size)}",
            code="invalid_page_size",
            size=_shown(size),
        )
    return size


@dataclass(frozen=True)
class Anchor:
    """線の端点を結ぶ要素の接続位置。`Element.anchor()` の戻り値で、要素を動かすと端点も追従する。"""

    element: "Element"
    side: str
    gap: float = 0.0
    at: float = 0.5

    def point(self):
        """接続位置の座標。辺は要素自身の向きを基準にし、要素の反転・回転を施した点を返す。画像は表示される範囲を基準にする。"""
        area = self.element._visible_area()
        point = area.anchor(self.side, gap=self.gap, at=self.at)
        data = self.element.data
        rotation, flip_h, flip_v = data.get("rotation", 0.0), data.get("flip_h", False), data.get("flip_v", False)
        if not (rotation or flip_h or flip_v):
            return point
        return _transformed(point, area.center, rotation=rotation, flip_h=flip_h, flip_v=flip_v)

    def reference(self) -> dict:
        """場面の start_ref・end_ref。gap・at は既定（0、0.5）以外のときだけ書く。"""
        result = {"element_id": self.element.id, "side": self.side}
        if self.gap != 0:
            result["gap"] = self.gap
        if self.at != 0.5:
            result["at"] = self.at
        return result


class Element:
    # 画像の縦横比に使う寸法（PNG・JPEG は画素数、PDF はページの寸法、SVG は width・height か viewBox）。分からなければ None。
    _intrinsic = None

    def __init__(self, slide, data):
        self.slide = slide
        self.data = data

    @property
    def id(self):
        return self.data["id"]

    @property
    def area(self):
        return as_rect(self.data["area"])

    @area.setter
    def area(self, value):
        try:
            value = as_rect(value)
        except InputError as error:
            raise self.slide._locate(error, self.id) from None
        if value.w is None or value.h is None:
            raise InputError(
                f"ページ {self.slide.id} の要素 {self.id}: 要素の領域の幅と高さは数値で指定してください",
                code="invalid_geometry",
                slide_id=self.slide.id,
                element_id=self.id,
            )
        previous = self.area
        self.data["area"] = list(value)
        if self.data["kind"] in ("line", "arrow") and hasattr(self, "_start"):
            dx, dy = value.x - previous.x, value.y - previous.y
            for key in ("_start", "_end"):
                point = getattr(self, key)
                if not isinstance(point, Anchor):
                    setattr(self, key, (point[0] + dx, point[1] + dy))
        if self.data["kind"] == "group":
            dx, dy = value.x - previous.x, value.y - previous.y
            for child in self.slide.elements:
                if child.data.get("parent_id") == self.id:
                    child.area = child.area.translated(dx, dy)
                    # まとまりの中で積み上げた子の領域も、まとまりと一緒に動かす。
                    if child.data.get("layout_bounds") is not None:
                        child.data["layout_bounds"] = list(as_rect(child.data["layout_bounds"]).translated(dx, dy))

    def _visible_area(self) -> Rect:
        """表示される範囲。contain の画像は縦横比を保って中央に置いた範囲（`Rect.contain`）、ほかは要素の領域。"""
        area = self.area
        if (
            self.data["kind"] == "image"
            and self.data.get("fit", "contain") == "contain"
            and self._intrinsic is not None
        ):
            return area.contain(*self._intrinsic)
        return area

    def anchor(self, side="center", *, gap=0, at=0.5):
        """線の端点を結ぶ接続位置。side は left・right・top・bottom・center、gap は辺から外側への距離（pt、0 以上）、
        at は辺の上の位置（上下の辺は左端 0〜右端 1、左右の辺は上端 0〜下端 1）。

        回転・反転した要素では、辺は要素自身の向きを、画像は表示される範囲を基準にする。
        """
        try:
            side, gap, at = anchor_options(side, gap, at)
            area = self.area
            if self.data["kind"] == "group" and area.w == 0 and area.h == 0:
                raise InputError(
                    "領域のないまとまりには接続できません。group に area を指定してください", code="invalid_anchor"
                )
        except InputError as error:
            raise self.slide._locate(error, self.id) from None
        return Anchor(self, side, gap, at)

    def measure(self):
        """作成済みの要素を今の書式と枠で計測する（`measure.layout_element` の文字の結果。配置の検査と同じ値）。

        戻り値は `measure.TextLayout`。文字を持たない要素（線・矢印・画像・まとまり）は計測できない。
        """
        kind = self.data["kind"]
        if kind not in _MEASURED_KINDS:
            raise self.slide._locate(
                InputError(f"文字を持たない要素は計測できません（{kind}）", code="input_error"), self.id
            )
        try:
            return layout_element(self.data, self.slide.project.fonts).text
        except ValueError as error:
            raise self.slide._locate(error, self.id) from None

    def move(self, dx=0, dy=0):
        try:
            area = self.area.translated(dx, dy)
        except InputError as error:
            raise self.slide._locate(error, self.id) from None
        self.area = area
        return self


class Group(Element):
    """まとまり。子は同じ名前の関数で作り、同じ検査を受ける。子の座標はページ上の座標へ展開して持つ。

    `group.stack(area)` の領域と、子の `area`・端点の座標は、まとまりの左上を原点とする（`group.references()` の area も同じ）。
    `group.measure(text, ...)` は `slide.measure()` と同じで、引数なしの `group.measure()` は要素の計測（まとまりは計測できない）。
    条件（`_pptx_groups`）を満たすまとまりは、PowerPoint のグループとして書き出す。
    """

    _CREATORS = frozenset(
        {"text", "box", "image", "equation", "line", "arrow", "table", "group", "stack", "measure", "references"}
    )

    def measure(self, *args, **kwargs):
        if not args and not kwargs:
            return super().measure()
        return self.slide.measure(*args, **kwargs)

    def __getattr__(self, name):
        if name not in Group._CREATORS:
            raise AttributeError(name)
        if name == "stack":

            def stack(area, *, gap=12, avoid=None):
                return Stack(self.slide, area, gap, owner=self, avoid=avoid)

            return stack

        def add(*, id, **arguments):
            try:
                _identifier(id)
            except InputError as error:
                raise self.slide._locate(error, None) from None
            element_id = f"{self.id}/{id}"
            arguments["_element_id"] = element_id
            arguments["_parent_id"] = self.id
            if arguments.get("area") is not None:
                # 子の領域の誤りも、直接作るときと同じく子のページと要素の識別子を付ける。
                try:
                    arguments["area"] = as_rect(arguments["area"]).translated(self.area.x, self.area.y)
                except InputError as error:
                    raise self.slide._locate(error, element_id) from None
            for key in ("start", "end"):
                point = arguments.get(key)
                if (
                    isinstance(point, (list, tuple))
                    and len(point) == 2
                    and all(isinstance(value, (int, float)) and not isinstance(value, bool) for value in point)
                ):
                    arguments[key] = (point[0] + self.area.x, point[1] + self.area.y)
            return getattr(self.slide, name)(id=id, **arguments)

        return add


class Stack:
    """要素を領域の上から順に、間隔を空けて置く補助（`slide.stack()`・`group.stack()` の戻り値）。

    `text`・`box`・`equation`・`image`・`table` は元の関数と同じ引数をとり、`area` の代わりに `height`・`width`・
    `indent`・`place`・`gap` を受け取る。幅の既定は領域の幅から `indent` を除いた幅（画像で高さだけを指定したときは縦横比から）、
    高さの既定は内容から決める。`place` は幅が領域より狭いときの横位置（left・center・right。文字の揃えの align とは別）、
    `gap` はその要素の後の間隔（既定は stack の gap）である。領域の高さが数値なら、置いた要素に `layout_bounds` を記録し、
    領域を超えた要素は配置の検査が `layout_bounds_overflow` にする。`take` で受け取った領域が超えた分は、場面の `issues` に
    同じ識別子の診断として入る。まとまりの stack では、座標の原点はまとまりの左上である。

    `avoid` は避ける領域（ページの座標の `Rect` か、"footer"（ページ番号の領域）か、その列）である。幅を指定しない要素が
    避ける領域に重なるときは、領域の左に収まるように幅を縮めて作り直す。左に領域の幅の半分以上が残らなければ、
    要素を避ける領域の下へ送る。幅を指定した要素は動かさない（重なりは配置の検査が示す）。
    """

    def __init__(self, slide, area, gap=12, *, owner=None, avoid=None):
        self._slide, self._owner = slide, owner
        try:
            area = as_rect(area)
        except InputError as error:
            raise slide._locate(error, None) from None
        if area.w is None or area.w <= 0:
            raise self._error(f"積み上げる領域の幅は正の数値で指定してください: {list(area)}", area=list(area))
        self._area = area
        self._gap = self._spacing(gap)
        self._cursor = area.y
        self._elements = []
        self._avoid = self._regions(avoid)

    # ---- 状態 ----

    @property
    def area(self) -> Rect:
        """積み上げる領域。"""
        return self._area

    @property
    def gap(self) -> float:
        """既定の間隔。"""
        return self._gap

    @property
    def cursor(self) -> float:
        """次の要素の上端。"""
        return self._cursor

    @property
    def remaining(self) -> Rect:
        """残りの領域。領域の高さが None なら高さも None。"""
        area = self._area
        height = None if area.h is None else max(0.0, area.y + area.h - self._cursor)
        return Rect(area.x, self._cursor, area.w, height)

    @property
    def elements(self) -> tuple:
        """置いた要素（表はまとまり）。"""
        return tuple(self._elements)

    # ---- 要素を置く ----

    def text(self, *, id, height=None, width=None, indent=0, place="left", gap=None, **arguments):
        return self._place("text", id, height, width, indent, place, gap, arguments)

    def box(self, *, id, height=None, width=None, indent=0, place="left", gap=None, **arguments):
        return self._place("box", id, height, width, indent, place, gap, arguments)

    def equation(self, *, id, height=None, width=None, indent=0, place="left", gap=None, **arguments):
        return self._place("equation", id, height, width, indent, place, gap, arguments)

    def image(self, *, id, height=None, width=None, indent=0, place="left", gap=None, **arguments):
        return self._place("image", id, height, width, indent, place, gap, arguments)

    def table(self, *, id, height=None, width=None, indent=0, place="left", gap=None, **arguments):
        return self._place("table", id, height, width, indent, place, gap, arguments)

    def take(self, height, *, gap=None) -> Rect:
        """現在の位置から高さ height の領域を返して進む（受け取った領域は自分で使う）。

        領域の高さが数値で、受け取った領域がその下端を超えるときも例外にはしない。超えた分は、場面の `issues` に
        `layout_bounds_overflow` として入る（置いた要素が超えたときと同じく、診断で報告する）。
        """
        height = self._size(height, "height")
        gap = self._spacing(gap, optional=True)
        area = self._area
        taken = Rect(area.x, self._cursor, area.w, height)
        if area.h is not None and taken.bottom > area.bottom + _EPSILON:
            self._slide._taken_overflows.append((self, taken))
        self._cursor += height + (self._gap if gap is None else gap)
        return taken

    def space(self, amount) -> None:
        """位置を amount だけ下げる。"""
        number = _option_number(amount)
        if number is None or number < 0:
            raise self._error(f"積み上げに足す間隔は 0 以上の数値で指定してください: {_shown(amount)}")
        self._cursor += number

    # ---- 内部 ----

    def _error(self, message, element_id=None, *, code="invalid_geometry", **details) -> InputError:
        return self._slide._locate(InputError(message, code=code, **details), element_id)

    def _spacing(self, value, element_id=None, *, optional=False):
        """間隔（0 以上の有限値）。optional なら None（stack の gap を使う）を受け付ける。"""
        if value is None and optional:
            return None
        number = _option_number(value)
        if number is None or number < 0:
            raise self._error(f"積み上げる間隔は 0 以上にしてください: {_shown(value)}", element_id)
        return number

    def _size(self, value, name, element_id=None, *, optional=False):
        """寸法（正の有限値）。optional なら None（内容から決める）を受け付ける。"""
        if value is None and optional:
            return None
        number = _option_number(value)
        if number is None or number <= 0:
            raise self._error(f"{name} は正の有限値にしてください: {_shown(value)}", element_id)
        return number

    def _regions(self, avoid) -> list[Rect]:
        """避ける領域（ページの座標）。"footer" はページ番号の領域にし、番号を置かないページでは除く。"""
        if avoid is None:
            return []
        items = [avoid] if isinstance(avoid, (str, Rect)) else list(avoid) if isinstance(avoid, (list, tuple)) else None
        if items is None or (
            isinstance(avoid, tuple) and len(avoid) == 4 and all(_option_number(v) is not None for v in avoid)
        ):
            items = [avoid]
        regions = []
        for item in items:
            if item == "footer":
                if self._slide.footer_area is not None:
                    regions.append(self._slide.footer_area)
                continue
            try:
                region = as_rect(item)
            except (InputError, TypeError, ValueError):
                raise self._error(f'avoid は Rect か "footer" か、その列で指定してください: {_shown(item)}') from None
            if region.w is None or region.h is None:
                raise self._error(f"避ける領域の幅と高さは数値で指定してください: {list(region)}")
            regions.append(region)
        return regions

    def _blocking(self, element) -> Rect | None:
        """要素（ページの座標）に重なる避ける領域のうち、最も上のもの。"""
        x, y = self._origin()
        rect = element.area.translated(x, y)
        hits = [
            region
            for region in self._avoid
            if min(rect.right, region.right) - max(rect.x, region.x) > _EPSILON
            and min(rect.bottom, region.bottom) - max(rect.y, region.y) > _EPSILON
        ]
        return min(hits, key=lambda region: region.y) if hits else None

    def _origin(self) -> tuple[float, float]:
        """領域の座標の原点（ページでは (0, 0)、まとまりではその左上）。"""
        if self._owner is None:
            return 0.0, 0.0
        area = self._owner.area
        return area.x, area.y

    def _place(self, kind, id, height, width, indent, place, gap, arguments):
        slide, owner = self._slide, self._owner
        element_id = f"{owner.id}/{id}" if owner is not None and isinstance(id, str) else id
        if "area" in arguments:
            raise slide._locate(
                ArgumentNameError(
                    f"area は stack の {kind} では使えません。寸法は height・width で指定してください",
                    code="unsupported_argument",
                    names=["area"],
                ),
                element_id,
            )
        height = self._size(height, "height", element_id, optional=True)
        width = self._size(width, "width", element_id, optional=True)
        gap = self._spacing(gap, element_id, optional=True)
        indent_value = _option_number(indent)
        if indent_value is None or indent_value < 0:
            raise self._error(f"indent は 0 以上の数値で指定してください: {_shown(indent)}", element_id)
        if place not in _PLACES:
            raise self._error(
                f"place は {'、'.join(_PLACES)} のいずれかにしてください: {_shown(place)}",
                element_id,
                code="input_error",
            )
        area = self._area
        available = area.w - indent_value
        if available <= 0:
            raise self._error(f"indent {indent_value:g} pt が積み上げる領域の幅 {area.w:g} pt 以上です", element_id)
        requested = width
        if width is None and not (kind == "image" and height is not None):
            width = available
        start = len(slide.elements)
        try:
            if width is not None and width > available + _EPSILON:
                raise self._error(
                    f"幅 {width:g} pt が積み上げる領域の幅（indent を除いて {available:g} pt）を超えています",
                    element_id,
                    width=width,
                    available=available,
                )
            create = getattr(slide if owner is None else owner, kind)
            element = create(id=id, area=Rect(area.x + indent_value, self._cursor, width, height), **arguments)
            if requested is None and self._avoid:
                element = self._clear(element, create, id, start, height, indent_value, available, arguments)
            placed = element.area.w
            if placed > available + _EPSILON:
                # 画像の縦横比や表の列幅から決まった幅が、領域に収まらない場合。
                label = "画像の縦横比から決めた幅" if kind == "image" else "置いた要素の幅"
                raise self._error(
                    f"{label} {placed:g} pt が積み上げる領域の幅（indent を除いて {available:g} pt）を"
                    "超えています。width を指定してください",
                    element_id,
                    width=placed,
                    available=available,
                )
            offset = {"left": 0.0, "center": (available - placed) / 2, "right": available - placed}[place]
            if offset:
                element.move(dx=offset)
            x, y = self._origin()
            if area.h is not None:
                element.data["layout_bounds"] = list(area.translated(x, y))
        except BaseException:
            # 失敗した呼び出しは要素を残さない（表のまとまりとセルも除く）。
            del slide.elements[start:]
            raise
        self._cursor = element.area.y + element.area.h - y + (self._gap if gap is None else gap)
        self._elements.append(element)
        return element

    def _clear(self, element, create, id, start, height, indent, available, arguments):
        """避ける領域に重なる要素を、幅を縮めるか領域の下へ送って作り直す（最大 4 回）。"""
        slide, area = self._slide, self._area
        ox, oy = self._origin()
        left = area.x + indent
        width, top = element.area.w, element.area.y
        for _ in range(4):
            region = self._blocking(element)
            if region is None:
                break
            room = region.x - ox - min(self._gap, 8.0) - left
            if region.x - ox > left + _EPSILON and room >= available / 2:
                width = room
            else:
                top = max(top, region.bottom - oy + self._gap)
                width = available
            del slide.elements[start:]
            element = create(id=id, area=Rect(left, top, width, height), **arguments)
        return element


class Deck:
    def __init__(self, context):
        self.project = context
        self.slides = []
        # collect_errors が真のときに集めた、続行できる入力の誤り（InputError）。
        self.issues = []
        self._bibliography = None

    @property
    def collect_errors(self) -> bool:
        """続行できる誤りを集めるか。資料の collect_errors（ない資料では偽）を使う。"""
        return bool(getattr(self.project, "collect_errors", False))

    @property
    def bibliography(self):
        """出典と文献一覧に使う文献（`bibliography.Bibliography`）。

        資料の `bibliography`（`Project.bibliography`。構築の流れが先に読む）を使い、それを持たない資料では
        inputs.bibliography から 1 回だけ読む。
        """
        if self._bibliography is None:
            found = getattr(self.project, "bibliography", None)
            if found is None:
                from .bibliography import load_bibliography

                found = load_bibliography(self.project.config.get("inputs", {}).get("bibliography", ()))
            self._bibliography = found
        return self._bibliography

    def slide(
        self,
        *,
        id,
        kind="content",
        title="",
        notes="",
        seconds=None,
        title_role=None,
        title_area=None,
        title_style=None,
        decorations=True,
        allow_roles=(),
    ):
        """ページを作る。title は文字列・断片・断片の列で、None は題名なし。notes の None は空のノート。

        title_role・title_area・title_style は、このページの題名の役割・領域・書式（ページ種別の値に重ねる）。
        decorations を偽にすると、このページだけページ種別の飾り（テーマの decorations）を置かない。
        allow_roles はこのページで使う「使う場所を限った役割」の宣言である。
        """
        _identifier(id, "ページの識別子", code="invalid_slide_id")
        if any(slide.id == id for slide in self.slides):
            raise InputError(f"ページの識別子が重複しています: {id}", code="invalid_slide_id", slide_id=id)
        if kind not in PAGE_KINDS:
            raise InputError(
                f"ページ {id}: 未対応のページ種別です: {_shown(kind)}"
                f"{styles.format_candidates(styles.suggest(kind, PAGE_KINDS))}",
                code="input_error",
                slide_id=id,
            )
        if not isinstance(decorations, bool):
            raise InputError(
                f"ページ {id}: decorations は真偽値で指定してください: {_shown(decorations)}",
                code="input_error",
                slide_id=id,
            )
        slide = Slide(
            self,
            id,
            kind,
            title,
            notes,
            seconds,
            title_role=title_role,
            title_area=title_area,
            title_style=title_style,
            decorations=decorations,
            allow_roles=allow_roles,
        )
        self.slides.append(slide)
        return slide

    def cited_keys(self) -> list[str]:
        """作ったページの出典の文献キー（bib_key）を、ページ順・要素順で初めて現れた順に返す。"""
        keys: dict[str, None] = {}
        for slide in self.slides:
            for element in slide.elements:
                for source in element.data.get("sources", ()):
                    key = source.get("bib_key") if isinstance(source, Mapping) else None
                    if key:
                        keys.setdefault(key, None)
        return list(keys)

    def to_dict(self):
        _page_size(self.project)
        scene = {
            "schema_version": 1,
            "size": list(self.project.size),
            "metadata": copy.deepcopy(self.project.config["metadata"]),
            "theme": copy.deepcopy(self.project.theme),
            "rules": copy.deepcopy(self.project.rules),
            "slides": [slide.to_dict(number, len(self.slides)) for number, slide in enumerate(self.slides, 1)],
        }
        has_math = any(
            fragment["kind"] == "math"
            for slide in scene["slides"]
            for element in slide["elements"]
            for fragment in element.get("content", [])
        )
        if (
            has_math
            and self.project.config["math"]["fallback"] == "off"
            and "libreoffice" in self.project.config["render"]["required"]
        ):
            raise InputError("LibreOffice が必須の数式資料では互換画像を off にできません", code="input_error")
        overflows = [(number, slide) for number, slide in enumerate(self.slides, 1) if slide._taken_overflows]
        if self.issues or overflows:
            from .validators import make_diagnostic, merge_diagnostics

            numbers = {slide.id: number for number, slide in enumerate(self.slides, 1)}
            issues = [
                make_diagnostic(
                    error.code,
                    str(error),
                    severity="error",
                    scope="input",
                    slide_id=error.slide_id,
                    element_id=error.element_id,
                    observations=error.details or None,
                    slide_number=numbers.get(error.slide_id),
                )
                for error in self.issues
            ]
            for number, slide in overflows:
                issues.extend(slide._taken_overflow_issues(number, make_diagnostic))
            scene["issues"] = merge_diagnostics(issues)
        return scene


class Slide:
    def __init__(
        self,
        deck,
        id,
        kind,
        title,
        notes,
        seconds,
        *,
        title_role=None,
        title_area=None,
        title_style=None,
        decorations=True,
        allow_roles=(),
    ):
        self.deck, self.id, self.kind = deck, id, kind
        self.notes, self.seconds = notes, seconds
        self._decorations = decorations
        self.elements = []
        self._allowed_warnings = []
        # stack.take で受け取った領域のうち、積み上げる領域を超えたもの（(stack, 受け取った領域)。座標は stack の座標）。
        self._taken_overflows = []
        theme, (w, h) = self.project.theme, _page_size(self.project)
        self.kind_settings = copy.deepcopy(theme["page_kinds"][kind])
        if "content_area_pt" in self.kind_settings:
            self.content = as_rect(self.kind_settings["content_area_pt"])
        elif "content_area" in self.kind_settings:
            self.content = self._fraction_area(self.kind_settings["content_area"])
        else:
            margins = theme["margins"]
            top = margins["left"] if kind == "blank" else margins["top"]
            self.content = Rect(
                margins["left"], top, w - margins["left"] - margins["right"], h - top - margins["bottom"]
            )
        self.allow_roles = self._allow_roles(allow_roles, None)
        # 題名の領域がないページでも、題名の役割の名前を確かめる（誤りを集めるときは、ページ種別の役割を使う）。
        if title_role is not None and (not isinstance(title_role, str) or title_role not in theme["roles"]):
            title_role = self._problem(self._unknown_role(title_role, "_theme/title"))
        self._title_role = title_role
        self._title_area = None if title_area is None else self._element_area(title_area, "text", "_theme/title")
        if title_style is not None and not isinstance(title_style, Mapping):
            raise InputError(
                f"ページ {id}: title_style は書式の辞書で指定してください: {_shown(title_style)}",
                code="invalid_style",
                slide_id=id,
                element_id="_theme/title",
            )
        self._title_style = self._known_names("text", title_style or {}, "_theme/title")
        self.title = title

    @property
    def project(self):
        return self.deck.project

    @property
    def default_role(self) -> str:
        """ページ種別の既定の役割。"""
        return self.kind_settings.get("default_role", "body")

    @property
    def title(self) -> str:
        """ページの題名（通常の文字列。数式の断片は $…$ で囲む）。"""
        return self._title_text

    @title.setter
    def title(self, value):
        self._title_text, self._title = self._resolve_title(value)

    # ---- 誤りの扱い ----

    def _context(self, element_id) -> str:
        return f"ページ {self.id} の要素 {element_id}" if element_id else f"ページ {self.id}"

    def _locate(self, error, element_id):
        """例外にページと要素の識別子を入れ、場所が付いていなければ文言の先頭に付ける。"""
        if getattr(error, "slide_id", None) is None:
            error.slide_id = self.id
        if element_id is not None and getattr(error, "element_id", None) is None:
            error.element_id = element_id
        context = self._context(error.element_id)
        message = str(error)
        if not message.startswith(f"ページ {self.id}"):
            error.args = (f"{context}: {message}", *error.args[1:])
        return error

    def _problem(self, error, fallback=None):
        """続行できる誤り。集めるときは記録して fallback を返し、集めないときは送出する。"""
        if not self.deck.collect_errors:
            raise error
        self.deck.issues.append(error)
        return fallback

    def _known_names(self, kind, values, element_id, *, extra=(), arguments=False) -> dict:
        """受け付けない書式名を確かめる。集めるときは、その名前を除いた辞書を返す。"""
        try:
            styles.check_names(
                kind,
                values,
                extra=extra,
                context=self._context(element_id),
                slide_id=self.id,
                element_id=element_id,
                arguments=arguments,
            )
        except ArgumentNameError as error:
            self._problem(error)
            rejected = set(error.details.get("names", ()))
            return {name: value for name, value in values.items() if name not in rejected}
        return dict(values)

    # ---- 書式 ----

    def _fraction_area(self, values):
        x, y, w, h = values
        sw, sh = self.project.size
        return Rect(x * sw, y * sh, w * sw, h * sh)

    def _restricted(self, role) -> bool:
        definition = self.project.theme["roles"].get(role, {})
        return bool(definition.get("restricted", False)) or definition.get("allowed_slide_ids") is not None

    def _role_setting(self, role) -> str:
        """役割の設定の場所（例：theme.toml の roles.formula）。"""
        config = getattr(self.project, "config", {})
        theme = config.get("theme") if isinstance(config, Mapping) else None
        name = Path(theme).name if isinstance(theme, str) and theme != "default" else "同梱テーマ"
        return f"{name} の roles.{role}"

    def _unknown_role(self, name, element_id) -> InputError:
        roles = self.project.theme["roles"]
        return InputError(
            f"{self._context(element_id)}: 未定義の文字の役割です: {_shown(name)}"
            f"{styles.format_candidates(styles.suggest(name, roles))}",
            code="unknown_role",
            slide_id=self.id,
            element_id=element_id,
            role=_shown(name),
        )

    def _resolve_role(self, role, *, content, element_id, declared, default=None) -> str:
        """役割の名前を確かめる。誤りを集めるときは、既定の役割に替えて続ける。"""
        roles = self.project.theme["roles"]
        fallback = self.default_role if default is None else default
        name = fallback if role is None else role
        context = self._context(element_id)
        if not isinstance(name, str) or name not in roles:
            name = self._problem(self._unknown_role(name, element_id), fallback)
        else:
            problems = role_usage_problems(name, roles[name], slide_id=self.id, declared=declared, content=content)
            if problems:
                setting = self._role_setting(name)
                error = InputError(
                    f"{context}: " + "。".join(problem.describe(setting=setting) for problem in problems),
                    code="role_usage",
                    slide_id=self.id,
                    element_id=element_id,
                    role=name,
                    setting=setting,
                    problems=[problem.kind for problem in problems],
                )
                name = self._problem(error, fallback)
        if name not in roles:
            raise InputError(
                f"{context}: 既定の文字の役割 {name} がテーマにありません",
                code="unknown_role",
                slide_id=self.id,
                element_id=element_id,
                role=name,
            )
        return name

    def _color_names(self) -> str:
        """色の誤りの文言に添える、テーマの色名の一覧。"""
        names = "、".join(sorted(self.project.theme.get("colors", {}))) or "なし"
        return f"（テーマの色: {names}）"

    def _value_error(self, name, problem, value, element_id, *, label=None) -> InputError:
        """書式の値の誤り（invalid_style）。label は文言の書式名（既定は name）。色の項目にはテーマの色名の一覧を添える。"""
        shown = _shown(value)
        detail = f": {shown}{self._color_names()}" if name in styles.COLOR_KEYS else f"（値: {shown}）"
        return InputError(
            f"{self._context(element_id)}: {name if label is None else label} {problem}{detail}",
            code="invalid_style",
            slide_id=self.id,
            element_id=element_id,
            setting=name,
            value=shown,
        )

    def _color_problem(self, name, value, element_id, *, label=None):
        """色の値（形は確かめた後）の誤り。none は fill・line_color だけに指定できる。"""
        context = self._context(element_id)
        label = name if label is None else label
        if value == "none":
            if name in styles.PAINT_KEYS:
                return None
            return InputError(
                f"{context}: {label} に none は指定できません。none は fill・line_color だけに使えます{self._color_names()}",
                code="invalid_style",
                slide_id=self.id,
                element_id=element_id,
                setting=name,
                value="none",
            )
        if styles.HEX_COLOR.fullmatch(value) or value in self.project.theme.get("colors", {}):
            return None
        return InputError(
            f"{context}: {label} に未定義の色を指定しています: {_shown(value)}{self._color_names()}",
            code="undefined_color",
            slide_id=self.id,
            element_id=element_id,
            setting=name,
            value=_shown(value),
        )

    def _clean(self, values, element_id, definition=None) -> dict:
        """要素が指定した書式の値を確かめる。集めるときは、間違った項目を除いた辞書を返す（None は未指定のまま残す）。

        definition（役割）を渡すと、役割と違う min_size_pt を誤りにする。
        """
        context = self._context(element_id)
        result = {}
        for name, value in values.items():
            error = None
            problem = styles.check_style_value(name, value)
            if problem:
                error = self._value_error(name, problem, value, element_id)
            elif value is not None and name in styles.COLOR_KEYS:
                error = self._color_problem(name, value, element_id)
            elif (
                value is not None
                and name == "min_size_pt"
                and definition is not None
                and value != definition.get("min_size_pt")
            ):
                error = InputError(
                    f"{context}: 要素から最小文字サイズを変更できません。資料側で役割を定義してください",
                    code="minimum_override",
                    slide_id=self.id,
                    element_id=element_id,
                    specified_minimum=value,
                    role_minimum=definition.get("min_size_pt"),
                )
            if error is not None:
                self._problem(error)
                continue
            result[name] = value
        return result

    def _hex(self, value, name, element_id):
        """色を 16 進数（大文字、# なし）にする。塗らない・線なし（None と "none"）は None。"""
        if value is None or value == "none":
            return None
        colors = self.project.theme.get("colors", {})
        resolved = colors.get(value, value) if isinstance(value, str) else value
        if not isinstance(resolved, str) or not _HEX.fullmatch(resolved.lstrip("#")):
            raise InputError(
                f"{self._context(element_id)}: {name} に未定義の色を指定しています: {_shown(value)}"
                f"{self._color_names()}",
                code="undefined_color",
                slide_id=self.id,
                element_id=element_id,
                setting=name,
                value=_shown(value),
            )
        return resolved.lstrip("#").upper()

    def _style(
        self, kind, role, layers=(), *, content=(), element_id=None, declared=(), default_role=None, role_layer=True
    ):
        """書式を解決し、(役割の名前, 書式) を返す。

        kind は書式の種類（`styles.STYLE_KEYS` の名前。None はテーマの背景）。layers は要素の層の列 (値の辞書, 確かめるか) で、
        後ろほど優先する。確かめる層では、値の誤りを `_clean` で扱う。padding_pt の一部の辺は下の層の値を引き継ぐ。
        role_layer が偽なら役割の書式を重ねない（ページ種別の飾りは種類の既定に飾りの指定を重ねる）。
        """
        role = self._resolve_role(role, content=content, element_id=element_id, declared=declared, default=default_role)
        definition = self.project.theme["roles"][role]
        theme = self.project.theme
        element_layers = [
            styles.element_layer(self._clean(values, element_id, definition) if checked else values, kind)
            for values, checked in layers
        ]
        resolved, padding = {}, 0.0
        for layer in (
            styles.library_layer(theme),
            *styles.theme_layers(theme, kind),
            *([styles.role_layer(definition)] if role_layer else []),
            *element_layers,
        ):
            for name, value in layer.items():
                if name == "padding_pt":
                    try:
                        padding = normalize_padding(value, base=padding)
                    except InputError as error:
                        raise self._locate(error, element_id) from None
                    value = padding
                resolved[name] = value
        if "min_size_pt" in definition:
            resolved["min_size_pt"] = definition["min_size_pt"]
        for name, value in resolved.items():
            if name == "padding_pt" or name in styles.COLOR_KEYS:
                continue
            problem = styles.check_style_value(name, value)
            if problem:
                raise self._value_error(name, problem, value, element_id)
        for name in _NUMBER_KEYS:
            if resolved.get(name) is not None:
                resolved[name] = float(resolved[name])
        for name in ("color", "fill", "line_color", "emphasis_color"):
            resolved[name] = self._hex(resolved.get(name), name, element_id)
        return role, resolved

    def _fragments(self, fragments, style, element_id):
        """断片の色を 16 進数にし、リンクの形式を確かめ、数式の強調（emphasis）を太字と要素の強調色に解決する。"""
        context = self._context(element_id)
        result = []
        for fragment in fragments:
            fragment = copy.deepcopy(fragment)
            own = fragment.pop("style", None) or {}
            if "color" in own:
                value, label = own["color"], "断片の color"
                problem = styles.check_style_value("color", value)
                if problem:
                    error = self._value_error("color", problem, value, element_id, label=label)
                else:
                    error = self._color_problem("color", value, element_id, label=label)
                if error is None:
                    own["color"] = self._hex(value, "color", element_id)
                else:
                    self._problem(error)
                    own.pop("color")
            if "link" in fragment and not valid_link(fragment["link"]):
                self._problem(
                    InputError(
                        f"{context}: link は https://・http://・mailto: で始まり、空白を含まない文字列で指定してください: "
                        f"{_shown(fragment['link'])}",
                        code="invalid_link",
                        slide_id=self.id,
                        element_id=element_id,
                        link=_shown(fragment["link"]),
                    )
                )
                fragment.pop("link")
            if fragment.get("kind") == "math" and fragment.pop("emphasis", False):
                own.setdefault("bold", True)
                if style.get("emphasis_color") is not None:
                    own.setdefault("color", style["emphasis_color"])
            if own:
                fragment["style"] = own
            result.append(fragment)
        return result

    # ---- 要素の共通の処理 ----

    def _allow_roles(self, value, element_id) -> tuple:
        """使う場所を限った役割の宣言を確かめる。重複は順序を保って除く。"""
        if value is None:
            return ()
        context = self._context(element_id)
        if isinstance(value, str) or not isinstance(value, (list, tuple)):
            hint = f'（1 つでも ["{value}"] のように書きます）' if isinstance(value, str) else ""
            raise InputError(
                f"{context}: allow_roles は役割名の list か tuple で指定してください{hint}: {_shown(value)}",
                code="input_error",
                slide_id=self.id,
                element_id=element_id,
            )
        roles = self.project.theme["roles"]
        result = []
        for name in value:
            if not isinstance(name, str) or name not in roles:
                self._problem(
                    InputError(
                        f"{context}: allow_roles の役割 {_shown(name)} はテーマにありません"
                        f"{styles.format_candidates(styles.suggest(name, roles))}",
                        code="unknown_role",
                        slide_id=self.id,
                        element_id=element_id,
                        role=_shown(name),
                    )
                )
            elif name not in result:
                result.append(name)
        return tuple(result)

    def _declared(self, own, parent) -> set:
        """要素で宣言済みの役割：ページ、要素、祖先のまとまりの allow_roles の和。"""
        declared = set(self.allow_roles) | set(own)
        lookup = {element.id: element for element in self.elements} if parent is not None else {}
        seen = set()
        while parent is not None and parent.id not in seen:
            seen.add(parent.id)
            declared.update(parent.data.get("allow_roles", ()))
            parent = lookup.get(parent.data.get("parent_id"))
        return declared

    def _element_area(self, area, kind, element_id, *, unknown=False):
        """要素の領域を確かめる。unknown が真なら、幅・高さの None（内容から決める寸法）を残して返す。"""
        try:
            area = as_rect(area)
        except InputError as error:
            raise self._locate(error, element_id) from None
        context = self._context(element_id)
        if (area.w is None or area.h is None) and not unknown:
            hint = (
                "（内容から決める寸法は text・box・equation・image と table で使えます）"
                if kind in ("line", "arrow", "group")
                else ""
            )
            raise InputError(
                f"{context}: 領域の幅と高さを数値で指定してください{hint}: {list(area)}",
                code="invalid_geometry",
                slide_id=self.id,
                element_id=element_id,
                area=list(area),
            )
        if kind not in ("line", "arrow", "group") and any(
            value is not None and value <= 0 for value in (area.w, area.h)
        ):
            raise InputError(
                f"{context}: 文字、図形、画像の幅と高さは正にしてください: {list(area)}",
                code="invalid_geometry",
                slide_id=self.id,
                element_id=element_id,
                area=list(area),
            )
        return area

    def _layout(self, kind, fragments, style, *, width=None, shape=None):
        """解決済みの書式と場面の断片を計測する（`measure.layout_text`）。equation で数式が 1 つだけなら独立した数式として計測する。"""
        from .math.office_math import math_mode

        display = kind == "equation" and math_mode({"kind": "equation", "content": fragments}) == "display"
        return layout_text(
            fragments,
            style,
            self.project.fonts,
            width=width,
            shape=(shape or "rectangle") if kind == "box" else "rectangle",
            display=display,
        )

    def _resolve_area(self, area, kind, element_id, *, fragments, style, shape, intrinsic, source):
        """領域の幅・高さの None を内容（文字・図形・数式）か縦横比（画像）から数値にする。

        幅は折り返さない内容の幅と左右の距離（折り返す要素は計測の誤差の帯を含む）、高さは指定した幅で計測した
        文字ブロックと上下の距離。図形の文字領域は図形の種類と余白から求める（`measure.frame_size`）。
        """
        if area.w is not None and area.h is not None:
            return area
        context = self._context(element_id)
        if kind == "image":
            if area.w is None and area.h is None:
                raise InputError(
                    f"{context}: 画像の幅か高さの一方を指定してください: {list(area)}",
                    code="invalid_geometry",
                    slide_id=self.id,
                    element_id=element_id,
                    area=list(area),
                )
            if intrinsic is None:
                label = "高さ" if area.h is None else "幅"
                raise InputError(
                    f"{context}: 画像の縦横比が分からないため、{label}を内容から決められません: {source}",
                    code="invalid_geometry",
                    slide_id=self.id,
                    element_id=element_id,
                    path=str(source),
                )
            image_w, image_h = intrinsic
            if area.h is None:
                return Rect(area.x, area.y, area.w, area.w * image_h / image_w)
            return Rect(area.x, area.y, area.h * image_w / image_h, area.h)
        try:
            layout = self._layout(kind, fragments, style, width=area.w, shape=shape)
        except ValueError as error:
            raise self._locate(error, element_id) from None
        if area.w is None and layout.content_width <= 0:
            raise InputError(
                f"{context}: 内容のない要素は幅を内容から決められません。幅を数値で指定してください: {list(area)}",
                code="invalid_geometry",
                slide_id=self.id,
                element_id=element_id,
                area=list(area),
            )
        return Rect(
            area.x, area.y, layout.width if area.w is None else area.w, layout.height if area.h is None else area.h
        )

    def _attributes(self, values, element_id) -> dict:
        """要素の属性（書式でない引数）を確かめる。None は未指定として除く。"""
        context = self._context(element_id)
        result = {}
        for name, value in values.items():
            if value is None:
                continue
            if name == "rotation":
                if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
                    raise InputError(
                        f"{context}: rotation は有限の数値（度）で指定してください: {_shown(value)}",
                        code="invalid_geometry",
                        slide_id=self.id,
                        element_id=element_id,
                    )
                result[name] = float(value)
            elif name in ("flip_h", "flip_v"):
                if not isinstance(value, bool):
                    raise InputError(
                        f"{context}: {name} は真偽値で指定してください: {_shown(value)}",
                        code="invalid_geometry",
                        slide_id=self.id,
                        element_id=element_id,
                    )
                result[name] = value
            elif name == "description":
                if not isinstance(value, str):
                    raise InputError(
                        f"{context}: description は文字列で指定してください: {_shown(value)}",
                        code="input_error",
                        slide_id=self.id,
                        element_id=element_id,
                    )
                result[name] = value
            elif name == "semantic":
                # 文字列でない値（list など）も、ほかの間違った値と同じく続行できる誤りにする。
                if isinstance(value, str) and value in styles.SEMANTIC_VALUES:
                    result[name] = value
                else:
                    self._problem(
                        InputError(
                            f"{context}: semantic は {'、'.join(sorted(styles.SEMANTIC_VALUES))} のどちらかで指定してください: "
                            f"{_shown(value)}",
                            code="input_error",
                            slide_id=self.id,
                            element_id=element_id,
                        )
                    )
            elif name == "allow_roles":
                roles = self._allow_roles(value, element_id)
                if roles:
                    result[name] = list(roles)
        return result

    def _sources(self, sources, element_id):
        if sources is None:
            return []
        if not isinstance(sources, (list, tuple)):
            sources = [sources]
        context = self._context(element_id)
        documents = {d["id"]: d for d in self.project.config["inputs"].get("documents", [])}
        result = []
        for source in sources:
            if isinstance(source, str):
                source = {"document_id": source} if source in documents else {"bib_key": source}
            if not isinstance(source, dict) or not (source.get("document_id") or source.get("bib_key")):
                raise InputError(
                    f"{context}: 出典には document_id または bib_key を指定してください: {_shown(source)}",
                    code="source_missing",
                    slide_id=self.id,
                    element_id=element_id,
                )
            if source.get("document_id") and source["document_id"] not in documents:
                self._problem(
                    InputError(
                        f"{context}: 未定義の元資料です: {source['document_id']}",
                        code="source_missing",
                        slide_id=self.id,
                        element_id=element_id,
                        source=copy.deepcopy(source),
                    )
                )
                continue
            # 文献キーの判定は変換を伴わない（文献項目の LaTeX 表記は文献一覧で参照したときに変換する）。
            if source.get("bib_key") and (
                not isinstance(source["bib_key"], str) or source["bib_key"] not in self.deck.bibliography
            ):
                self._problem(
                    InputError(
                        f"{context}: 未定義の文献キーです: {source['bib_key']}",
                        code="source_missing",
                        slide_id=self.id,
                        element_id=element_id,
                        source=copy.deepcopy(source),
                    )
                )
                continue
            result.append(copy.deepcopy(source))
        return result

    def _add(
        self,
        kind,
        *,
        id,
        area,
        role=None,
        content=None,
        sources=None,
        style=None,
        _parent_id=None,
        _element_id=None,
        _style_kind=None,
        _shape=None,
        _intrinsic=None,
        _source=None,
        **arguments,
    ):
        """要素を作る共通の処理。arguments は属性と書式のキーワード引数である。

        _parent_id・_element_id はまとまりの子の親と識別子、_style_kind は書式の種類（既定は kind。表のセルは
        table_cell・table_header）、_shape は box の図形、_intrinsic・_source は画像の縦横比に使う寸法と素材のパスである。
        領域の幅・高さの None は書式を解決した後に `_resolve_area` で数値にする。
        """
        if kind not in _KINDS:
            raise InputError(
                f"ページ {self.id}: 未対応の要素です: {_shown(kind)}", code="unknown_element", slide_id=self.id
            )
        try:
            _identifier(id)
        except InputError as error:
            raise self._locate(error, None) from None
        element_id = _element_id or id
        context = self._context(element_id)
        if any(element.id == element_id for element in self.elements):
            raise InputError(
                f"{context}: 要素の識別子が重複しています: {element_id}",
                code="invalid_element_id",
                slide_id=self.id,
                element_id=element_id,
            )
        parent = None
        if _parent_id is not None:
            parent = next((e for e in self.elements if e.id == _parent_id and e.data["kind"] == "group"), None)
            if parent is None:
                raise InputError(
                    f"{context}: 親のまとまりが見つかりません: {_parent_id}",
                    code="parent_missing",
                    slide_id=self.id,
                    element_id=element_id,
                    parent_id=_parent_id,
                )
        area = self._element_area(area, kind, element_id, unknown=kind in _SIZED_KINDS)
        style_kind = _style_kind or kind
        attribute_names = styles.attribute_keys(kind)
        if kind == "group":
            transformed = [name for name in _TRANSFORMS if arguments.get(name) not in (None, 0, False)]
            if transformed:
                raise InputError(
                    f"{context}: まとまりは回転・反転できません（{'、'.join(transformed)}）",
                    code="unsupported_argument",
                    slide_id=self.id,
                    element_id=element_id,
                    names=transformed,
                )
            for name in _TRANSFORMS:
                arguments.pop(name, None)
        extra = (
            attribute_names
            | set(_SPECIFIC[kind])
            | set(_COMMON)
            | ({"area"} if kind not in ("line", "arrow") else set())
        )
        arguments = self._known_names(style_kind, arguments, element_id, extra=extra, arguments=True)
        attributes = self._attributes(
            {name: arguments.pop(name) for name in list(arguments) if name in attribute_names}, element_id
        )
        if style is not None and not isinstance(style, Mapping):
            raise InputError(
                f"{context}: style は書式の辞書で指定してください: {_shown(style)}",
                code="invalid_style",
                slide_id=self.id,
                element_id=element_id,
            )
        style = self._known_names(style_kind, style or {}, element_id)
        if (
            style_kind == "box"
            and _shape != "rounded_rectangle"
            and any(values.get("corner_radius_pt") is not None for values in (style, arguments))
        ):
            self._problem(
                InputError(
                    f"{context}: corner_radius_pt は shape が rounded_rectangle の box だけに指定できます（shape: {_shape}）",
                    code="invalid_style",
                    slide_id=self.id,
                    element_id=element_id,
                    setting="corner_radius_pt",
                )
            )
            style.pop("corner_radius_pt", None)
            arguments.pop("corner_radius_pt", None)
        try:
            raw = content_fragments(content)
        except InputError as error:
            raise self._locate(error, element_id) from None
        declared = self._declared(attributes.get("allow_roles", ()), parent)
        role, resolved = self._style(
            style_kind, role, [(style, True), (arguments, True)], content=raw, element_id=element_id, declared=declared
        )
        fragments = self._fragments(raw, resolved, element_id)
        if any(attributes.get(name) for name in _TRANSFORMS) and any(
            fragment["kind"] == "math" for fragment in fragments
        ):
            raise InputError(
                f"{context}: 数式を含む要素は回転・反転できません",
                code="math_transform",
                slide_id=self.id,
                element_id=element_id,
            )
        area = self._resolve_area(
            area,
            kind,
            element_id,
            fragments=fragments,
            style=resolved,
            shape=_shape,
            intrinsic=_intrinsic,
            source=_source,
        )
        data = {
            "id": element_id,
            "kind": kind,
            "area": list(area),
            "role": role,
            "style": resolved,
            "content": fragments,
            **attributes,
        }
        if parent is not None:
            data["parent_id"] = parent.id
        if sources:
            found = self._sources(sources, element_id)
            if found:
                data["sources"] = found
        if kind == "box":
            data["shape"] = _shape
        element = (Group if kind == "group" else Element)(self, data)
        if kind == "image":
            element._intrinsic = None if _intrinsic is None else tuple(_intrinsic)
        self.elements.append(element)
        return element

    # ---- 要素を作る関数 ----

    def text(self, *, id, text, area, role=None, **arguments):
        return self._add("text", id=id, area=area, role=role, content=text, **arguments)

    def box(self, *, id, area, text="", shape="rectangle", role=None, **arguments):
        shape = "rectangle" if shape is None else shape
        if shape not in styles.SHAPES:
            element_id = arguments.get("_element_id") or id
            raise InputError(
                f"{self._context(element_id)}: 未対応の図形です: {_shown(shape)}（候補: {'、'.join(styles.SHAPES)}）",
                code="input_error",
                slide_id=self.id,
                element_id=element_id,
            )
        return self._add("box", id=id, area=area, role=role, content=text, _shape=shape, **arguments)

    def equation(self, *, id, area, expression=None, role=None, **arguments):
        element_id = arguments.get("_element_id") or id
        context = self._context(element_id)
        if "text" in arguments:
            raise ArgumentNameError(
                f"{context}: text は equation では使えません。数式は expression で指定してください",
                code="unsupported_argument",
                slide_id=self.id,
                element_id=element_id,
                names=["text"],
            )
        if not isinstance(expression, str) or not expression:
            raise InputError(
                f"{context}: 数式を expression に空でない文字列で指定してください: {_shown(expression)}",
                code="invalid_fragment",
                slide_id=self.id,
                element_id=element_id,
            )
        # 役割を指定しない数式は、テーマの役割 equation（同梱テーマは body を継承）を使う。
        if role is None and "equation" in self.project.theme["roles"]:
            role = "equation"
        return self._add("equation", id=id, area=area, role=role, content=Math(expression), **arguments)

    def group(self, *, id, area=None, role=None, pptx_group=None, **arguments):
        """まとまりを作る。子は戻り値の同じ名前の関数で作る（座標はまとまりの左上が原点）。

        pptx_group が None か真なら、条件（描く子孫がある、子孫がまとまりの直後に続く、子孫に数式がない）を満たすときに
        PowerPoint のグループとして書き出す。偽なら書き出さない（子はスライドの図形として書き出す）。
        """
        if pptx_group is not None and not isinstance(pptx_group, bool):
            element_id = arguments.get("_element_id") or id
            raise InputError(
                f"{self._context(element_id)}: pptx_group は真偽値か None で指定してください: {_shown(pptx_group)}",
                code="input_error",
                slide_id=self.id,
                element_id=element_id,
            )
        element = self._add("group", id=id, area=Rect(0, 0, 0, 0) if area is None else area, role=role, **arguments)
        if pptx_group is False:
            element.data["pptx_group"] = False
        return element

    def _line(self, kind, *, id, start, end, **arguments):
        element_id = arguments.get("_element_id") or id
        context = self._context(element_id)
        if "area" in arguments:
            raise ArgumentNameError(
                f"{context}: area は {kind} では使えません。位置は start と end で指定してください",
                code="unsupported_argument",
                slide_id=self.id,
                element_id=element_id,
                names=["area"],
            )

        def point(value, name):
            if isinstance(value, Anchor):
                if value.element.slide is not self:
                    raise InputError(
                        f"{context}: 別のページの要素へ接続できません: {value.element.id}",
                        code="connection_missing",
                        slide_id=self.id,
                        element_id=element_id,
                    )
                return value.point()
            if isinstance(value, (list, tuple)) and len(value) == 2:
                try:
                    return (finite_number(value[0], f"{name} x"), finite_number(value[1], f"{name} y"))
                except ValueError:
                    pass
            raise InputError(
                f"{context}: 線の端点は (x, y) または要素の anchor で指定してください: {_shown(value)}",
                code="invalid_geometry",
                slide_id=self.id,
                element_id=element_id,
            )

        a, b = point(start, "start"), point(end, "end")
        if a == b:
            raise InputError(
                f"{context}: 線の始点と終点を別にしてください",
                code="invalid_geometry",
                slide_id=self.id,
                element_id=element_id,
            )
        element = self._add(
            kind, id=id, area=Rect(min(a[0], b[0]), min(a[1], b[1]), abs(b[0] - a[0]), abs(b[1] - a[1])), **arguments
        )
        element._start, element._end = start, end
        return element

    def line(self, *, id, start, end, **arguments):
        return self._line("line", id=id, start=start, end=end, **arguments)

    def arrow(self, *, id, start, end, **arguments):
        return self._line("arrow", id=id, start=start, end=end, **arguments)

    def image(self, *, id, path, area, fit="contain", page=None, raster_scale=None, vector=None, **arguments):
        """画像を置く。area の幅か高さの一方を None にすると、もう一方を画像の縦横比から決める。

        raster_scale（PDF・SVG を PNG にするときの配置寸法 1 pt あたりの画素数、1 以上 16 以下）と vector（SVG でも埋め込むか）は
        PDF・SVG の素材だけに指定でき、PDF の素材には vector=True を指定できない。None は deck.toml の [images] の値を使う。
        """
        element_id = arguments.get("_element_id") or id
        context = self._context(element_id)
        fit = "contain" if fit is None else fit
        if fit not in ("contain", "cover"):
            raise InputError(
                f"{context}: 画像の fit は contain または cover にしてください: {_shown(fit)}",
                code="image_fit",
                slide_id=self.id,
                element_id=element_id,
            )
        try:
            resolved = self.project.resolve_input(path)
        except FileNotFoundError as error:
            raise InputError(
                f"{context}: {error}", code="asset_missing", slide_id=self.id, element_id=element_id, path=str(path)
            ) from error
        except InputError as error:
            raise self._locate(error, element_id) from None
        except ValueError as error:
            raise InputError(
                f"{context}: {error}", code="input_error", slide_id=self.id, element_id=element_id, path=str(path)
            ) from error
        suffix = resolved.suffix.lower()
        if suffix not in _IMAGE_FORMATS:
            raise InputError(
                f"{context}: 未対応の画像形式です: {suffix}",
                code="input_error",
                slide_id=self.id,
                element_id=element_id,
            )
        if suffix == ".pdf" and (isinstance(page, bool) or not isinstance(page, int) or page <= 0):
            raise InputError(
                f"{context}: PDF の画像には 1 から数える page を指定してください",
                code="input_error",
                slide_id=self.id,
                element_id=element_id,
            )
        options = self._image_options(suffix, raster_scale, vector, element_id)
        try:
            intrinsic = _intrinsic_size(resolved, page)
        except Exception as error:  # 画像を読み込むことができない（形式の誤り、破損したファイル、PDF にないページ）
            raise InputError(
                f"{context}: 画像を読み込めません: {resolved}（{error}）",
                code="input_error",
                slide_id=self.id,
                element_id=element_id,
                path=str(resolved),
            ) from error
        element = self._add("image", id=id, area=area, _intrinsic=intrinsic, _source=resolved, **arguments)
        element.data.update(path=str(resolved), fit=fit)
        if page is not None:
            element.data["page"] = page
        if suffix in _RASTER_FORMATS:
            element.data["image_size"] = list(intrinsic)
        element.data.update(options)
        return element

    def _image_options(self, suffix, raster_scale, vector, element_id) -> dict:
        """画像ごとの raster_scale・vector を確かめ、場面に書く項目（指定したものだけ）を返す。"""
        context = self._context(element_id)

        def failure(message, **details):
            return InputError(
                f"{context}: {message}", code="invalid_image_option", slide_id=self.id, element_id=element_id, **details
            )

        given = [name for name, value in (("raster_scale", raster_scale), ("vector", vector)) if value is not None]
        if given and suffix not in (".pdf", ".svg"):
            raise failure(
                f"{'・'.join(given)} は PDF・SVG の素材だけに指定できます（{suffix} の素材は画素数のまま使います）",
                names=given,
            )
        options = {}
        if raster_scale is not None:
            number = _option_number(raster_scale)
            if number is None or not 1 <= number <= 16:
                raise failure(
                    f"raster_scale は 1 以上 16 以下の数で指定してください: {_shown(raster_scale)}",
                    value=_shown(raster_scale),
                )
            options["raster_scale"] = number
        if vector is not None:
            if not isinstance(vector, bool):
                raise failure(f"vector は真偽値で指定してください: {_shown(vector)}", value=_shown(vector))
            if vector and suffix == ".pdf":
                from .assets import PDF_VECTOR_MESSAGE

                raise failure(PDF_VECTOR_MESSAGE, names=["vector"])
            options["vector"] = vector
        return options

    def table(self, **arguments):
        """表を作る。引数は tables.add_table と同じ。"""
        from .tables import add_table

        return add_table(self, **arguments)

    def references(
        self,
        *,
        id,
        keys=None,
        area=None,
        formatter=None,
        gap_pt=8,
        role=None,
        pptx_group=False,
        _parent_id=None,
        _element_id=None,
        **style,
    ):
        """文献一覧を作り、項目をまとめたまとまりを返す。

        keys は載せる文献キーの順で、None なら `deck.cited_keys()`（それまでに作ったページの出典の文献キー）。
        空・重複・未定義のキーは誤り（`invalid_reference`）にする。area は既定でページの内容の領域で、高さは None でもよい。
        各項目は幅 area.w で高さを計測し（`Rect(x, y, w, None)` の文字と同じ）、gap_pt の間隔で上から積む。領域の下端を越えても
        縮めず、配置の検査（`layout_bounds_overflow`）にまかせる。formatter は `bibliography.Reference` を受け取って断片の列を
        返す関数で、既定は `bibliography.format_reference`。
        項目の識別子は文献キー（/ は _ に置き換える）で、出典にその文献キーを付ける。role の既定はページの既定の役割で、
        style は各項目の書式である。まとまりは既定で PowerPoint のグループにしない（項目を個別に編集しやすくするため）。
        戻り値のまとまりの area は実際に積んだ範囲（幅は area.w）である。
        """
        from .bibliography import format_reference

        element_id = _element_id or id
        context = self._context(element_id)

        def failure(message, code="invalid_reference", **details):
            return InputError(f"{context}: {message}", code=code, slide_id=self.id, element_id=element_id, **details)

        reserved = sorted(set(style) & {"text", "sources", "height", "width", "indent", "place", "gap"})
        if reserved:
            raise ArgumentNameError(
                f"{context}: {'、'.join(reserved)} は references では使えません（項目の文章と出典は文献から作り、"
                "位置と間隔は area と gap_pt で決めます）",
                code="unsupported_argument",
                slide_id=self.id,
                element_id=element_id,
                names=reserved,
            )
        if keys is None:
            keys = self.deck.cited_keys()
            if not keys:
                raise failure(
                    "文献一覧に載せる文献がありません。keys を指定するか、要素の出典に文献キーを付けてください"
                )
        elif isinstance(keys, str) or not isinstance(keys, (list, tuple)):
            raise failure(f"keys は文献キーの list か tuple で指定してください: {_shown(keys)}")
        keys = list(keys)
        if not keys:
            raise failure("keys に文献キーを 1 つ以上指定してください")
        if any(not isinstance(key, str) or not key for key in keys):
            raise failure(f"文献キーは空でない文字列で指定してください: {_shown(keys)}")
        repeated = sorted({key for key in keys if keys.count(key) > 1})
        if repeated:
            raise failure(f"文献キーが重複しています: {'、'.join(repeated)}", keys=repeated)
        bibliography = self.deck.bibliography
        for key in keys:
            if key not in bibliography:
                try:
                    bibliography[key]  # 候補付きの誤り（BibliographyKeyError）
                except KeyError as error:
                    raise self._locate(error, element_id) from None
        names = {key: key.replace("/", "_") for key in keys}
        collided = sorted({name for name in names.values() if list(names.values()).count(name) > 1})
        if collided:
            raise failure(f"/ を _ に置き換えた項目の識別子が重複しています: {'、'.join(collided)}", names=collided)
        if formatter is None:
            formatter = format_reference
        elif not callable(formatter):
            raise failure(
                f"formatter は文献項目から断片の列を返す関数で指定してください: {_shown(formatter)}", code="input_error"
            )
        gap = _option_number(gap_pt)
        if gap is None or gap < 0:
            raise failure(f"gap_pt は 0 以上の数値で指定してください: {_shown(gap_pt)}", code="invalid_geometry")
        area = self.content if area is None else self._element_area(area, "text", element_id, unknown=True)
        if area.w is None:
            raise failure(
                f"文献一覧の領域の幅を数値で指定してください: {list(area)}", code="invalid_geometry", area=list(area)
            )
        start = len(self.elements)
        try:
            group = self.group(
                id=id,
                area=Rect(area.x, area.y, area.w, 0),
                pptx_group=pptx_group,
                _parent_id=_parent_id,
                _element_id=_element_id,
            )
            stack = Stack(self, Rect(0, 0, area.w, area.h), gap, owner=group)
            for key in keys:
                item_id = f"{group.id}/{names[key]}"
                try:
                    content = formatter(bibliography[key])
                except InputError as error:
                    raise self._locate(error, item_id) from None
                stack.text(id=names[key], text=content, role=role, sources=[{"bib_key": key}], **style)
            bottom = max(element.area.bottom for element in stack.elements)
            # まとまりの領域は data に直接書く（area の代入は子を動かすため）。
            group.data["area"] = [area.x, area.y, area.w, bottom - area.y]
        except BaseException:
            # 失敗した呼び出しは、まとまりも作った項目も残さない。
            del self.elements[start:]
            raise
        return group

    # ---- 寸法を決める補助 ----

    def measure(self, text, *, kind="text", width=None, role=None, shape="rectangle", style=None, **arguments):
        """要素を作るときと同じ書式の解決で内容を計測し、`measure.TextLayout` を返す。要素は作らない。

        text は `slide.text` の text と同じ形（kind が equation なら式の文字列）。width は枠の幅（pt）で、None なら折り返さない。
        結果の height は `Rect(x, y, w, None)` で決まる高さ、width は `Rect(x, y, None, h)` で決まる幅である。
        shape と書式の corner_radius_pt は kind が box のときの図形。書式は style= の辞書とキーワード引数で指定する。
        使う場所を限った役割も宣言なしで計測できる（要素を置くときに宣言する）。
        """
        context = self._context(None)
        if kind not in _MEASURED_KINDS:
            raise InputError(
                f"{context}: kind は text、box、equation のいずれかにしてください: {_shown(kind)}",
                code="input_error",
                slide_id=self.id,
            )
        if width is not None:
            number = _option_number(width)
            if number is None or number <= 0:
                raise InputError(
                    f"{context}: 幅は正の有限値にしてください: {_shown(width)}",
                    code="invalid_geometry",
                    slide_id=self.id,
                )
            width = number
        shape = "rectangle" if shape is None else shape
        if kind == "box" and shape not in styles.SHAPES:
            raise InputError(
                f"{context}: 未対応の図形です: {_shown(shape)}（候補: {'、'.join(styles.SHAPES)}）",
                code="input_error",
                slide_id=self.id,
            )
        if kind != "box" and shape != "rectangle":
            raise InputError(
                f"{context}: shape は kind が box のときだけ指定できます: {_shown(shape)}",
                code="input_error",
                slide_id=self.id,
            )
        if kind == "equation":
            if not isinstance(text, str) or not text:
                raise InputError(
                    f"{context}: 数式を空でない文字列で指定してください: {_shown(text)}",
                    code="invalid_fragment",
                    slide_id=self.id,
                )
            text = Math(text)
            if role is None and "equation" in self.project.theme["roles"]:
                role = "equation"
        if style is not None and not isinstance(style, Mapping):
            raise InputError(
                f"{context}: style は書式の辞書で指定してください: {_shown(style)}",
                code="invalid_style",
                slide_id=self.id,
            )
        style = self._known_names(kind, style or {}, None)
        arguments = self._known_names(kind, arguments, None, arguments=True)
        if (
            kind == "box"
            and shape != "rounded_rectangle"
            and any(values.get("corner_radius_pt") is not None for values in (style, arguments))
        ):
            self._problem(
                InputError(
                    f"{context}: corner_radius_pt は shape が rounded_rectangle の box だけに指定できます（shape: {shape}）",
                    code="invalid_style",
                    slide_id=self.id,
                    setting="corner_radius_pt",
                )
            )
            style.pop("corner_radius_pt", None)
            arguments.pop("corner_radius_pt", None)
        try:
            raw = content_fragments(text)
        except InputError as error:
            raise self._locate(error, None) from None
        _, resolved = self._style(
            kind, role, [(style, True), (arguments, True)], content=raw, declared=set(self.project.theme["roles"])
        )
        fragments = self._fragments(raw, resolved, None)
        try:
            return self._layout(kind, fragments, resolved, width=width, shape=shape)
        except ValueError as error:
            raise self._locate(error, None) from None

    @property
    def footer_area(self) -> Rect | None:
        """ページ番号の領域（ページ種別の footer_area_pt、なければページの右下の幅 36 pt・高さ 20 pt）。番号を置かないページは None。"""
        settings = self.kind_settings
        if not settings.get("page_number", False):
            return None
        w, h = _page_size(self.project)
        return as_rect(settings["footer_area_pt"]) if "footer_area_pt" in settings else Rect(w - 72, h - 28, 36, 20)

    def stack(self, area, *, gap=12, avoid=None):
        """area の上から順に、gap の間隔で要素を置く補助（`Stack`）を返す。area の高さは None でもよい。

        avoid は避ける領域（`Rect` か "footer"（ページ番号の領域）か、その列）。幅を指定しない要素が重なるときは、
        幅を縮めるか、領域の下へ送る。
        """
        return Stack(self, area, gap, avoid=avoid)

    def _taken_overflow_issues(self, number, make_diagnostic) -> list[dict]:
        """stack.take で受け取った領域が積み上げる領域を超えた分の診断（layout_bounds_overflow）。

        座標はページの座標で、まとまりの stack では、場面にするときのまとまりの位置で換算する（後で動かしても合う）。
        observations は配置の検査の layout_bounds_overflow と同じ形（area・bounds・overflow_pt）である。
        """
        issues = []
        for stack, taken in self._taken_overflows:
            x, y = stack._origin()
            area, bounds = taken.translated(x, y), stack.area.translated(x, y)
            overflow = {"side": "bottom", "amount": round(area.bottom - bounds.bottom, 2)}
            issues.append(
                make_diagnostic(
                    "layout_bounds_overflow",
                    "stack.take で受け取った領域が積み上げる領域を超えている",
                    slide_id=self.id,
                    slide_number=number,
                    observations={"area": list(area), "bounds": list(bounds), "overflow_pt": overflow},
                )
            )
        return issues

    # ---- 題名とテーマの要素 ----

    def _resolve_title(self, value):
        """題名を (場面の通常の文字列, _theme/title の要素か None) にする。"""
        if value is None:
            return "", None
        element_id = "_theme/title"
        try:
            fragments = content_fragments(value)
        except InputError as error:
            raise self._locate(error, element_id) from None
        plain = _plain_text(fragments)
        settings = self.kind_settings
        if self._title_area is not None:
            area = self._title_area
        elif "title_area_pt" in settings:
            area = as_rect(settings["title_area_pt"])
        elif "title_area" in settings:
            area = self._fraction_area(settings["title_area"])
        else:
            area = None
        if not plain or area is None:
            return plain, None
        designated = settings.get("title_role", "slide_title")
        overridden = self._title_role is not None
        role = self._title_role if overridden else designated
        # ページ種別が決めた題名の役割は、テーマ自身の指定なので宣言済みとする。上書きした役割は宣言が要る。
        declared = set(self.allow_roles) | (set() if overridden else {designated})
        role, style = self._style(
            "text",
            role,
            [(settings.get("title_style") or {}, False), (self._title_style, True)],
            content=fragments,
            element_id=element_id,
            declared=declared,
            default_role=designated,
        )
        element = {
            "id": element_id,
            "kind": "text",
            "area": list(area),
            "role": role,
            "content": self._fragments(fragments, style, element_id),
            "style": style,
            "semantic": "title",
        }
        if not overridden and self._restricted(role):
            element["allow_roles"] = [role]
        return plain, element

    def _theme_elements(self, number, total=None):
        """テーマの要素を描く順（背景、ページ種別の飾り、題名、ページ番号）に返す。

        ページ番号の文字は、ページ種別の page_number_format の {page} をページ番号に、{total} を全体のページ数に置き換えて作る。
        total を省くと、全体のページ数をページ番号と同じとみなす。

        ページ種別の background が "none" なら背景の図形を作らない（雛形のマスターの図形を見せるとき）。
        """
        w, h = _page_size(self.project)
        settings = self.kind_settings
        result = []
        fill = settings.get("background", "background")
        if fill != "none":
            _, background_style = self._style(
                None,
                "body",
                [({"fill": fill, "line_color": "none"}, False)],
                element_id="_theme/background",
                declared={"body"},
            )
            background = {
                "id": "_theme/background",
                "kind": "box",
                "area": [0, 0, w, h],
                "role": "body",
                "content": [],
                "shape": "rectangle",
                "semantic": "background",
                "style": background_style,
            }
            if self._restricted("body"):
                background["allow_roles"] = ["body"]
            result.append(background)
        if self._decorations:
            result.extend(self._decoration_elements())
        if self._title is not None:
            result.append(copy.deepcopy(self._title))
        if settings.get("page_number", False):
            label = settings.get("page_number_format", "{page}").format(
                page=number, total=number if total is None else total
            )
            content = content_fragments(label)
            area = self.footer_area
            _, style = self._style(
                "text",
                "note",
                [({"align": "right"}, False), (settings.get("footer_style") or {}, False)],
                content=content,
                element_id="_theme/footer",
                declared={"note"},
            )
            footer = {
                "id": "_theme/footer",
                "kind": "text",
                "area": list(area),
                "role": "note",
                "content": content,
                "style": style,
                "semantic": "footer",
            }
            if self._restricted("note"):
                footer["allow_roles"] = ["note"]
            result.append(footer)
        return result

    def _decoration_elements(self) -> list[dict]:
        """ページ種別の飾り（テーマの page_kinds.<種別>.decorations）の要素。識別子は _theme/decoration/<id>。

        位置は box が area（ページの幅・高さに対する比率）か area_pt、線が start・end（比率）か start_pt・end_pt である。
        書式は種類の既定（ライブラリ、[defaults]、[defaults.<種類>]）に飾りの指定を重ね、役割の書式は重ねない。
        設定の形と色の参照はテーマの読み込み（`schema.check_decorations`）で確かめてある。
        """
        w, h = _page_size(self.project)
        result = []
        for item in self.kind_settings.get("decorations") or ():
            kind = item["kind"]
            element_id = f"_theme/decoration/{item['id']}"
            values = {name: value for name, value in item.items() if name in styles.STYLE_KEYS[kind]}
            _, style = self._style(
                kind, "body", [(values, False)], element_id=element_id, declared={"body"}, role_layer=False
            )
            element = {
                "id": element_id,
                "kind": kind,
                "role": "body",
                "content": [],
                "style": style,
                "semantic": "decoration",
            }
            if kind == "box":
                area = as_rect(item["area_pt"]) if "area_pt" in item else self._fraction_area(item["area"])
                element.update(area=list(area), shape=item.get("shape", "rectangle"))
            else:
                if "start_pt" in item:
                    start, end = item["start_pt"], item["end_pt"]
                else:
                    start, end = [(point[0] * w, point[1] * h) for point in (item["start"], item["end"])]
                start, end = [float(start[0]), float(start[1])], [float(end[0]), float(end[1])]
                element.update(
                    area=[min(start[0], end[0]), min(start[1], end[1]), abs(end[0] - start[0]), abs(end[1] - start[1])],
                    start=start,
                    end=end,
                )
            if self._restricted("body"):
                element["allow_roles"] = ["body"]
            result.append(element)
        return result

    # ---- 警告の許容の宣言 ----

    def allow_warning(self, code, *, element=None, related=None, reason):
        """このページの警告 code を、理由を付けて許容すると宣言する。宣言した警告は、確認記録なしで最終検査を通る。

        element・related は要素か要素の識別子で、related は組の警告（element_overlap など）だけに指定できる。
        組の警告では element が必須である。それ以外の警告で element を省くと、そのページの同じ種類の警告すべてが対象になる。
        宣言の誤りは集めず、常にその場で `InputError` にする（許容の範囲が変わらないように）。
        """
        # 診断の表はここで読み込む（`import slidekit` が読み込むと `python -m slidekit.diagnostics` が警告を出すため）。
        from .diagnostics import DIAGNOSTICS, WARNING_CODES

        context = f"ページ {self.id}"

        def failure(message, **details):
            return InputError(f"{context}: {message}", code="input_error", slide_id=self.id, **details)

        if not isinstance(code, str) or code not in WARNING_CODES:
            if isinstance(code, str) and code in DIAGNOSTICS and DIAGNOSTICS[code].severity == "error":
                raise failure(f"エラーは許容できない: {code}", diagnostic=code)
            raise failure(
                f"許容できる警告の種類ではない: {_shown(code)}{styles.format_candidates(styles.suggest(code, WARNING_CODES))}",
                diagnostic=_shown(code),
            )
        pair = DIAGNOSTICS[code].evidence.pair
        element_id = self._declared_target(element, "element", failure)
        related_id = self._declared_target(related, "related", failure)
        if pair and element_id is None:
            raise failure(
                f"{code} は 2 つの要素の組の警告なので、element に対象の要素を指定してください（ページ全体には宣言できません）",
                diagnostic=code,
            )
        if related_id is not None and not pair:
            pairs = sorted(name for name in WARNING_CODES if DIAGNOSTICS[name].evidence.pair)
            raise failure(f"related は組の警告（{'、'.join(pairs)}）だけに指定できます: {code}", diagnostic=code)
        if related_id is not None and related_id == element_id:
            raise failure(f"element と related に同じ要素を指定しています: {element_id}", diagnostic=code)
        if not isinstance(reason, str) or not reason.strip():
            raise failure(f"reason に許容する理由を書いてください: {_shown(reason)}", diagnostic=code)
        targets = {element_id, related_id}
        if any(
            item["code"] == code and {item["element_id"], item["related_element_id"]} == targets
            for item in self._allowed_warnings
        ):
            raise failure(f"同じ警告の許容がすでに宣言されています: {code}", diagnostic=code)
        frame = sys._getframe(1)
        self._allowed_warnings.append(
            {
                "code": code,
                "element_id": element_id,
                "related_element_id": related_id,
                "reason": reason,
                "location": self._location(frame.f_code.co_filename, frame.f_lineno),
            }
        )

    def _declared_target(self, value, name, failure):
        if value is None:
            return None
        if isinstance(value, Element):
            if value.slide is not self:
                raise failure(f"{name} に別のページの要素を指定しています: {value.id}")
            return value.id
        if isinstance(value, str) and value.strip():
            return value
        raise failure(f"{name} は要素か要素の識別子で指定してください: {_shown(value)}")

    def _location(self, filename, line) -> str:
        """宣言した場所。資料のディレクトリからの相対パスと行（資料の外なら絶対パス）。"""
        path = Path(filename)
        root = getattr(self.project, "root", None)
        if root is not None and path.is_absolute():
            try:
                relative = Path(os.path.relpath(path, Path(root)))
            except ValueError:
                relative = path
            if not str(relative).startswith(".."):
                path = relative
        return f"{path.as_posix()}:{line}"

    # ---- 場面のページ ----

    def to_dict(self, number, total=None):
        notes = "" if self.notes is None else self.notes
        if not isinstance(notes, str):
            raise InputError(
                f"ページ {self.id}: 発表者ノートは文字列にしてください", code="input_error", slide_id=self.id
            )
        try:
            seconds = None if self.seconds is None else finite_number(self.seconds, "説明時間")
        except ValueError as error:
            raise InputError(f"ページ {self.id}: {error}", code="input_error", slide_id=self.id) from None
        if seconds is not None and seconds < 0:
            raise InputError(f"ページ {self.id}: 説明時間は 0 以上にしてください", code="input_error", slide_id=self.id)
        elements = self._theme_elements(number, total)
        lookup = {element.id: element for element in self.elements}
        for element in self.elements:
            data = copy.deepcopy(element.data)
            context = self._context(element.id)
            # 移動後に文字枠・画像の寸法をもう一度確かめる。
            area = element.area
            if data["kind"] not in ("line", "arrow", "group") and (area.w <= 0 or area.h <= 0):
                raise InputError(
                    f"{context}: 要素の寸法が不正です: {list(area)}",
                    code="invalid_geometry",
                    slide_id=self.id,
                    element_id=element.id,
                )
            if data["kind"] in ("line", "arrow"):
                for key, value in (("start", element._start), ("end", element._end)):
                    if isinstance(value, Anchor):
                        data[key] = list(value.point())
                        data[key + "_ref"] = value.reference()
                    else:
                        data[key] = list(value)
                a, b = data["start"], data["end"]
                if a == b:
                    raise InputError(
                        f"{context}: 移動後の線の始点と終点が一致しています",
                        code="invalid_geometry",
                        slide_id=self.id,
                        element_id=element.id,
                    )
                data["area"] = [min(a[0], b[0]), min(a[1], b[1]), abs(a[0] - b[0]), abs(a[1] - b[1])]
            if data["kind"] == "image" and "image_size" in data:
                # PNG・JPEG の表示範囲（PDF・SVG は素材の準備が変換後の画素数で決める）。
                data["actual_area"] = list(element._visible_area())
            style = data.get("style") or {}
            if data["kind"] in ("text", "box", "equation") and style.get("phrase_wrap") and style.get("word_wrap"):
                # 文節の境で折り返した位置に明示改行を入れる（PowerPoint は和文の文字の間ならどこでも折り返すため）。
                try:
                    data["content"] = phrase_wrapped_content(data, self.project.fonts)
                except InputError:
                    pass  # 計測できない要素は、配置の検査が診断で示す
            if any(fragment["kind"] == "math" for fragment in data["content"]):
                current, seen = element, set()
                while current:
                    if current.id in seen:
                        raise InputError(
                            f"{context}: 親子関係が循環しています",
                            code="parent_cycle",
                            slide_id=self.id,
                            element_id=element.id,
                        )
                    seen.add(current.id)
                    if (
                        current.data.get("rotation", 0)
                        or current.data.get("flip_h", False)
                        or current.data.get("flip_v", False)
                    ):
                        raise InputError(
                            f"{context}: 数式を含む要素または親は回転・反転できません",
                            code="math_transform",
                            slide_id=self.id,
                            element_id=element.id,
                            transformed_element=current.id,
                        )
                    current = lookup.get(current.data.get("parent_id"))
            elements.append(data)
        # PowerPoint のグループとして書き出すまとまりに pptx_group を書く（利用者が無効にしたまとまりは偽のまま）。
        grouped = _pptx_groups(elements)
        for data in elements:
            if data["id"] in grouped:
                data["pptx_group"] = True
        page = {
            "id": self.id,
            "number": number,
            "kind": self.kind,
            "title": self.title,
            "notes": notes,
            "seconds": seconds,
            "content_area": list(self.content),
            "elements": elements,
        }
        if self.allow_roles:
            page["allow_roles"] = list(self.allow_roles)
        if self._allowed_warnings:
            ids = [item["id"] for item in elements]
            for declaration in self._allowed_warnings:
                for key in ("element_id", "related_element_id"):
                    target = declaration[key]
                    if target is not None and target not in ids:
                        raise InputError(
                            f"ページ {self.id}: 許容の宣言（{declaration['code']}、{declaration['location']}）の要素 "
                            f"{target} がページにありません{styles.format_candidates(styles.suggest(target, ids))}",
                            code="input_error",
                            slide_id=self.id,
                            element_id=target,
                        )
            page["allowed_warnings"] = copy.deepcopy(self._allowed_warnings)
        return page
