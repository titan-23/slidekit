"""ポイント単位の領域計算。内容や文字サイズは変更しない。

`Rect` の幅・高さの `None` は「内容から決める寸法」を表す。その寸法を使う操作（`right`・`bottom`・`center`・`anchor`・
`contain`、`columns` の幅、`rows` の高さ、その向きの `inset`）はエラーにする。領域の誤りと
接続位置（`anchor`）の指定の誤りは、それぞれ `invalid_geometry`・`invalid_anchor` の `InputError` にする。
"""

from __future__ import annotations

from dataclasses import dataclass
from math import isfinite
from typing import Iterable, Sequence

from .errors import InputError


def finite_number(value: float, name: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not isfinite(value):
        raise ValueError(f"{name} は有限の数値で指定してください")
    return float(value)


def _geometry_error(message: str, **details) -> InputError:
    return InputError(message, code="invalid_geometry", **details)


def _shown(value) -> str:
    text = repr(value)
    return text if len(text) <= 40 else text[:39] + "…"


def _coordinate(value, name: str) -> float:
    """座標・寸法・移動量の数値。有限の数でなければ invalid_geometry の InputError にする。"""
    try:
        return finite_number(value, name)
    except ValueError as error:
        raise _geometry_error(f"{error}: {_shown(value)}", value=_shown(value)) from None


ANCHOR_SIDES = ("left", "right", "top", "bottom", "center")


def anchor_options(side, gap=0, at=0.5) -> tuple[str, float, float]:
    """接続位置の指定（辺、辺から外側への距離 gap、辺の上の位置 at）を確かめて返す。誤りは invalid_anchor にする。

    at は上下の辺では左端 0〜右端 1、左右の辺では上端 0〜下端 1 である。center には gap・at を指定できない。
    """
    if not isinstance(side, str) or side not in ANCHOR_SIDES:
        raise InputError(
            f"未対応の接続位置です: {_shown(side)}（候補: {'、'.join(ANCHOR_SIDES)}）",
            code="invalid_anchor",
            side=_shown(side),
        )
    try:
        gap_value = finite_number(gap, "gap")
    except ValueError:
        gap_value = -1.0
    if gap_value < 0:
        raise InputError(
            f"gap は 0 以上の数値で指定してください: {_shown(gap)}", code="invalid_anchor", gap=_shown(gap)
        )
    try:
        at_value = finite_number(at, "at")
    except ValueError:
        at_value = -1.0
    if not 0 <= at_value <= 1:
        raise InputError(f"at は 0 以上 1 以下で指定してください: {_shown(at)}", code="invalid_anchor", at=_shown(at))
    if side == "center" and (gap_value != 0 or at_value != 0.5):
        raise InputError(
            "center には gap・at を指定できません（辺の接続位置だけに指定できます）",
            code="invalid_anchor",
            side=side,
            gap=gap_value,
            at=at_value,
        )
    return side, gap_value, at_value


@dataclass(frozen=True)
class Rect:
    x: float
    y: float
    w: float | None
    h: float | None

    def __post_init__(self):
        for name in ("x", "y", "w", "h"):
            value = getattr(self, name)
            if value is None and name in ("w", "h"):
                continue
            object.__setattr__(self, name, _coordinate(value, name))
        if (self.w is not None and self.w < 0) or (self.h is not None and self.h < 0):
            raise _geometry_error("領域の幅と高さは負にできません")

    def __iter__(self):
        return iter((self.x, self.y, self.w, self.h))

    def _known(self, name: str, operation: str) -> float:
        """寸法の値。内容から決める寸法（None）なら、その寸法を使う操作としてエラーにする。"""
        value = getattr(self, name)
        if value is None:
            label = "幅" if name == "w" else "高さ"
            raise _geometry_error(
                f"{label}を内容から決める領域（{name}=None）では {operation} を求められません。"
                "要素を作成した後の element.area を使ってください",
                dimension=name,
                operation=operation,
            )
        return value

    @property
    def right(self):
        return self.x + self._known("w", "right")

    @property
    def bottom(self):
        return self.y + self._known("h", "bottom")

    @property
    def center(self):
        return (self.x + self._known("w", "center") / 2, self.y + self._known("h", "center") / 2)

    def inset(self, amount=0, *, left=None, right=None, top=None, bottom=None):
        """指定した余白を四辺から除く。内容から決める寸法の向きには、余白を指定できない。"""
        values = [amount if v is None else v for v in (left, right, top, bottom)]
        l, r, t, b = (finite_number(v, "余白") for v in values)
        if min(l, r, t, b) < 0:
            raise ValueError("余白は 0 以上にしてください")
        w = None if self.w is None and not (l or r) else self._known("w", "inset") - l - r
        h = None if self.h is None and not (t or b) else self._known("h", "inset") - t - b
        return Rect(self.x + l, self.y + t, w, h)

    def columns(self, weights: Sequence[float] | int, gap=12):
        return self._split(weights, gap, horizontal=True)

    def rows(self, weights: Sequence[float] | int, gap=12):
        return self._split(weights, gap, horizontal=False)

    def _split(self, weights, gap, horizontal):
        if isinstance(weights, int):
            weights = [1] * weights
        weights = [finite_number(v, "分割比率") for v in weights]
        gap = finite_number(gap, "間隔")
        if not weights or min(weights) <= 0 or gap < 0:
            raise ValueError("分割比率は正、間隔は 0 以上にしてください")
        length = self._known("w", "columns") if horizontal else self._known("h", "rows")
        available = length - gap * (len(weights) - 1)
        if available <= 0:
            raise ValueError("分割する領域より間隔が大きくなっています")
        cursor = self.x if horizontal else self.y
        result = []
        total = sum(weights)
        for weight in weights:
            length = available * weight / total
            # 分けない向きの寸法は、内容から決める寸法（None）もそのまま引き継ぐ。
            result.append(Rect(cursor, self.y, length, self.h) if horizontal else Rect(self.x, cursor, self.w, length))
            cursor += length + gap
        return tuple(result)

    def translated(self, dx=0, dy=0):
        return Rect(self.x + _coordinate(dx, "横移動"), self.y + _coordinate(dy, "縦移動"), self.w, self.h)

    def anchor(self, side="center", *, gap=0, at=0.5):
        """辺または中央の点 (x, y)。gap は辺から外側への距離、at は辺の上の位置（`anchor_options`）。

        left は (x - gap, y + h × at)、right は (x + w + gap, y + h × at)、top は (x + w × at, y - gap)、
        bottom は (x + w × at, y + h + gap)、center は中央である。既定値では 0.1 と同じ点を返す。
        """
        side, gap, at = anchor_options(side, gap, at)
        w, h = self._known("w", "anchor"), self._known("h", "anchor")
        if side == "center":
            return (self.x + w / 2, self.y + h / 2)
        if side == "left":
            return (self.x - gap, self.y + h * at)
        if side == "right":
            return (self.x + w + gap, self.y + h * at)
        if side == "top":
            return (self.x + w * at, self.y - gap)
        return (self.x + w * at, self.y + h + gap)

    def contain(self, width, height):
        """縦横比を保ち、中央に収まる画像の領域を返す。"""
        width, height = finite_number(width, "画像幅"), finite_number(height, "画像高")
        area_w, area_h = self._known("w", "contain"), self._known("h", "contain")
        if width <= 0 or height <= 0 or area_w <= 0 or area_h <= 0:
            raise _geometry_error("画像と配置領域の寸法は正にしてください")
        ratio = min(area_w / width, area_h / height)
        w, h = width * ratio, height * ratio
        return Rect(self.x + (area_w - w) / 2, self.y + (area_h - h) / 2, w, h)


def as_rect(value) -> Rect:
    if isinstance(value, Rect):
        return value
    try:
        values = tuple(value)
    except TypeError:
        values = None
    if isinstance(value, (str, bytes)) or values is None or len(values) != 4:
        raise _geometry_error(
            f"領域は Rect か [x, y, 幅, 高さ] の 4 つの数で指定してください: {_shown(value)}", value=_shown(value)
        )
    return Rect(*values)


def align(rectangles: Iterable[Rect], area: Rect, *, horizontal=None, vertical=None):
    """各領域を指定した辺または中央へそろえる。"""
    if horizontal not in (None, "left", "center", "right") or vertical not in (None, "top", "center", "bottom"):
        raise ValueError("未対応の整列位置です")
    area = as_rect(area)
    result = []
    for rect in map(as_rect, rectangles):
        x, y = rect.x, rect.y
        # 計算の順序は 0.1 と同じにする（同じ入力で同じ座標になるように）。
        if horizontal == "left":
            x = area.x
        elif horizontal == "center":
            x = (area.x + area._known("w", "align") / 2) - rect._known("w", "align") / 2
        elif horizontal == "right":
            x = area.right - rect._known("w", "align")
        if vertical == "top":
            y = area.y
        elif vertical == "center":
            y = (area.y + area._known("h", "align") / 2) - rect._known("h", "align") / 2
        elif vertical == "bottom":
            y = area.bottom - rect._known("h", "align")
        result.append(Rect(x, y, rect.w, rect.h))
    return tuple(result)


def distribute(rectangles: Iterable[Rect], area: Rect, *, axis="horizontal", gap=None):
    """寸法を変えず、領域内に等間隔で並べる。"""
    rectangles, area = tuple(map(as_rect, rectangles)), as_rect(area)
    if axis not in ("horizontal", "vertical"):
        raise ValueError("axis は horizontal または vertical です")
    if not rectangles:
        return ()
    horizontal = axis == "horizontal"
    name = "w" if horizontal else "h"
    lengths = [rect._known(name, "distribute") for rect in rectangles]
    total = sum(lengths)
    available = area._known(name, "distribute")
    if gap is None:
        gap = (available - total) / (len(rectangles) - 1) if len(rectangles) > 1 else 0
    gap = finite_number(gap, "間隔")
    if gap < 0 or total + gap * (len(rectangles) - 1) > available + 1e-7:
        raise ValueError("指定した領域に要素が収まりません")
    cursor, result = area.x if horizontal else area.y, []
    for rect, length in zip(rectangles, lengths):
        result.append(Rect(cursor, rect.y, rect.w, rect.h) if horizontal else Rect(rect.x, cursor, rect.w, rect.h))
        cursor += length + gap
    return tuple(result)
