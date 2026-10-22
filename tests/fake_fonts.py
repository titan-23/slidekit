"""試験用の FontSource（measure.FontSource の偽物）。

既定の値：送り幅は全文字 0.5 em（文字数 × 文字サイズ ÷ 2）、ascent 0.9 em、descent 0.3 em（usWin の値）、
kern なし、字形の範囲は (−0.1, 0.7) em（空白は輪郭なし）。どの書体も同じ値なので、行の配置のモデルでは、行の上側が
行間 1 以上で 行間 × 0.9 em（行間がちょうど 1 なら 0.9 em）、最後の基線より下が 0.3 em になる。
数式は 1 文字 0.5 em、インクは上 0.7 em・下 0.1 em で、行の配置に使う上下もインクと同じ。`\\unknown` を含む式は
MeasureError にする。
字形がない文字は `missing` に、すべての書体に共通の文字の集合か、書体名から文字の集合への辞書で指定する。
実フォントの規則（kern 表、輪郭、OS/2 の usWin）を確かめる試験は `synthetic_fonts` の合成フォントを使う。
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping

from slidekit.measure import MEASURE_VERSION, FaceMetrics, MathBox, MeasureError, layout_text, missing_glyphs

DEFAULT_ROLES = {"japanese": "Yu Gothic UI", "latin": "Arial", "math": "Cambria Math"}


class FakeFonts:
    def __init__(
        self,
        *,
        missing: Iterable[str] | Mapping[str, Iterable[str]] = (),
        advance: float = 0.5,
        bold_advance: float | None = None,
        ascent: float = 0.9,
        descent: float = 0.3,
        bounds: tuple[float, float] = (-0.1, 0.7),
        kerning: Mapping[tuple[str, str], float] | None = None,
        math_advance: float = 0.5,
        math_ascent: float = 0.7,
        math_descent: float = 0.1,
        math_line_ascent: float | None = None,
        math_line_descent: float | None = None,
        math_boxes: Mapping[str, tuple[float, float, float, float, float]] | None = None,
        roles: Mapping[str, str] | None = None,
    ):
        """`math_boxes` は式から (幅, 上, 下, 行の上, 行の下) の em への辞書で、その式だけ値を替える。"""
        if isinstance(missing, Mapping):
            self._missing = {family: frozenset(characters) for family, characters in missing.items()}
            self._missing_everywhere = frozenset()
        else:
            self._missing = {}
            self._missing_everywhere = frozenset(missing)
        self._advance, self._bold_advance = advance, advance if bold_advance is None else bold_advance
        self._ascent, self._descent, self._bounds = ascent, descent, tuple(bounds)
        self._kerning = dict(kerning or {})
        self._math = (
            math_advance,
            math_ascent,
            math_descent,
            math_ascent if math_line_ascent is None else math_line_ascent,
            math_descent if math_line_descent is None else math_line_descent,
        )
        self._math_boxes = dict(math_boxes or {})
        self.roles = {**DEFAULT_ROLES, **(roles or {})}
        self.calls: list[dict] = []  # math_box の呼び出し（式・大きさ・太字・書体・置き方）

    # FontSource

    def face_metrics(self, family: str, bold: bool = False) -> FaceMetrics:
        return FaceMetrics(family, bold, self._ascent, self._descent, bool(self._kerning))

    def has_glyph(self, char: str, family: str, bold: bool = False) -> bool:
        return char not in self._missing_everywhere and char not in self._missing.get(family, ())

    def advance(self, char: str, family: str, bold: bool = False) -> float:
        if not self.has_glyph(char, family, bold):
            raise MeasureError(
                f"{family} に字形がない文字です: {char!r}", code="font_glyph_missing", characters=[char], fonts=[family]
            )
        return self._bold_advance if bold else self._advance

    def kerning(self, left: str, right: str, family: str, bold: bool = False) -> float:
        return self._kerning.get((left, right), 0.0)

    def glyph_bounds(self, char: str, family: str, bold: bool = False) -> tuple[float, float]:
        return (0.0, 0.0) if char.isspace() else self._bounds

    def math_box(
        self,
        expr: str,
        size: float,
        bold: bool = False,
        *,
        family: str | None = None,
        display: bool = False,
        japanese: str | None = None,
        latin: str | None = None,
    ) -> MathBox:
        self.calls.append(
            {
                "expression": expr,
                "size": size,
                "bold": bold,
                "family": family,
                "display": display,
                "japanese": japanese,
                "latin": latin,
            }
        )
        if r"\unknown" in expr:
            raise MeasureError(r"未対応の数式命令です: \unknown", expression=expr)
        if expr in self._math_boxes:
            values = self._math_boxes[expr]
        else:
            values = (len(expr) * self._math[0], *self._math[1:])
        return MathBox(*(value * size for value in values))

    # 互換用の計測の関数（Project.measure_text・measure_math などが呼ぶ）

    def measure_text(
        self, text: str, size: float, bold: bool = False, *, japanese: str | None = None, latin: str | None = None
    ) -> tuple[float, float]:
        style = {
            "size_pt": size,
            "bold": bold,
            "line_spacing": 1.0,
            "word_wrap": False,
            "font_latin": latin or self.roles["latin"],
            "font_japanese": japanese or self.roles["japanese"],
        }
        layout = layout_text([{"kind": "text", "text": str(text)}], style, self)
        return layout.content_width, layout.content_height

    def measure_math_metrics(
        self, expr: str, size: float, bold: bool = False, *, family: str | None = None, **kwargs
    ) -> tuple[float, float, float]:
        box = self.math_box(str(expr), float(size), bold, family=family or self.roles["math"], **kwargs)
        return box.width, box.ascent + box.descent, box.descent

    def measure_math(
        self, expr: str, size: float, bold: bool = False, *, family: str | None = None, **kwargs
    ) -> tuple[float, float]:
        return self.measure_math_metrics(expr, size, bold, family=family, **kwargs)[:2]

    def missing_glyphs(
        self, text: str, *, bold: bool = False, japanese: str | None = None, latin: str | None = None
    ) -> list[str]:
        style = {"font_latin": latin or self.roles["latin"], "font_japanese": japanese or self.roles["japanese"]}
        return missing_glyphs(text, style, self, bold=bold)

    def inventory(self, *, include_unused: bool = False) -> dict:
        return {"fonts": {}, "roles": dict(self.roles), "measure_version": MEASURE_VERSION}

    def close(self) -> None:
        pass
