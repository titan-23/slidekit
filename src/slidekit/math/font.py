"""数式フォント（OpenType の MATH 表）の値を読む。

`MathFont` は字形の送り幅と描画範囲、MATH 表の定数、斜体補正、上に付ける記号の取り付け位置、
添字の食い込み（MathKern）、伸縮する字形の候補と組み立て、GSUB の数式用の字形の置き換え（ssty・dtls・flac）を返す。
値はすべて em（1 em は文字サイズ）である。
フォント集合（TTC）は `fontNumber` で直接読み、書体を取り出して保存しない（fontTools の警告を出さず、一時ファイルを作らない）。

このモジュールは slidekit のほかのモジュールに依存しない。`MathFont` はスレッド間で共有しない（`metrics.MathMeasurer` が排他する）。
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import unicodedata

from fontTools.ttLib import TTFont, TTLibError

__all__ = ["Assembly", "Glyph", "MathFont", "Part", "math_character"]

# 値が em ではなく割合（0〜1）の定数
_PERCENT_CONSTANTS = frozenset(
    {"ScriptPercentScaleDown", "ScriptScriptPercentScaleDown", "RadicalDegreeBottomRaisePercent"}
)
_CORNERS = ("TopRight", "TopLeft", "BottomRight", "BottomLeft")


@dataclass(frozen=True)
class Glyph:
    """字形 1 つの値（em。基線が 0、上が正）。"""

    name: str
    advance: float  # 送り幅
    bounds: tuple[float, float, float, float] | None  # 描画範囲 (左, 下, 右, 上)。インクのない字形（空白）は None
    italic: float = 0.0  # 斜体補正
    accent: float | None = None  # 上に付ける記号の取り付け位置（なければ None）


@dataclass(frozen=True)
class Part:
    """伸縮する字形の組み立ての部品。並びは縦なら下から上、横なら左から右である。"""

    glyph: str  # 部品の字形名
    start: float  # 始まりの側のつなぎの長さ（em）
    end: float  # 終わりの側のつなぎの長さ（em）
    advance: float  # 部品の全体の送り（em）
    extender: bool  # 0 回以上繰り返す部品か


@dataclass(frozen=True)
class Assembly:
    """伸縮する字形の組み立て（MathGlyphConstruction の GlyphAssembly）。"""

    italic: float
    parts: tuple[Part, ...]


# ---- 数式用の英数字（Unicode の Mathematical Alphanumeric Symbols）----

# 書体（Office Math の m:scr）と太字・斜体から、文字名の前に付ける語を決める
_PREFIXES = {
    (None, False, False): None,
    (None, False, True): "ITALIC",
    (None, True, False): "BOLD",
    (None, True, True): "BOLD ITALIC",
    ("script", False, False): "SCRIPT",
    ("script", False, True): "SCRIPT",
    ("script", True, False): "BOLD SCRIPT",
    ("script", True, True): "BOLD SCRIPT",
    ("fraktur", False, False): "FRAKTUR",
    ("fraktur", False, True): "FRAKTUR",
    ("fraktur", True, False): "BOLD FRAKTUR",
    ("fraktur", True, True): "BOLD FRAKTUR",
    ("double-struck", False, False): "DOUBLE-STRUCK",
    ("double-struck", False, True): "DOUBLE-STRUCK",
    ("double-struck", True, False): "DOUBLE-STRUCK",
    ("double-struck", True, True): "DOUBLE-STRUCK",
    ("sans-serif", False, False): "SANS-SERIF",
    ("sans-serif", False, True): "SANS-SERIF ITALIC",
    ("sans-serif", True, False): "SANS-SERIF BOLD",
    ("sans-serif", True, True): "SANS-SERIF BOLD ITALIC",
    ("monospace", False, False): "MONOSPACE",
    ("monospace", False, True): "MONOSPACE",
    ("monospace", True, False): "MONOSPACE",
    ("monospace", True, True): "MONOSPACE",
}
# 字形の書式（構文木の Run.style）を（太字, 斜体）にする
_STYLES = {
    "italic": (False, True),
    "upright": (False, False),
    "bold": (True, False),
    "bold-italic": (True, True),
}
# 文字名の規則に当てはまらない記号の、数式用の英数字での名前
_SYMBOL_NAMES = {
    "ϵ": "EPSILON SYMBOL",
    "ϑ": "THETA SYMBOL",
    "ϰ": "KAPPA SYMBOL",
    "ϕ": "PHI SYMBOL",
    "ϱ": "RHO SYMBOL",
    "ϖ": "PI SYMBOL",
    "ϴ": "CAPITAL THETA SYMBOL",
    "∇": "NABLA",
    "∂": "PARTIAL DIFFERENTIAL",
    "ı": "SMALL DOTLESS I",
    "ȷ": "SMALL DOTLESS J",
}
# 歴史的な理由で Letterlike Symbols に置かれた文字（数式用の英数字の該当の位置は空き）
_EXCEPTIONS = {
    "ITALIC": {"h": "ℎ"},
    "SCRIPT": {
        "B": "ℬ",
        "E": "ℰ",
        "F": "ℱ",
        "H": "ℋ",
        "I": "ℐ",
        "L": "ℒ",
        "M": "ℳ",
        "R": "ℛ",
        "e": "ℯ",
        "g": "ℊ",
        "o": "ℴ",
    },
    "FRAKTUR": {"C": "ℭ", "H": "ℌ", "I": "ℑ", "R": "ℜ", "Z": "ℨ"},
    "DOUBLE-STRUCK": {"C": "ℂ", "H": "ℍ", "N": "ℕ", "P": "ℙ", "Q": "ℚ", "R": "ℝ", "Z": "ℤ"},
}


def _base_name(character: str) -> str | None:
    """数式用の英数字の名前の後半（"CAPITAL A"、"SMALL ALPHA"、"DIGIT ZERO" など）。対応しない文字は None。"""
    if character in _SYMBOL_NAMES:
        return _SYMBOL_NAMES[character]
    name = unicodedata.name(character, "")
    for prefix in ("LATIN ", "GREEK "):
        if name.startswith(prefix) and " LETTER " in name:
            case, _, letter = name.removeprefix(prefix).partition(" LETTER ")
            if case in ("CAPITAL", "SMALL") and letter:
                return f"{case} {letter}"
            return None
    if name.startswith("DIGIT "):
        return name
    return None


def math_character(character: str, style: str = "upright", variant: str | None = None) -> str:
    """文字を書式と書体に対応する数式用の英数字にする。対応する文字がなければ元の文字を返す。

    `style` は italic・upright・bold・bold-italic、`variant` は None・script・fraktur・double-struck・sans-serif・monospace。
    """
    bold, italic = _STYLES.get(style, (False, False))
    prefix = _PREFIXES.get((variant, bold, italic))
    if prefix is None or len(character) != 1:
        return character
    exception = _EXCEPTIONS.get(prefix, {}).get(character)
    if exception is not None:
        return exception
    base = _base_name(character)
    if base is None:
        return character
    try:
        return unicodedata.lookup(f"MATHEMATICAL {prefix} {base}")
    except KeyError:
        return character


# ---- MATH 表 ----


def _value(record) -> float:
    """MathValueRecord か整数の値。"""
    return float(getattr(record, "Value", record))


def _coverage_map(coverage, values) -> dict[str, object]:
    if coverage is None or values is None:
        return {}
    return dict(zip(coverage.glyphs, values))


class MathFont:
    """MATH 表のある数式フォント 1 書体。"""

    def __init__(self, path, index=0, family="Cambria Math"):
        """TTC は fontNumber で直接読む（取り出して保存しない）。書体名と MATH 表を検査する。"""
        self.path = Path(path).expanduser().resolve(strict=True)
        if not isinstance(index, int) or isinstance(index, bool) or index < 0:
            raise ValueError("フォント集合の位置は 0 以上の整数で指定する")
        self.index = index
        self.family = family
        with self.path.open("rb") as stream:
            collection = stream.read(4) == b"ttcf"
        if not collection and index != 0:
            raise ValueError("単一フォントの位置は 0 を指定する")
        try:
            self._face = TTFont(self.path, fontNumber=index if collection else -1, lazy=True)
        except TTLibError as error:
            raise ValueError(f"数式フォントを読めない: {self.path} (index={index}): {error}") from error
        try:
            self._load()
        except Exception:
            self._face.close()
            raise
        self._glyphs: dict[str, Glyph] = {}
        self._characters: dict[tuple[str, str, str | None], Glyph] = {}
        self._closed = False

    def _load(self) -> None:
        face = self._face
        names = {item.toUnicode() for item in face["name"].names if item.nameID in (1, 16)}
        if self.family not in names:
            raise ValueError(f"指定した書体 {self.family} ではない: {self.path} (index={self.index})")
        if "MATH" not in face:
            raise ValueError("数式用の MATH テーブルがない: " + str(self.path))
        self.units_per_em = int(face["head"].unitsPerEm)
        table = face["MATH"].table
        constants = table.MathConstants
        if constants is None:
            raise ValueError("数式用の MATH テーブルに定数（MathConstants）がない: " + str(self.path))
        self._constants: dict[str, float] = {}
        for converter in constants.getConverters():
            self._constants[converter.name] = _value(getattr(constants, converter.name))
        info = table.MathGlyphInfo
        italics = getattr(info, "MathItalicsCorrectionInfo", None)
        self._italics = {
            name: _value(value)
            for name, value in _coverage_map(
                getattr(italics, "Coverage", None), getattr(italics, "ItalicsCorrection", None)
            ).items()
        }
        accents = getattr(info, "MathTopAccentAttachment", None)
        self._accents = {
            name: _value(value)
            for name, value in _coverage_map(
                getattr(accents, "TopAccentCoverage", None), getattr(accents, "TopAccentAttachment", None)
            ).items()
        }
        shapes = getattr(info, "ExtendedShapeCoverage", None)
        self._extended = frozenset(shapes.glyphs) if shapes is not None else frozenset()
        kerns = getattr(info, "MathKernInfo", None)
        self._kerns = _coverage_map(
            getattr(kerns, "MathKernCoverage", None), getattr(kerns, "MathKernInfoRecords", None)
        )
        variants = table.MathVariants
        self._min_overlap = float(getattr(variants, "MinConnectorOverlap", 0) or 0) if variants is not None else 0.0
        self._vertical = _coverage_map(
            getattr(variants, "VertGlyphCoverage", None), getattr(variants, "VertGlyphConstruction", None)
        )
        self._horizontal = _coverage_map(
            getattr(variants, "HorizGlyphCoverage", None), getattr(variants, "HorizGlyphConstruction", None)
        )
        self._cmap = face.getBestCmap() or {}
        self._metrics = face["hmtx"].metrics
        self._outlines = face["glyf"] if "glyf" in face else None
        self._glyph_set = None if self._outlines is not None else face.getGlyphSet()
        self._substitutions: dict[str, dict[str, tuple[str, ...]]] | None = None

    def _features(self) -> dict[str, dict[str, tuple[str, ...]]]:
        """GSUB の ssty（添字用の字形）・dtls（点のない字形）・flac（平らな記号）の置き換え。初回の参照で読む。"""
        if self._substitutions is not None:
            return self._substitutions
        self._check()
        result: dict[str, dict[str, tuple[str, ...]]] = {"ssty": {}, "dtls": {}, "flac": {}}
        table = self._face["GSUB"].table if "GSUB" in self._face else None
        if table is not None and table.FeatureList and table.LookupList:
            # 数式の用字（math）があれば、その既定の言語の機能だけを使う
            indices = None
            for record in getattr(table.ScriptList, "ScriptRecord", None) or []:
                if record.ScriptTag == "math" and record.Script.DefaultLangSys is not None:
                    indices = set(record.Script.DefaultLangSys.FeatureIndex)
            for position, feature in enumerate(table.FeatureList.FeatureRecord):
                mapping = result.get(feature.FeatureTag)
                if mapping is None or (indices is not None and position not in indices):
                    continue
                for lookup_index in feature.Feature.LookupListIndex:
                    lookup = table.LookupList.Lookup[lookup_index]
                    for subtable in lookup.SubTable:
                        subtable = getattr(subtable, "ExtSubTable", subtable)
                        if hasattr(subtable, "alternates"):
                            for source, targets in subtable.alternates.items():
                                mapping.setdefault(source, tuple(targets))
                        elif hasattr(subtable, "mapping"):
                            for source, target in subtable.mapping.items():
                                mapping.setdefault(source, (target,))
        self._substitutions = result
        return result

    def script_glyph(self, glyph: str, level: int) -> str:
        """添字の段階（1 以上）に合う字形名（GSUB の ssty）。置き換えがなければ元の字形名。"""
        if level <= 0:
            return glyph
        targets = self._features()["ssty"].get(glyph)
        if not targets:
            return glyph
        return targets[min(level, len(targets)) - 1]

    def dotless_glyph(self, glyph: str) -> str:
        """点のない字形名（GSUB の dtls）。上に記号を付ける i・j に使う。置き換えがなければ元の字形名。"""
        targets = self._features()["dtls"].get(glyph)
        return targets[0] if targets else glyph

    def flattened_glyph(self, glyph: str) -> str:
        """平らな記号の字形名（GSUB の flac）。高い文字の上の記号に使う。置き換えがなければ元の字形名。"""
        targets = self._features()["flac"].get(glyph)
        return targets[0] if targets else glyph

    def _check(self) -> None:
        if self._closed:
            raise RuntimeError("閉じた数式フォントは使用できない")

    def _em(self, value: float) -> float:
        return value / self.units_per_em

    def constant(self, name: str) -> float:
        """MathConstants の値（em）。ScriptPercentScaleDown などの割合の値は 0〜1。MinConnectorOverlap も返す。"""
        if name == "MinConnectorOverlap":
            return self._em(self._min_overlap)
        if name not in self._constants:
            raise KeyError(f"MATH 表の定数ではない: {name}")
        value = self._constants[name]
        return value / 100 if name in _PERCENT_CONSTANTS else self._em(value)

    def has_character(self, char: str) -> bool:
        return len(char) == 1 and ord(char) in self._cmap

    def glyph(self, name: str) -> Glyph:
        """字形名から字形の値を返す。"""
        self._check()
        cached = self._glyphs.get(name)
        if cached is not None:
            return cached
        if name not in self._metrics:
            raise ValueError(f"{self.family} に字形がない: {name}")
        advance = self._em(self._metrics[name][0])
        bounds = self._bounds(name)
        accent = self._accents.get(name)
        glyph = Glyph(
            name,
            advance,
            None if bounds is None else tuple(self._em(value) for value in bounds),
            self._em(self._italics.get(name, 0.0)),
            None if accent is None else self._em(accent),
        )
        self._glyphs[name] = glyph
        return glyph

    def _bounds(self, name: str) -> tuple[float, float, float, float] | None:
        """字形の描画範囲（設計単位。左、下、右、上）。インクがなければ None。"""
        if self._outlines is not None:
            outline = self._outlines[name]
            if getattr(outline, "numberOfContours", 0) == 0 or not hasattr(outline, "xMin"):
                return None
            return outline.xMin, outline.yMin, outline.xMax, outline.yMax
        from fontTools.pens.boundsPen import BoundsPen

        pen = BoundsPen(self._glyph_set)
        self._glyph_set[name].draw(pen)
        return pen.bounds

    def glyph_for(self, char: str, style: str = "upright", variant: str | None = None) -> Glyph:
        """文字を数式用の英数字（U+1D400 以降、ℎ などの例外を含む）へ対応させた字形を返す。

        数式用の英数字の字形がなければ元の文字の字形を返す。元の文字にもなければ「字形がない」の ValueError。
        """
        self._check()
        key = (char, style, variant)
        cached = self._characters.get(key)
        if cached is not None:
            return cached
        mapped = math_character(char, style, variant)
        name = self._cmap.get(ord(mapped)) if len(mapped) == 1 else None
        if name is None and len(char) == 1:
            name = self._cmap.get(ord(char))
        if name is None:
            code = f"U+{ord(char):04X}" if len(char) == 1 else "、".join(f"U+{ord(c):04X}" for c in char)
            raise ValueError(f"{self.family} に字形がない: {char!r} ({code})")
        glyph = self.glyph(name)
        self._characters[key] = glyph
        return glyph

    def is_extended_shape(self, glyph: str) -> bool:
        return glyph in self._extended

    def _variants(self, table, glyph: str) -> list[tuple[str, float]]:
        construction = table.get(glyph)
        if construction is None:
            return []
        records = getattr(construction, "MathGlyphVariantRecord", None) or []
        return [(record.VariantGlyph, self._em(record.AdvanceMeasurement)) for record in records]

    def _assembly(self, table, glyph: str) -> Assembly | None:
        construction = table.get(glyph)
        assembly = getattr(construction, "GlyphAssembly", None) if construction is not None else None
        if assembly is None or not assembly.PartRecords:
            return None
        parts = tuple(
            Part(
                part.glyph,
                self._em(part.StartConnectorLength),
                self._em(part.EndConnectorLength),
                self._em(part.FullAdvance),
                bool(part.PartFlags & 1),
            )
            for part in assembly.PartRecords
        )
        return Assembly(self._em(_value(assembly.ItalicsCorrection)), parts)

    def vertical_variants(self, glyph: str) -> list[tuple[str, float]]:
        """縦に伸びる候補の (字形名, 送りの長さ em) の列。小さい順。"""
        return self._variants(self._vertical, glyph)

    def vertical_assembly(self, glyph: str) -> Assembly | None:
        return self._assembly(self._vertical, glyph)

    def horizontal_variants(self, glyph: str) -> list[tuple[str, float]]:
        """横に伸びる候補の (字形名, 送りの長さ em) の列。小さい順。"""
        return self._variants(self._horizontal, glyph)

    def horizontal_assembly(self, glyph: str) -> Assembly | None:
        return self._assembly(self._horizontal, glyph)

    def kern(self, glyph: str, corner: str, height: float) -> float:
        """MathKern（em）。corner は TopRight・TopLeft・BottomRight・BottomLeft。height は字形の基線からの高さ（em）。

        値がなければ 0。高さが最初の区切りより低ければ最初の値、最後の区切り以上なら最後の値を使う。
        """
        if corner not in _CORNERS:
            raise ValueError("MathKern の角は TopRight・TopLeft・BottomRight・BottomLeft のいずれかにする")
        record = self._kerns.get(glyph)
        table = getattr(record, corner + "MathKern", None) if record is not None else None
        if table is None:
            return 0.0
        heights = [self._em(_value(value)) for value in table.CorrectionHeight]
        values = [self._em(_value(value)) for value in table.KernValue]
        if not values:
            return 0.0
        position = 0
        while position < len(heights) and height >= heights[position]:
            position += 1
        return values[min(position, len(values) - 1)]

    def close(self) -> None:
        if not self._closed:
            self._closed = True
            self._glyphs.clear()
            self._characters.clear()
            self._face.close()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()
