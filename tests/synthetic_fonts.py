"""試験用の合成フォント。字形は長方形の輪郭で、送り幅・上下の範囲・kern 表を指定できる。

Windows のフォントがない環境でも、実フォントの表（cmap・hmtx・OS/2・kern・輪郭）を読む規則を試験できる。
`make_font` は hhea と OS/2 の usWin に同じ ascent・descent を書く。`SyntheticFonts` は FontRegistry と同じく usWin を読む。
`make_font` でフォントを作り、`SyntheticFonts` で measure.FontSource として読む。数式は FakeFonts と同じ規則で計測する。
"""

from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path

from fontTools.fontBuilder import FontBuilder
from fontTools.pens.boundsPen import BoundsPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTFont, newTable
from fontTools.ttLib.tables._k_e_r_n import KernTable_format_0

from fake_fonts import FakeFonts
from slidekit.measure import FaceMetrics, MathBox, MeasureError


def _glyph_name(char: str) -> str:
    return f"u{ord(char):04X}"


def make_font(
    path,
    family: str,
    glyphs: Mapping[str, tuple[int, int, int]],
    *,
    upm: int = 1000,
    ascent: int = 900,
    descent: int = 200,
    kerning: Mapping[tuple[str, str], int] | None = None,
    style: str = "Regular",
    win_ascent: int | None = None,
    win_descent: int | None = None,
) -> Path:
    """`glyphs` は文字から（送り幅, y_min, y_max）への辞書（単位は unitsPerEm の目盛り）。y_max ≤ y_min なら輪郭なし。

    `kerning` は文字の対から値への辞書で、kern 表（format 0、横書き）にする。`ascent`・`descent` は hhea・sTypo・usWin に
    書く。usWin だけを違う値にするときは `win_ascent`・`win_descent` を指定する。
    """
    path = Path(path)
    names = [".notdef"] + [_glyph_name(char) for char in glyphs]
    builder = FontBuilder(upm, isTTF=True)
    builder.setupGlyphOrder(names)
    builder.setupCharacterMap({ord(char): _glyph_name(char) for char in glyphs})
    outlines, metrics = {".notdef": TTGlyphPen(None).glyph()}, {".notdef": (upm // 2, 0)}
    for char, (advance, y_min, y_max) in glyphs.items():
        pen = TTGlyphPen(None)
        if y_max > y_min:
            pen.moveTo((0, y_min))
            pen.lineTo((0, y_max))
            pen.lineTo((advance, y_max))
            pen.lineTo((advance, y_min))
            pen.closePath()
        outlines[_glyph_name(char)] = pen.glyph()
        metrics[_glyph_name(char)] = (advance, 0)
    builder.setupGlyf(outlines)
    builder.setupHorizontalMetrics(metrics)
    builder.setupHorizontalHeader(ascent=ascent, descent=-descent)
    builder.setupNameTable(
        {
            "familyName": family,
            "styleName": style,
            "fullName": f"{family} {style}",
            "psName": f"{family.replace(' ', '')}-{style}",
        }
    )
    builder.setupOS2(
        sTypoAscender=ascent,
        sTypoDescender=-descent,
        usWinAscent=ascent if win_ascent is None else win_ascent,
        usWinDescent=descent if win_descent is None else win_descent,
    )
    builder.setupPost()
    if kerning:
        table = newTable("kern")
        table.version = 0
        subtable = KernTable_format_0()
        subtable.version, subtable.coverage, subtable.format = 0, 1, 0
        subtable.kernTable = {
            (_glyph_name(left), _glyph_name(right)): value for (left, right), value in kerning.items()
        }
        table.kernTables = [subtable]
        builder.font["kern"] = table
    path.parent.mkdir(parents=True, exist_ok=True)
    builder.save(str(path))
    return path


def _group(characters: str, advance: int, y_min: int, y_max: int) -> dict[str, tuple[int, int, int]]:
    return {char: (advance, y_min, y_max) for char in characters}


# 欧文の書体 Test Latin。大文字と ascender のある小文字は 0〜700、x の高さの小文字は 0〜500、descender は −200〜500。
LATIN_GLYPHS = {
    **_group("ABCDEFGHIJKLMNOPQRSTUVWXYZbdfhiklt0123456789!?%()[]/", 500, 0, 700),
    **_group("acemnorsuvwxz", 500, 0, 500),
    **_group("gjpqy", 500, -200, 500),
    **_group(".,:;", 250, -100, 100),
    **_group("-‐", 300, 250, 330),
    " ": (250, 0, 0),
    "→": (1000, 100, 500),
    "“": (400, 500, 700),
    "”": (400, 500, 700),
    "漢": (1000, -80, 820),  # 和文の文字は和文の書体で計測することを確かめるため、欧文の書体にも置く
}
LATIN_KERNING = {("T", "A"): -80, ("A", "V"): -60, ("A", "T"): -70, ("T", "漢"): -300}

# 和文の書体 Test Japanese。全角の文字は送り幅 1000、−80〜820。全角空白は輪郭なし。
JAPANESE_GLYPHS = {
    **_group(
        "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん"
        "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン"
        "ょっゃゅーァィゥェォッャュョ漢字日本語文章計測幅行改折返"
        "。、「」（）『』！？・※①△☆",
        1000,
        -80,
        820,
    ),
    "　": (1000, 0, 0),
    " ": (250, 0, 0),
}
JAPANESE_KERNING = {("あ", "い"): -100}


def make_standard_fonts(directory) -> dict[str, dict[str, str]]:
    """欧文 Test Latin（ascent 900・descent 200、kern 表あり）と和文 Test Japanese（ascent 1080・descent 250）の
    通常と太字を作る。太字は送り幅を 20 広げる。戻り値は実行環境の fonts の形（書体名から path・bold_path）。"""
    directory = Path(directory)
    bold_latin = {char: (advance + 20, y_min, y_max) for char, (advance, y_min, y_max) in LATIN_GLYPHS.items()}
    bold_japanese = {char: (advance + 20, y_min, y_max) for char, (advance, y_min, y_max) in JAPANESE_GLYPHS.items()}
    return {
        "Test Latin": {
            "path": str(make_font(directory / "latin.ttf", "Test Latin", LATIN_GLYPHS, kerning=LATIN_KERNING)),
            "bold_path": str(
                make_font(directory / "latin-bold.ttf", "Test Latin", bold_latin, kerning=LATIN_KERNING, style="Bold")
            ),
        },
        "Test Japanese": {
            "path": str(
                make_font(
                    directory / "japanese.ttf",
                    "Test Japanese",
                    JAPANESE_GLYPHS,
                    ascent=1080,
                    descent=250,
                    kerning=JAPANESE_KERNING,
                )
            ),
            "bold_path": str(
                make_font(
                    directory / "japanese-bold.ttf",
                    "Test Japanese",
                    bold_japanese,
                    ascent=1080,
                    descent=250,
                    kerning=JAPANESE_KERNING,
                    style="Bold",
                )
            ),
        },
    }


class _Face:
    def __init__(self, path: Path):
        font = TTFont(str(path))
        self.upm = font["head"].unitsPerEm
        # 行の配置の割合は OS/2 の usWinAscent・usWinDescent で求める（FontRegistry.face_metrics と同じ）。
        os2 = font["OS/2"]
        self.ascent = os2.usWinAscent / self.upm
        self.descent = os2.usWinDescent / self.upm
        self.cmap = font.getBestCmap() or {}
        self.advances = {name: advance for name, (advance, _lsb) in font["hmtx"].metrics.items()}
        self.kerning = {}
        if "kern" in font:
            for subtable in font["kern"].kernTables:
                if getattr(subtable, "format", None) == 0 and subtable.coverage & 1 and not subtable.coverage & 6:
                    for pair, value in subtable.kernTable.items():
                        self.kerning[pair] = self.kerning.get(pair, 0) + value
        glyphs = font.getGlyphSet()
        self.bounds = {}
        for name in self.advances:
            pen = BoundsPen(glyphs)
            glyphs[name].draw(pen)
            bounds = pen.bounds
            self.bounds[name] = (0.0, 0.0) if bounds is None else (bounds[1] / self.upm, bounds[3] / self.upm)


class SyntheticFonts:
    """合成フォントのファイルを読む FontSource。`faces` は（書体名、太字）からフォントのパスへの辞書。"""

    def __init__(self, faces: Mapping[tuple[str, bool], str | Path], *, math: FakeFonts | None = None):
        self._faces = {key: _Face(Path(path)) for key, path in faces.items()}
        self._math = math or FakeFonts()

    @classmethod
    def from_environment(cls, fonts: Mapping[str, Mapping[str, str]], *, math: FakeFonts | None = None):
        faces = {}
        for family, specification in fonts.items():
            faces[(family, False)] = specification["path"]
            faces[(family, True)] = specification.get("bold_path", specification["path"])
        return cls(faces, math=math)

    def _face(self, family: str, bold: bool) -> _Face:
        if (family, bold) not in self._faces:
            raise MeasureError(f"試験の書体がありません: {family}")
        return self._faces[(family, bold)]

    def face_metrics(self, family: str, bold: bool = False) -> FaceMetrics:
        face = self._face(family, bold)
        return FaceMetrics(family, bold, face.ascent, face.descent, bool(face.kerning))

    def has_glyph(self, char: str, family: str, bold: bool = False) -> bool:
        return ord(char) in self._face(family, bold).cmap

    def advance(self, char: str, family: str, bold: bool = False) -> float:
        face = self._face(family, bold)
        if ord(char) not in face.cmap:
            raise MeasureError(
                f"{family} に字形がない文字です: {char!r}", code="font_glyph_missing", characters=[char], fonts=[family]
            )
        return face.advances[face.cmap[ord(char)]] / face.upm

    def kerning(self, left: str, right: str, family: str, bold: bool = False) -> float:
        face = self._face(family, bold)
        pair = (face.cmap.get(ord(left)), face.cmap.get(ord(right)))
        return face.kerning.get(pair, 0) / face.upm

    def glyph_bounds(self, char: str, family: str, bold: bool = False) -> tuple[float, float]:
        face = self._face(family, bold)
        return face.bounds.get(face.cmap.get(ord(char)), (0.0, 0.0))

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
        return self._math.math_box(expr, size, bold, family=family, display=display, japanese=japanese, latin=latin)
