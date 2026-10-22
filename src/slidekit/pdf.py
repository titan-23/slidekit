"""PDF の読み取りと画像化。

ページの要素・文字の位置・色・描画は PDFium（pypdfium2）で、文字を描く命令の文字コードは pypdf でページの描画命令から、
字形の送り幅と輪郭は PDF に書かれた幅と埋め込まれた書体（fontTools）から求める。単位は pt で、原点はページの左上、
y は下向きとする。

読み取りの値は、以前使っていた PyMuPDF の値と同じになるようにしてある（PowerPoint 16.0 の PDF で照合した）。
文字の原点と範囲は、単精度の計算の順序が PyMuPDF と違うので、0.001 pt 未満の差がある。

- 描画順の番号（`seqno`）：ページの要素を描く順に振る。フォーム XObject の中も展開する。塗りと輪郭の両方で描く文字と、
  1 つの要素で塗りと線を描く図形は 2 つの番号を使う（PyMuPDF と同じ）。
- 字形の範囲：`bbox` は横が字形の送り幅、縦が書体のアセンダーとディセンダーを合計 1 em に縮めた範囲（PyMuPDF の
  texttrace の範囲）。`ink` は横が PDF に書かれた幅（`/W`・`/Widths`）の送り、縦が字形の輪郭（制御点を含む）の範囲
  （PyMuPDF の rawdict の正確な範囲）。字形の送り幅は、CID の書体では埋め込まれた書体の値、単純な書体では PDF に書かれた
  幅を使う。検査に使う `Page.text_spans()` の字の `bbox` は、PyMuPDF の get_text("dict") の範囲にする。
- 描画：塗りの直後に同じ形の枠線を描く 2 つの要素は、1 つ（`kind` が `"fs"`）にまとめる。範囲は最初の移動点と、
  線と曲線の点（曲線は制御点も）から求め、2 つ目以降の移動点と部分経路を閉じる線は含めない。
- PDFium が文字として返さない字（空白や、近い位置に同じ文字列を描く文字の要素の字。PDFium は重ね描きとみなして除く）は、
  描画命令の文字コードから補う。文字は `/ToUnicode`・符号化・埋め込まれた書体の対応表の順に推定する。

字形番号を求めることができない文字（埋め込まれていない書体、TrueType でない書体、対応していない符号化）は、
PDFium の値で近似する。範囲が少し広くなることがある。
"""

from __future__ import annotations

from collections.abc import Iterator
import ctypes
import dataclasses
from dataclasses import dataclass
import hashlib
import io
import math
from pathlib import Path
import re
import struct
from typing import Any

# 画素の範囲の丸めの余裕（MuPDF の fz_round_rect と同じ）。
_ROUND_EPSILON = 0.001
# 単精度の浮動小数点数の計算機イプシロン（C の FLT_EPSILON）。
_FLOAT_EPSILON = 1.1920928955078125e-07
# 部分集合の書体名の接頭辞（ABCDEF+）。
_SUBSET_PREFIX = re.compile(r"^[A-Z]{6}\+")
# 書体の埋め込みを読む上限（不正な PDF で巨大な領域を確保しない）。
_FONT_DATA_LIMIT = 64 * 1024 * 1024


@dataclass(frozen=True)
class Box:
    """矩形（x0, y0, x1, y1）。x0 ≥ x1 か y0 ≥ y1 の矩形は空とする。"""

    x0: float
    y0: float
    x1: float
    y1: float

    @classmethod
    def of(cls, value) -> Box:
        """Box か、4 つの数（x0, y0, x1, y1）の列から矩形を作る。"""
        if isinstance(value, Box):
            return value
        x0, y0, x1, y1 = (float(v) for v in value)
        return cls(x0, y0, x1, y1)

    @property
    def width(self) -> float:
        return self.x1 - self.x0

    @property
    def height(self) -> float:
        return self.y1 - self.y0

    @property
    def is_empty(self) -> bool:
        return self.x0 >= self.x1 or self.y0 >= self.y1

    def __iter__(self) -> Iterator[float]:
        return iter((self.x0, self.y0, self.x1, self.y1))

    def __or__(self, other) -> Box:
        """両方を含む最小の矩形。空の矩形は含めない（MuPDF の fz_union_rect と同じ）。"""
        other = Box.of(other)
        if other.is_empty:
            return self
        if self.is_empty:
            return other
        return Box(min(self.x0, other.x0), min(self.y0, other.y0), max(self.x1, other.x1), max(self.y1, other.y1))

    def __and__(self, other) -> Box:
        """共通部分（なければ空の矩形）。"""
        other = Box.of(other)
        return Box(max(self.x0, other.x0), max(self.y0, other.y0), min(self.x1, other.x1), min(self.y1, other.y1))

    def intersects(self, other) -> bool:
        """共通部分が空でないか（辺が接するだけなら偽）。"""
        return not (self & other).is_empty

    def expanded(self, amount: float) -> Box:
        return Box(self.x0 - amount, self.y0 - amount, self.x1 + amount, self.y1 + amount)

    def pixels(self, scale: float) -> tuple[int, int, int, int]:
        """scale 倍した画素の範囲（MuPDF の fz_round_rect と同じ丸め）。"""
        return (
            math.floor(self.x0 * scale + _ROUND_EPSILON),
            math.floor(self.y0 * scale + _ROUND_EPSILON),
            math.ceil(self.x1 * scale - _ROUND_EPSILON),
            math.ceil(self.y1 * scale - _ROUND_EPSILON),
        )


@dataclass(frozen=True)
class Char:
    """描かれた 1 字。"""

    unicode: int  # 文字の番号（基本多言語面の外の文字も 1 つの番号）
    origin: tuple[float, float]  # 原点
    bbox: tuple[float, float, float, float]  # 横は送り幅、縦はアセンダー・ディセンダーの範囲
    ink: tuple[float, float, float, float]  # 横は送り幅、縦は字形の輪郭の範囲
    gid: int | None = None  # 字形番号（求めることができなければ None）


@dataclass(frozen=True)
class Span:
    """1 回の文字の描画（PDF の文字の要素 1 つ）。"""

    seqno: int
    font: str  # 書体名（部分集合の接頭辞を除く）
    size: float  # 文字サイズ（pt。文字の行列の倍率を含む）
    kind: int  # 0 は塗り、1 は輪郭、3 は見えない文字
    opacity: float  # 0〜1
    color: tuple[float, float, float]  # RGB（0〜1）。塗りは塗りの色、輪郭は線の色
    chars: tuple[Char, ...]
    ascender: float = 0.8  # 書体のアセンダー（em。PyMuPDF の get_text("dict") の ascender）
    descender: float = -0.2  # 書体のディセンダー（em。負の値）

    @property
    def text(self) -> str:
        return "".join(chr(char.unicode) for char in self.chars)

    @property
    def bbox(self) -> Box:
        """文字の `bbox` の外接矩形（文字がなければ空）。"""
        if not self.chars:
            return Box(0.0, 0.0, 0.0, 0.0)
        boxes = [char.bbox for char in self.chars]
        return Box(
            min(b[0] for b in boxes), min(b[1] for b in boxes), max(b[2] for b in boxes), max(b[3] for b in boxes)
        )


@dataclass(frozen=True)
class Drawing:
    """1 つの図形の描画。"""

    seqno: int
    kind: str  # "f"（塗り）・"s"（線）・"fs"（塗りと線）
    rect: Box
    fill: tuple[float, float, float] | None  # RGB（0〜1）
    color: tuple[float, float, float] | None  # 線の色
    width: float | None  # 線幅（pt）
    fill_opacity: float
    stroke_opacity: float


def _pdfium():
    import pypdfium2 as pdfium
    import pypdfium2.raw as raw

    return pdfium, raw


def _address(handle) -> int:
    return ctypes.cast(handle, ctypes.c_void_p).value or 0


def _compose(inner, outer):
    """inner の後に outer を適用する行列（PDF の行列 [a b c d e f]）。"""
    a, b, c, d, e, f = inner
    A, B, C, D, E, F = outer
    return (a * A + b * C, a * B + b * D, c * A + d * C, c * B + d * D, e * A + f * C + E, e * B + f * D + F)


def _rgba(getter, handle) -> tuple[tuple[float, float, float], float] | None:
    r, g, b, a = (ctypes.c_uint() for _ in range(4))
    if not getter(handle, r, g, b, a):
        return None
    return (r.value / 255, g.value / 255, b.value / 255), a.value / 255


class Document:
    """開いた PDF。`with open_pdf(path) as document:` で使い、ページは `document[index]`（0 始まり）で読む。"""

    def __init__(self, path: str | Path):
        pdfium, _raw = _pdfium()
        self.path = Path(path)
        self._pdf = pdfium.PdfDocument(str(self.path))
        self._pages: dict[int, Page] = {}
        self._reader = None
        self._fonts: dict[str, Any] = {}

    def __len__(self) -> int:
        return len(self._pdf)

    def __getitem__(self, index: int) -> Page:
        if not 0 <= index < len(self):
            raise IndexError(f"PDF にそのページがない: {index + 1} ページ目（全 {len(self)} ページ）")
        if index not in self._pages:
            self._pages[index] = Page(self, index)
        return self._pages[index]

    def __iter__(self) -> Iterator[Page]:
        return (self[index] for index in range(len(self)))

    def __enter__(self) -> Document:
        return self

    def __exit__(self, *_exc) -> None:
        self.close()

    def close(self) -> None:
        for page in self._pages.values():
            page._close()
        self._pages.clear()
        self._pdf.close()

    def _content_reader(self):
        """描画命令を読む pypdf の読み取り（必要になったときに 1 回だけ開く）。"""
        if self._reader is None:
            import pypdf

            self._reader = pypdf.PdfReader(str(self.path))
        return self._reader

    def _font(self, handle):
        """PDFium の書体から、埋め込まれた書体（fontTools の TTFont）を返す。読むことができなければ None。"""
        _pdfium_module, raw = _pdfium()
        size = ctypes.c_size_t()
        if not raw.FPDFFont_GetFontData(handle, None, 0, size) or not 0 < size.value <= _FONT_DATA_LIMIT:
            return None
        buffer = (ctypes.c_uint8 * size.value)()
        if not raw.FPDFFont_GetFontData(handle, buffer, size.value, size):
            return None
        data = bytes(buffer)
        key = hashlib.sha256(data).hexdigest()
        if key not in self._fonts:
            self._fonts[key] = _Glyphs.load(data)
        glyphs = self._fonts[key]
        return None if glyphs is None or glyphs.unusable else glyphs


class _Glyphs:
    """埋め込まれた書体の字形の範囲（em。上向き）。"""

    def __init__(self, font):
        self.font = font
        self.order = font.getGlyphOrder()
        self.units = font["head"].unitsPerEm
        self._glyphs = font.getGlyphSet()
        self._bounds: dict[int, tuple[float, float] | None] = {}
        self._reverse: dict[int, int] | None = None
        # 表を読む途中で失敗した書体は使わない（その文字の要素から PDFium の値で近似する）。
        self.unusable = False

    @classmethod
    def load(cls, data: bytes) -> _Glyphs | None:
        from fontTools.ttLib import TTFont, TTLibError

        try:
            font = TTFont(io.BytesIO(data), lazy=True)
            if "head" not in font or not ("glyf" in font or "CFF " in font):
                return None
            return cls(font)
        except (TTLibError, KeyError, ValueError, AssertionError, struct.error):
            return None

    def _guarded(self, read, default):
        """書体の表を読む。読むことができなければ、書体を使わない印を付けて default を返す。"""
        if self.unusable:
            return default
        try:
            return read()
        except Exception:  # noqa: BLE001  不正な表を持つ書体でも、PDF の読み取りは PDFium の値で続ける
            self.unusable = True
            return default

    def vertical(self, gid: int) -> tuple[float, float] | None:
        """字形の輪郭の縦の範囲（下端、上端。em）。輪郭がなければ None。"""
        if gid not in self._bounds:
            from fontTools.pens.boundsPen import ControlBoundsPen

            def read():
                if not 0 <= gid < len(self.order):
                    return None
                pen = ControlBoundsPen(self._glyphs)
                self._glyphs[self.order[gid]].draw(pen)
                return None if pen.bounds is None else (pen.bounds[1] / self.units, pen.bounds[3] / self.units)

            self._bounds[gid] = self._guarded(read, None)
        return self._bounds[gid]

    def advance(self, gid: int) -> float | None:
        """字形の送り幅（em）。"""

        def read():
            if not 0 <= gid < len(self.order) or "hmtx" not in self.font:
                return None
            return self.font["hmtx"][self.order[gid]][0] / self.units

        return self._guarded(read, None)

    def reverse_cmap(self) -> dict[int, int]:
        """字形番号から文字（対応表で最初に見つかったもの）。"""
        if self._reverse is None:

            def read():
                reverse: dict[int, int] = {}
                for unicode, name in sorted((self.font.getBestCmap() or {}).items()):
                    gid = self._index(name)
                    if gid is not None:
                        reverse.setdefault(gid, unicode)
                return reverse

            self._reverse = self._guarded(read, {})
        return self._reverse

    def simple_gid(self, code: int, char: str | None, name: str | None, symbolic: bool) -> int | None:
        """単純な TrueType の書体の符号から字形番号を引く（PDF 1.7 の 9.6.6.4）。記号でない書体は、符号化の文字で
        (3,1) の表を、字形名で書体の字形を、符号で (1,0) の表を引く。記号の書体と、それで見つからない字形は、符号と
        0xF000 + 符号で (3,0) の表を、符号で (1,0) の表を引く。"""
        return self._guarded(lambda: self._simple_gid(code, char, name, symbolic), None)

    def _simple_gid(self, code: int, char: str | None, name: str | None, symbolic: bool) -> int | None:
        tables = self.cmap_tables()
        if not symbolic:
            if char is not None and (3, 1) in tables and ord(char) in tables[(3, 1)]:
                return self._index(tables[(3, 1)][ord(char)])
            if name is not None and name in self.font.getGlyphOrder():
                return self._index(name)
            if (1, 0) in tables and code in tables[(1, 0)]:
                return self._index(tables[(1, 0)][code])
        if (3, 0) in tables:
            for value in (code, 0xF000 + code):
                if value in tables[(3, 0)]:
                    return self._index(tables[(3, 0)][value])
        if (1, 0) in tables and code in tables[(1, 0)]:
            return self._index(tables[(1, 0)][code])
        if char is not None and (3, 1) in tables and ord(char) in tables[(3, 1)]:
            return self._index(tables[(3, 1)][ord(char)])
        return None

    def cmap_tables(self) -> dict:
        if "cmap" not in self.font:
            return {}
        return {(table.platformID, table.platEncID): table.cmap for table in self.font["cmap"].tables}

    def _index(self, name: str) -> int | None:
        try:
            return self.font.getGlyphID(name)
        except KeyError:
            return None


class _FontInfo:
    """PDF の書体の辞書から、文字コードの区切り方・字形番号・送り幅を求める（pypdf の辞書を読む）。"""

    def __init__(self, font):
        from pypdf.generic import NameObject

        self.subtype = font.get("/Subtype")
        self.two_byte = False
        self.identity = True
        self.cid_to_gid: list[int] | None = None
        self.default_width = 1000.0
        self.widths: dict[int, float] = {}
        self.known = False  # 字形番号を求める方法が分かるか
        self.to_unicode = _read_to_unicode(_resolve(font.get("/ToUnicode")))
        if self.subtype == "/Type0":
            # Identity 以外の CMap（符号の長さが変わるもの）は読まない。その文字の要素は PDFium の値で近似する。
            if font.get("/Encoding") not in ("/Identity-H", "/Identity-V"):
                raise ValueError(f"CMap に対応していない: {font.get('/Encoding')}")
            self.two_byte = True
            descendant = _resolve(_resolve(font.get("/DescendantFonts"))[0])
            self.known = self.two_byte and descendant.get("/Subtype") == "/CIDFontType2"
            self.default_width = float(descendant.get("/DW", 1000))
            items = list(_resolve(descendant.get("/W", [])) or [])
            index = 0
            while index + 1 < len(items):
                first = int(items[index])
                value = _resolve(items[index + 1])
                if isinstance(value, list):
                    for offset, width in enumerate(value):
                        self.widths[first + offset] = float(_resolve(width))
                    index += 2
                elif index + 2 < len(items):
                    for code in range(first, int(value) + 1):
                        self.widths[code] = float(_resolve(items[index + 2]))
                    index += 3
                else:
                    break
            mapping = descendant.get("/CIDToGIDMap")
            mapping = _resolve(mapping) if mapping is not None else None
            if mapping is not None and not isinstance(mapping, NameObject):
                data = mapping.get_data()
                self.cid_to_gid = [int.from_bytes(data[i : i + 2], "big") for i in range(0, len(data) - 1, 2)]
        else:
            self.known = self.subtype == "/TrueType"
            first = int(font.get("/FirstChar", 0))
            for offset, width in enumerate(_resolve(font.get("/Widths", [])) or []):
                self.widths[first + offset] = float(_resolve(width))
            descriptor = _resolve(font.get("/FontDescriptor"))
            if descriptor is not None:
                self.default_width = float(descriptor.get("/MissingWidth", 0))
                self.symbolic = bool(int(descriptor.get("/Flags", 0)) & 4)
            self._read_encoding(_resolve(font.get("/Encoding")))

    # 単純な書体の符号化（記号の書体で符号化がなければ、文字コードで書体の対応表を引く）。
    symbolic = False
    base_encoding: str | None = None
    differences: dict[int, str] = {}

    def _read_encoding(self, encoding) -> None:
        from pypdf.generic import DictionaryObject, NameObject

        self.differences = {}
        if isinstance(encoding, NameObject):
            self.base_encoding = str(encoding)
        elif isinstance(encoding, DictionaryObject):
            base = encoding.get("/BaseEncoding")
            self.base_encoding = str(base) if base is not None else None
            code = 0
            for item in _resolve(encoding.get("/Differences", [])) or []:
                item = _resolve(item)
                if isinstance(item, NameObject):
                    self.differences[code] = str(item)[1:]
                    code += 1
                else:
                    code = int(item)

    def _encoded_char(self, code: int) -> str | None:
        """単純な書体の符号の文字（`/Differences` の字形名、基本の符号化の順）。"""
        if code in self.differences:
            from fontTools.agl import toUnicode

            text = toUnicode(self.differences[code])
            return text if len(text) == 1 else None
        codec = {"/MacRomanEncoding": "mac_roman"}.get(self.base_encoding or "", "cp1252")
        if self.symbolic and self.base_encoding is None:
            return None
        try:
            return bytes([code]).decode(codec)
        except UnicodeDecodeError:
            return None

    def codes(self, data: bytes) -> list[int]:
        if self.two_byte:
            return [int.from_bytes(data[i : i + 2], "big") for i in range(0, len(data) - 1, 2)]
        return list(data)

    def width(self, code: int) -> float | None:
        """符号の送り幅（1/1000 em）。単純な書体で幅の表がなければ None。"""
        if self.subtype != "/Type0" and not self.widths:
            return None
        return self.widths.get(code, self.default_width)

    def unicode(self, code: int, glyphs: _Glyphs | None) -> int | None:
        """符号の文字の推定（`/ToUnicode`、単純な書体は符号化、CID の書体は埋め込まれた書体の対応表の逆引きの順）。
        分からなければ None。"""
        if code in self.to_unicode:
            return self.to_unicode[code]
        if self.subtype != "/Type0":
            char = self._encoded_char(code)
            if char is not None:
                return ord(char)
        gid = self.gid(code, glyphs)
        if gid is None or glyphs is None:
            return None
        return glyphs.reverse_cmap().get(gid)

    def gid(self, code: int, glyphs: _Glyphs | None) -> int | None:
        if not self.known or glyphs is None:
            return None
        if self.subtype == "/Type0":
            if self.cid_to_gid is None:
                return code
            return self.cid_to_gid[code] if code < len(self.cid_to_gid) else None
        return glyphs.simple_gid(code, self._encoded_char(code), self.differences.get(code), self.symbolic)


def _resolve(value):
    from pypdf.generic import IndirectObject

    return value.get_object() if isinstance(value, IndirectObject) else value


# /ToUnicode の CMap の対応（bfchar の 1 組と、bfrange の範囲）。
_BFCHAR = re.compile(r"beginbfchar(.*?)endbfchar", re.S)
_BFRANGE = re.compile(r"beginbfrange(.*?)endbfrange", re.S)
_HEX_PAIR = re.compile(r"<([0-9A-Fa-f\s]*)>\s*<([0-9A-Fa-f\s]*)>")
_HEX_RANGE = re.compile(r"<([0-9A-Fa-f\s]*)>\s*<([0-9A-Fa-f\s]*)>\s*(?:<([0-9A-Fa-f\s]*)>|\[([^\]]*)\])")
# 1 つの bfrange で読む符号の数の上限（不正な CMap で巨大な表を作らない）。
_RANGE_LIMIT = 0x10000


def _read_to_unicode(stream) -> dict[int, int]:
    """`/ToUnicode` の CMap から、1 つの文字に対応する符号と文字（合字など、2 字以上に対応する符号は含めない）。"""
    if stream is None or not hasattr(stream, "get_data"):
        return {}
    try:
        text = stream.get_data().decode("latin-1")
    except Exception:  # noqa: BLE001  読むことができない CMap は使わない（ほかの方法で推定する）
        return {}

    def number(value: str) -> int | None:
        value = re.sub(r"\s", "", value)
        return int(value, 16) if value else None

    def char(value: str) -> int | None:
        try:
            decoded = bytes.fromhex(re.sub(r"\s", "", value)).decode("utf-16-be")
        except (ValueError, UnicodeDecodeError):
            return None
        return ord(decoded) if len(decoded) == 1 else None

    result: dict[int, int] = {}
    for block in _BFCHAR.findall(text):
        for source, target in _HEX_PAIR.findall(block):
            code, unicode = number(source), char(target)
            if code is not None and unicode is not None:
                result[code] = unicode
    for block in _BFRANGE.findall(text):
        for low, high, target, items in _HEX_RANGE.findall(block):
            first, last = number(low), number(high)
            if first is None or last is None or not 0 <= last - first < _RANGE_LIMIT:
                continue
            if items:
                for offset, item in enumerate(re.findall(r"<([0-9A-Fa-f\s]*)>", items)[: last - first + 1]):
                    unicode = char(item)
                    if unicode is not None:
                        result[first + offset] = unicode
                continue
            start = char(target)
            if start is not None:
                for offset in range(last - first + 1):
                    if start + offset <= 0x10FFFF:
                        result[first + offset] = start + offset
    return result


def _string_bytes(value) -> bytes:
    from pypdf.generic import TextStringObject

    return value.original_bytes if isinstance(value, TextStringObject) else bytes(value)


@dataclass(frozen=True)
class _Run:
    """文字を描く 1 つの命令。"""

    font: _FontInfo | None
    codes: list[int]
    shifts: list[float]  # 各文字の前の TJ の数（1/1000 em。正なら左へ戻す）
    char_spacing: float  # Tc（文字空間の単位）
    word_spacing: float  # Tw（1 バイトの符号 32 だけに加わる）
    scaling: float  # Tz / 100
    rise: float  # Ts（文字空間の単位）


def _text_runs(reader, index: int) -> list[_Run]:
    """ページの文字を描く命令を描画順に返す。フォーム XObject の中もたどる。"""
    from pypdf.generic import ArrayObject, ByteStringObject, ContentStream, IndirectObject, TextStringObject

    page = reader.pages[index]
    cache: dict[Any, _FontInfo | None] = {}
    runs: list[_Run] = []

    def font_info(resources, name):
        fonts = _resolve(resources.get("/Font", {})) if resources else {}
        reference = fonts.get(name) if fonts else None
        key = reference.idnum if isinstance(reference, IndirectObject) else id(reference)
        if key not in cache:
            try:
                cache[key] = _FontInfo(_resolve(reference)) if reference is not None else None
            except (KeyError, TypeError, ValueError, IndexError):
                cache[key] = None
        return cache[key]

    def number(value) -> float:
        try:
            return float(value)
        except (TypeError, ValueError):
            return 0.0

    def walk(content, resources, depth, inherited=None):
        if content is None or depth > 8:
            return
        stack = []
        # 文字の状態（書体、Tc、Tw、Tz、Ts）。q と Q で保存して戻す。フォーム XObject は呼んだ時点の状態を受け継ぐ。
        state = list(inherited) if inherited is not None else [None, 0.0, 0.0, 1.0, 0.0]

        def show(items):
            font = state[0]
            codes: list[int] = []
            shifts: list[float] = []
            pending = 0.0
            drawn = False
            for item in items:
                if isinstance(item, (ByteStringObject, TextStringObject)):
                    data = _string_bytes(item)
                    drawn = drawn or bool(data)
                    for code in font.codes(data) if font is not None else []:
                        codes.append(code)
                        shifts.append(pending)
                        pending = 0.0
                else:
                    pending += number(item)
            # 空の文字列だけを描く命令からは、PDFium は文字の要素を作らない。
            if drawn:
                runs.append(_Run(font, codes, shifts, state[1], state[2], state[3], state[4]))

        for operands, operator in ContentStream(content, reader).operations:
            if operator == b"q":
                stack.append(list(state))
            elif operator == b"Q":
                state = stack.pop() if stack else state
            elif operator == b"Tf" and operands:
                state[0] = font_info(resources, operands[0])
            elif operator == b"Tc" and operands:
                state[1] = number(operands[0])
            elif operator == b"Tw" and operands:
                state[2] = number(operands[0])
            elif operator == b"Tz" and operands:
                state[3] = number(operands[0]) / 100
            elif operator == b"Ts" and operands:
                state[4] = number(operands[0])
            elif operator in (b"Tj", b"'") and operands:
                show([operands[0]])
            elif operator == b'"' and len(operands) >= 3:
                state[2], state[1] = number(operands[0]), number(operands[1])
                show([operands[2]])
            elif operator == b"TJ" and operands and isinstance(operands[0], ArrayObject):
                show(list(operands[0]))
            elif operator == b"Do" and operands:
                xobjects = _resolve(resources.get("/XObject", {})) if resources else {}
                target = _resolve(xobjects.get(operands[0])) if xobjects else None
                if target is not None and target.get("/Subtype") == "/Form":
                    walk(target, _resolve(target.get("/Resources")) or resources, depth + 1, state)

    walk(page.get_contents(), _resolve(page.get("/Resources")), 0)
    return runs


class Page:
    """PDF の 1 ページ。"""

    def __init__(self, document: Document, index: int):
        _pdfium_module, raw = _pdfium()
        self._document = document
        self.index = index
        self._page = document._pdf[index]
        self.width = float(self._page.get_width())
        self.height = float(self._page.get_height())
        self._spans: list[Span] | None = None
        self._drawings: list[Drawing] | None = None
        # 描画順の要素（要素、種類、ページへの行列）と、描画順の番号。
        self._objects: list[tuple[Any, int, tuple]] = []
        numbers: list[int] = []
        counter = 0

        def walk(container, count, get, outer):
            nonlocal counter
            for position in range(count(container)):
                handle = get(container, position)
                kind = raw.FPDFPageObj_GetType(handle)
                m = raw.FS_MATRIX()
                raw.FPDFPageObj_GetMatrix(handle, m)
                matrix = _compose((m.a, m.b, m.c, m.d, m.e, m.f), outer)
                if kind == raw.FPDF_PAGEOBJ_FORM:
                    walk(handle, raw.FPDFFormObj_CountObjects, raw.FPDFFormObj_GetObject, matrix)
                    continue
                self._objects.append((handle, kind, matrix))
                numbers.append(counter)
                counter += self._calls(handle, kind)

        walk(self._page.raw, raw.FPDFPage_CountObjects, raw.FPDFPage_GetObject, (1.0, 0.0, 0.0, 1.0, 0.0, 0.0))
        self._numbers = numbers
        self._seqno = {_address(handle): number for (handle, _k, _m), number in zip(self._objects, numbers)}

    @property
    def rect(self) -> Box:
        return Box(0.0, 0.0, self.width, self.height)

    @staticmethod
    def _calls(handle, kind) -> int:
        """要素を描く命令の数。塗りと輪郭の両方で描く文字と、1 つの要素で塗りと線を描く図形は 2、切り抜きだけの文字と、
        塗りも線もない図形は 0（切り抜きには番号を振らない）、ほかは 1。"""
        _pdfium_module, raw = _pdfium()
        if kind == raw.FPDF_PAGEOBJ_TEXT:
            mode = raw.FPDFTextObj_GetTextRenderMode(handle)
            if mode in (raw.FPDF_TEXTRENDERMODE_FILL_STROKE, raw.FPDF_TEXTRENDERMODE_FILL_STROKE_CLIP):
                return 2
            return 0 if mode == raw.FPDF_TEXTRENDERMODE_CLIP else 1
        if kind == raw.FPDF_PAGEOBJ_PATH:
            fillmode, stroke = ctypes.c_int(), ctypes.c_int()
            raw.FPDFPath_GetDrawMode(handle, fillmode, stroke)
            return int(fillmode.value != 0) + int(bool(stroke.value))
        return 1

    def _close(self) -> None:
        try:
            self._page.close()
        except Exception:  # noqa: BLE001  閉じる処理の失敗は、読み取りの結果に影響しない
            pass

    # 文字

    def spans(self) -> list[Span]:
        """描画順の文字の描画。塗りと輪郭の両方で描く文字は、塗りと輪郭の 2 つの描画にする。"""
        if self._spans is None:
            self._spans = self._read_spans()
        return self._spans

    def text_spans(self) -> list[Span]:
        """文字の検査に使う、見える文字の描画（PyMuPDF の get_text("dict") の文字片に当たる）。

        見えない文字と、塗りと輪郭の両方で描く文字の輪郭の描画は除く。字の `bbox` は PyMuPDF の文字の抽出の範囲にする（横は
        `ink` と同じ送り、縦は書体のアセンダーとディセンダー。合計が 1 em に満たない書体は 1 em に広げる）。字形の輪郭が
        ページの完全に外にある字は除く（PyMuPDF の TEXT_MEDIABOX_CLIP と同じ）。
        """
        width, height = self.width, self.height
        result = []
        previous = None
        for span in self.spans():
            # 塗りと輪郭の両方で描く文字の要素は、塗りと輪郭の 2 つの描画が同じ字の組を持つ。
            paired = previous is not None and previous.kind == 0 and previous.chars is span.chars
            previous = span
            if span.kind not in (0, 1) or (span.kind == 1 and paired):
                continue
            top, bottom = _extent(span.ascender, span.descender)
            chars = tuple(
                dataclasses.replace(
                    char,
                    bbox=(
                        char.ink[0],
                        _single(char.origin[1] - top * span.size),
                        char.ink[2],
                        _single(char.origin[1] - bottom * span.size),
                    ),
                )
                for char in span.chars
                # 横は送りで、辺に接する字も外とする。縦は字形の輪郭で、辺に接する字は中とする。
                if char.ink[0] < width and char.ink[2] > 0 and char.ink[1] <= height and char.ink[3] >= 0
            )
            if chars:
                result.append(dataclasses.replace(span, chars=chars))
        return result

    def _read_spans(self) -> list[Span]:
        _pdfium_module, raw = _pdfium()
        height = self.height
        textpage = self._page.get_textpage()
        handle = textpage.raw
        try:
            # PDFium の文字（生成した空白などを除く）を、文字の要素ごとに並べる。代用符号の組は 1 字にまとめる。
            listed = []
            for index in range(raw.FPDFText_CountChars(handle)):
                if not raw.FPDFText_IsGenerated(handle, index):
                    owner = _address(raw.FPDFText_GetTextObject(handle, index))
                    listed.append([index, raw.FPDFText_GetUnicode(handle, index), owner])
            members: dict[int, list[tuple[int, int]]] = {}
            position = 0
            while position < len(listed):
                index, unicode, owner = listed[position]
                step = 1
                if 0xD800 <= unicode <= 0xDFFF and position + 1 < len(listed):
                    # 右から左に書く文字が続くと、PDFium は組を逆の順（下位、上位）で返すことがある。
                    _next_index, other, next_owner = listed[position + 1]
                    high, low = (unicode, other) if unicode <= 0xDBFF else (other, unicode)
                    if 0xD800 <= high <= 0xDBFF and 0xDC00 <= low <= 0xDFFF and next_owner == owner:
                        unicode = 0x10000 + ((high - 0xD800) << 10) + (low - 0xDC00)
                        step = 2
                members.setdefault(owner, []).append((index, unicode))
                position += step
            runs = self._runs()
            texts = [handle_ for handle_, kind, _m in self._objects if kind == raw.FPDF_PAGEOBJ_TEXT]
            if runs is not None and len(runs) != len(texts):
                runs = None
            spans = []
            text_number = -1
            for (obj, kind, matrix), seqno in zip(self._objects, self._numbers):
                if kind != raw.FPDF_PAGEOBJ_TEXT:
                    continue
                text_number += 1
                entries = members.get(_address(obj), [])
                run = runs[text_number] if runs is not None else None
                font = raw.FPDFTextObj_GetFont(obj)
                info = run.font if run is not None else None
                glyphs = self._document._font(font) if info is not None and info.known else None
                chars, size = self._object_chars(handle, obj, matrix, entries, run, font, glyphs)
                if glyphs is not None and glyphs.unusable:
                    # 埋め込まれた書体の表を読む途中で失敗したら、その書体を使わずに読み直す。
                    chars, size = self._object_chars(handle, obj, matrix, entries, run, font, None)
                name = (ctypes.c_char * 256)()
                raw.FPDFFont_GetBaseFontName(font, name, 256)
                font_name = _SUBSET_PREFIX.sub("", name.value.decode("latin-1", "replace"))
                ascender, descender = _raw_metrics(font)
                mode = raw.FPDFTextObj_GetTextRenderMode(obj)
                fill = _rgba(raw.FPDFPageObj_GetFillColor, obj) or ((0.0, 0.0, 0.0), 1.0)
                stroke = _rgba(raw.FPDFPageObj_GetStrokeColor, obj) or ((0.0, 0.0, 0.0), 1.0)
                common = dict(font=font_name, size=size, chars=chars, ascender=ascender, descender=descender)
                if mode == raw.FPDF_TEXTRENDERMODE_FILL_STROKE:
                    spans.append(Span(seqno, kind=0, color=fill[0], opacity=fill[1], **common))
                    spans.append(Span(seqno + 1, kind=1, color=stroke[0], opacity=stroke[1], **common))
                elif mode == raw.FPDF_TEXTRENDERMODE_STROKE:
                    spans.append(Span(seqno, kind=1, color=stroke[0], opacity=stroke[1], **common))
                elif mode == raw.FPDF_TEXTRENDERMODE_INVISIBLE:
                    spans.append(Span(seqno, kind=3, color=fill[0], opacity=fill[1], **common))
                elif mode in (raw.FPDF_TEXTRENDERMODE_FILL, raw.FPDF_TEXTRENDERMODE_FILL_CLIP):
                    spans.append(Span(seqno, kind=0, color=fill[0], opacity=fill[1], **common))
                elif mode == raw.FPDF_TEXTRENDERMODE_STROKE_CLIP:
                    spans.append(Span(seqno, kind=1, color=stroke[0], opacity=stroke[1], **common))
                elif mode == raw.FPDF_TEXTRENDERMODE_FILL_STROKE_CLIP:
                    spans.append(Span(seqno, kind=0, color=fill[0], opacity=fill[1], **common))
                    spans.append(Span(seqno + 1, kind=1, color=stroke[0], opacity=stroke[1], **common))
                # 切り抜きだけの文字（モード 7）は描かれないので読まない。
            return spans
        finally:
            textpage.close()

    def _object_chars(self, handle, obj, matrix, entries, run, font, glyphs) -> tuple[tuple[Char, ...], float]:
        """1 つの文字の要素の字と文字サイズ。glyphs は埋め込まれた書体（使わなければ None）。"""
        _pdfium_module, raw = _pdfium()
        info = run.font if run is not None else None
        aligned = _align(entries, run, glyphs) if run is not None else None
        if aligned is None:
            run, info, glyphs = None, None, None
            aligned = entries
        metrics = _metrics(font)
        object_size = _object_size(obj)
        size = object_size * math.sqrt(abs(matrix[0] * matrix[3] - matrix[1] * matrix[2]))
        # 原点：PDFium は幅の表を整数に丸めて文字を置くので、幅が小数の書体と、PDFium が返さない字（空白や、
        # 重ね描きとみなして除いた文字の要素の字）は、描画命令から求める。幅の表のない字は PDFium の幅を使う。
        # ほかの字は PDFium の原点（PyMuPDF と同じ単精度の計算）を使う。
        origins, exact = None, True
        if run is not None:
            missing = any(index is None for index, _unicode in aligned)
            widths = [run.font.width(code) for code in run.codes] if run.font is not None else []
            exact = all(width is None or float(width).is_integer() for width in widths)
            if missing or not exact:
                fallback = [_pdfium_width(font, unicode) for _index, unicode in aligned] if missing else None
                origins = _origins(run, matrix, object_size, self.height, fallback)
        chars: list[Char] = []
        for position, (char_index, unicode) in enumerate(aligned):
            code = run.codes[position] if run is not None else None
            if char_index is not None:
                size = float(raw.FPDFText_GetFontSize(handle, char_index))
            origin = origins[position] if origins is not None and (char_index is None or not exact) else None
            if origin is None and char_index is None:
                continue
            chars.append(self._char(handle, char_index, obj, origin, unicode, size, info, code, glyphs, metrics))
        return tuple(chars), size

    def _runs(self):
        """描画命令の文字コード。読むことができなければ None（PDFium の値で近似する）。"""
        try:
            return _text_runs(self._document._content_reader(), self.index)
        except Exception:  # noqa: BLE001  描画命令を読めない PDF でも、PDFium の値で読み取りを続ける
            return None

    def _char(self, handle, index, obj, origin, unicode, size, info, code, glyphs, metrics) -> Char:
        """1 字。origin（描画命令から求めた原点）がなければ PDFium の原点を使う。index が None なら PDFium が文字として
        返さない字である。"""
        _pdfium_module, raw = _pdfium()
        ascent, descent = metrics
        if origin is not None:
            ox, oy = origin
        else:
            x, y = ctypes.c_double(), ctypes.c_double()
            raw.FPDFText_GetCharOrigin(handle, index, x, y)
            # 原点は単精度に丸める（PyMuPDF と同じ精度にし、0.1 pt に丸めた基線が丸めの境目で分かれないようにする）。
            ox, oy = _single(float(x.value)), _single(self.height - float(y.value))
        gid = info.gid(code, glyphs) if info is not None and code is not None else None
        # 送り幅：PDF の幅の表（文字を置く送り）と、字形の送り（CID の書体はフォント本体の値）。PyMuPDF は rawdict の範囲に
        # 前者、texttrace の範囲に後者を使う。
        pen = None
        if info is not None and code is not None:
            width = info.width(code)
            pen = width / 1000 * size if width is not None else None
        if pen is None:
            pen = self._pdfium_advance(handle, index, obj, unicode, size, ox)
        trace = pen
        if info is not None and info.subtype == "/Type0" and glyphs is not None and gid is not None:
            em = glyphs.advance(gid)
            trace = em * size if em is not None else pen
        vertical = glyphs.vertical(gid) if glyphs is not None and gid is not None else None
        if vertical is None and index is None and (gid is None or glyphs is None):
            vertical = _pdfium_vertical(raw.FPDFTextObj_GetFont(obj), unicode)
        if vertical is not None:
            ink_top, ink_bottom = oy - vertical[1] * size, oy - vertical[0] * size
        elif (gid is not None and glyphs is not None) or index is None:
            # 輪郭のない字形（空白など）は、原点の高さの幅 0 の範囲にする。
            ink_top = ink_bottom = oy
        else:
            left, right, bottom, top = (ctypes.c_double() for _ in range(4))
            raw.FPDFText_GetCharBox(handle, index, left, right, bottom, top)
            ink_top, ink_bottom = self.height - float(top.value), self.height - float(bottom.value)
        # 範囲は単精度に丸める（PyMuPDF の矩形と同じ精度）。
        return Char(
            unicode=unicode,
            origin=(ox, oy),
            bbox=tuple(map(_single, (ox, oy - ascent * size, ox + trace, oy - descent * size))),
            ink=tuple(map(_single, (ox, ink_top, ox + pen, ink_bottom))),
            gid=gid,
        )

    @staticmethod
    def _pdfium_advance(handle, index, obj, unicode, size, origin_x) -> float:
        """PDF の幅の表がない文字の送り幅（PDFium の値）。"""
        _pdfium_module, raw = _pdfium()
        measured = ctypes.c_float()
        font = raw.FPDFTextObj_GetFont(obj)
        if 0 <= unicode <= 0xFFFF and raw.FPDFFont_GetGlyphWidth(font, unicode, size, measured):
            return float(measured.value)
        if index is None:
            return 0.0
        loose = raw.FS_RECTF()
        raw.FPDFText_GetLooseCharBox(handle, index, loose)
        return float(loose.right) - origin_x

    # 描画

    def drawings(self) -> list[Drawing]:
        """描画順の図形の描画（塗り・線のどちらもない図形は除く）。"""
        if self._drawings is None:
            self._drawings = self._read_drawings()
        return self._drawings

    def _read_drawings(self) -> list[Drawing]:
        _pdfium_module, raw = _pdfium()
        height = self.height
        result: list[Drawing] = []
        shapes: list[list | None] = []
        point_x, point_y = ctypes.c_float(), ctypes.c_float()
        for (obj, kind, matrix), seqno in zip(self._objects, self._numbers):
            if kind != raw.FPDF_PAGEOBJ_PATH:
                continue
            fillmode, stroke = ctypes.c_int(), ctypes.c_int()
            raw.FPDFPath_GetDrawMode(obj, fillmode, stroke)
            filled, stroked = fillmode.value != 0, bool(stroke.value)
            if not filled and not stroked:
                continue
            a, b, c, d, e, f = matrix
            points = []
            for position in range(raw.FPDFPath_CountSegments(obj)):
                segment = raw.FPDFPath_GetPathSegment(obj, position)
                raw.FPDFPathSegment_GetPoint(segment, point_x, point_y)
                px, py = float(point_x.value), float(point_y.value)
                points.append(
                    (
                        raw.FPDFPathSegment_GetType(segment),
                        bool(raw.FPDFPathSegment_GetClose(segment)),
                        round(a * px + c * py + e, 3),
                        round(height - (b * px + d * py + f), 3),
                    )
                )
            xs, ys = _path_points(points, raw.FPDF_SEGMENT_MOVETO)
            if not xs:
                continue
            fill = _rgba(raw.FPDFPageObj_GetFillColor, obj) if filled else None
            line = _rgba(raw.FPDFPageObj_GetStrokeColor, obj) if stroked else None
            width = ctypes.c_float()
            raw.FPDFPageObj_GetStrokeWidth(obj, width)
            expansion = math.sqrt(abs(a * d - b * c))
            drawing = Drawing(
                seqno=seqno,
                kind="fs" if filled and stroked else ("f" if filled else "s"),
                rect=Box(min(xs), min(ys), max(xs), max(ys)),
                fill=fill[0] if fill else None,
                color=line[0] if line else None,
                width=float(width.value) * expansion if stroked else None,
                fill_opacity=fill[1] if fill else 1.0,
                stroke_opacity=line[1] if line else 1.0,
            )
            # 塗りの直後に同じ形の枠線を描く 2 つの要素は、1 つにまとめる。
            if drawing.kind == "s" and result and result[-1].kind == "f" and shapes[-1] == points:
                previous = result[-1]
                result[-1] = Drawing(
                    seqno=previous.seqno,
                    kind="fs",
                    rect=previous.rect,
                    fill=previous.fill,
                    color=drawing.color,
                    width=drawing.width,
                    fill_opacity=previous.fill_opacity,
                    stroke_opacity=drawing.stroke_opacity,
                )
                shapes[-1] = None
                continue
            result.append(drawing)
            shapes.append(points)
        return result

    # 画像化

    def render(self, scale: float):
        """ページを scale 倍（1 pt あたりの画素数）の RGB の画像（Pillow）にする。背景は白。

        画素数は MuPDF と同じく、ページの範囲を `Box.pixels` で丸めた大きさにする。
        """
        if not math.isfinite(scale) or scale <= 0:
            raise ValueError(f"画像化の倍率は正の有限値にする: {scale}")
        pdfium, raw = _pdfium()
        left, top, right, bottom = self.rect.pixels(scale)
        width, height = right - left, bottom - top
        # PDFium の既定の画像化（pypdfium2 の render）は、切り上げた画素数に合わせてページを引き伸ばす。倍率を正確に保つため、
        # 画素数を決めた画像に行列で描く（回転したページも PDFium がページの向きに直す）。
        bitmap = pdfium.PdfBitmap.new_native(width, height, raw.FPDFBitmap_BGR)
        try:
            bitmap.fill_rect((255, 255, 255, 255), 0, 0, width, height)
            matrix = raw.FS_MATRIX(scale, 0, 0, scale, -left, -top)
            clip = raw.FS_RECTF(0, 0, width, height)
            raw.FPDF_RenderPageBitmapWithMatrix(bitmap, self._page, matrix, clip, raw.FPDF_ANNOT)
            return bitmap.to_pil().convert("RGB")
        finally:
            bitmap.close()


def _single(value: float) -> float:
    """単精度の浮動小数点数に丸めた値。"""
    return struct.unpack("f", struct.pack("f", value))[0]


def _same_char(listed: int, expected: int) -> bool:
    """PDFium の文字と、文字コードから推定した文字が同じ字か（PDFium は行末のハイフンを U+0002 で返す）。"""
    return listed == expected or (listed == 0x02 and expected in (0x2D, 0xAD))


def _repaired(listed: int, expected: int | None) -> int:
    """PDFium の文字を、推定した文字で直す（制御文字と、組にならなかった代用符号）。"""
    if listed < 0x20 or 0xD800 <= listed <= 0xDFFF:
        if expected is not None:
            return expected
        return 0x2D if listed == 0x02 else listed
    return listed


def _align(entries: list[tuple[int, int]], run: _Run, glyphs) -> list[tuple[int | None, int]] | None:
    """文字コードの列に、PDFium の文字を順に対応させる。PDFium が返さない字（空白など）は (None, 推定した文字) にする。
    PDFium の文字のほうが多いか、対応させきれなければ None。"""
    codes = run.codes
    if run.font is None or len(entries) > len(codes):
        return None
    expected = [run.font.unicode(code, glyphs) for code in codes]
    result: list[tuple[int | None, int]] = []
    taken = 0
    for position, code in enumerate(codes):
        guess = expected[position]
        remaining, available = len(codes) - position, len(entries) - taken
        if available == 0:
            take = False
        elif available >= remaining:
            # 残りをすべて対応させる。推定した字と食い違う普通の字があれば（右から左に書く文字の並べ替えや、合字を
            # 2 字にした /ToUnicode など）、対応させずに PDFium の値を使う。
            listed = entries[taken][1]
            if (
                guess is not None
                and listed >= 0x20
                and not 0xD800 <= listed <= 0xDFFF
                and not _same_char(listed, guess)
            ):
                return None
            take = True
        elif guess is None:
            take = entries[taken][1] != 0x20
        else:
            take = _same_char(entries[taken][1], guess)
        if take:
            index, unicode = entries[taken]
            taken += 1
            result.append((index, _repaired(unicode, guess)))
        else:
            result.append((None, guess if guess is not None else 0xFFFD))
    return result if taken == len(entries) else None


def _origins(
    run: _Run, matrix, font_size: float, height: float, fallback: list[float | None] | None = None
) -> list[tuple[float, float]] | None:
    """描画命令から求めた各字の原点（PDF の仕様の 9.4.4 の式。幅は PDF の幅の表を小数のまま使う）。

    PDFium は幅の表を整数に丸めて文字を置くので、幅が小数の PDF では位置がずれていく。描画命令から求めると、PyMuPDF と同じ
    位置になる。幅の表のない字は fallback（各字の幅。1/1000 em）を使い、それもなければ None（PDFium の原点を使う）。
    matrix は文字の要素の行列（文字の行列とページへの行列）で、その位置は文字の上下の移動（Ts）と、最初の字の前の TJ の
    数を含む。
    """
    info = run.font
    if info is None:
        return None
    a, b, c, d, e, f = matrix
    x, y = e, f
    result = []
    for position, code in enumerate(run.codes):
        if position:
            shift = -run.shifts[position] / 1000 * font_size * run.scaling
            x, y = x + shift * a, y + shift * b
        result.append((_single(x), _single(height - y)))
        width = info.width(code)
        if width is None and fallback is not None:
            width = fallback[position]
        if width is None:
            return None
        spacing = run.char_spacing + (run.word_spacing if not info.two_byte and code == 32 else 0.0)
        advance = (width / 1000 * font_size + spacing) * run.scaling
        x, y = x + advance * a, y + advance * b
    return result


def _pdfium_width(font, unicode: int) -> float | None:
    """PDFium の書体の、文字の送り幅（1/1000 em）。分からなければ None。"""
    _pdfium_module, raw = _pdfium()
    measured = ctypes.c_float()
    if 0 <= unicode <= 0xFFFF and raw.FPDFFont_GetGlyphWidth(font, unicode, 1000.0, measured):
        return float(measured.value)
    return None


def _pdfium_vertical(font, unicode: int) -> tuple[float, float] | None:
    """PDFium の書体の字形の輪郭（制御点を含む）の縦の範囲（下端、上端。em）。輪郭がなければ None。

    埋め込まれていない書体の字で、PDFium が文字として返さないものに使う。"""
    _pdfium_module, raw = _pdfium()
    if not 0 <= unicode <= 0xFFFF:
        return None
    path = raw.FPDFFont_GetGlyphPath(font, unicode, 1.0)
    if not path:
        return None
    x, y = ctypes.c_float(), ctypes.c_float()
    values = []
    for index in range(raw.FPDFGlyphPath_CountGlyphSegments(path)):
        if raw.FPDFPathSegment_GetPoint(raw.FPDFGlyphPath_GetGlyphPathSegment(path, index), x, y):
            values.append(float(y.value))
    return (min(values), max(values)) if values else None


def _raw_metrics(font) -> tuple[float, float]:
    """書体のアセンダーとディセンダー（em）。"""
    _pdfium_module, raw = _pdfium()
    ascent, descent = ctypes.c_float(), ctypes.c_float()
    raw.FPDFFont_GetAscent(font, 1.0, ascent)
    raw.FPDFFont_GetDescent(font, 1.0, descent)
    return float(ascent.value), float(descent.value)


def _extent(ascender: float, descender: float) -> tuple[float, float]:
    """PyMuPDF の文字の抽出の字の範囲の縦（上端と下端。em）。アセンダーとディセンダーの合計が 1 em 以上ならそのまま、
    満たなければ 1 em に広げる（PyMuPDF の JM_char_quad と同じ）。"""
    if ascender - descender + _FLOAT_EPSILON >= 1:
        return ascender, descender
    if ascender < 1e-3:
        return 0.9, -0.1
    total = ascender - descender + _FLOAT_EPSILON
    return ascender / total, descender / total


def _metrics(font) -> tuple[float, float]:
    """書体のアセンダーとディセンダーを、合計 1 em に縮めた値（PyMuPDF の texttrace と同じ）。"""
    high, low = _raw_metrics(font)
    if high < 1e-3:
        high, low = 0.9, -0.1
    if high - low <= 0:
        return 0.8, -0.2
    return high / (high - low), low / (high - low)


def _object_size(obj) -> float:
    _pdfium_module, raw = _pdfium()
    size = ctypes.c_float()
    raw.FPDFTextObj_GetFontSize(obj, size)
    return float(size.value)


def _path_points(points, moveto) -> tuple[list[float], list[float]]:
    """範囲に入れる経路の点（PyMuPDF と同じ規則）：最初の移動点と、線と曲線の点（曲線は制御点も）。2 つ目以降の移動点と、
    部分経路を閉じる区間（始点へ戻る閉じた区間）は含めない。"""
    xs: list[float] = []
    ys: list[float] = []
    start = None
    for kind, close, x, y in points:
        if kind == moveto:
            start = (x, y)
            if not xs:
                xs.append(x)
                ys.append(y)
            continue
        if close and start is not None and (x, y) == start:
            continue
        xs.append(x)
        ys.append(y)
    return xs, ys


def open_pdf(path: str | Path) -> Document:
    """PDF を開く。開くことができなければ OSError か ValueError にする。"""
    try:
        return Document(path)
    except FileNotFoundError:
        raise
    except Exception as error:  # noqa: BLE001  PDFium の例外を、呼び出し側が扱う種類にそろえる
        if isinstance(error, (OSError, ValueError)):
            raise
        raise ValueError(f"PDF を開くことができない: {path}（{error}）") from error


def page_size(path: str | Path, page: int) -> tuple[float, float]:
    """PDF の page ページ目（1 始まり）の寸法（pt）。ページがなければ ValueError。"""
    with open_pdf(path) as document:
        if not 1 <= page <= len(document):
            raise ValueError(f"PDF にページ {page} がありません（全 {len(document)} ページ）")
        target = document[page - 1]
        return target.width, target.height


def render_to_pixels(path: str | Path, page: int, pixels: tuple[int, int]):
    """PDF の page ページ目（1 始まり）を、縦横比を保って pixels（幅、高さ）を覆う大きさの RGB の画像にする。"""
    with open_pdf(path) as document:
        if not 1 <= page <= len(document):
            raise ValueError(f"PDF 素材のページが存在しない: {path} ({page})")
        target = document[page - 1]
        factor = max(pixels[0] / target.width, pixels[1] / target.height)
        return target.render(factor)
