"""実フォントを解決し、文字と数式の寸法・利用記録を返す。

`FontRegistry` は `measure.FontSource` を実装する。文字の値（送り幅・kern・輪郭・OS/2 の usWinAscent・usWinDescent）は
fontTools で表を直接読み、行の組み立てと折り返しは `measure` が行う。数式の寸法は `math.metrics`（数式フォントの
OpenType の MATH 表）で計算し、数式の中の `\\text` には文字の計測（`text_line_metrics`）を使う。
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field
from functools import partial
import os
from pathlib import Path
import platform
import struct
import threading
from typing import Any

from fontTools.pens.boundsPen import BoundsPen
from fontTools.ttLib import TTFont

from . import measure
from .files import file_hash
from .math import metrics as math_metrics

# 既定で探す書体。書体名: (通常のファイル, 書体番号, 太字のファイル, 書体番号)
WINDOWS_FONT_FILES: dict[str, tuple[str, int, str | None, int | None]] = {
    "Yu Gothic UI": ("YuGothM.ttc", 1, "YuGothB.ttc", 1),
    "Yu Gothic": ("YuGothM.ttc", 0, "YuGothB.ttc", 0),
    "Arial": ("arial.ttf", 0, "arialbd.ttf", 0),
    "Cambria Math": ("cambria.ttc", 1, None, None),
    "Calibri": ("calibri.ttf", 0, "calibrib.ttf", 0),
    "Segoe UI": ("segoeui.ttf", 0, "segoeuib.ttf", 0),
    "Meiryo": ("meiryo.ttc", 0, "meiryob.ttc", 0),
    "BIZ UDGothic": ("BIZ-UDGothicR.ttc", 0, "BIZ-UDGothicB.ttc", 0),
}

# WSL で PowerShell の実行に失敗したときに探すフォルダー。
WSL_FONT_DIRECTORY = Path("/mnt/c/Windows/Fonts")

# 自動探索の結果。キーは PowerShell の実行ファイル（Windows の Python では "nt"）。
_DETECTED: dict[str, tuple[Path, ...]] = {}


def _is_wsl() -> bool:
    return bool(os.environ.get("WSL_DISTRO_NAME")) or "microsoft" in platform.release().lower()


def _existing(folders: Iterable[Path]) -> tuple[Path, ...]:
    """存在するフォルダーだけを、重複を除いて順に返す。"""
    result: list[Path] = []
    for folder in folders:
        if folder not in result and folder.is_dir():
            result.append(folder)
    return tuple(result)


def _fixed_font_directories() -> list[Path]:
    """PowerShell を使わずに求めるフォルダー。Windows の Python では環境変数から、WSL では WSL_FONT_DIRECTORY。"""
    if os.name == "nt":
        folders = []
        if os.environ.get("WINDIR"):
            folders.append(Path(os.environ["WINDIR"]) / "Fonts")
        if os.environ.get("LOCALAPPDATA"):
            folders.append(Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "Windows" / "Fonts")
        return folders
    return [WSL_FONT_DIRECTORY] if _is_wsl() else []


def _wsl_font_directories(environment: Mapping) -> list[Path]:
    """PowerShell でシステムと利用者単位のフォントのフォルダーを求める。使うことができなければ
    WSL_FONT_DIRECTORY に戻る。"""
    from . import renderers

    try:
        context = renderers.windows_context(environment)
        folders = []
        if context.get("fonts"):
            folders.append(renderers._unix_path(context["fonts"]))
        if context.get("local_app_data"):
            folders.append(renderers._unix_path(context["local_app_data"]) / "Microsoft" / "Windows" / "Fonts")
        return folders
    except renderers.RendererError:
        return [WSL_FONT_DIRECTORY]


def _detect_font_directories(environment: Mapping) -> tuple[Path, ...]:
    """既定の書体を探すフォルダーを自動で求める。結果はプロセス内で記憶する（試験はこの関数を差し替える）。"""
    if os.name == "nt":
        key = "nt"
    elif _is_wsl():
        key = str(environment.get("renderers", {}).get("powerpoint", {}).get("powershell", "powershell.exe"))
    else:
        return ()
    if key not in _DETECTED:
        _DETECTED[key] = _existing(_fixed_font_directories() if key == "nt" else _wsl_font_directories(environment))
    return _DETECTED[key]


def font_directories(environment: Mapping) -> tuple[Path, ...]:
    """既定の書体を探すフォルダー。設定があれば設定を、なければ自動で求めた結果を返す。"""
    if "font_directories" in environment:
        return tuple(Path(folder) for folder in environment["font_directories"])
    return _detect_font_directories(environment)


def _find_file(name: str, folders: tuple[Path, ...]) -> Path | None:
    """同じ名前のファイルは前のフォルダーを優先する。"""
    for folder in folders:
        path = folder / name
        if path.is_file():
            return path.resolve()
    return None


def default_font(family: str, environment: Mapping) -> dict | None:
    """候補表の書体を font_directories() の順に探す。

    見つかれば {"path", "index"}（太字があれば "bold_path", "bold_index" も）を返し、なければ None。
    """
    files = WINDOWS_FONT_FILES.get(family)
    if files is None:
        return None
    regular, index, bold, bold_index = files
    folders = font_directories(environment)
    path = _find_file(regular, folders)
    if path is None:
        return None
    specification = {"path": str(path), "index": index}
    bold_path = _find_file(bold, folders) if bold else None
    if bold_path is not None:
        specification.update(bold_path=str(bold_path), bold_index=bold_index)
    return specification


_COLLECTION_SUFFIXES = frozenset({".ttc", ".otc"})


def _open_font(path: Path, index: int) -> TTFont:
    """表を必要なときに読む形で開く。単独のフォントの書体番号は 0 だけを受け付ける。"""
    collection = path.suffix.lower() in _COLLECTION_SUFFIXES
    if not collection and index != 0:
        raise ValueError(f"単独フォントの index は 0 にする: {path}")
    return TTFont(path, fontNumber=index if collection else -1, lazy=True)


def _font_info(font: TTFont, path: Path, index: int, family: str) -> dict:
    """生成記録に残す値（書体名・版・ハッシュ値・MATH 表の有無）。書体名が指定と違えば、ハッシュ値を求めずに止める。"""
    names = sorted({entry.toUnicode() for entry in font["name"].names if entry.nameID in (1, 4, 16)})
    # 太字では family 名の後に Bold などが付く nameID 4 も存在する。
    if family.casefold() not in {name.casefold() for name in names}:
        raise ValueError(f"指定書体と実フォントが一致しない: {family}: {path} (index={index})")
    versions = sorted({entry.toUnicode() for entry in font["name"].names if entry.nameID == 5})
    return {
        "path": str(path),
        "index": index,
        "sha256": file_hash(path),
        "font_names": names,
        "version": "; ".join(versions),
        "has_math_table": "MATH" in font,
        "family": family,
    }


def _kern_pairs(data: bytes, glyph_order: list[str]) -> dict[tuple[str, str], int]:
    """kern 表（Windows の形式、版 0）の format 0 で横書きの小表の値を、字形名の対ごとに合計する。

    最小値の小表と交差方向の小表は除く。同じ小表の中で重なる対は後の値を使う（fontTools と同じ）。
    次の小表へは宣言の長さで進み、format 0 の対は宣言の長さの範囲で対の数まで読む（後ろの余分なバイトは読まない）。
    対の多い小表では 16 ビットの長さの欄があふれていることがある（Calibri など）ので、宣言の長さが対の数から求めた長さより
    65536 の倍数だけ短いときと最後の小表では、対の数から求めた長さを使う。Apple の形式（版 1.0）は読まない（kern なしとみなす）。
    """
    pairs: dict[tuple[str, str], int] = {}
    if len(data) < 4:
        return pairs
    version, count = struct.unpack_from(">HH", data)
    if version != 0:
        return pairs
    offset = 4
    for position in range(count):
        if offset + 6 > len(data):
            break
        _version, length, form, coverage = struct.unpack_from(">HHBB", data, offset)
        if form == 0:
            if offset + 14 > len(data):
                break
            (number,) = struct.unpack_from(">H", data, offset + 6)
            computed = 14 + 6 * number
            if length < computed and (position == count - 1 or (computed - length) % 0x10000 == 0):
                length = computed
            if length < 14:
                break
            # coverage の下位 3 ビットが「横書き」だけの小表を使う。
            if coverage & 0x07 == 0x01:
                body = data[offset + 14 : offset + min(length, computed)]
                found = {}
                for left, right, value in struct.iter_unpack(">HHh", body[: len(body) - len(body) % 6]):
                    if left < len(glyph_order) and right < len(glyph_order):
                        found[(glyph_order[left], glyph_order[right])] = value
                for pair, value in found.items():
                    pairs[pair] = pairs.get(pair, 0) + value
        elif length < 6:
            break
        offset += length
    return pairs


@dataclass
class _Tables:
    """書体 1 つの、文字の計測に使う値。輪郭の範囲は字形ごとに初めて使うときに求めて覚える。"""

    units: int  # unitsPerEm
    ascent: float  # OS/2 の usWinAscent（em）
    descent: float  # OS/2 の usWinDescent（em）
    cmap: Mapping[int, str]  # 符号位置から字形名
    advances: Mapping[str, tuple[int, int]]  # hmtx（字形名から送り幅と左の余白）
    kerning: Mapping[tuple[str, str], int]  # kern 表の値（字形名の対）
    glyphs: Any  # 輪郭を描く字形の集合（fontTools の glyph set）
    # 輪郭の範囲 (x_min, y_min, x_max, y_max)（em）。輪郭がなければ None
    bounds: dict[str, tuple[float, float, float, float] | None] = field(default_factory=dict)


def _read_tables(font: TTFont, path: Path) -> _Tables:
    """計測に使う表を読む。字形の集合もここで作るので、ファイルを閉じた後もメモリ上の表から輪郭を描くことができる。

    上下は OS/2 の usWinAscent・usWinDescent にする。PowerPoint（Windows）は行の上下をこの値で決める（Calibri など、
    hhea と値が違う書体の最初の基線で確かめた）。OS/2 の表がないか値が 0 の書体だけ、hhea の値を使う。
    """
    units = font["head"].unitsPerEm
    if units <= 0:
        raise ValueError(f"フォントの unitsPerEm が不正です: {path}")
    os2 = font["OS/2"] if "OS/2" in font else None
    if os2 is not None and os2.usWinAscent + os2.usWinDescent > 0:
        ascent, descent = os2.usWinAscent, os2.usWinDescent
    else:
        ascent, descent = font["hhea"].ascent, abs(font["hhea"].descent)
    kerning = _kern_pairs(font.getTableData("kern"), font.getGlyphOrder()) if "kern" in font else {}
    return _Tables(
        units,
        ascent / units,
        descent / units,
        font.getBestCmap() or {},
        font["hmtx"].metrics,
        kerning,
        font.getGlyphSet(),
    )


class FontRegistry:
    """設定した書体だけを使う、資料ごとの寸法計算器。`measure.FontSource` を実装する。

    計測に使う表は、書体ごとにファイルを 1 回だけ開いて読み、すぐに閉じる。表の読み込みと字形の展開は排他にするので、
    スレッドの間で共有できる。
    """

    def __init__(self, project):
        self.project = project
        self.roles = dict(project.theme.get("fonts", {}))
        self.roles.setdefault("japanese", "Yu Gothic UI")
        self.roles.setdefault("latin", "Arial")
        self.roles.setdefault("math", "Cambria Math")
        self._specifications = dict(project.environment.get("fonts", {}))
        self._faces: dict[tuple[str, bool], dict] = {}
        self._tables: dict[tuple[str, bool], _Tables] = {}
        self._math = {}
        self._lock = threading.RLock()
        self.functions = tuple(project.config.get("math", {}).get("functions", ()))

    def _specification(self, family: str) -> dict:
        if family not in self._specifications:
            # 固定した別書体へ置き換えず、指定した書体の実ファイルだけを探す。
            specification = default_font(family, self.project.environment)
            if specification is None:
                searched = "、".join(map(str, font_directories(self.project.environment))) or "なし"
                raise FileNotFoundError(
                    f'実フォントが見つからない: {family}。environment.toml の [fonts."{family}"] に path を指定する（探したフォルダー: {searched}）'
                )
            self._specifications[family] = specification
        return dict(self._specifications[family])

    def _load(self, family: str, bold: bool, *, tables: bool) -> dict:
        """書体の記録の値を返す。tables が真なら、計測に使う表も同じ読み込みで覚える。"""
        key = (family, bold)
        with self._lock:
            info = self._faces.get(key)
            if info is not None and (not tables or key in self._tables):
                return info
            specification = self._specification(family)
            if bold and not specification.get("bold_path"):
                raise FileNotFoundError(f"太字の実フォントを bold_path に指定する必要がある: {family}")
            path = Path(specification["bold_path" if bold else "path"]).expanduser().resolve()
            index = int(specification.get("bold_index" if bold else "index", 0))
            if index < 0:
                raise ValueError(f"フォントの index は 0 以上にする: {family}")
            if not path.is_file():
                raise FileNotFoundError(f"実フォントがない: {path}")
            with _open_font(path, index) as font:
                if info is None:
                    info = _font_info(font, path, index, family)
                if tables:
                    self._tables[key] = _read_tables(font, path)
            self._faces[key] = info
            return info

    def _face(self, family: str, bold: bool = False) -> dict:
        """書体の記録の値。指定した書体名と実ファイルの書体名を照らし合わせる。"""
        return self._load(family, bold, tables=False)

    def _font_tables(self, family: str, bold: bool) -> _Tables:
        tables = self._tables.get((family, bold))
        if tables is None:
            self._load(family, bold, tables=True)
            tables = self._tables[(family, bold)]
        return tables

    # measure.FontSource

    def face_metrics(self, family: str, bold: bool = False) -> measure.FaceMetrics:
        """OS/2 の usWinAscent・usWinDescent（em。OS/2 がなければ hhea）と、kern 表（format 0、横書き）の対があるか。"""
        tables = self._font_tables(family, bold)
        return measure.FaceMetrics(family, bold, tables.ascent, tables.descent, bool(tables.kerning))

    def has_glyph(self, char: str, family: str, bold: bool = False) -> bool:
        return ord(char) in self._font_tables(family, bold).cmap

    def advance(self, char: str, family: str, bold: bool = False) -> float:
        """送り幅（em）。字形がなければ MeasureError（font_glyph_missing）。"""
        tables = self._font_tables(family, bold)
        name = tables.cmap.get(ord(char))
        if name is None:
            raise measure.MeasureError(
                f"{family} に字形がない文字です: '{char}' (U+{ord(char):04X})",
                code="font_glyph_missing",
                characters=[char],
                fonts=[family],
            )
        return tables.advances[name][0] / tables.units

    def kerning(self, left: str, right: str, family: str, bold: bool = False) -> float:
        """kern 表の値（em）。対がなければ 0。"""
        tables = self._font_tables(family, bold)
        pair = (tables.cmap.get(ord(left)), tables.cmap.get(ord(right)))
        return tables.kerning.get(pair, 0) / tables.units

    def glyph_bounds(self, char: str, family: str, bold: bool = False) -> tuple[float, float]:
        """字形の輪郭の (y_min, y_max)（em）。輪郭か字形がなければ (0.0, 0.0)。"""
        box = self._glyph_box(char, family, bold)
        return (0.0, 0.0) if box is None else (box[1], box[3])

    def _glyph_box(self, char: str, family: str, bold: bool = False) -> tuple[float, float, float, float] | None:
        """字形の輪郭の (x_min, y_min, x_max, y_max)（em。曲線の極値）。輪郭か字形がなければ None。"""
        tables = self._font_tables(family, bold)
        name = tables.cmap.get(ord(char))
        if name is None:
            return None
        if name not in tables.bounds:
            # 字形の展開は fontTools の表を書き換えるので、スレッドの間で同時に行わない。
            with self._lock:
                pen = BoundsPen(tables.glyphs)
                tables.glyphs[name].draw(pen)
                box = pen.bounds
                tables.bounds[name] = None if box is None else tuple(value / tables.units for value in box)
        return tables.bounds[name]

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
    ) -> measure.MathBox:
        """数式 1 断片の寸法（pt）。math.metrics で組み、送り幅・インクの上下・行の配置に使う上下を返す。

        display は独立した数式の組み方、family は数式の書体、japanese・latin は \\text の書体（既定はテーマの書体）。
        インクは根号・分数線・伸びた括弧を含む描画範囲、行の配置に使う上下は根号の上の余白などを含む組版上の上下である。
        """
        box = self.math_layout(expr, size, bold, family=family, display=display, japanese=japanese, latin=latin)
        return measure.MathBox(
            box.width, max(0.0, box.ink[2]), max(0.0, box.ink[3]), max(0.0, box.height), max(0.0, box.depth)
        )

    def math_layout(
        self,
        expr: str,
        size: float,
        bold: bool = False,
        *,
        family: str | None = None,
        display: bool = False,
        japanese: str | None = None,
        latin: str | None = None,
    ) -> math_metrics.Box:
        """数式 1 断片の組んだ寸法（`math.metrics.Box`。pt）。引数は `math_box` と同じ。

        送り幅・組版上の上下に加えて、描画範囲の左右（原点は式の左端）と末尾の斜体補正を持つ。送り幅と描画の右端の差は、
        式の末尾の空き（添字の後の SpaceAfterScript など）である（中央揃えの独立した数式の描画の位置を求めるときに使う）。
        """
        return self._math_measurer(family).box(
            str(expr), float(size), bool(bold), display=bool(display), text_fonts=self._text_fonts(japanese, latin)
        )

    def math_constant(self, name: str, *, family: str | None = None) -> float:
        """数式の書体の MATH 表の定数（長さは em、ScriptPercentScaleDown などの割合は 0〜1）。family を省くとテーマの数式の書体。

        未知の名前は KeyError にする。中央揃えの独立した数式の末尾の空き（`writer.display_trailing_space`）が SpaceAfterScript を読む。
        """
        return self._math_measurer(family).font.constant(str(name))

    # 互換用の計測の関数（Project.measure_text・measure_math などが呼ぶ）

    def measure_text(
        self, text: str, size: float, bold: bool = False, *, japanese: str | None = None, latin: str | None = None
    ) -> tuple[float, float]:
        """明示改行だけで区切り、最大の行幅と文字ブロックの高さを返す（measure.layout_text を行間 1.0・折り返しなしで呼ぶ）。

        行末の空白は幅に含めない（PowerPoint の行の幅と同じ）。行の上下は行の書体の組で決まり、欧文の文字だけの行は
        欧文の書体だけで計測する（和文の文字を含む行では和文の書体を読む）。
        """
        style = {
            "size_pt": float(size),
            "bold": bool(bold),
            "line_spacing": 1.0,
            "word_wrap": False,
            "font_latin": latin or self.roles["latin"],
            "font_japanese": japanese or self.roles["japanese"],
        }
        layout = measure.layout_text([{"kind": "text", "text": str(text)}], style, self)
        return layout.content_width, layout.content_height

    def measure_math_metrics(
        self,
        expr: str,
        size: float,
        bold: bool = False,
        *,
        family: str | None = None,
        display: bool = False,
        japanese: str | None = None,
        latin: str | None = None,
    ) -> tuple[float, float, float]:
        """式の幅・全高・基線より下の高さ（pt）。既定は文中の数式の組み方で、\\text はテーマの文字の書体で計測する。"""
        return self._math_measurer(family).measure(
            str(expr), float(size), bool(bold), display=bool(display), text_fonts=self._text_fonts(japanese, latin)
        )

    def measure_math(
        self,
        expr: str,
        size: float,
        bold: bool = False,
        *,
        family: str | None = None,
        display: bool = False,
        japanese: str | None = None,
        latin: str | None = None,
    ) -> tuple[float, float]:
        """式の幅と全高（pt）。context.measure_math() が呼ぶ（呼び方と戻り値は 0.1 と同じ）。"""
        return self.measure_math_metrics(
            expr, size, bold, family=family, display=display, japanese=japanese, latin=latin
        )[:2]

    def text_line_metrics(
        self, text: str, size: float, bold: bool = False, *, japanese: str | None = None, latin: str | None = None
    ) -> tuple[float, float, float, float, float]:
        """1 行の文字列の (送り幅, インクの上, インクの下, 描画の左, 描画の右)（pt）。数式の \\text の計測に使う。

        書体の選び方と kern は measure.layout_text（行間 1.0・折り返しなし）と同じにする。行末の空白も送り幅に含める
        （数式の中の文章では、後ろの空白も次の文字までの間になる）。上下は字形の描画範囲（インク）。描画の左右は
        文字列の左端を原点にした最初と最後の空白でない文字の字形の描画範囲で、描画する文字がなければ (0, 送り幅)。
        """
        text, size, bold = str(text), float(size), bool(bold)
        style = {
            "size_pt": size,
            "bold": bold,
            "line_spacing": 1.0,
            "word_wrap": False,
            "font_latin": latin or self.roles["latin"],
            "font_japanese": japanese or self.roles["japanese"],
        }
        # 文字ごとの書体（どちらの書体にも字形がない文字は、ここで MeasureError になる）
        families = [run.family for run in measure.font_runs(text, style, self, bold=bold) for _char in run.text]

        def advances(start: int, stop: int) -> float:
            return sum(self.advance(text[index], families[index], bold) for index in range(start, stop)) * size

        body = text.rstrip()
        trailing = advances(len(body), len(text))
        if not body:
            return trailing, 0.0, 0.0, 0.0, trailing
        line = measure.layout_text([{"kind": "text", "text": body}], style, self).lines[0]
        first, last = len(body) - len(body.lstrip()), len(body) - 1
        left = advances(0, first)
        first_box = self._glyph_box(body[first], families[first], bold)
        if first_box is not None:
            left += first_box[0] * size
        right = line.width
        last_box = self._glyph_box(body[last], families[last], bold)
        if last_box is not None:
            right -= (self.advance(body[last], families[last], bold) - last_box[2]) * size
        return line.width + trailing, line.ink_top, line.ink_bottom, left, right

    def _math_measurer(self, family: str | None = None) -> math_metrics.MathMeasurer:
        """数式の書体ごとの計測器（初回に作り、資料の計測の間で共有する）。"""
        family = family or self.roles["math"]
        with self._lock:
            measurer = self._math.get(family)
            if measurer is None:
                info = self._face(family)
                if not info["has_math_table"]:
                    raise ValueError(f"数式フォントに MATH テーブルがない: {family}")
                measurer = self._math[family] = math_metrics.MathMeasurer(
                    info["path"], index=info["index"], family=family, functions=self.functions
                )
        return measurer

    def _text_fonts(self, japanese: str | None = None, latin: str | None = None) -> math_metrics.TextFonts:
        """数式の \\text を計測する文字の書体（既定はテーマの書体）。"""
        japanese, latin = japanese or self.roles["japanese"], latin or self.roles["latin"]
        return math_metrics.TextFonts(
            (japanese, latin), partial(self.text_line_metrics, japanese=japanese, latin=latin)
        )

    def missing_glyphs(
        self, text: str, *, bold: bool = False, japanese: str | None = None, latin: str | None = None
    ) -> list[str]:
        """欧文と和文のどちらの書体にも字形がない文字を返す（measure.missing_glyphs と同じ）。"""
        style = {"font_latin": latin or self.roles["latin"], "font_japanese": japanese or self.roles["japanese"]}
        return measure.missing_glyphs(str(text), style, self, bold=bool(bold))

    def inventory(self, *, include_unused: bool = False) -> dict:
        """実際のフォント名・版・書体番号・ハッシュ値と、文字と数式の計測の方式の版を記録する。"""
        with self._lock:
            families = {family for family, _bold in self._faces}
            if include_unused:
                families.update(self.roles.values())
            fonts = {}
            for family in sorted(families):
                info = dict(self._face(family))
                info["roles"] = sorted(role for role, assigned in self.roles.items() if assigned == family)
                if self._specification(family).get("bold_path"):
                    info["bold"] = dict(self._face(family, True))
                fonts[family] = info
        return {
            "fonts": fonts,
            "roles": self.roles,
            "measure_version": measure.MEASURE_VERSION,
            "math_engine": math_metrics.ENGINE_VERSION,
        }

    def close(self):
        """数式の計測器を閉じる（数式フォントのファイルを閉じる）。文字の計測の表はメモリ上にあり、閉じた後も
        使うことができる。

        閉じた後に数式を計測すると、新しい計測器を作る。
        """
        with self._lock:
            measurers = list(self._math.values())
            self._math.clear()
        for measurer in measurers:
            measurer.close()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()
