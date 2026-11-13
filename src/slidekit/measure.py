"""文字と数式の寸法・折り返し・位置を、PowerPoint の配置に合わせて計算する。単位は pt。フォントの値は em で扱う。入力は変更しない。

作成 API（表の自動寸法、内容から決める寸法、`slide.measure()`）、配置と PDF の検査、書き出しの書体の区切りは、
すべてこのモジュールで計算する。フォントは引数の `FontSource` で受け取り、`fonts`・`model`・`validators` を読み込まない。
断片の実効の文字サイズと太字は `styles.fragment_style()`、文章で受け付けない文字は `fragments.control_character()` で決める
（書き出しと検査と同じ規則）。数式の置き方（独立した数式か文中の数式か）は `math.office_math.math_mode()` で、
数式の構造（分数などを含むか）は `math.syntax.parse()` で決める（どちらも関数の中で読み込む）。
"""

from __future__ import annotations

import bisect
from collections import OrderedDict
from collections.abc import Mapping, Sequence
from dataclasses import astuple, dataclass
import difflib
import json
import math
import threading
from typing import Protocol, runtime_checkable
import unicodedata
import weakref

from .errors import InputError
from .fragments import control_character
from .layout import Rect, as_rect
from .styles import fragment_style

__all__ = [
    "MEASURE_VERSION",
    "NO_LINE_START",
    "NO_LINE_END",
    "EAST_ASIAN_FONT_CHARACTERS",
    "EAST_ASIAN_FONT_RANGES",
    "MODEL",
    "Insets",
    "FaceMetrics",
    "MathBox",
    "FontRun",
    "TextLine",
    "TextLayout",
    "WRAP_KINDS",
    "ElementLayout",
    "LayoutModel",
    "MeasureError",
    "FontSource",
    "uses_east_asian_font",
    "normalize_padding",
    "padding_insets",
    "text_area",
    "frame_size",
    "font_runs",
    "missing_glyphs",
    "layout_text",
    "layout_element",
    "layout_scene",
    "phrase_wrapped_content",
    "fit_findings",
    "layout_record",
]

MEASURE_VERSION = "9"  # 計測の方式を変えたら上げる。診断の証拠と生成記録に含める

# 行頭禁則と行末禁則。PowerPoint の折り返しを予測する基準の表で、書き出しの p:kinsoku にも同じ表を書く。
# Office の日本語の「標準」の禁則文字に、PowerPoint 16.0（build 20430）が行頭に置かなかった小書きの仮名と長音記号を加えた
# （校正用の資料で確かめたのは「ょ」「っ」「ー」。ほかは同じ種類の文字として加えた）。全角空白は入れない（空白として扱う）。
# 明示改行の検査（writing）はこの表との差（中点を除く、閉じ括弧などを加える）を定数で示す。設定項目は設けない。
NO_LINE_START = frozenset(
    "!%),.:;?]}¢°’”‰′″℃、。々〉》」』】〕゛゜ゝゞ・ヽヾ！％），．：；？］｝｡｣､･ﾞﾟ￠"
    "ぁぃぅぇぉっゃゅょゎゕゖァィゥェォッャュョヮヵヶㇰㇱㇲㇳㇴㇵㇶㇷㇸㇹㇺㇻㇼㇽㇾㇿー"
    "ｧｨｩｪｫｬｭｮｯｰ"
)
NO_LINE_END = frozenset("$([\\{£¥‘“〈《「『【〔＄（［｛｢￡￥")

# 幅が W・F・H でない文字のうち、PowerPoint が lang="ja-JP" の文字で和文の書体（a:ea）で描く文字（個別の文字と範囲）。
# Unicode の性質（East Asian Width・一般カテゴリー）では決まらないので一覧で持つ。PowerPoint 16.0.20430.20182 で、
# Arial と Yu Gothic UI の組の 20 pt の文字（BMP の 2,878 文字）を、書体の区切りなしで描いた PDF の書体で決めた。
# 一覧の外の幅の曖昧な文字（→ … — – Ł é · ¶ † ‡ ≠ ∞ ′ ″、ギリシャ文字など）は欧文の書体で描かれる。
# “ ” ‘ ’ は lang="en-US" なら欧文の書体になるが、書き出しは常に lang="ja-JP" を書き、forced で両方の書体を指定する。
# ꞵ ꭓ は理由が分からないが和文の書体で描かれたので入れた。書き出しの書体の区切り（font_runs）も同じ一覧で決まる。
EAST_ASIAN_FONT_CHARACTERS = frozenset("§¨°±´×÷‘’‚‛“”„⁇⁈⹀ꞵꭓ")
EAST_ASIAN_FONT_RANGES = (
    (0x2460, 0x24FF),  # 囲み英数字
    (0x25A0, 0x25FF),  # 幾何学模様
    (0x2600, 0x26FF),  # その他の記号
    (0x2700, 0x27BF),  # 装飾記号
    (0x3200, 0x32FF),  # 囲み CJK 文字・月
)

_SIDES = ("left", "top", "right", "bottom")
_SHAPES = ("rectangle", "rounded_rectangle", "ellipse", "diamond")
_TEXT_KINDS = frozenset({"text", "box", "equation"})
_FRAGMENT_KINDS = frozenset({"text", "emphasis", "math"})
_ALIGNS = ("left", "center", "right", "justify")
_VALIGNS = ("top", "middle", "bottom")
_HYPHENS = frozenset("-‐")
# 折り返しで空白として扱う文字（前で分けず、前の行の末尾に残し、行末では幅に含めない）。
_SPACES = frozenset(" 　")
# 行間をちょうど 1 とみなす差。書き出しの行間（a:spcPct）は 1/100000 単位なので、その丸めの範囲にする。
_EXACT_SPACING = 5e-6
_MATH_LINE_ABOVE = 0.75  # 行間を掛けた全高のうち基線より上に置く割合（背の高い式では上限）
_ROUNDED_ADJUST = 0.16667  # 角丸の大きさの既定（ECMA-376 の roundRect の adj）
_ROUNDED_TEXT = 0.29289  # 角丸の文字領域の係数（同 rect の il）
_DIAGONAL = math.sqrt(0.5)  # cos 45°（楕円の文字領域）
_EPSILON = 1e-6
_CACHE_SIZE = 4096
_PASSES = 50


class MeasureError(InputError):
    """計測できない入力。`code` は診断の識別子と同じ名前にする。

    - `font_glyph_missing`：欧文と和文のどちらの書体にも字形がない文字がある
    - `invalid_fragment`：制御文字など、文章の断片が不正
    - `invalid_style`：余白・文字サイズ・幅などの値が不正
    - `text_measurement_failed`：それ以外（数式の誤り、フォントを読むことができないなど）

    `details` は診断の観測値に入れる。ほかの例外から作ったときは `__cause__` に元の例外を持つ。
    """

    code = "text_measurement_failed"
    CODES = frozenset({"font_glyph_missing", "invalid_fragment", "invalid_style", "text_measurement_failed"})

    def __init__(self, message: str, *, code: str | None = None, **kwargs) -> None:
        if code is not None and code not in self.CODES:
            raise ValueError(f"計測の誤りの code が不正です: {code}")
        super().__init__(message, code=code, **kwargs)


@dataclass(frozen=True)
class Insets:
    """枠から内側までの距離（pt）。"""

    left: float
    top: float
    right: float
    bottom: float


@dataclass(frozen=True)
class FaceMetrics:
    """フォント 1 書体の値（em）。"""

    family: str
    bold: bool
    ascent: float  # OS/2 の usWinAscent ÷ unitsPerEm（行の配置の割合に使う）
    descent: float  # OS/2 の usWinDescent ÷ unitsPerEm
    has_kerning: bool  # kern 表（format 0、横書き）があるか


@dataclass(frozen=True)
class MathBox:
    """数式 1 断片の寸法（pt、基線基準）。"""

    width: float  # 送り幅
    ascent: float  # 基線より上のインク（根号・分数線・伸びた括弧を含む。0 以上）
    descent: float  # 基線より下のインク（0 以上）
    line_ascent: float  # 行の配置に使う上側。行間 1.0 での値
    line_descent: float  # 行の配置に使う下側。行間 1.0 での値
    # 描画範囲の左右（式の左端を原点とする。左は送りの始まりより左なら負）。文字の送りには使わず、はみ出しの検査に使う。
    # None は送り幅の範囲（0、width）と同じ。入れ子の分数の線は送り幅より左右へ張り出す。
    ink_left: float | None = None
    ink_right: float | None = None


@dataclass(frozen=True)
class FontRun:
    """出力する通常文字の区切り。"""

    text: str
    family: str  # この区切りで使う書体
    forced: bool  # 真なら a:latin と a:ea の両方を family にする


@dataclass(frozen=True)
class TextLine:
    text: str  # 行の内容。数式は $…$ で表す（記録と診断の表示用）
    width: float  # 行末の空白を除いた幅（pt）。両端揃えで折り返した行は内側の幅
    baseline: float  # 文字ブロックの上端から基線まで（pt）
    ink_top: float  # 基線からインクの上端まで（pt、上向きが正。インクのない行は 0）
    ink_bottom: (
        float  # 基線からインクの下端まで（pt、下向きが正。基線より上だけにある字形の行では負。インクのない行は 0）
    )
    hard_break: bool  # 明示改行（\n）で終わる行か
    near_limit: bool  # 折り返しの判定が計測の誤差の範囲にある行か
    split: bool = False  # 内側の幅より長い語を、PowerPoint と同じく語の途中で分けた行か
    # 数式の描画が行の送りの範囲（左端から width まで）より左・右へ張り出す量（pt、0 以上）。入れ子の分数の線など
    overhang_left: float = 0.0
    overhang_right: float = 0.0
    # 行末の自動の折り返しの種類（明示改行で終わる行と最後の行は None）。WRAP_KINDS の値
    wrap: str | None = None
    characters: int = 0  # 空白を除いた文字の数（数式は 1 つと数える）
    origin: tuple[int, int] | None = None  # 行の最初の単位の、断片の番号と断片の文字列の中の位置（空の行は None）


# 自動の折り返しの種類。space は空白の後、math は数式の前後、punctuation は約物や欧文と和文の境、phrase は和文の文節の境、
# mid_phrase は和文の文節の途中、hyphen は欧文の語のハイフンの後、split は内側より長い語の途中、other はそれ以外
# （文節の区切りを判定できない場合を含む）。
WRAP_KINDS = ("space", "math", "punctuation", "phrase", "mid_phrase", "hyphen", "split", "other")


@dataclass(frozen=True)
class TextLayout:
    """公開する計測結果（slide.measure() の戻り値）。"""

    lines: tuple[TextLine, ...]
    content_width: float  # 最も広い行の幅
    content_height: float  # 文字ブロックの高さ
    available_width: float | None  # 折り返しに使った内側の幅（折り返さない場合は None）
    unbreakable_width: float  # 途中に分ける位置のない部分の最大の幅
    insets: Insets  # 図形の文字領域の差と余白の和（求めた枠の寸法での値）
    width: float  # 枠の幅（指定した幅、なければ内容から）
    height: float  # 必要な枠の高さ（文字ブロックと上下の距離の和）
    word_wrap: bool
    natural_width: float = 0.0  # 折り返さないときの最も広い行の幅（明示改行だけで分けた行。1 行に収めるのに要る幅）


@dataclass(frozen=True)
class ElementLayout:
    """枠の中に置いた結果（検査用。絶対座標）。"""

    element_id: str
    frame: Rect  # 要素の領域
    text_area: Rect  # 図形の文字領域
    inner: Rect  # 内側
    text: TextLayout
    block_top: float  # 文字ブロックの上端（valign を反映）
    line_left: tuple[float, ...]  # 各行の左端（align を反映）
    ink: Rect | None  # 予測したインクの外接矩形（文字がなければ None）
    visible_frame: bool  # 塗りか線がある枠か


@dataclass(frozen=True)
class LayoutModel:
    """PowerPoint の配置を近似する定数。PowerPoint で変換した PDF の実測から決め、校正用の資料の照合で確かめる。"""

    line_pitch_em: float = 1.2  # 行送り = 1.2 × 文字サイズ × 行間（同じ書体の組の行どうし）
    line_top_limit_em: float = 0.9035  # 行間が 1 より大きいときの、行の上側の書体の値の上限（計測した値 0.900〜0.908）
    math_stack_top_ratio: float = (
        1.2  # 背の高い数式（インクの上下の和が math_stack_full_em 未満）の行の上側 ÷（行間 × インクの上）
    )
    math_stack_gap_em: float = (
        0.147  # 背の高い数式の行の、行間 1.0 での上側 − インクの上（em。ほかの数式の行間 1.0 の下限にも使う）
    )
    math_stack_full_em: float = (
        1.3  # 行間で増える分を数式の行の高さ h から求める、背の高い数式のインクの上下の和の下限（em）
    )
    math_stack_above_share: float = (
        0.75  # 同じく、行間で増える分のうち行の上に置く割合（行間で増える分が下の境を超える式）
    )
    math_stack_above_full: float = 1.17  # 同じく、行間で増える分がこの値 × インクの下 d 以下なら、すべて行の上に置く
    math_radical_top_ratio: float = 1.17  # 文中の背の高い根号の行の上側 ÷（行間 × インクの上）。行間が 1 より大きいとき
    math_radical_single_ratio: float = 1.22  # 同じく、行間がちょうど 1 の行の上側 ÷ インクの上
    math_descent_above: float = 0.5  # 文中の数式の下が文字の行の下より深い分のうち、行間 × この値を行の上に加える
    kerning_min_size_pt: float = 12.0  # 欧文の kern 表を使う最小の文字サイズ（出力の kern="1200" と対応）
    wrap_band_ratio: float = 0.004  # 折り返しの判定の誤差（行の幅に比例する分）
    wrap_band_pt: float = 0.25  # 同（一定の分）
    visible_tolerance_pt: float = 0.5  # 塗りか線のある枠からのはみ出しの許容
    hidden_tolerance_pt: float = 1.5  # 塗りも線もない枠からのはみ出しの許容
    padding_overlap_min_pt: float = 1.0  # 余白へのはみ出しを警告する最小の量
    padding_overlap_ratio: float = 0.25  # 同（その辺の余白に対する割合）
    superscript_offset_em: float = 0.3  # 上付きの断片の基準線のずらし（文字サイズに対する割合。書き出しの baseline）
    superscript_scale: float = 2 / 3  # 上付きの断片を PowerPoint が描く大きさの倍率
    short_line_characters: int = 2  # 自動で折り返した段落の最後の行を短いとみなす文字の数（以下）
    short_line_ratio: float = 0.2  # 同じく、内側の幅に対する割合（未満）


MODEL = LayoutModel()


@runtime_checkable
class FontSource(Protocol):
    """計測に使うフォントの値。`fonts.FontRegistry` と試験用の偽物が実装する。

    `face_metrics` の ascent・descent は OS/2 の usWinAscent・usWinDescent（em）で、行の配置の割合に使う。
    `advance`・`kerning`・`glyph_bounds` は em で返す。字形がない文字の `advance` は `MeasureError`。
    `math_box` は pt で返す。`display` は独立した数式の組み方、`japanese`・`latin` は `\\text` の書体。
    """

    def face_metrics(self, family: str, bold: bool = False) -> FaceMetrics: ...

    def has_glyph(self, char: str, family: str, bold: bool = False) -> bool: ...

    def advance(self, char: str, family: str, bold: bool = False) -> float: ...

    def kerning(self, left: str, right: str, family: str, bold: bool = False) -> float: ...

    def glyph_bounds(self, char: str, family: str, bold: bool = False) -> tuple[float, float]: ...

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
    ) -> MathBox: ...


# ---- 値の検査 ----


def _number(value) -> bool:
    return not isinstance(value, bool) and isinstance(value, (int, float)) and math.isfinite(value)


def _style_error(message: str, **details) -> MeasureError:
    return MeasureError(message, code="invalid_style", **details)


def _subject(name: str) -> str:
    # 英字の名前と日本語の間にだけ空白を入れる。
    return f"{name} は" if name.isascii() else f"{name}は"


def _positive(value, name: str) -> float:
    if not _number(value) or value <= 0:
        raise _style_error(f"{_subject(name)}正の有限値にしてください", setting=name, value=value)
    return float(value)


def _nonnegative(value, name: str) -> float:
    if not _number(value) or value < 0:
        raise _style_error(f"{_subject(name)} 0 以上の有限値にしてください", setting=name, value=value)
    return float(value)


def _shape(shape) -> str:
    if shape not in _SHAPES:
        raise _style_error(f"未対応の図形です: {shape}（候補: {'、'.join(_SHAPES)}）", setting="shape", value=shape)
    return shape


def _check_text(text: str) -> None:
    """文章で受け付けない文字（\\n 以外の制御文字など。作成 API の断片の検査と同じ規則）があれば MeasureError にする。"""
    char = control_character(text)
    if char is not None:
        raise MeasureError(
            f"文章にタブなどの制御文字は使えません（改行は \\n だけを使えます）: U+{ord(char):04X}",
            code="invalid_fragment",
            character=f"U+{ord(char):04X}",
        )


def _east_asian(char: str) -> bool:
    return unicodedata.east_asian_width(char) in {"W", "F", "H"}


def uses_east_asian_font(char: str) -> bool:
    """幅が W・F・H でない文字のうち、PowerPoint が既定で和文の書体（a:ea）で描く文字か（EAST_ASIAN_FONT_CHARACTERS・
    EAST_ASIAN_FONT_RANGES）。幅が W・F・H の文字は対象にしない（偽を返す）。"""
    if len(char) != 1 or _east_asian(char):
        return False
    if char in EAST_ASIAN_FONT_CHARACTERS:
        return True
    code = ord(char)
    return any(start <= code <= end for start, end in EAST_ASIAN_FONT_RANGES)


def _band(model: LayoutModel, width: float) -> float:
    return model.wrap_band_ratio * width + model.wrap_band_pt


def _ceil(value: float) -> float:
    """小数第 2 位で切り上げる（0.1 pt 単位）。計算の誤差だけの差では切り上げない。"""
    return math.ceil(value * 10 - _EPSILON) / 10


def _round(value: float) -> float:
    return round(value, 2)


def _rect_list(rect: Rect) -> list[float]:
    return [_round(rect.x), _round(rect.y), _round(rect.w), _round(rect.h)]


def _bounds_list(rect: Rect | None) -> list[float] | None:
    if rect is None:
        return None
    return [_round(rect.x), _round(rect.y), _round(rect.right), _round(rect.bottom)]


# ---- 余白と図形 ----


def _side_value(value) -> float:
    if not _number(value) or value < 0:
        raise _style_error("padding_pt の値は 0 以上の有限値にしてください", setting="padding_pt", value=value)
    return float(value)


def _padding_sides(value, inherited: dict[str, float]) -> dict[str, float]:
    sides = dict(inherited)
    if value is None:
        return sides
    if isinstance(value, Mapping):
        unknown = [key for key in value if key not in _SIDES]
        if unknown:
            names = []
            for key in unknown:
                candidates = difflib.get_close_matches(str(key), _SIDES, n=3, cutoff=0.6)
                names.append(f"{key}（候補: {'、'.join(candidates)}）" if candidates else str(key))
            raise _style_error(
                f"padding_pt に未知の辺があります: {'、'.join(names)}",
                setting="padding_pt",
                unknown=[str(key) for key in unknown],
            )
        for side, item in value.items():
            sides[side] = _side_value(item)
        return sides
    if _number(value):
        return dict.fromkeys(_SIDES, _side_value(value))
    raise _style_error(
        "padding_pt は 0 以上の数値か、left・top・right・bottom を持つ辞書で指定してください",
        setting="padding_pt",
        value=value,
    )


def normalize_padding(value, *, base=0.0) -> float | dict[str, float]:
    """padding_pt の指定を検査して正規化する。

    数値は 4 辺に同じ値、辞書は一部の辺だけでもよく、ない辺は `base`（優先順位の低い層で解決済みの値）を引き継ぐ。
    4 辺が同じなら数値、違えば 4 辺を持つ辞書を返す。数値の組（タプル・リスト）は辺の順序を間違えやすいので受け付けない。
    """
    inherited = _padding_sides(0.0 if base is None else base, dict.fromkeys(_SIDES, 0.0))
    sides = _padding_sides(value, inherited)
    values = [sides[side] for side in _SIDES]
    if all(item == values[0] for item in values):
        return values[0]
    return {side: sides[side] for side in _SIDES}


def padding_insets(style: Mapping) -> Insets:
    """場面の書式の padding_pt（数値か辺の辞書。ない辺は 0）を Insets にする。"""
    sides = _padding_sides(style.get("padding_pt", 0), dict.fromkeys(_SIDES, 0.0))
    return Insets(sides["left"], sides["top"], sides["right"], sides["bottom"])


def _rounded_inset(shortest: float, corner_radius_pt: float | None) -> float:
    if shortest <= 0:
        return 0.0
    adjust = _ROUNDED_ADJUST if corner_radius_pt is None else min(0.5, max(0.0, corner_radius_pt / shortest))
    return _ROUNDED_TEXT * adjust * shortest


def text_area(area: Rect, shape: str = "rectangle", corner_radius_pt: float | None = None) -> Rect:
    """図形の文字領域。ECMA-376 の図形の定義（presetShapeDefinitions）の rect に合わせる。

    角丸の大きさは書き出しと同じく `min(0.5, corner_radius_pt / min(w, h))`（指定がなければ 0.16667）。
    """
    area = as_rect(area)
    shape = _shape(shape)
    if corner_radius_pt is not None:
        corner_radius_pt = _nonnegative(corner_radius_pt, "corner_radius_pt")
    if shape == "rectangle":
        return area
    if shape == "rounded_rectangle":
        dx = dy = _rounded_inset(min(area.w, area.h), corner_radius_pt)
    elif shape == "ellipse":
        dx, dy = area.w * (1 - _DIAGONAL) / 2, area.h * (1 - _DIAGONAL) / 2
    else:
        dx, dy = area.w / 4, area.h / 4
    return Rect(area.x + dx, area.y + dy, max(0.0, area.w - 2 * dx), max(0.0, area.h - 2 * dy))


def frame_size(
    shape: str,
    content_width: float,
    content_height: float,
    padding: Insets,
    *,
    width: float | None = None,
    corner_radius_pt: float | None = None,
) -> tuple[float, float]:
    """内容の寸法から枠の寸法を求める（text_area の逆）。`width` を指定すると幅はその値にし、高さだけを求める。

    角丸の文字領域は min(w, h) に依存するので、どちらが短いかの場合を含めて解く。
    """
    shape = _shape(shape)
    content_width = _nonnegative(content_width, "内容の幅")
    content_height = _nonnegative(content_height, "内容の高さ")
    if corner_radius_pt is not None:
        corner_radius_pt = _nonnegative(corner_radius_pt, "corner_radius_pt")
    if width is not None:
        width = _nonnegative(width, "幅")
    horizontal = content_width + padding.left + padding.right
    vertical = content_height + padding.top + padding.bottom
    if shape == "rectangle":
        return (horizontal if width is None else width), vertical
    if shape == "ellipse":
        return (horizontal / _DIAGONAL if width is None else width), vertical / _DIAGONAL
    if shape == "diamond":
        return (horizontal * 2 if width is None else width), vertical * 2
    # 角丸：差 il は min(w, h) について単調で傾きが 0.15 未満なので、下から繰り返すと最小の解に収束する。
    w = horizontal if width is None else width
    h = vertical
    for _ in range(200):
        inset = _rounded_inset(min(w, h), corner_radius_pt)
        new_w = horizontal + 2 * inset if width is None else width
        new_h = vertical + 2 * inset
        if abs(new_w - w) < 1e-12 and abs(new_h - h) < 1e-12:
            break
        w, h = new_w, new_h
    return new_w, new_h


# ---- 書式と断片 ----


@dataclass(frozen=True)
class _Style:
    size: float
    bold: bool
    latin: str
    japanese: str
    math: str | None
    line_spacing: float | None
    line_spacing_pt: float | None
    padding: Insets
    word_wrap: bool
    align: str
    valign: str
    corner_radius: float | None
    phrase_wrap: bool = False  # 和文の文字の間では、文節の境（と幅より長い文節の中）だけで折り返す


def _family(style: Mapping, key: str, *, required: bool = True) -> str | None:
    value = style.get(key)
    if value is None and not required:
        return None
    if not isinstance(value, str) or not value.strip():
        raise _style_error(f"{key} は書体名の文字列で指定してください", setting=key, value=value)
    return value


def _text_style(style: Mapping) -> _Style:
    """解決済みの書式を検査する。項目がない場面（古い版で作った場面など）は、word_wrap を偽、
    bold を偽、align を left、valign を top、余白を 0 とみなす。"""
    if not isinstance(style, Mapping):
        raise _style_error("書式は辞書で指定してください", value=style)
    line_spacing_pt = style.get("line_spacing_pt")
    line_spacing_pt = None if line_spacing_pt is None else _positive(line_spacing_pt, "line_spacing_pt")
    line_spacing = style.get("line_spacing")
    if line_spacing is None and line_spacing_pt is None:
        raise _style_error("line_spacing は正の有限値にしてください", setting="line_spacing", value=None)
    line_spacing = None if line_spacing is None else _positive(line_spacing, "line_spacing")
    word_wrap = style.get("word_wrap", False)
    if word_wrap is None:
        word_wrap = False
    if not isinstance(word_wrap, bool):
        raise _style_error("word_wrap は真偽値で指定してください", setting="word_wrap", value=word_wrap)
    bold = style.get("bold", False)
    if not isinstance(bold, bool):
        raise _style_error("bold は真偽値で指定してください", setting="bold", value=bold)
    align = style.get("align") or "left"
    if align not in _ALIGNS:
        raise _style_error(f"align は {'、'.join(_ALIGNS)} のいずれかにしてください", setting="align", value=align)
    valign = style.get("valign") or "top"
    if valign not in _VALIGNS:
        raise _style_error(f"valign は {'、'.join(_VALIGNS)} のいずれかにしてください", setting="valign", value=valign)
    radius = style.get("corner_radius_pt")
    phrase_wrap = style.get("phrase_wrap") or False
    if not isinstance(phrase_wrap, bool):
        raise _style_error("phrase_wrap は真偽値で指定してください", setting="phrase_wrap", value=phrase_wrap)
    return _Style(
        size=_positive(style.get("size_pt"), "size_pt"),
        bold=bold,
        latin=_family(style, "font_latin"),
        japanese=_family(style, "font_japanese"),
        math=_family(style, "font_math", required=False),
        line_spacing=line_spacing,
        line_spacing_pt=line_spacing_pt,
        padding=padding_insets(style),
        word_wrap=word_wrap,
        align=align,
        valign=valign,
        corner_radius=None if radius is None else _nonnegative(radius, "corner_radius_pt"),
        phrase_wrap=phrase_wrap,
    )


def _fragments(content) -> list[tuple[str, str, Mapping, Mapping]]:
    if isinstance(content, (str, bytes)) or not isinstance(content, Sequence):
        raise MeasureError("文章は断片（kind と text を持つ辞書）の列で指定してください", code="invalid_fragment")
    result = []
    for fragment in content:
        if not isinstance(fragment, Mapping):
            raise MeasureError("文章の断片は kind と text を持つ辞書にしてください", code="invalid_fragment")
        kind = fragment.get("kind", "text")
        if kind not in _FRAGMENT_KINDS:
            raise MeasureError(f"未対応の文章の断片です: {kind}", code="invalid_fragment", kind=str(kind))
        text = fragment.get("text", "")
        if not isinstance(text, str):
            raise MeasureError("文章の断片の text は文字列にしてください", code="invalid_fragment")
        own = fragment.get("style") or {}
        if not isinstance(own, Mapping):
            raise MeasureError("文章の断片の style は辞書にしてください", code="invalid_fragment")
        if kind != "math":
            _check_text(text)
        result.append((kind, text, own, fragment))
    return result


def _fragment_font(own: Mapping, fragment: Mapping, style: _Style) -> tuple[float, bool]:
    """断片の実効の文字サイズと太字。値を確かめてから、書き出しと検査と同じ styles.fragment_style で求める
    （断片の指定 > 要素。強調の断片と、emphasis が真の数式は、指定がなければ太字）。"""
    size = own.get("size_pt")
    if size is not None and (not _number(size) or size <= 0):
        raise MeasureError("文章の断片の size_pt は正の有限値にしてください", code="invalid_fragment", value=size)
    bold = own.get("bold")
    if bold is not None and not isinstance(bold, bool):
        raise MeasureError("文章の断片の bold は真偽値にしてください", code="invalid_fragment", value=bold)
    # 要素の文字サイズと太字は _text_style で確かめた値を渡す（書式に bold がない古い場面は偽）。
    values = fragment_style({"size_pt": style.size, "bold": style.bold}, fragment)
    return float(values["size_pt"]), values["bold"]


# ---- 文字ごとの書体 ----


def _char_font(char: str, latin: str, japanese: str, fonts: FontSource, bold: bool) -> tuple[str, bool] | None:
    """文字を計測する書体と、a:latin・a:ea の両方に書くか（forced）。どちらの書体にも字形がなければ None。

    - 幅が W・F・H の文字：和文の書体。和文の書体になければ、欧文の書体で forced。
    - PowerPoint が既定で和文の書体で描く文字（uses_east_asian_font。○ × ± “ など）：和文の書体にあれば和文の書体、
      なければ欧文の書体で、どちらも forced（PowerPoint の既定の書体と、言語の設定によらず同じ書体で描かせる）。
    - それ以外：欧文の書体にあれば欧文の書体（幅が A なら forced）、なければ和文の書体で forced。
    """
    if _east_asian(char):
        if fonts.has_glyph(char, japanese, bold):
            return japanese, False
        # 和文の書体にない全角の文字は、欧文の書体にあればそれで表示させる。
        return (latin, True) if fonts.has_glyph(char, latin, bold) else None
    if uses_east_asian_font(char):
        if fonts.has_glyph(char, japanese, bold):
            return japanese, True
        return (latin, True) if fonts.has_glyph(char, latin, bold) else None
    if fonts.has_glyph(char, latin, bold):
        return latin, unicodedata.east_asian_width(char) == "A"
    if fonts.has_glyph(char, japanese, bold):
        return japanese, True
    return None


def _glyph_missing(characters: Sequence[str], latin: str, japanese: str) -> MeasureError:
    listed = "、".join(f"'{char}' (U+{ord(char):04X})" for char in characters)
    return MeasureError(
        f"{latin} と {japanese} のどちらにも字形がない文字です: {listed}",
        code="font_glyph_missing",
        characters=list(characters),
        fonts=[latin, japanese],
    )


def _plain_text(text: str) -> str:
    if "\n" in text:
        raise MeasureError("改行を含む文字列は行ごとに分けて渡してください", code="invalid_fragment")
    _check_text(text)
    return text


def font_runs(text: str, style: Mapping, fonts: FontSource, *, bold: bool) -> list[FontRun]:
    """通常文字を同じ書体と forced の区切りにまとめる。どちらの書体にも字形がない文字は MeasureError。"""
    text = _plain_text(text)
    latin, japanese = _family(style, "font_latin"), _family(style, "font_japanese")
    runs: list[list] = []
    missing = []
    for char in text:
        font = _char_font(char, latin, japanese, fonts, bold)
        if font is None:
            if char not in missing:
                missing.append(char)
            continue
        if runs and (runs[-1][1], runs[-1][2]) == font:
            runs[-1][0] += char
        else:
            runs.append([char, *font])
    if missing:
        raise _glyph_missing(missing, latin, japanese)
    return [FontRun(part, family, forced) for part, family, forced in runs]


def missing_glyphs(text: str, style: Mapping, fonts: FontSource, *, bold: bool) -> list[str]:
    """欧文と和文のどちらの書体にも字形がない文字を、重複なく並べて返す。改行は対象にしない。"""
    latin, japanese = _family(style, "font_latin"), _family(style, "font_japanese")
    found = set()
    for line in str(text).split("\n"):
        _check_text(line)
        found.update(char for char in line if _char_font(char, latin, japanese, fonts, bold) is None)
    return sorted(found)


# ---- 行の組み立て ----


@dataclass(slots=True)
class _Unit:
    """折り返しの単位（文字 1 つか数式 1 つ）。"""

    text: str
    math: bool
    width: float  # 送り幅（pt）
    kern: float  # 直前の単位との kern（pt）。同じ行に並ぶときだけ加える
    size: float
    bold: bool
    space: bool  # 折り返しで空白として扱う文字（U+0020 と U+3000）
    east_asian: bool  # W・F・H
    ink: tuple[float, float] | None  # 基線より上・下のインク（pt）。インクがなければ None
    box: MathBox | None
    family: str | None = None  # 通常文字を計測した書体（数式は None）
    stacked: bool = False  # 背の高い数式の規則（_math_line）に当たる数式
    under: bool = False  # 同じく、下に置くもの（\underset・\limits の下の添字）を含む数式
    radical: bool = False  # 文中の背の高い根号の規則（_math_line）に当たる数式
    origin: tuple[int, int] | None = None  # 断片の番号と、その断片の文字列の中の位置（明示改行を入れる位置に使う）


@dataclass(slots=True)
class _HardLine:
    units: list
    size: float  # 空の行の文字サイズ
    bold: bool


@dataclass(slots=True)
class _Piece:
    units: Sequence[_Unit]
    width: float
    soft: bool  # 折り返しで終わる行か
    hard_break: bool
    near: bool
    size: float
    bold: bool
    split: bool  # 語の途中で分けた行か
    wrap: str | None = None  # 行末の自動の折り返しの種類


class _Faces:
    """1 回の計測で使う書体の値を覚える。"""

    def __init__(self, fonts: FontSource, style: _Style, model: LayoutModel):
        self.fonts, self.style, self.model = fonts, style, model
        self._metrics: dict[tuple[str, bool], FaceMetrics] = {}
        self._fonts: dict[tuple[str, bool], tuple[str, bool] | None] = {}
        self._natural: dict[tuple[frozenset[str], bool], tuple[float, float]] = {}

    def metrics(self, family: str, bold: bool) -> FaceMetrics:
        key = (family, bold)
        if key not in self._metrics:
            self._metrics[key] = self.fonts.face_metrics(family, bold)
        return self._metrics[key]

    def natural(self, families: frozenset[str], bold: bool) -> tuple[float, float]:
        """行の書体の組での、行の上側と最後の基線より下の割合（em。和は line_pitch_em）。

        書体ごとに A_f = 1.2 × usWinAscent ÷（usWinAscent＋usWinDescent）、D_f = 1.2 − A_f とし、
        A* = max A_f、D* = max D_f から、上側 1.2 × A* ÷（A*＋D*）と下 1.2 − 上側を返す。
        """
        key = (families, bold)
        if key not in self._natural:
            pitch = self.model.line_pitch_em
            ascents, descents = [], []
            for family in sorted(families):
                face = self.metrics(family, bold)
                total = face.ascent + face.descent if _number(face.ascent) and _number(face.descent) else 0.0
                if total <= 0 or face.ascent < 0 or face.descent < 0:
                    raise MeasureError(
                        f"書体の上下の値が不正です: {family}（ascent {face.ascent}、descent {face.descent}）",
                        family=family,
                        ascent=face.ascent,
                        descent=face.descent,
                    )
                ascents.append(pitch * face.ascent / total)
                descents.append(pitch - ascents[-1])
            ascent = pitch * max(ascents) / (max(ascents) + max(descents))
            self._natural[key] = (ascent, pitch - ascent)
        return self._natural[key]

    def font(self, char: str, bold: bool) -> tuple[str, bool] | None:
        key = (char, bold)
        if key not in self._fonts:
            self._fonts[key] = _char_font(char, self.style.latin, self.style.japanese, self.fonts, bold)
        return self._fonts[key]


def _math_box(fonts: FontSource, expr: str, size: float, bold: bool, style: _Style, display: bool) -> MathBox:
    try:
        box = fonts.math_box(
            expr, size, bold, family=style.math, display=display, japanese=style.japanese, latin=style.latin
        )
    except MeasureError as error:
        error.details.setdefault("expression", expr)
        raise
    except ValueError as exc:
        raise MeasureError(f"数式の寸法を計算できません: {exc}", expression=expr, reason=str(exc)) from exc
    values = [getattr(box, name, None) for name in ("width", "ascent", "descent", "line_ascent", "line_descent")]
    if not all(_number(value) for value in values) or values[0] < 0:
        raise MeasureError("数式の寸法が不正です", expression=expr, box=repr(box))
    sides = [getattr(box, name, None) for name in ("ink_left", "ink_right")]
    if not all(value is None or _number(value) for value in sides):
        raise MeasureError("数式の描画範囲が不正です", expression=expr, box=repr(box))
    return MathBox(
        *(max(0.0, float(value)) for value in values), *(None if value is None else float(value) for value in sides)
    )


def _structure(expr: str) -> str | None:
    """行の置き方を決める数式の構造。添字（上付き・下付き）の中と、上下に置くもの（\\overset・\\underset・\\limits）の中は調べない。

    - "under"：下に置くもの（\\underset、\\limits の下の添字）を含む（分数などの中を除く）。
    - "stacked"：下に置くものを含まず、分数（横書きの / を除く）・\\binom、伸びる括弧（\\left・\\right）、大型演算子の
      どれかを含む。
    - "radical"：上のどれも含まず、根号を含み、根号の外に上へ出る部分（上付きの添字、上に置くもの、アクセント、上線）がない
      （根号の規則は式のインクの上を根号の上として使うので、根号より上へ出うる部分がある式には当てない）。
    - None：それ以外。構文の誤りは数式の計測が先に報告するので、ここでは None にする。

    添字の中の分数（a^{\\frac{a}{c}}）と上だけに置くもの（\\overset{n}{x}）には、背が高くても通常の数式の行の規則を
    当てる（_math_line）。下に置くものは \\underset{x \\in S}{\\arg\\min} f(x)・
    \\underset{i}{\\max} x_i・\\lim\\limits_{n \\to \\infty} a_n で確かめた。添字の土台の中の分数（(\\frac{a}{b})^2）は規則に当たる。
    """
    from .math import syntax

    found = {"stacked": False, "under": False, "radical": False, "raised": False}

    def visit(nodes, inside: bool) -> None:
        for node in nodes:
            if (
                isinstance(node, syntax.Fraction)
                and node.kind != "lin"
                or isinstance(node, syntax.Fenced)
                and node.stretchy
                or isinstance(node, syntax.Nary)
            ):
                found["stacked"] = True
                continue
            if isinstance(node, syntax.Limits):
                if node.under:
                    found["stacked"] = found["under"] = True
                    continue
                found["raised"] = found["raised"] or not inside
                continue
            if isinstance(node, syntax.Scripts):
                found["raised"] = found["raised"] or (node.sup is not None and not inside)
                visit(node.base, inside)
                continue
            if isinstance(node, syntax.Radical):
                found["radical"] = True
                visit(node.degree, True)
                visit(node.body, True)
                continue
            if isinstance(node, syntax.Accent) or isinstance(node, syntax.Bar) and node.position == "top":
                found["raised"] = found["raised"] or not inside
            for group in syntax.children(node):
                visit(group, inside)

    try:
        visit(syntax.parse(expr), False)
    except (ValueError, TypeError):
        return None
    if found["stacked"]:
        return "under" if found["under"] else "stacked"
    return "radical" if found["radical"] and not found["raised"] else None


def _math_line(
    unit: "_Unit", spacing: float, model: LayoutModel, floor: float = 0.0, bottom: float | None = None
) -> tuple[float, float, bool]:
    """数式の単位の、行の上側と下側（pt）と、規則で置いたか（下側を行間によらない値にしたか）。

    背の高い数式（unit.stacked）は、行間 1 以上で次のように置く（s は数式の断片の文字サイズ、sp は行間）。
    a = max(インクの上 + math_stack_gap_em × s, floor)、d = インクの下、h = a + d とする。floor は行の書体の組の行間 1.0 の
    上側（pt）で、下が深い \\underset と下が浅い \\underset の行で確かめた。

    - インクの上下の和が math_stack_full_em × s 以上の式と、下に置くもの（unit.under）を含む式：行間で増える分
      e = (sp − 1) × h が math_stack_above_full × d 以下なら e をすべて行の上に、超えれば math_stack_above_share × e を
      行の上に置き、残りを行の下に置く（上側 a + 上の分、下側 d + 残り）。切り替わる行間は式の d と h で決まり、
      一定の行間や h の割合では分けることができない（20 pt の独立した分数は行間 1.3 まですべて上、総和は 1.4 まで
      すべて上で、どちらも次の行間から 4 分の 3 を上）。行間が 1 より大きいときは、上側を全高 sp × h の
      4 分の 3 以下に抑え、減らした分を下に回す。上に偏る入れ子の分数でも、行全体の高さは変えない。
    - 和がそれより小さく、下に置くものを含まない式：上側 math_stack_top_ratio × sp × インクの上、下側 d。
    - どちらも上側は数式の計測の line_ascent × sp より小さくしない。d には行間を掛けない。最後の行の下側は _arrange で決める
      （独立した数式だけの行は d と残り、文字の段落の行は d）。

    文中の背の高い根号（unit.radical）は、行間 1 以上で次のように置く。行間 1.0 の上側を
    T1 = max(line_ascent, math_radical_single_ratio × インクの上) とし、行の上側は、行間がちょうど 1 なら T1、1 より大きければ
    max(line_ascent × sp, math_radical_top_ratio × sp × インクの上) にする。下側（次の行への行送りに使う値）は、行間 1.0 の
    行の高さに行間を掛けた値から上側を引いた sp × (T1 + d) − 上側（行間 1.0 では d）にする。最後の行の下側は d。

    ほかの数式の行間 1.0 の上側は A = max(line_ascent, インクの上 + math_stack_gap_em × s) にする。
    A が floor より高く、行間が 1 より大きいときは、H = A + max(line_descent, bottom) を行間 1.0 の全高とし、
    行間を掛けた全高 sp × H の 4 分の 3 を上、残りを下に置く。bottom は文字の行間 1.0 の下側で、インクの下とは別である。
    独立した数式だけの行では bottom を None にして、この規則を当てない。
    それ以外は line_ascent・line_descent に行間を掛ける（行間 1.0 の上側は A）。行間 1 未満の背の高い数式は
    確かめていないので、上側を max(line_ascent, math_stack_top_ratio × インクの上) × sp、下側を line_descent × sp にする。

    PowerPoint 16.0.20430.20182 の校正用の資料で確かめたのは 18〜28 pt、行間 1.0・1.15・1.2・1.3・1.4・1.5・2.0 である。
    行間で増える分の分け方の境は、e ÷ d が 1.134 以下（すべて上）と 1.214 以上（4 分の 3 を上）の間にあり、その近くは
    確かめていない。floor と下に置くものを含む式の規則は、校正用の資料の 20 pt の、下が深い \\underset（行間 1.0・1.15・
    1.5・2.0）、下が浅い \\underset（1.15）、\\lim\\limits_{n \\to \\infty} a_n（1.15・1.2）と、実際の資料の PDF の 24 pt の
    \\underset（1.15）で確かめた。根号の規則は実際の資料の PDF の文中の O(\\sqrt{n\\log n}) の形の式（14・18・20 pt、
    行間 1.2）と、校正用の資料の 18〜24 pt の行間 1.2 と 20 pt の行間 1.0・1.5 の上揃え・中揃え・下揃えで確かめ、根号の行の
    次の行への行送りは 20 pt の行間 1.2 の 1 要素、行間がちょうど 1 の上側の下限は \\overset{n}{x}（20 pt）の 1 要素で
    確かめた。通常の数式で広がる行は、2026-10-23 の校正用の資料の上付きの横分数・斜線・上下の分数を含む式で、
    12・20・32 pt、行間 1.0・1.15・1.2・1.5、欧文・和文・複数行と縦の揃えを変えて確かめた。
    上に偏る入れ子の分数の上側の上限は、12・20・32 pt、行間 1.0〜2.0 の文中・独立した数式と、縦の揃え・複数行で確かめた。
    """
    box, size = unit.box, unit.size
    top, lower = spacing * box.line_ascent, spacing * box.line_descent
    exact = abs(spacing - 1.0) <= _EXACT_SPACING
    if unit.radical and spacing >= 1.0 - _EXACT_SPACING:
        single = max(box.line_ascent, model.math_radical_single_ratio * box.ascent)  # 行間 1.0 の上側
        above = single if exact else max(top, model.math_radical_top_ratio * spacing * box.ascent)
        return above, spacing * (single + box.descent) - above, True
    if not unit.stacked:
        single = max(box.line_ascent, box.ascent + model.math_stack_gap_em * size)
        if exact:
            top = single
        elif spacing > 1.0 and bottom is not None and single > floor:
            whole = spacing * (single + max(bottom, box.line_descent))
            top = _MATH_LINE_ABOVE * whole
            lower = whole - top
        return top, lower, False
    if spacing < 1.0 - _EXACT_SPACING:
        return max(top, model.math_stack_top_ratio * spacing * box.ascent), lower, False
    if unit.under or box.ascent + box.descent >= model.math_stack_full_em * size:
        above_ink = max(box.ascent + model.math_stack_gap_em * size, floor)
        whole = above_ink + box.descent
        extra = max(0.0, spacing - 1.0) * whole
        if extra <= model.math_stack_above_full * box.descent + _EPSILON:
            above = extra
        else:
            above = model.math_stack_above_share * extra
        if not exact:
            above = min(above, _MATH_LINE_ABOVE * spacing * whole - above_ink)
        return max(top, above_ink + above), box.descent + extra - above, True
    return max(top, model.math_stack_top_ratio * spacing * box.ascent), box.descent, True


def _prepare(content, style: _Style, faces: _Faces, display: bool) -> list[_HardLine]:
    """断片を明示改行ごとの単位の列にする。どちらの書体にも字形がない文字は、すべて集めてから MeasureError にする。

    12 pt 以上では、欧文の書体で計測する隣どうしの欧文の文字に kern 表の値を加える。断片（a:r）の境界をまたぐ対にも、
    書体・大きさ・太さと forced（書き出しの書体の区切り）が同じなら加える。PowerPoint は forced の区切りの境界と、
    大きさか太さが違う断片の境界をまたぐ対を詰めない（校正用の資料で確かめた）。数式と明示改行の前後の対には加えない。
    """
    fragments = _fragments(content)
    first = (style.size, style.bold)
    if fragments:
        _kind, _text, own, fragment = fragments[0]
        first = _fragment_font(own, fragment, style)
    lines = [_HardLine([], *first)]
    missing: list[str] = []
    fonts, model = faces.fonts, faces.model
    previous: tuple[str, str, bool, float, bool] | None = None  # 直前の通常文字の（文字、書体、太字、大きさ、forced）
    for position, (kind, text, own, fragment) in enumerate(fragments):
        size, bold = _fragment_font(own, fragment, style)
        if kind == "math":
            box = _math_box(fonts, text, size, bold, style, display)
            # インクの上下の和が文字サイズを超える式だけを、構造で分ける。根号の規則は文中の数式だけに当てる
            # （独立した数式の背の高い根号は、行間を掛けた計測の上下で合った）。
            structure = _structure(text) if box.ascent + box.descent > size + _EPSILON else None
            lines[-1].units.append(
                _Unit(
                    text,
                    True,
                    box.width,
                    0.0,
                    size,
                    bold,
                    False,
                    False,
                    (box.ascent, box.descent),
                    box,
                    stacked=structure in ("stacked", "under"),
                    under=structure == "under",
                    radical=structure == "radical" and not display,
                    origin=(position, 0),
                )
            )
            previous = None
            continue
        kerning = size >= model.kerning_min_size_pt
        # 上付き（書き出しの baseline）：送り幅とインクは縮めた大きさで求め、インクを基準線のずらしの分だけ上げる。
        # 行の配置に使う文字サイズ（_Unit.size）は断片の大きさのままにする。上付きの文字は前後と詰めない。
        raised = bool(own.get("superscript")) if isinstance(own, Mapping) else False
        scale = model.superscript_scale if raised else 1.0
        lift = model.superscript_offset_em * size if raised else 0.0
        cursor = 0  # 断片の文字列の中の位置
        for index, part in enumerate(text.split("\n")):
            if index:
                lines.append(_HardLine([], size, bold))
                previous = None
                cursor += 1  # 明示改行の分
            for column, char in enumerate(part, cursor):
                font = faces.font(char, bold)
                if font is None:
                    if char not in missing:
                        missing.append(char)
                    previous = None
                    continue
                family, forced = font
                east_asian = _east_asian(char)
                kern = 0.0
                if (
                    kerning
                    and not raised
                    and previous is not None
                    and previous[1:] == (family, bold, size, forced)
                    and family == style.latin
                    and not east_asian
                    and not _east_asian(previous[0])
                    and faces.metrics(family, bold).has_kerning
                ):
                    kern = fonts.kerning(previous[0], char, family, bold) * size
                previous = None if raised else (char, family, bold, size, forced)
                space = char in _SPACES
                ink = None
                if not space:
                    y_min, y_max = fonts.glyph_bounds(char, family, bold)
                    if (y_min, y_max) != (0.0, 0.0):
                        ink = (y_max * size * scale + lift, -y_min * size * scale - lift)
                lines[-1].units.append(
                    _Unit(
                        char,
                        False,
                        fonts.advance(char, family, bold) * size * scale,
                        kern,
                        size,
                        bold,
                        space,
                        east_asian,
                        ink,
                        None,
                        family,
                        origin=(position, column),
                    )
                )
            cursor += len(part)
    if missing:
        raise _glyph_missing(missing, style.latin, style.japanese)
    return lines


def _kinsoku_allows(before: _Unit, after: _Unit) -> bool:
    """空白の前でなく、行頭禁則の文字の前・行末禁則の文字の後でもない位置か。"""
    if after.space:
        return False
    return not ((not after.math and after.text in NO_LINE_START) or (not before.math and before.text in NO_LINE_END))


def _can_break(before: _Unit, after: _Unit) -> bool:
    """before と after の間で行を分けることができるか。

    空白（U+0020 と U+3000）の前と、行頭禁則の文字の前・行末禁則の文字の後では分けない。空白の後、数式の前後、
    和文の文字の前後、ハイフン（- と ‐）と英数字の間で分ける。それ以外（欧文の語の途中）は分けない。
    """
    if not _kinsoku_allows(before, after):
        return False
    if before.space:
        return True
    if before.math or after.math:
        return True
    if before.east_asian or after.east_asian:
        return True
    return before.text in _HYPHENS and after.text.isalnum()


class _Widths:
    """単位の列の部分の幅（行末の空白を除き、行の中の kern を加える）。"""

    def __init__(self, units: Sequence[_Unit]):
        self.advance, self.kern, self.last = [0.0], [0.0], [-1]
        for index, unit in enumerate(units):
            self.advance.append(self.advance[-1] + unit.width)
            self.kern.append(self.kern[-1] + unit.kern)
            self.last.append(self.last[-1] if unit.space else index)

    def width(self, start: int, end: int) -> float:
        last = self.last[end]
        if last < start:
            return 0.0
        return self.advance[last + 1] - self.advance[start] + self.kern[last + 1] - self.kern[start + 1]


def _breaks(units: Sequence[_Unit]) -> list[bool]:
    """breaks[k]：units[k] の前で分けることができるか。末尾の位置 len(units) は真。"""
    return [False] + [_can_break(units[k - 1], units[k]) for k in range(1, len(units))] + [True]


def _split(units: Sequence[_Unit], start: int, index: int) -> int:
    """分ける位置のないまま units[index] で内側の幅を超えた行の終わり。

    PowerPoint と同じく、収まる最後の単位の後で分ける（1 行に 1 単位以上）。その位置が空白の前・行頭禁則・行末禁則に
    当たれば、当たらない位置まで前へずらす。前にずらすことができなければ（行の最初の単位だけで幅を超え、その後が空白か
    行頭禁則の文字の場合など）、当たらない位置まで後ろへずらす（空白と行頭禁則の文字は幅を超えて前の行に残る）。
    """
    count = len(units)
    end = max(index, start + 1)
    for candidate in range(end, start, -1):
        if candidate >= count or _kinsoku_allows(units[candidate - 1], units[candidate]):
            return candidate
    for candidate in range(end + 1, count + 1):
        if candidate == count or _kinsoku_allows(units[candidate - 1], units[candidate]):
            return candidate
    return count


def _phrase_limits(units: Sequence[_Unit], breaks: list[bool], widths: _Widths, available: float) -> list[bool]:
    """phrase_wrap の分ける位置：和文の文字どうしの間は、文節の境だけにする。内側の幅より長い文節の中は今までどおり分ける。"""
    candidates = [
        k for k in range(1, len(units)) if breaks[k] and _japanese_letter(units[k - 1]) and _japanese_letter(units[k])
    ]
    if not candidates:
        return breaks
    boundaries = _phrase_boundaries("".join(_OBJECT if unit.math else unit.text for unit in units))
    if boundaries is None:
        return breaks
    cuts = sorted(boundaries | {0, len(units)})
    result = list(breaks)
    for k in candidates:
        if k in boundaries:
            continue
        index = bisect.bisect_right(cuts, k)
        if widths.width(cuts[index - 1], cuts[index]) <= available + _EPSILON:
            result[k] = False
    return result


def _wrap(
    units: Sequence[_Unit], widths: _Widths, available: float | None, model: LayoutModel, phrase: bool = False
) -> list[tuple[int, int, bool, bool]]:
    """明示改行で分けた 1 行を、内側の幅に収まる限り前から詰めて分ける。

    戻り値は（始まり、終わり、誤差の範囲か、語の途中で分けたか）。分ける位置のないまま内側の幅を超えた場合
    （内側より長い語など）は、_split の位置で分ける。phrase が真なら、和文の文字の間は文節の境だけで分ける。
    """
    count = len(units)
    if available is None or count == 0:
        return [(0, count, False, False)]
    breaks = _breaks(units)
    if phrase:
        breaks = _phrase_limits(units, breaks, widths, available)
    following = [count] * (count + 1)  # following[k]：k 以降で breaks が真になる最初の位置
    for k in range(count - 1, 0, -1):
        following[k] = k if breaks[k] else following[k + 1]
    result = []
    start = 0
    while start < count:
        end, last, forced = count, None, False
        for index in range(start, count):
            if index > start and breaks[index]:
                last = index
            if not units[index].space and widths.width(start, index + 1) > available + _EPSILON:
                if last is not None:
                    end = last
                else:
                    end, forced = _split(units, start, index), True
                break
        width = widths.width(start, end)
        # PowerPoint が最後の単位を次の行へ送るか、次の行の最初の単位をこの行に残す可能性がある行。
        # 計算の誤差だけの差は寸法の切り上げ（_ceil）と同じく区別しない。
        near = available - _band(model, width) + _EPSILON < width <= available + _EPSILON
        if end < count:
            joined = widths.width(start, end + 1 if forced else following[end + 1])
            near = near or joined <= available + _band(model, joined) + _EPSILON
        split = forced and end < count and not breaks[end] and not units[end - 1].math and not units[end].math
        result.append((start, end, near, split))
        start = end
    return result


def _unbreakable(lines: Sequence[_HardLine]) -> float:
    widest = 0.0
    for line in lines:
        widths, breaks = _Widths(line.units), _breaks(line.units)
        start = 0
        for k in range(1, len(line.units) + 1):
            if breaks[k]:
                widest = max(widest, widths.width(start, k))
                start = k
    return widest


def _line_text(units: Sequence[_Unit]) -> str:
    return "".join(f"${unit.text}$" if unit.math else unit.text for unit in units)


def _overhang(units: Sequence[_Unit], width: float) -> tuple[float, float]:
    """数式の描画が、行の送りの範囲（0 から width）より左・右へ張り出す量（pt）。

    文字の送りは送り幅で決め、はみ出しの検査だけに描画範囲を使う（入れ子の分数の線は式の送りの始まりより左に出る）。
    通常文字は字形の左右を計測しないので、送りの範囲のままにする。
    """
    x = left = right = 0.0
    for index, unit in enumerate(units):
        if index:
            x += unit.kern
        box = unit.box
        if unit.math and box is not None and unit.ink is not None:
            if box.ink_left is not None:
                left = max(left, -(x + box.ink_left))
            if box.ink_right is not None:
                right = max(right, x + box.ink_right - width)
        x += unit.width
    return left, right


_phrase_parser = None
_phrase_lock = threading.Lock()
_OBJECT = "￼"  # 文節の判定で数式の代わりに置く文字


def _phrase_boundaries(text: str) -> frozenset[int] | None:
    """和文の文節の境の位置（文字の番号。0 と末尾を除く）。BudouX の日本語のモデルで判定する。

    BudouX を読み込むことができなければ None にする（折り返しの種類を other にし、診断を出さない）。
    """
    global _phrase_parser
    with _phrase_lock:
        if _phrase_parser is None:
            try:
                import budoux

                _phrase_parser = budoux.load_default_japanese_parser()
            except (ImportError, OSError, ValueError):
                _phrase_parser = False
        parser = _phrase_parser
    if parser is False:
        return None
    positions, offset = set(), 0
    for chunk in parser.parse(text)[:-1]:
        offset += len(chunk)
        positions.add(offset)
    return frozenset(positions)


def _japanese_letter(unit: _Unit) -> bool:
    """文節の判定の対象にする和文の文字（漢字・仮名・々・ー など。約物と記号を除く）。"""
    return not unit.math and unit.east_asian and unicodedata.category(unit.text).startswith("L")


def _wrap_kinds(units: Sequence[_Unit], wrapped: Sequence[tuple[int, int, bool, bool]]) -> list[str | None]:
    """明示改行で分けた 1 行の、各行の末尾の折り返しの種類（WRAP_KINDS。最後の行は None）。"""
    kinds: list[str | None] = []
    boundaries: frozenset[int] | None | bool = False  # 未計算
    for index, (_start, end, _near, split) in enumerate(wrapped):
        if index == len(wrapped) - 1:
            kinds.append(None)
            continue
        before, after = units[end - 1], units[end]
        if split:
            kind = "split"
        elif before.space:
            kind = "space"
        elif before.math or after.math:
            kind = "math"
        elif before.text in _HYPHENS and after.text.isalnum():
            kind = "hyphen" if end >= 2 and units[end - 2].text.isalnum() and not units[end - 2].east_asian else "other"
        elif _japanese_letter(before) and _japanese_letter(after):
            if boundaries is False:
                boundaries = _phrase_boundaries("".join(_OBJECT if unit.math else unit.text for unit in units))
            kind = "other" if boundaries is None else "phrase" if end in boundaries else "mid_phrase"
        elif before.east_asian or after.east_asian:
            kind = "punctuation"
        else:
            kind = "other"
        kinds.append(kind)
    return kinds


def _line_top(ascent: float, spacing: float, model: LayoutModel) -> float:
    """行の上側（em）。ascent は行の書体の組の上側の割合、spacing は行間。

    PowerPoint は行間がちょうど 1 のときだけ書体の値で行を置き、1 より大きいときは行間に比例させ、書体の値を上限で抑える
    （1.0 から 1.02 で値が飛ぶ）。1 未満は、上側から 1.2 ×（1 − 行間）を除いた値と、1 より大きいときと同じ式の値の
    大きいほうにする（1 から下では値が飛ばない）。確かめたのは行間 0.9 だけで、書体の値が上限より大きい組（Arial・Calibri・
    Yu Gothic UI・BIZ UDGothic の行）は前者に、上限より小さい組（Meiryo と Arial の和文の行）は後者に合った。
    """
    if abs(spacing - 1.0) <= _EXACT_SPACING:
        return ascent
    proportional = spacing * min(ascent, model.line_top_limit_em)
    if spacing > 1.0:
        return proportional
    return max(ascent - model.line_pitch_em * (1.0 - spacing), proportional)


def _arrange(
    lines: Sequence[_HardLine], available: float | None, style: _Style, faces: _Faces, model: LayoutModel, display: bool
) -> tuple[tuple[TextLine, ...], float, float]:
    """折り返して行を置く。戻り値は（行、内容の幅、文字ブロックの高さ）。

    PowerPoint の配置のモデルは次のとおり（s は文字サイズ、sp は行間。line_spacing_pt は 1.2 × s で割って行間にする）。

    - 行の書体の組は欧文の書体と、行の通常文字を計測した書体。和文の文字（W・F・H）を含む行には和文の書体も加える。
      空の行と U+0020 だけの行は欧文の書体だけの組にする（PowerPoint で、和文の書体が Meiryo の要素の末尾の空の行と
      空白だけの行が、欧文の書体だけの位置に置かれた）。組の上側と下の割合 A・D を _Faces.natural で求める。
    - 行の上側は _line_top(A, sp) × s。行送りは前の行の下側（1.2 × sp × s − 上側）と次の行の上側の和
      （同じ書体の組の行どうしは 1.2 × sp × s）。文字ブロックの下端は最後の基線から D × s（sp によらない）。
    - 数式は _math_line の上側と下側で行の上下を広げる。通常の数式が文字の上側を超える段落では、行間 1.0 の行全体の
      高さに行間を掛け、4 分の 3 を上に置く。独立した数式だけの行と低い式は line_ascent・line_descent に行間を掛ける。
      分数などを含む背の高い式と文中の背の高い根号を行間 1 以上で置いた行は、最後の行の文字ブロックの下端を、
      独立した数式だけの行では _math_line の下側（d と行間で増えた分の残り）、文字の段落の行では d にする。ほかの数式で
      最後の行の数式が下側より深い分は、文字ブロックの下端も下げる。
    - 行間が 1 より大きい文字の段落の行で、規則に当たらない数式か文中の背の高い根号のインクの下 d が、行の書体の組の下の
      割合 D × s より深いときは、行の上側（書体の組の値）に math_descent_above × sp × (d − D × s) を加え、最後の行の下側を
      d にする（実際の資料の PDF の、和文と添字・根号の行の行間 1.2 と、校正用の資料の \\sqrt{y_j} の行で確かめた。
      行間 1.2 以外は確かめていない）。根号の行の上側は根号の規則の上側との大きいほうになる。次の行への行送りは今までどおり。
    - 独立した数式だけの行（display）は、最後の基線から行の下側までを文字ブロックにする。文中の数式だけの行
      （書き出しが U+200B の区切りを加える段落）は、欧文の書体の文字の行として扱う。
    - 行のインクの上下は字形の範囲のままにする（基線より上だけにある字形 × “ などの行は、インクの下が負になる）。
    """
    pieces = []
    for number, line in enumerate(lines):
        widths = _Widths(line.units)
        wrapped = _wrap(line.units, widths, available, model, style.phrase_wrap)
        kinds = _wrap_kinds(line.units, wrapped)
        for index, (start, end, near, split) in enumerate(wrapped):
            last = index == len(wrapped) - 1
            pieces.append(
                _Piece(
                    line.units[start:end],
                    widths.width(start, end),
                    not last,
                    last and number < len(lines) - 1,
                    near,
                    line.size,
                    line.bold,
                    split,
                    kinds[index],
                )
            )
    pitch = model.line_pitch_em
    rows = []
    for piece in pieces:
        units = piece.units
        size = max((unit.size for unit in units), default=piece.size)
        spacing = style.line_spacing_pt / (pitch * size) if style.line_spacing_pt is not None else style.line_spacing
        text_units = [unit for unit in units if not unit.math]
        if text_units:
            sources = dict.fromkeys((unit.size, unit.bold) for unit in text_units)
        else:
            sources = {(size, piece.bold if not units else style.bold): None}
        families = {style.latin, *(unit.family for unit in text_units)}
        if any(unit.east_asian for unit in text_units):
            families.add(style.japanese)
        families = frozenset(families)
        tops, lowers, bottoms, floors = [], [], [], []
        for scale, bold in sources:
            ascent_em, descent_em = faces.natural(families, bold)
            top = _line_top(ascent_em, spacing, model) * scale
            tops.append(top)
            lowers.append(pitch * spacing * scale - top)
            bottoms.append(descent_em * scale)
            floors.append(_line_top(ascent_em, 1.0, model) * scale)
        text_bottom = None if display and not text_units else max(bottoms)
        placed = [(unit, _math_line(unit, spacing, model, max(floors), text_bottom)) for unit in units if unit.math]
        maths = [line for _unit, line in placed]
        math_only = display and bool(maths) and not text_units
        base_top, base_descent, bottom = max(tops), max(lowers), max(bottoms)
        # 規則に当たらない数式と文中の背の高い根号のインクの下のうち、書体の組の下（D × s）より深い分。
        deep = max(
            [0.0] + [unit.box.descent - bottom for unit, (_top, _lower, rule) in placed if not rule or unit.radical]
        )
        spaced = spacing > 1.0 + _EXACT_SPACING and not math_only
        if spaced:
            base_top += model.math_descent_above * spacing * deep
        ascent = max([base_top] + [top for top, _lower, _rule in maths])
        descent = max([base_descent] + [lower for _top, lower, _rule in maths])
        if math_only:
            last = None  # 最後の基線から行の下側までを文字ブロックにする
        else:
            # 最後の行の下側：行間が 1 より大きい行は、書体の組の下と規則に当たらない数式の d の深いほう。
            # それ以外は、規則に当たらない数式の行間を掛けた下側が行の下側より深い分を加える。
            other = max([base_descent] + [lower for _top, lower, rule in maths if not rule])
            last = bottom + (deep if spaced else max(0.0, other - base_descent))
            # 規則で置いた数式（背の高い数式と文中の背の高い根号）は d を文字ブロックの下端にする（下揃えの行で確かめた）。
            last = max([last] + [unit.box.descent for unit, (_top, _lower, rule) in placed if rule])
        rows.append((ascent, descent, last))
    baselines = [rows[0][0]]
    for previous, row in zip(rows, rows[1:]):
        baselines.append(baselines[-1] + previous[1] + row[0])
    _ascent, descent, last = rows[-1]
    height = baselines[-1] + (descent if last is None else last)
    justify = style.align == "justify" and available is not None
    result = []
    for piece, baseline in zip(pieces, baselines):
        inks = [unit.ink for unit in piece.units if unit.ink is not None]
        width = max(piece.width, available) if justify and piece.soft else piece.width
        overhang_left, overhang_right = _overhang(piece.units, piece.width)
        result.append(
            TextLine(
                _line_text(piece.units),
                width,
                baseline,
                max((ink[0] for ink in inks), default=0.0),
                max((ink[1] for ink in inks), default=0.0),
                piece.hard_break,
                piece.near,
                piece.split,
                overhang_left,
                overhang_right,
                piece.wrap,
                sum(1 for unit in piece.units if not unit.space),
                piece.units[0].origin if piece.units else None,
            )
        )
    return tuple(result), max(line.width for line in result), height


def _insets(shape: str, width: float, height: float, style: _Style) -> Insets:
    area = text_area(Rect(0, 0, width, height), shape, style.corner_radius)
    padding = style.padding
    return Insets(
        area.x + padding.left,
        area.y + padding.top,
        width - area.right + padding.right,
        height - area.bottom + padding.bottom,
    )


def _available(width: float, insets: Insets, style: _Style) -> float | None:
    return max(0.0, width - insets.left - insets.right) if style.word_wrap else None


_caches: "weakref.WeakKeyDictionary[object, OrderedDict[str, TextLayout]]" = weakref.WeakKeyDictionary()
_cache_lock = threading.Lock()


def _cache_key(content, style: Mapping, width, height, shape, display, model) -> str | None:
    used = (
        "size_pt",
        "bold",
        "font_latin",
        "font_japanese",
        "font_math",
        "line_spacing",
        "line_spacing_pt",
        "padding_pt",
        "word_wrap",
        "phrase_wrap",
        "align",
        "corner_radius_pt",
    )
    try:
        return json.dumps(
            [content, {key: style.get(key) for key in used}, width, height, shape, display, astuple(model)],
            ensure_ascii=False,
            sort_keys=True,
            allow_nan=False,
        )
    except (TypeError, ValueError):
        return None


def _text_layout(content, style: Mapping, fonts: FontSource, *, width, height, shape, display, model) -> TextLayout:
    """layout_text と layout_element の共通部分。height が None なら必要な高さを求め、数値なら枠を固定する。"""
    shape = _shape(shape)
    parsed = _text_style(style)
    if height is None:
        width = None if width is None else _positive(width, "幅")
    else:
        width, height = _nonnegative(width, "幅"), _nonnegative(height, "高さ")
    key = _cache_key(content, style, width, height, shape, display, model)
    entries = None
    if key is not None:
        try:
            with _cache_lock:
                entries = _caches.setdefault(fonts, OrderedDict())
                if key in entries:
                    entries.move_to_end(key)
                    return entries[key]
        except TypeError:  # 弱参照を作成できない FontSource の結果は保持しない
            entries = None
    faces = _Faces(fonts, parsed, model)
    lines = _prepare(content, parsed, faces, display)
    unbreakable = _unbreakable(lines)
    natural = max((_Widths(line.units).width(0, len(line.units)) for line in lines), default=0.0)
    if width is None:
        arranged, content_width, content_height = _arrange(lines, None, parsed, faces, model, display)
        # 折り返す要素は、計測の誤差で PowerPoint が折り返さないよう、誤差の帯の分だけ広げる。
        needed = content_width + (_band(model, content_width) if parsed.word_wrap else 0.0)
        frame_width, frame_height = frame_size(
            shape, needed, content_height, parsed.padding, corner_radius_pt=parsed.corner_radius
        )
        frame_width, frame_height = _ceil(frame_width), _ceil(frame_height)
        # 切り上げで角丸の文字領域の差が広がり、内側が足りなくなる場合は 0.1 pt ずつ広げる。
        for _ in range(_PASSES):
            insets = _insets(shape, frame_width, frame_height, parsed)
            short_width = needed - (frame_width - insets.left - insets.right)
            short_height = content_height - (frame_height - insets.top - insets.bottom)
            if short_width <= _EPSILON and short_height <= _EPSILON:
                break
            if short_width > _EPSILON:
                frame_width = _ceil(frame_width + short_width)
            if short_height > _EPSILON:
                frame_height = _ceil(frame_height + short_height)
        result = TextLayout(
            arranged,
            content_width,
            content_height,
            None,
            unbreakable,
            insets,
            frame_width,
            frame_height,
            parsed.word_wrap,
            natural,
        )
    else:
        # 角丸の文字領域は高さにも依存するので、折り返しが変わらなくなるまで高さを求め直す。
        depends = parsed.word_wrap and shape == "rounded_rectangle" and height is None
        frame_height = 0.0 if height is None else height
        for _ in range(_PASSES):
            insets = _insets(shape, width, frame_height, parsed)
            available = _available(width, insets, parsed)
            arranged, content_width, content_height = _arrange(lines, available, parsed, faces, model, display)
            needed = _ceil(
                frame_size(
                    shape,
                    content_width,
                    content_height,
                    parsed.padding,
                    width=width,
                    corner_radius_pt=parsed.corner_radius,
                )[1]
            )
            if not depends or needed <= frame_height + _EPSILON:
                break
            frame_height = needed
        if height is None:
            frame_height = max(frame_height, needed)
            insets = _insets(shape, width, frame_height, parsed)
            needed = frame_height
        result = TextLayout(
            arranged,
            content_width,
            content_height,
            available,
            unbreakable,
            insets,
            width,
            needed,
            parsed.word_wrap,
            natural,
        )
    if entries is not None:
        with _cache_lock:
            entries[key] = result
            while len(entries) > _CACHE_SIZE:
                entries.popitem(last=False)
    return result


def layout_text(
    content: Sequence[Mapping],
    style: Mapping,
    fonts: FontSource,
    *,
    width: float | None = None,
    shape: str = "rectangle",
    display: bool = False,
    model: LayoutModel = MODEL,
) -> TextLayout:
    """場面の断片の列（kind が text・emphasis・math の辞書）を、解決済みの書式で計測する。

    `style` には size_pt・font_latin・font_japanese と、line_spacing か line_spacing_pt が要る。
    `width` は枠の幅（pt）。None なら折り返さずに、内容から枠の幅を決める（折り返す要素は誤差の帯の分を加える）。
    `style["word_wrap"]` が偽なら幅があっても折り返さない。角丸の大きさは corner_radius_pt、余白は padding_pt を使う。
    枠の高さは内容から求め、0.1 pt 単位に切り上げる。`display` は数式を独立した数式として計測するか
    （要素では `math.office_math.math_mode()` で決める）。同じ入力の結果は FontSource ごとに 4,096 件まで保持する。
    """
    return _text_layout(content, style, fonts, width=width, height=None, shape=shape, display=display, model=model)


def _display(element: Mapping) -> bool:
    content = element.get("content")
    if not isinstance(content, (list, tuple)):
        return False
    if not any(isinstance(fragment, Mapping) and fragment.get("kind") == "math" for fragment in content):
        return False
    from .math.office_math import math_mode

    return math_mode(element) == "display"


def _frame(area) -> Rect:
    try:
        return as_rect(area)
    except (TypeError, ValueError) as exc:
        raise MeasureError(f"要素の領域が不正です: {area}", area=repr(area), reason=str(exc)) from exc


def _visible(style: Mapping) -> bool:
    def painted(value) -> bool:
        if value is None or value is False:
            return False
        return str(value).strip().lower() not in {"", "none", "transparent"}

    width = style.get("line_width_pt", 1)
    return painted(style.get("fill")) or (painted(style.get("line_color")) and (not _number(width) or width > 0))


def layout_element(element: Mapping, fonts: FontSource, *, model: LayoutModel = MODEL) -> ElementLayout:
    """場面の要素（text・box・equation）を枠に置き、valign・align を反映した絶対座標の結果を返す。"""
    element_id = element.get("id")
    try:
        kind = element.get("kind")
        if kind not in _TEXT_KINDS:
            raise MeasureError(f"文字を持たない要素は計測できません: {kind}")
        frame = _frame(element.get("area"))
        style = element.get("style") or {}
        shape = element.get("shape", "rectangle") if kind == "box" else "rectangle"
        text = _text_layout(
            element.get("content", []),
            style,
            fonts,
            width=frame.w,
            height=frame.h,
            shape=shape,
            display=_display(element),
            model=model,
        )
        parsed = _text_style(style)
        area = text_area(frame, shape, parsed.corner_radius)
        padding = parsed.padding
        inner = Rect(
            area.x + padding.left,
            area.y + padding.top,
            max(0.0, area.w - padding.left - padding.right),
            max(0.0, area.h - padding.top - padding.bottom),
        )
        offset = {"top": 0.0, "middle": (inner.h - text.content_height) / 2, "bottom": inner.h - text.content_height}[
            parsed.valign
        ]
        block_top = inner.y + offset
        lefts = []
        for line in text.lines:
            if parsed.align == "center":
                lefts.append(inner.x + (inner.w - line.width) / 2)
            elif parsed.align == "right":
                lefts.append(inner.x + inner.w - line.width)
            else:
                lefts.append(inner.x)
        inked = [(left, line) for left, line in zip(lefts, text.lines) if line.ink_top > 0 or line.ink_bottom > 0]
        ink = None
        if inked:
            # 左右は送りの範囲に、数式の描画の張り出し（入れ子の分数の線など）を加える。
            x0 = min(left - line.overhang_left for left, line in inked)
            x1 = max(left + line.width + line.overhang_right for left, line in inked)
            y0 = min(block_top + line.baseline - line.ink_top for _left, line in inked)
            y1 = max(block_top + line.baseline + line.ink_bottom for _left, line in inked)
            ink = Rect(x0, y0, max(0.0, x1 - x0), max(0.0, y1 - y0))
        return ElementLayout(
            "" if element_id is None else str(element_id),
            frame,
            area,
            inner,
            text,
            block_top,
            tuple(lefts),
            ink,
            _visible(style),
        )
    except MeasureError as error:
        if error.element_id is None and element_id is not None:
            error.element_id = str(element_id)
        raise


def phrase_wrapped_content(element: Mapping, fonts: FontSource, *, model: LayoutModel = MODEL) -> list[dict]:
    """phrase_wrap の要素の断片の列に、計測した自動の折り返しの位置で明示改行（\n）を入れた写しを返す。

    PowerPoint は和文の文字の間ならどこでも折り返すので、文節の境で折り返した行を明示改行で固定する。
    折り返しのない要素はそのままの写しを返す。数式の前で折り返す行は、直前の文字の断片の末尾か、改行だけの断片で分ける。
    """
    content = [dict(fragment) for fragment in element.get("content") or []]
    lines = layout_element(element, fonts, model=model).text.lines
    breaks = [
        line.origin
        for previous, line in zip(lines, lines[1:])
        if previous.wrap is not None and not previous.hard_break and line.origin is not None
    ]
    for index, column in sorted(breaks, reverse=True):
        fragment = content[index]
        if fragment.get("kind") != "math":
            text = fragment.get("text", "")
            content[index] = {**fragment, "text": text[:column] + "\n" + text[column:]}
        elif index > 0 and content[index - 1].get("kind") != "math":
            content[index - 1] = {**content[index - 1], "text": content[index - 1].get("text", "") + "\n"}
        else:
            content.insert(index, {"kind": "text", "text": "\n"})
    return content


def _has_text(content) -> bool:
    if not isinstance(content, (list, tuple)):
        return bool(content)
    return any(not isinstance(fragment, Mapping) or str(fragment.get("text", "")).strip() for fragment in content)


def layout_scene(
    scene: Mapping, fonts: FontSource, *, model: LayoutModel = MODEL
) -> dict[tuple[str, str], ElementLayout | MeasureError]:
    """全ページの、文字を持つ要素（text・box・equation で通常文字か数式がある要素）を計測する。

    キーは（ページの識別子、要素の識別子）。計測できない要素は、ページと要素の識別子を付けた MeasureError を値にする。
    """
    result: dict[tuple[str, str], ElementLayout | MeasureError] = {}
    for slide in scene.get("slides", []) or []:
        slide_id = slide.get("id")
        for element in slide.get("elements", []) or []:
            if element.get("kind") not in _TEXT_KINDS or not _has_text(element.get("content")):
                continue
            element_id = element.get("id")
            try:
                value = layout_element(element, fonts, model=model)
            except MeasureError as error:
                value = error
            except (ValueError, RuntimeError, OSError, ImportError) as exc:
                value = MeasureError(f"文字または数式の寸法を計算できません: {exc}", reason=str(exc))
                value.__cause__ = exc
            if isinstance(value, MeasureError):
                value.slide_id = value.slide_id if value.slide_id is not None else slide_id
                value.element_id = value.element_id if value.element_id is not None else element_id
            result[(slide_id, element_id)] = value
    return result


# ---- 判定と記録 ----


def fit_findings(layout: ElementLayout, *, model: LayoutModel = MODEL) -> list[dict]:
    """はみ出しなどの判定結果を {"code", "severity", "message", "observations"} の列で返す。判定はインクで行う。

    - text_overflow（エラー）：インクが文字領域から許容幅を超えてはみ出す（塗りか線のある枠 0.5 pt、ない枠 1.5 pt）。
      折り返す要素で、途中に分ける位置のない部分（長い語・数式）が内側の幅より長い場合も、インクによらず報告する
      （observations の unbreakable_width_pt。PowerPoint は長い語を語の途中で改行し、計測も同じ位置で分ける。
      その行を split_lines に示す。数式は途中で改行しないので、内側より広い数式ははみ出す）。
    - text_padding_overlap（警告）：はみ出しではないが、余白が 0 より大きい辺で、インクが余白に
      max(1 pt, 余白の 25 %) を超えて入り込む。
    - text_wrap_near_limit（参考）：折り返しの判定が計測の誤差の範囲にある行がある。
    - text_short_line（参考）：自動で折り返した段落の最後の行が 2 文字以下か、内側の幅の short_line_ratio 未満である。
    - text_wrap_mid_phrase（参考）：和文の文節の途中で自動で折り返した行がある（BudouX の日本語のモデルで判定する）。
    - text_wrap_hyphen（参考）：欧文の語のハイフンの後で自動で折り返した行がある。

    text_overflow と text_padding_overlap の observations には、折り返しで増えた行数と、高さの不足と、
    明示改行だけで分けたときの最も広い行の幅（single_line_width_pt。1 行に収めるのに要る内側の幅）を加える。
    """
    text, area, inner, ink = layout.text, layout.text_area, layout.inner, layout.ink
    common = {
        "required_width_pt": _round(text.content_width),
        "available_width_pt": _round(inner.w),
        "lines": len(text.lines),
        "line_widths_pt": [_round(line.width) for line in text.lines],
        "word_wrap": text.word_wrap,
    }
    hard_lines = 1 + sum(1 for line in text.lines[:-1] if line.hard_break)
    shortfall = text.content_height - inner.h
    sizing = {"hard_lines": hard_lines}
    if len(text.lines) > hard_lines:
        sizing["single_line_width_pt"] = _round(text.natural_width)
    if shortfall > _EPSILON:
        sizing["height_shortfall_pt"] = _round(shortfall)
    located = {
        **common,
        "required_height_pt": _round(text.content_height),
        "available_height_pt": _round(inner.h),
        "ink_pt": _bounds_list(ink),
        "text_area_pt": _rect_list(area),
        "inner_pt": _rect_list(inner),
    }
    findings = []
    unbreakable = text.word_wrap and text.unbreakable_width > inner.w + _EPSILON
    overflow = {}
    if ink is not None:
        tolerance = model.visible_tolerance_pt if layout.visible_frame else model.hidden_tolerance_pt
        overflow = {
            "left": area.x - ink.x,
            "top": area.y - ink.y,
            "right": ink.right - area.right,
            "bottom": ink.bottom - area.bottom,
        }
        side = max(overflow, key=overflow.get)
        if overflow[side] <= tolerance + _EPSILON:
            overflow = {}
    if overflow or unbreakable:
        messages = []
        observations = {**located, **sizing}
        if overflow:
            messages.append("文章が文字領域からはみ出している")
            observations["overflow_pt"] = {"side": side, "amount": _round(overflow[side])}
        if unbreakable:
            observations["unbreakable_width_pt"] = _round(text.unbreakable_width)
            split = [number for number, line in enumerate(text.lines, 1) if line.split]
            if split:
                observations["split_lines"] = split
                messages.append("内側の幅より長い語があり、PowerPoint は語の途中で改行する")
            else:
                messages.append("途中で分けられない部分（数式など）が内側の幅より長い")
        observations["suggestion"] = (
            "領域を広げるか文章を短くする。高さは Rect(x, y, w, None) で内容から決められる"
            if not unbreakable
            else "領域を広げるか、長い語や数式を短くするか分ける。高さは Rect(x, y, w, None) で内容から決められる"
        )
        findings.append(
            {"code": "text_overflow", "severity": "error", "message": "。".join(messages), "observations": observations}
        )
    elif ink is not None:
        frame = layout.frame
        padding = {
            "left": text.insets.left - (area.x - frame.x),
            "top": text.insets.top - (area.y - frame.y),
            "right": text.insets.right - (frame.right - area.right),
            "bottom": text.insets.bottom - (frame.bottom - area.bottom),
        }
        intrusion = {
            "left": inner.x - ink.x,
            "top": inner.y - ink.y,
            "right": ink.right - inner.right,
            "bottom": ink.bottom - inner.bottom,
        }
        excess = {}
        for name in _SIDES:
            if padding[name] <= _EPSILON:
                continue
            amount = min(padding[name], max(0.0, intrusion[name]))
            if amount > max(model.padding_overlap_min_pt, model.padding_overlap_ratio * padding[name]) + _EPSILON:
                excess[name] = amount
        if excess:
            side = max(excess, key=excess.get)
            findings.append(
                {
                    "code": "text_padding_overlap",
                    "severity": "warning",
                    "message": "文章が余白に入り込んでいる",
                    "observations": {
                        **located,
                        **sizing,
                        "overflow_pt": {"side": side, "amount": _round(excess[side])},
                        "suggestion": "余白を小さくするか領域を広げる",
                    },
                }
            )
    near = [number for number, line in enumerate(text.lines, 1) if line.near_limit]
    if near:
        # 次の行の単位を残す可能性がある行も含めるため、内側の幅での帯を示す。
        band = _band(model, max([inner.w] + [text.lines[number - 1].width for number in near]))
        findings.append(
            {
                "code": "text_wrap_near_limit",
                "severity": "info",
                "message": "折り返しの位置が計測の誤差の範囲にある",
                "observations": {
                    **common,
                    "near_limit_lines": near,
                    "band_pt": _round(band),
                    "suggestion": "幅を band_pt 以上広げるか、改行を明示する",
                },
            }
        )
    findings.extend(_wrap_findings(text, inner.w, common, model))
    return findings


def _break_context(line: TextLine, following: TextLine, size: int = 6) -> str:
    """折り返しの前後の文字（「…届いたた / め、…」の形の表示用）。"""
    before, after = line.text.rstrip(), following.text.lstrip()
    head = before if len(before) <= size else "…" + before[-size:]
    tail = after if len(after) <= size else after[:size] + "…"
    return f"{head} / {tail}"


def _wrap_findings(text: TextLayout, available: float, common: dict, model: LayoutModel) -> list[dict]:
    """自動の折り返しの位置の参考（短い最後の行、和文の文節の途中、欧文の語のハイフンの後）。"""
    if not text.word_wrap or available <= 0:
        return []
    lines = text.lines
    findings = []
    short, short_text = [], []
    start = 0
    for index, line in enumerate(lines):
        if not (line.hard_break or index == len(lines) - 1):
            continue
        last = lines[index]
        if (
            index > start
            and last.characters
            and (last.characters <= model.short_line_characters or last.width < model.short_line_ratio * available)
        ):
            short.append(index + 1)
            short_text.append(last.text.strip())
        start = index + 1
    if short:
        findings.append(
            {
                "code": "text_short_line",
                "severity": "info",
                "message": "自動で折り返した段落の最後の行が短い",
                "observations": {
                    **common,
                    "short_lines": short,
                    "short_texts": short_text,
                    "ratio": model.short_line_ratio,
                    "suggestion": "改行を明示するか、幅か文言を変えて行の長さをそろえる",
                },
            }
        )
    for kind, finding in (
        (
            "mid_phrase",
            {"code": "text_wrap_mid_phrase", "severity": "info", "message": "和文の文節の途中で折り返している"},
        ),
        (
            "hyphen",
            {"code": "text_wrap_hyphen", "severity": "info", "message": "欧文の語のハイフンの後で折り返している"},
        ),
    ):
        numbers = [number for number, line in enumerate(lines[:-1], 1) if line.wrap == kind]
        if numbers:
            suggestion = (
                "文節の境に改行を明示するか、幅か文言を変える"
                if kind == "mid_phrase"
                else "語の前で改行を明示するか、幅を広げる"
            )
            observations = {
                **common,
                "wrapped_lines": numbers,
                "breaks": [_break_context(lines[number - 1], lines[number]) for number in numbers],
                "suggestion": suggestion,
            }
            findings.append({**finding, "observations": observations})
    return findings


def layout_record(layout: ElementLayout) -> dict:
    """layout.json の 1 要素分。座標は絶対の pt で、行の基線は絶対の y、行のインクは基線からの距離
    （上は上向き、下は下向きが正。基線より上だけにある字形の行では、下が負になる）。"""
    text = layout.text
    insets = text.insets
    return {
        "id": layout.element_id,
        "word_wrap": text.word_wrap,
        "frame_pt": _rect_list(layout.frame),
        "text_area_pt": _rect_list(layout.text_area),
        "inner_pt": _rect_list(layout.inner),
        "block_top_pt": _round(layout.block_top),
        "content_size_pt": [_round(text.content_width), _round(text.content_height)],
        "required_size_pt": [
            _round(text.content_width + insets.left + insets.right),
            _round(text.content_height + insets.top + insets.bottom),
        ],
        "ink_pt": _bounds_list(layout.ink),
        "unbreakable_width_pt": _round(text.unbreakable_width),
        "lines": [
            {
                "text": line.text,
                "width_pt": _round(line.width),
                "baseline_pt": _round(layout.block_top + line.baseline),
                "left_pt": _round(left),
                "ink_top_pt": _round(line.ink_top),
                "ink_bottom_pt": _round(line.ink_bottom),
                "hard_break": line.hard_break,
                "near_limit": line.near_limit,
                "wrap": line.wrap,
            }
            for line, left in zip(text.lines, layout.line_left)
        ],
    }
