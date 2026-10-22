"""診断の表。識別子ごとに重大度・区分・判定・直し方と、確認の証拠に入れる範囲を持つ。

`docs/diagnostics.md` の表はこの表から作る（`python -m slidekit.diagnostics --write docs/diagnostics.md`）。
診断を出す処理を加えるときは、識別子をこの表に登録してから文書を作り直す。
このモジュールは標準ライブラリだけに依存する（要約・検査・確認記録から軽く読み込むため）。
"""

from __future__ import annotations

import argparse
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
import re
import sys
from types import MappingProxyType

# 文書の区分。診断の scope（検査の段階）とは別の分類で、並びは文書の節の順である。
CATEGORIES: dict[str, str] = {
    "input": "入力",
    "layout": "配置",
    "writing": "記述規則",
    "pptx": "PPTX の検査",
    "pdf": "PDF の検査",
    "build": "生成の失敗",
    "check": "整合検査・最終検査",
}
SEVERITIES: dict[str, str] = {"error": "エラー", "warning": "警告", "info": "参考"}
EVIDENCE_ELEMENTS = frozenset({"all", "content"})
EVIDENCE_PAGE_VALUES = frozenset({"size", "notes", "content_area"})
START_MARKER = "<!-- slidekit:diagnostics:start -->"
END_MARKER = "<!-- slidekit:diagnostics:end -->"
_CODE = re.compile(r"[a-z][a-z0-9_]*")


@dataclass(frozen=True)
class EvidenceScope:
    """警告の証拠に入れるもの。既定は 0.1 と同じ広い範囲（記述規則の全体・フォント・要素の全体・寸法とノート）である。

    許容できるのは警告だけなので、範囲を狭めるのは警告の種類だけにする。
    """

    rules: tuple[str, ...] | None = None  # 記述規則の最上位の項目。None は全体
    roles: bool = False  # 要素の役割の定義（テーマ）を入れる
    fonts: bool = True  # 計測に使ったフォントの一覧を入れる
    element: str = "all"  # "all" は要素の全体、"content" は kind・content・role・semantic だけ
    page: tuple[str, ...] = ("size", "notes")  # 入れるページの値
    pair: bool = False  # 二つの要素の組の診断か（宣言の照合に使う）
    version: int = 1  # 判定の方法を変えたら上げる

    def __post_init__(self) -> None:
        if self.rules is not None and (
            not isinstance(self.rules, tuple) or not all(isinstance(name, str) and name for name in self.rules)
        ):
            raise ValueError(f"証拠の rules は None か項目名のタプルにする: {self.rules!r}")
        if self.element not in EVIDENCE_ELEMENTS:
            raise ValueError(f"証拠の element は all か content にする: {self.element!r}")
        if not isinstance(self.page, tuple) or not set(self.page) <= EVIDENCE_PAGE_VALUES:
            raise ValueError(f"証拠の page は size・notes・content_area のタプルにする: {self.page!r}")
        for name in ("roles", "fonts", "pair"):
            if not isinstance(getattr(self, name), bool):
                raise ValueError(f"証拠の {name} は真偽値にする")
        if isinstance(self.version, bool) or not isinstance(self.version, int) or self.version < 1:
            raise ValueError(f"証拠の version は 1 以上の整数にする: {self.version!r}")


@dataclass(frozen=True)
class DiagnosticSpec:
    """診断の種類ごとの定義。summary は判定（である調の句）、fix は直し方である。"""

    severity: str
    category: str
    summary: str
    fix: str
    evidence: EvidenceScope = EvidenceScope()

    def __post_init__(self) -> None:
        if self.severity not in SEVERITIES:
            raise ValueError(f"診断の重大度は error・warning・info のどれかにする: {self.severity!r}")
        if self.category not in CATEGORIES:
            raise ValueError(f"診断の区分が不正である: {self.category!r}")
        for name in ("summary", "fix"):
            value = getattr(self, name)
            if not isinstance(value, str) or not value.strip() or "\n" in value:
                raise ValueError(f"診断の {name} は 1 行の文字列にする")
        if not isinstance(self.evidence, EvidenceScope):
            raise ValueError("診断の evidence は EvidenceScope にする")


def _registry(entries: Iterable[tuple[str, DiagnosticSpec]]) -> Mapping[str, DiagnosticSpec]:
    table: dict[str, DiagnosticSpec] = {}
    for code, spec in entries:
        if not _CODE.fullmatch(code):
            raise ValueError(f"診断の識別子は英小文字・数字・_ にする: {code!r}")
        if code in table:
            raise ValueError(f"診断の識別子が重複している: {code}")
        table[code] = spec
    return MappingProxyType(table)


# 警告の証拠の範囲（範囲を決めていない種類は EvidenceScope() の広い範囲になる）。
# PDF の警告のページの値は、validate_pdf に渡すページの表示の証拠を使うので、ここでは入れない。
_PDF_WARNING = EvidenceScope(rules=(), fonts=False, element="all", page=())

DIAGNOSTICS: Mapping[str, DiagnosticSpec] = _registry(
    [
        # 入力
        (
            "invalid_page_size",
            DiagnosticSpec("error", "input", "ページ寸法が正の有限値でない", "`deck.toml` の `size` を直す"),
        ),
        (
            "invalid_slide_id",
            DiagnosticSpec(
                "error",
                "input",
                "ページの識別子が文字列でないか、空であるか、`/`・制御文字・予約した接頭辞 `_theme` を含むか、ほかのページと重複している",
                "ページの `id` を、規則に合う一意の名前にする",
            ),
        ),
        (
            "invalid_element_id",
            DiagnosticSpec(
                "error",
                "input",
                "要素の識別子（`key_id` のキーと表の行キー・列キーを含む）が文字列でないか、空であるか、`/`・制御文字・予約した接頭辞 `_theme` を含むか、同じページか同じ表で重複している。表の行キー `_header` は見出し行のために予約している",
                "要素の `id` とキーを、規則に合う一意の名前にする",
            ),
        ),
        ("unknown_element", DiagnosticSpec("error", "input", "要素の種類が未対応である", "公開 API で要素を作る")),
        (
            "parent_missing",
            DiagnosticSpec("error", "input", "親の `group` を参照できない", "親の `group` を同じページに作る"),
        ),
        ("parent_cycle", DiagnosticSpec("error", "input", "親子関係が循環している", "`group` の入れ子を直す")),
        (
            "invalid_geometry",
            DiagnosticSpec(
                "error",
                "input",
                "座標・寸法・線の端点・回転・反転の値が不正である（有限でない、負、幅か高さが 0、線の両端が同じ、型の誤り）。または、寸法が内容から決まらない（縦横比の分からない画像、内容のない文字、線・まとまり・題名の領域、内容から決める寸法を使う `Rect` の操作）か、`stack` と表の寸法（幅・高さ・間隔・`indent`、列幅・行の高さとその数）が不正である",
                "`area`・`start`・`end` を直す。内容から決まらない寸法は数値で指定する。`stack` と表は、メッセージの寸法を直す",
            ),
        ),
        (
            "connection_missing",
            DiagnosticSpec(
                "error",
                "input",
                "線の接続先の要素を参照できない（`anchor()` の要素が線と別のページにある）",
                "`anchor()` の元の要素を同じページに置く",
            ),
        ),
        (
            "invalid_anchor",
            DiagnosticSpec(
                "error",
                "input",
                "`anchor()` の辺・間隔・位置が不正である（`center` に `gap`・`at` を指定した場合を含む）か、領域のないまとまりに接続している",
                "`side`・`gap`・`at` を直すか、接続先のまとまりに領域を与える",
            ),
        ),
        (
            "math_transform",
            DiagnosticSpec(
                "error",
                "input",
                "数式を含む要素か、その親が回転・反転している。PPTX の検査でも同じ識別子を使う",
                "数式を含む要素と親の回転・反転を外す",
            ),
        ),
        (
            "unknown_argument",
            DiagnosticSpec(
                "error",
                "input",
                "要素を作る関数の引数か書式の辞書（`style=` など）に、未知の名前がある",
                "メッセージの候補の名前に直す",
            ),
        ),
        (
            "unsupported_argument",
            DiagnosticSpec(
                "error",
                "input",
                "その種類の要素が受け付けない引数か書式がある（`z_index`、線と `stack` の要素の `area`、`equation` の `text`、まとまりの回転・反転など）",
                "引数を外し、メッセージの代わりの方法を使う",
            ),
        ),
        (
            "invalid_style",
            DiagnosticSpec(
                "error",
                "input",
                '書式の値が不正である（型、範囲、余白の辺、`"none"` の誤用など）',
                "メッセージの書式名と値を直す",
            ),
        ),
        (
            "undefined_color",
            DiagnosticSpec(
                "error",
                "input",
                "書式の色の名前がテーマの `colors` にない",
                "テーマの `colors` に色を定義するか、6 桁の 16 進数で指定する",
            ),
        ),
        (
            "invalid_fragment",
            DiagnosticSpec(
                "error",
                "input",
                "文章の断片が不正である（文字列でない値、空の数式、タブなどの制御文字、断片の書式の型）",
                "断片の文字と書式を直す。改行は `\\n` だけを使う",
            ),
        ),
        (
            "invalid_link",
            DiagnosticSpec(
                "error",
                "input",
                "断片のリンクが `https://`・`http://`・`mailto:` で始まらないか、空白を含む",
                "リンク先の URL を直す",
            ),
        ),
        (
            "invalid_markup",
            DiagnosticSpec(
                "error",
                "input",
                "`markup.parse()` に渡した記法か引数が不正である（`$`・`**` が閉じていない、空の数式・強調、数式の中の改行、文字列・真偽値でない引数）",
                "メッセージの位置の記号を閉じるか、`\\$`・`\\*` と書く。引数は文字列と真偽値で渡す",
            ),
        ),
        (
            "invalid_table_key",
            DiagnosticSpec(
                "error",
                "input",
                "表の役割・書式の指定か `cell()` に、存在しない行キー・列キーがある（`row_styles` の `_header` を含む）",
                "メッセージの候補のキーに直す。見出し行の書式は `header_style` で指定する",
            ),
        ),
        ("asset_missing", DiagnosticSpec("error", "input", "画像ファイルを参照できない", "`path` を直す")),
        (
            "image_fit",
            DiagnosticSpec("error", "input", "画像の `fit` が `contain`・`cover` 以外である", "`fit` を直す"),
        ),
        (
            "invalid_image_option",
            DiagnosticSpec(
                "error",
                "input",
                "画像の `raster_scale`・`vector` の値が不正であるか、PDF・SVG 以外の素材に指定しているか、"
                "PDF の素材に `vector=True` を指定している",
                "値を直すか、PDF・SVG の素材だけに指定する。PDF の図を劣化させずに置くなら SVG に変換して使う",
            ),
        ),
        (
            "unknown_role",
            DiagnosticSpec(
                "error",
                "input",
                "文字の役割（`allow_roles` で宣言した役割を含む）がテーマにない",
                "テーマに役割を定義するか、既存の役割を使う",
            ),
        ),
        (
            "role_usage",
            DiagnosticSpec(
                "error",
                "input",
                "使う場所を限った役割を宣言のない場所で使っている。または、数式専用の役割に通常の文字が、文字専用の役割に数式がある",
                "`allow_roles` で宣言するか、内容に合う役割を使う",
            ),
        ),
        (
            "minimum_override",
            DiagnosticSpec(
                "error", "input", "要素で役割の最小文字サイズを下げようとしている", "最小値の低い役割をテーマに定義する"
            ),
        ),
        (
            "font_glyph_missing",
            DiagnosticSpec(
                "error",
                "input",
                "和文と欧文のどちらのフォントにも、必要な字形がない",
                "字形のあるフォントを指定するか、文字を変える",
            ),
        ),
        (
            "text_measurement_failed",
            DiagnosticSpec(
                "error", "input", "文字か数式の寸法を計算できない", "観測値の理由を見て、数式やフォントの指定を直す"
            ),
        ),
        (
            "math_syntax_error",
            DiagnosticSpec(
                "error",
                "input",
                "数式の記法が間違っている（未対応の命令、閉じていない括弧など）",
                "観測値の位置と助言を見て数式を直す。対応する命令は数式の一覧（`docs/math.md`）にある",
            ),
        ),
        (
            "source_missing",
            DiagnosticSpec(
                "error",
                "input",
                "出典の `document_id`・`bib_key` を参照できない",
                "`inputs.documents`・`inputs.bibliography` に加えるか、識別子を直す",
            ),
        ),
        (
            "invalid_reference",
            DiagnosticSpec(
                "error",
                "input",
                "文献ファイルの項目か、文献一覧・出典の文献キーが不正である（解析できない項目、重複したキー・項目、未対応の LaTeX 命令、エスケープしていない `%`、未定義のキーなど）",
                "メッセージの文献ファイルの項目か文献キーを直す",
            ),
        ),
        (
            "invalid_selection",
            DiagnosticSpec(
                "error",
                "input",
                "`layout`・`build` の `--slides` に、範囲外のページ番号か、ないページの識別子がある（書式の誤りは処理を始める前に引数の誤りにする）",
                "メッセージの候補を見て `--slides` の指定を直す",
            ),
        ),
        (
            "input_error",
            DiagnosticSpec(
                "error",
                "input",
                "資料の入力が不正である（ほかの識別子に当たらない誤り）",
                "メッセージと観測値の `location` を見て、編集元を直す",
            ),
        ),
        (
            "allow_warning_unused",
            DiagnosticSpec("info", "input", "許容の宣言に対応する警告がない", "宣言を消すか、要素の識別子を直す"),
        ),
        # 配置
        (
            "page_overflow",
            DiagnosticSpec(
                "error", "layout", "要素（線は線幅と矢印の先を含む）がページの外へはみ出している", "`area` を直す"
            ),
        ),
        (
            "content_area_overflow",
            DiagnosticSpec(
                "warning",
                "layout",
                "要素が本文領域（`slide.content`）からはみ出している。まとまりと、テーマが置く要素は除く",
                "`slide.content` の中に置く。意図した配置なら許容する",
                EvidenceScope(rules=(), fonts=False, element="all", page=("size", "content_area")),
            ),
        ),
        (
            "layout_bounds_overflow",
            DiagnosticSpec(
                "error",
                "layout",
                "高さを数値で指定した `stack` で、積み上げた要素か `stack.take` で受け取った領域が、積み上げる領域を超えている",
                "領域を広げるか、要素や間隔を減らす",
            ),
        ),
        (
            "font_too_small",
            DiagnosticSpec(
                "error",
                "layout",
                "文字サイズ（断片の大きさを含む）が役割の最小値より小さい",
                "文字サイズを上げるか、最小値の低い役割を使う",
            ),
        ),
        (
            "text_overflow",
            DiagnosticSpec(
                "error",
                "layout",
                "文字のインクが、図形の文字領域（余白を含む）から許容幅（線か塗りのある枠は 0.5 pt、ない枠は 1.5 pt）を超えてはみ出している。または、折り返す要素で、途中に分ける位置のない部分（長い語・数式）が余白を除いた幅より長い（PowerPoint は長い語を語の途中で改行し、数式は改行しない。インクがはみ出さなくても報告する）",
                "領域を広げるか、文章を短くする。長い語や数式は短くするか分ける。高さを内容から決めるには `Rect(x, y, w, None)` を使う",
            ),
        ),
        (
            "text_padding_overlap",
            DiagnosticSpec(
                "warning",
                "layout",
                "文字のインクが、余白のある辺で余白に入り込んでいる（入り込みが 1 pt とその辺の余白の 25 % の大きい方を超える）。`text_overflow` に当たる要素には出さない",
                "余白を小さくするか、領域を広げる。意図した配置なら許容する",
                EvidenceScope(rules=(), fonts=True, element="all", page=("size",)),
            ),
        ),
        (
            "text_wrap_near_limit",
            DiagnosticSpec(
                "info",
                "layout",
                "折り返しの位置が幅の限界に近く、PowerPoint の行数が計測と変わるおそれがある",
                "幅を広げるか、改行を明示する。行数は生成後の PDF の検査で確かめる",
            ),
        ),
        (
            "text_occluded",
            DiagnosticSpec(
                "error",
                "layout",
                "後に描く不透明な要素（塗りのある箱か画像）が、文章のある要素の領域全体を覆っている",
                "要素の作成順か位置を直す",
            ),
        ),
        (
            "element_overlap",
            DiagnosticSpec(
                "warning",
                "layout",
                "要素どうしが重なっている。箱の内側の要素（先に描いた箱の文章と重ならないものと、塗りも文章もない箱の中のもの）、まとまりとその子孫、線どうし、接続した線と接続先、端が辺に接する線、背景と飾りは除く",
                "位置を直す。意図した重なりなら `anchor()` で接続するか、`slide.allow_warning()` で宣言するか、理由を付けて許容する",
                EvidenceScope(rules=(), fonts=False, element="all", page=("size",), pair=True),
            ),
        ),
        # 記述規則
        (
            "notes_missing",
            DiagnosticSpec("error", "writing", "`require_notes` が真で、発表者ノートが空である", "ノートを書く"),
        ),
        (
            "notes_style",
            DiagnosticSpec(
                "warning",
                "writing",
                '`notes_style = "polite"` で、です・ます調でない文末がある（ひらがなで終わる文。括弧書き・引用・名詞止めは除く）',
                "文末を直す。意図した文なら `slide.allow_warning()` で宣言するか、許容する",
                EvidenceScope(rules=("notes_style",), fonts=False, page=("notes",)),
            ),
        ),
        (
            "forbidden_character",
            DiagnosticSpec("error", "writing", "`forbidden_characters` の文字が本文にある", "その文字を除く"),
        ),
        (
            "math_not_marked",
            DiagnosticSpec(
                "warning",
                "writing",
                "`math_candidates` の候補（英字と添字、数字、ギリシャ文字、数式の命令、指定した名前）が通常の文字にある。URL・DOI・arXiv の番号とリンクの中は除く",
                "`Math` で書く。数式でなければ `exclude_patterns`・`exclude_roles`・`exclude_page_kinds` で外すか、許容する",
                EvidenceScope(rules=("math_candidates",), fonts=False, element="content", page=()),
            ),
        ),
        (
            "line_break",
            DiagnosticSpec(
                "warning",
                "writing",
                "明示改行の直後の行頭に置かない文字（閉じ括弧、句読点、小書きの仮名、長音、後置の単位記号など）か、直前の行末に置かない文字（開き括弧、前置の通貨記号）がある。数式は 1 文字として扱い、中点で始まる行と自動折り返しは対象外",
                "改行の位置を直す",
                EvidenceScope(rules=(), fonts=False, element="content", page=()),
            ),
        ),
        # PPTX の検査
        (
            "package_duplicate",
            DiagnosticSpec(
                "error", "pptx", "PPTX 内に同じ名前の部品が重複している", "再生成する。続く場合は雛形を確かめる"
            ),
        ),
        (
            "invalid_xml",
            DiagnosticSpec("error", "pptx", "PPTX 内の XML を解釈できない", "再生成する。続く場合は雛形を確かめる"),
        ),
        (
            "relationship_missing",
            DiagnosticSpec(
                "error", "pptx", "図形などが、存在しない関係を参照している", "再生成する。続く場合は雛形を確かめる"
            ),
        ),
        (
            "package_reference_missing",
            DiagnosticSpec("error", "pptx", "関係の参照先の部品がない", "再生成する。続く場合は雛形を確かめる"),
        ),
        (
            "metadata_mismatch",
            DiagnosticSpec(
                "error", "pptx", "文書情報が `deck.toml` の `metadata` と一致しない", "`metadata` を確かめて再生成する"
            ),
        ),
        ("presentation_missing", DiagnosticSpec("error", "pptx", "PPTX に資料本体がない", "再生成する")),
        (
            "page_size_mismatch",
            DiagnosticSpec("error", "pptx", "PPTX のページ寸法が指定と違う", "雛形と `size` を確かめる"),
        ),
        ("slide_count_mismatch", DiagnosticSpec("error", "pptx", "PPTX のページ数が違う", "再生成する")),
        ("slide_missing", DiagnosticSpec("error", "pptx", "PPTX のページを読み込むことができない", "再生成する")),
        ("slide_order_mismatch", DiagnosticSpec("error", "pptx", "PPTX のページの識別子か順序が違う", "再生成する")),
        (
            "element_id_duplicate",
            DiagnosticSpec("error", "pptx", "PPTX 内で要素の名前が重複している", "要素の識別子を確かめて再生成する"),
        ),
        ("element_missing", DiagnosticSpec("error", "pptx", "PPTX 内に要素がない", "再生成する")),
        (
            "geometry_mismatch",
            DiagnosticSpec("error", "pptx", "PPTX の図形の位置か寸法が、指定と 0.1 pt を超えて違う", "再生成する"),
        ),
        (
            "text_mismatch",
            DiagnosticSpec("error", "pptx", "PPTX の通常の文字か明示改行（`a:br`）の数が編集元と違う", "再生成する"),
        ),
        (
            "automatic_text_adjustment",
            DiagnosticSpec(
                "error",
                "pptx",
                "文字枠に `a:noAutofit` がなく、PowerPoint が文字を自動で縮小・調整しうる",
                "再生成する",
            ),
        ),
        (
            "text_wrap_mismatch",
            DiagnosticSpec(
                "error", "pptx", "文字枠の折り返しの設定（`wrap`）が編集元の `word_wrap` と一致しない", "再生成する"
            ),
        ),
        (
            "font_mismatch",
            DiagnosticSpec(
                "error",
                "pptx",
                "PPTX の通常文字の区切り（数と各区切りの文字）か、区切りごとの書式（太字・下線・色）・フォントが書き出しの規則と違う。計測が書体を決めた区切りは、欧文・和文の両方がその書体であることを求める",
                "フォントの指定を確かめて再生成する",
            ),
        ),
        (
            "font_size_mismatch",
            DiagnosticSpec(
                "error", "pptx", "PPTX の文字サイズ（明示改行の大きさを含む）が、要素か断片の指定と違う", "再生成する"
            ),
        ),
        (
            "editable_math_missing",
            DiagnosticSpec("error", "pptx", "編集可能な数式（`m:oMath`）の数が、編集元の数式の数と違う", "再生成する"),
        ),
        (
            "math_content_mismatch",
            DiagnosticSpec(
                "error",
                "pptx",
                "数式の構造か書式（断片の大きさ・色・太字を含む）が編集元と違うか、期待する数式を作成できない",
                "再生成する",
            ),
        ),
        (
            "math_layout_mismatch",
            DiagnosticSpec(
                "error",
                "pptx",
                "数式の置き方（独立・文中と、その段落の構成）か揃え（`m:jc` と中央揃えの補正）が書き出しの規則と一致しない",
                "再生成する",
            ),
        ),
        (
            "math_grouped",
            DiagnosticSpec(
                "error", "pptx", "数式が PowerPoint のグループの中にあり、互換画像を作成できない", "再生成する"
            ),
        ),
        (
            "math_fallback_missing",
            DiagnosticSpec("error", "pptx", "必要な数式の互換画像がない", "`math.fallback` を確かめて再生成する"),
        ),
        (
            "group_mismatch",
            DiagnosticSpec(
                "error",
                "pptx",
                "書き出すまとまり（`pptx_group` が真）に対応する PowerPoint のグループがないか、子孫の図形が最も近い書き出すまとまりのグループの中にない。または、グループの入れ子が編集元と違うか、座標変換が恒等でないか、書き出さないまとまりか場面にない名前のグループがある",
                "再生成する",
            ),
        ),
        (
            "connection_mismatch",
            DiagnosticSpec(
                "error",
                "pptx",
                "線の端の PPTX の接続（接続先と接続点）が、`anchor()` から決めた接続と一致しない。接続しない端に接続がある場合を含む",
                "再生成する",
            ),
        ),
        (
            "line_style_mismatch",
            DiagnosticSpec(
                "error",
                "pptx",
                "線・矢印と、枠線のある箱・文字・数式の要素で、破線か実線かが編集元の `dashed` と一致しない",
                "再生成する",
            ),
        ),
        (
            "image_vector_missing",
            DiagnosticSpec("error", "pptx", "SVG でも埋め込む画像に、SVG の部品か参照がない", "再生成する"),
        ),
        (
            "fallback_crop_clipped",
            DiagnosticSpec(
                "warning",
                "pptx",
                "数式の描画が、互換画像に切り出す範囲の上限（数式の図形の矩形から 16 pt まで。ページの内側）を超えている。互換表示では超えた分が欠ける",
                "文字枠を広げる。互換表示で欠けていなければ許容する",
                EvidenceScope(rules=(), fonts=False, element="all", page=()),
            ),
        ),
        ("notes_mismatch", DiagnosticSpec("error", "pptx", "PPTX の発表者ノートが編集元と違う", "再生成する")),
        ("pptx_unreadable", DiagnosticSpec("error", "pptx", "PPTX を検査できない", "観測値の理由を確かめて再生成する")),
        # PDF の検査
        ("pdf_page_count", DiagnosticSpec("error", "pdf", "PDF のページ数が違う", "変換をやり直す")),
        ("pdf_page_size", DiagnosticSpec("error", "pdf", "PDF のページ寸法が違う", "雛形と `size` を確かめる")),
        (
            "pdf_text_missing",
            DiagnosticSpec(
                "warning",
                "pdf",
                "PDF から抽出した文字に欠落の候補があるか、数式が描かれていない要素がある（文字と数式が混じる要素で数式だけが描かれない場合を含む）",
                "該当ページの画像で表示を確かめる。欠けていなければ許容する",
                _PDF_WARNING,
            ),
        ),
        ("pdf_text_overflow", DiagnosticSpec("error", "pdf", "表示した文字がページの外にある", "配置を直す")),
        (
            "pdf_font_substitution",
            DiagnosticSpec(
                "warning",
                "pdf",
                "PDF に指定外のフォントがある",
                "フォントの導入と指定を確かめる。意図した置き換えなら許容する",
                _PDF_WARNING,
            ),
        ),
        (
            "pdf_element_text_overflow",
            DiagnosticSpec(
                "error",
                "pdf",
                "表示した文字と数式の線のインクが、図形の文字領域（余白を含む）から許容幅（線か塗りのある枠は 0.5 pt、ない枠は 1.5 pt）を超えてはみ出している",
                "領域を広げるか、文章を短くする",
            ),
        ),
        (
            "pdf_line_count_mismatch",
            DiagnosticSpec(
                "warning",
                "pdf",
                "PDF の行数（文字か数式のある行）が計測と異なる",
                "枠の幅を広げるか、改行を明示する。表示を確かめて問題がなければ許容する",
                _PDF_WARNING,
            ),
        ),
        (
            "pdf_text_unassigned",
            DiagnosticSpec(
                "info",
                "pdf",
                "PDF の字形の一部が、どの要素にも割り当てられていない",
                "雛形の文字なら対応は要らない。要素の文字なら表示を確かめる",
            ),
        ),
        ("pdf_unreadable", DiagnosticSpec("error", "pdf", "PDF を検査できない", "変換をやり直す")),
        # 生成の失敗
        (
            "build_failed",
            DiagnosticSpec(
                "error",
                "build",
                "生成中に例外が起きた（生成コードの誤りなど）",
                "メッセージと観測値の `location` を見て、編集元を直す",
            ),
        ),
        (
            "build_interrupted",
            DiagnosticSpec(
                "error",
                "build",
                "利用者が生成を中断した",
                "必要なら再実行する。途中の実行ディレクトリは `clean --failed` で消すことができる",
            ),
        ),
        (
            "settings_invalid",
            DiagnosticSpec(
                "error",
                "build",
                "資料の設定（`deck.toml`）・テーマ・記述規則・実行環境の設定名か値が不正である",
                "表示されたファイルと位置の設定を直す",
            ),
        ),
        (
            "final_form_mismatch",
            DiagnosticSpec(
                "error",
                "build",
                "最終 PPTX と、標準表示の確認に変換した最終形の PPTX で、スライドの XML か関係が一致しない",
                "再生成する。続く場合は雛形と互換画像の処理を確かめる",
            ),
        ),
        (
            "renderer_failed",
            DiagnosticSpec("error", "build", "表示ソフトでの変換に失敗した", "詳細を見て、表示ソフトの状態を確かめる"),
        ),
        (
            "renderer_busy",
            DiagnosticSpec(
                "error",
                "build",
                "別の変換が実行中で、待ち時間（`--wait`・実行環境の設定の `wait_seconds`）の間に終わらなかった",
                "観測値の `holder` か `renderer status` で使用中の処理を確かめ、終わってから再実行するか、待ち時間を延ばす",
            ),
        ),
        (
            "renderer_pending",
            DiagnosticSpec(
                "error",
                "build",
                "前回の変換の残存記録 `pending.json` がある",
                "前回の処理の終了を確かめてから再実行する。解除は `renderer clear` で行う",
            ),
        ),
        (
            "renderer_timeout",
            DiagnosticSpec(
                "error",
                "build",
                "変換か事前確認が制限時間を超えたか、作業プロセスが準備を完了しなかった",
                "`renderer status` で状態を確かめる。続く場合は `timeout` を延ばすか、表示ソフトの状態を確かめる",
            ),
        ),
        (
            "renderer_interrupted",
            DiagnosticSpec(
                "error",
                "build",
                "利用者が表示ソフトでの変換を中断した",
                "`renderer status` で状態を確かめてから再実行する",
            ),
        ),
        (
            "renderer_protocol",
            DiagnosticSpec(
                "error",
                "build",
                "変換の作業プロセスの出力（`renderer clear` の結果を含む）が不正であるか、作業プロセスが途中で終了して使用できない",
                "`renderer status` で状態を確かめてから再実行する。続く場合は表示ソフトの状態を確かめる",
            ),
        ),
        (
            "renderer_missing",
            DiagnosticSpec(
                "error", "build", "表示ソフトの実行ファイルがない", "実行環境の設定の `executable`・`powershell` を直す"
            ),
        ),
        ("renderer_unavailable", DiagnosticSpec("error", "build", "表示ソフトを実行できない", "実行環境を確かめる")),
        (
            "invalid_renderer",
            DiagnosticSpec(
                "error",
                "build",
                "表示ソフトの設定が不正である（表形式でない、制限時間・待ち時間の値が不正）",
                "実行環境の設定の `renderers` を直す",
            ),
        ),
        (
            "unsupported_renderer",
            DiagnosticSpec("error", "build", "未対応の表示ソフトである", "`powerpoint` か `libreoffice` を指定する"),
        ),
        (
            "font_unreadable",
            DiagnosticSpec("error", "build", "フォントを読み取ることができない", "実行環境の設定の `fonts` を直す"),
        ),
        (
            "font_missing",
            DiagnosticSpec(
                "error", "build", "フォントのファイル指定がない", "実行環境の設定の `fonts` に `path` を書く"
            ),
        ),
        (
            "font_family_mismatch",
            DiagnosticSpec("error", "build", "書体名とフォントの内容が一致しない", "`path`・`index` を直す"),
        ),
        (
            "windows_font_mismatch",
            DiagnosticSpec(
                "error",
                "build",
                "寸法計算のフォントと Windows に登録された書体を照合できない",
                "同じフォントを Windows に導入する",
            ),
        ),
        ("invalid_pdf", DiagnosticSpec("error", "build", "表示ソフトの出力が PDF でない", "変換をやり直す")),
        ("input_missing", DiagnosticSpec("error", "build", "変換する PPTX がない", "再生成する")),
        ("output_exists", DiagnosticSpec("error", "build", "PDF の出力先が既にある", "新しい実行で再生成する")),
        # 整合検査・最終検査
        ("artifact_unrecorded", DiagnosticSpec("error", "check", "必須の成果物が生成記録にない", "再生成する")),
        (
            "artifact_changed",
            DiagnosticSpec(
                "error", "check", "成果物が生成時と違う", "実行ディレクトリのファイルを変更しない。再生成する"
            ),
        ),
        ("input_changed", DiagnosticSpec("error", "check", "入力ファイルが生成時と違う", "再生成する")),
        ("build_incomplete", DiagnosticSpec("error", "check", "生成が完了していない", "再生成する")),
        (
            "build_not_exportable",
            DiagnosticSpec(
                "error",
                "check",
                "下書き・部分生成の実行は、確認記録と完成物の対象にできない",
                "全ページを通常の生成で作り直す",
            ),
        ),
        (
            "legacy_run",
            DiagnosticSpec("error", "check", "古い形式（`schema_version` 1）の実行である", "再生成してから確認する"),
        ),
        (
            "record_error",
            DiagnosticSpec(
                "error",
                "check",
                "`check`・`review`・`export`・`clean` の操作を進めることができない。操作の指定が間違っている（実行・ページ・記録の識別子、確認者、確認の種類など）、生成記録・確認記録などを読むことができないか形式が不正である、実行が前提を満たさない（`runs/` の外へ複製した。生成の未完了と成果物の不一致は `build_incomplete`・`artifact_changed`・`artifact_unrecorded` で示す）、完成物の保存先が不正である、反映の途中で状態が変わったか中断から自動で復旧できない、または資料単位のロックを待ち時間の間に取得できなかった（別の処理が実行中）",
                "メッセージに合わせて直す。指定の誤りは指定を、完成物の保存先は `deck.toml` の `deliverables` を直す。記録が破損しているか前提を満たさない実行は再生成する。ロックを取得できなかったときは、別の処理が終わってから再実行する",
            ),
        ),
        (
            "review_missing",
            DiagnosticSpec("error", "check", "必要な確認が記録されていない", "確認してから `review record` で記録する"),
        ),
        (
            "warning_unaccepted",
            DiagnosticSpec(
                "error",
                "check",
                "警告が直されず、許容も宣言もされていない",
                "直して再生成するか、`review accept` か `slide.allow_warning()` で許容する",
            ),
        ),
    ]
)

WARNING_CODES: frozenset[str] = frozenset(code for code, spec in DIAGNOSTICS.items() if spec.severity == "warning")


def _cell(value: str) -> str:
    return value.replace("|", "\\|")


def markdown() -> str:
    """`docs/diagnostics.md` の印の間に置く部分（区分ごとの見出しと表）を返す。"""
    sections = []
    for category, title in CATEGORIES.items():
        rows = [
            f"| `{code}` | {SEVERITIES[spec.severity]} | {_cell(spec.summary)} | {_cell(spec.fix)} |"
            for code, spec in DIAGNOSTICS.items()
            if spec.category == category
        ]
        if rows:
            sections.append(
                "\n".join([f"## {title}", "", "| 診断 | 重大度 | 判定 | 直し方 |", "|---|---|---|---|", *rows])
            )
    return "\n\n".join(sections) + "\n"


def _split(text: str) -> tuple[str, str, str]:
    """文書を開始の印の前・印の間・終了の印の後に分ける。"""
    start = text.find(START_MARKER)
    end = text.find(END_MARKER, start + len(START_MARKER)) if start >= 0 else -1
    if start < 0 or end < 0 or text.count(START_MARKER) != 1 or text.count(END_MARKER) != 1:
        raise ValueError(f"生成部分の印 {START_MARKER} と {END_MARKER} が 1 組だけある文書にしてください")
    return text[:start], text[start + len(START_MARKER) : end], text[end + len(END_MARKER) :]


def main(argv: Sequence[str] | None = None) -> int:
    """`python -m slidekit.diagnostics [--write PATH | --check PATH]`。引数がなければ生成部分を標準出力へ出す。"""
    parser = argparse.ArgumentParser(
        prog="python -m slidekit.diagnostics", description="診断の一覧の Markdown を出力・更新・検査する"
    )
    target = parser.add_mutually_exclusive_group()
    target.add_argument(
        "--write", type=Path, metavar="PATH", help="PATH の印の間を、診断の表から作った一覧に置き換える"
    )
    target.add_argument(
        "--check", type=Path, metavar="PATH", help="PATH の印の間が、診断の表から作った一覧と一致するかを確かめる"
    )
    args = parser.parse_args(argv)
    generated = markdown()
    path = args.write or args.check
    if path is None:
        sys.stdout.write(generated)
        return 0
    try:
        text = path.read_text(encoding="utf-8")
        before, current, after = _split(text)
    except (OSError, UnicodeError, ValueError) as error:
        print(f"{path}: {error}", file=sys.stderr)
        return 2
    if args.check is not None:
        if current.strip() == generated.strip():
            return 0
        print(
            f"{path} の診断の一覧が表と一致しません。python -m slidekit.diagnostics --write {path} で更新してください。",
            file=sys.stderr,
        )
        return 1
    updated = f"{before}{START_MARKER}\n\n{generated.strip()}\n\n{END_MARKER}{after}"
    if updated != text:
        path.write_text(updated, encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
