"""テーマ・記述規則・`deck.toml`・実行環境（`environment.toml`）の設定名と値の検査と、役割の使い方の判定。

問題は `Problem` の列で返し、読み込む側がまとめて 1 つの `SettingsError` にする。
`check_settings` は 1 つのファイルの設定名と値を確かめ、`check_theme_references`・`check_rules_references` は
合成した後の参照（色、役割、継承）を、`check_deck_requirements` は合成した後の `deck.toml` の必須の設定と設定どうしの関係を確かめる。
合成した後の検査は形の間違った値を飛ばし、同じ誤りを二重に報告しない。ファイルの有無と参照範囲は読み込む側（`config`）が確かめる。
書式の名前と値の規則は `styles` の一覧を使い、二重に持たない。依存は標準ライブラリと `styles`・`errors` だけにする。
"""

from __future__ import annotations

from collections.abc import Callable, Collection, Iterable, Mapping, Sequence
import dataclasses
from dataclasses import dataclass
import json
import math
import re
from types import MappingProxyType
from typing import Any, Literal

from . import styles
from .errors import InputError
from .styles import suggest

PAGE_KINDS = ("cover", "section", "content", "references", "blank")

THEME_KEYS = ("schema_version", "fonts", "colors", "margins", "defaults", "roles", "page_kinds")
FONT_KEYS = ("japanese", "latin", "math")
MARGIN_KEYS = ("left", "right", "top", "bottom")
PAGE_KIND_KEYS = (
    "title_role",
    "default_role",
    "page_number",
    "page_number_format",
    "background",
    "title_area",
    "content_area",
    "title_area_pt",
    "content_area_pt",
    "footer_area_pt",
    "title_style",
    "footer_style",
    "decorations",
)
RULES_KEYS = ("schema_version", "forbidden_characters", "require_notes", "notes_style", "math_candidates")
NOTES_STYLES = ("unrestricted", "polite")
MATH_CANDIDATE_KEYS = (
    "enabled",
    "numbers",
    "variables",
    "operations",
    "exclude_patterns",
    "exclude_roles",
    "exclude_page_kinds",
)
ROLE_CONTENTS = ("math", "text")
DECORATION_KINDS = ("line", "arrow", "box")
_LINE_DECORATION_KEYS = frozenset({"kind", "id", "start", "end", "start_pt", "end_pt"}) | styles.STROKE_KEYS
_DECORATION_KEYS = {
    "box": frozenset({"kind", "id", "area", "area_pt", "shape", "corner_radius_pt"}) | styles.FRAME_KEYS,
    "line": _LINE_DECORATION_KEYS,
    "arrow": _LINE_DECORATION_KEYS,
}
_ROLE_NAME = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")

# deck.toml の設定名。表の位置（最上位は ""）から、その表で受け付ける名前へ。inputs.documents は配列の各表の名前。
DECK_KEYS = MappingProxyType(
    {
        "": (
            "schema_version",
            "extends",
            "entrypoint",
            "size",
            "theme",
            "rules",
            "output_dir",
            "metadata",
            "inputs",
            "render",
            "math",
            "images",
            "deliverables",
            "template",
            "environment_override",
        ),
        "metadata": (
            "title",
            "subject",
            "author",
            "keywords",
            "comments",
            "last_modified_by",
            "category",
            "content_status",
            "identifier",
            "language",
            "version",
        ),
        "inputs": ("read_roots", "python", "bibliography", "asset_dirs", "documents"),
        "inputs.documents": ("id", "path", "affects_content"),
        "render": ("primary", "required"),
        "math": ("fallback", "fallback_renderer", "fallback_scale", "functions"),
        "images": ("raster_scale", "vector"),
        "deliverables": ("pptx", "pdf", "notes", "record"),
        "template": ("path", "layout_index", "show_master_shapes"),
    }
)
# extends で引き継がない設定。実行保存先と完成物を、継承した資料どうしで共有しないためである。
DECK_LOCAL_KEYS = ("output_dir", "deliverables")
# 必ず明示する文書情報。
REQUIRED_METADATA = ("title", "author", "last_modified_by")
PAGE_SIZES = ("16:9", "4:3")
RENDERERS = ("powerpoint", "libreoffice")
FALLBACK_MODES = ("auto", "required", "off")

# 実行環境（environment.toml）の設定名。fonts.* は書体ごとの表（キーは書体名）。
ENVIRONMENT_KEYS = MappingProxyType(
    {
        "": ("schema_version", "jobs", "font_directories", "fonts", "renderers"),
        "fonts.*": ("path", "index", "bold_path", "bold_index"),
        "renderers": RENDERERS,
        "renderers.powerpoint": ("powershell", "timeout", "wait_seconds"),
        "renderers.libreoffice": ("executable", "timeout"),
    }
)

_URL = re.compile(r"^[a-zA-Z][a-zA-Z0-9+.-]*://")
_BARE_KEY = re.compile(r"[A-Za-z0-9_-]+")


@dataclass(frozen=True)
class Problem:
    """設定の問題 1 件。文字列にすると「ファイル: 位置 問題（候補: …）」の 1 行になる。"""

    file: str  # 表示用のファイル名
    key: str  # 設定の位置。例：math_candidates.exclude_patterns
    message: str  # 「は不明な設定です」のように、位置に続く文
    suggestions: tuple[str, ...] = ()

    def __str__(self) -> str:
        text = f"{self.key} {self.message}" if self.key else self.message
        return f"{self.file}: {text}{styles.format_candidates(self.suggestions)}"


class SettingsError(InputError):
    """設定の問題をまとめて報告する例外。1 件を 1 行で書き、details["problems"] に各件の辞書を入れる。"""

    code = "settings_invalid"

    def __init__(self, problems: Sequence[Problem]) -> None:
        self.problems = tuple(problems)
        super().__init__(
            "\n".join(map(str, self.problems)),
            code="settings_invalid",
            problems=[dataclasses.asdict(problem) for problem in self.problems],
        )


@dataclass(frozen=True)
class RoleProblem:
    """役割の使い方の誤り。restricted は宣言のない場所での使用、math_only・text_only は内容の違反。"""

    kind: Literal["restricted", "math_only", "text_only"]
    role: str

    def describe(self, *, setting: str | None = None) -> str:
        """文言（ページと要素の識別子は呼び出し側が前に付ける）。setting は設定の場所（例：theme.toml の roles.formula）。"""
        if self.kind == "math_only":
            return f"役割 {self.role} は数式だけに使えます"
        if self.kind == "text_only":
            return f"役割 {self.role} に数式は使えません"
        where = f"（設定: {setting}）" if setting else ""
        return (
            f'役割 {self.role} はこのページでは使えません。deck.slide(..., allow_roles=["{self.role}"]) か'
            f"要素の allow_roles で宣言してください{where}"
        )


def role_usage_problems(
    name: str,
    role: Mapping[str, Any],
    *,
    slide_id: str,
    declared: Collection[str],
    content: Sequence[Mapping[str, Any]],
) -> list[RoleProblem]:
    """役割 name の使い方の誤りを返す。構築（作成 API）が使う。

    role は継承を解決した役割、declared はページ・要素・祖先のまとまりの allow_roles の和、content は場面の断片の列。
    使う場所を限った役割（restricted が真か、allowed_slide_ids がある役割）は、allowed_slide_ids のページか、宣言した場所でだけ使うことができる。
    allowed_content が math の役割は空白以外の通常の文字を、text の役割は数式を含めることができない。
    """
    problems = []
    allowed = role.get("allowed_slide_ids")
    pages = allowed if isinstance(allowed, (list, tuple)) else ()
    if (role.get("restricted", False) or allowed is not None) and name not in declared and slide_id not in pages:
        problems.append(RoleProblem("restricted", name))
    if role.get("allowed_content") == "math" and any(
        fragment.get("kind") != "math" and str(fragment.get("text", "")).strip() for fragment in content
    ):
        problems.append(RoleProblem("math_only", name))
    if role.get("allowed_content") == "text" and any(fragment.get("kind") == "math" for fragment in content):
        problems.append(RoleProblem("text_only", name))
    return problems


def _value(value: Any) -> str:
    text = repr(value)
    return f"（値: {text if len(text) <= 40 else text[:39] + '…'}）"


def _key(prefix: str, name: Any) -> str:
    return f"{prefix}.{name}" if prefix else str(name)


def _mapping(value: Any) -> Mapping:
    return value if isinstance(value, Mapping) else {}


def _is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def _not_table(file: str, key: str, value: Any) -> Problem:
    return Problem(file, key, ("は" if key else "設定は") + "表形式で指定してください" + _value(value))


# 値の検査。問題を名前に続く文で返し、問題がなければ None を返す。


def _version(value: Any) -> str | None:
    return None if value == 1 and not isinstance(value, bool) else "は 1 で指定してください"


def _boolean(value: Any) -> str | None:
    return None if isinstance(value, bool) else "は真偽値で指定してください"


def _text(value: Any) -> str | None:
    return None if isinstance(value, str) and value.strip() else "は空でない文字列で指定してください"


def _page_number_format(value: Any) -> str | None:
    """ページ番号の書式。{page} を必ず含み、波括弧の欄は {page} と {total} だけにする。"""
    problem = "は {page} を含み、波括弧の欄が {page} と {total} だけの文字列で指定してください"
    if not isinstance(value, str) or "{page}" not in value:
        return problem
    try:
        value.format(page=1, total=1)
    except (KeyError, IndexError, ValueError):
        return problem
    return None


def _non_negative(value: Any) -> str | None:
    return None if _is_number(value) and value >= 0 else "は 0 以上の有限の数で指定してください"


def _skippable_check(value: Any) -> str | None:
    """役割の skip_checks の項目：警告か参考の診断の識別子（エラーは外せない）。"""
    from .diagnostics import DIAGNOSTICS

    spec = DIAGNOSTICS.get(value) if isinstance(value, str) else None
    if spec is None:
        return "は診断の識別子で指定してください（docs/diagnostics.md）"
    if spec.severity == "error":
        return f"にエラーの診断 {value} は指定できません（警告と参考だけを外すことができます）"
    return None


def _choice(values: tuple[str, ...]) -> Callable[[Any], str | None]:
    def check(value: Any) -> str | None:
        return (
            None if isinstance(value, str) and value in values else f"は {'、'.join(values)} のどれかで指定してください"
        )

    return check


def _character(value: Any) -> str | None:
    return None if isinstance(value, str) and len(value) == 1 else "は 1 文字で指定してください"


def _word(value: Any) -> str | None:
    if isinstance(value, str) and value and not any(character.isspace() for character in value):
        return None
    return "は空白を含まない空でない文字列で指定してください"


def _pattern(value: Any) -> str | None:
    if not isinstance(value, str):
        return f"は文字列で指定してください{_value(value)}"
    try:
        re.compile(value)
    except re.error as error:
        return f"は正規表現として解釈できません（{error}）"
    return None


def _numbers(value: Any, count: int) -> bool:
    return isinstance(value, (list, tuple)) and len(value) == count and all(map(_is_number, value))


def _fraction_area(value: Any) -> str | None:
    if not _numbers(value, 4):
        return "は [x, y, 幅, 高さ] の 4 つの比で指定してください"
    x, y, w, h = value
    if min(x, y) < 0 or w <= 0 or h <= 0 or x + w > 1 + 1e-9 or y + h > 1 + 1e-9:
        return "は 0 以上 1 以下の比で、幅と高さを正、x + 幅と y + 高さを 1 以下にしてください"
    return None


def _pt_area(value: Any) -> str | None:
    if not _numbers(value, 4):
        return "は [x, y, 幅, 高さ] の 4 つの数（pt）で指定してください"
    return None if value[2] > 0 and value[3] > 0 else "の幅と高さは正にしてください"


def _fraction_point(value: Any) -> str | None:
    if not _numbers(value, 2):
        return "は [x, y] の 2 つの比で指定してください"
    return None if all(0 <= number <= 1 for number in value) else "は 0 以上 1 以下の比で指定してください"


def _pt_point(value: Any) -> str | None:
    return None if _numbers(value, 2) else "は [x, y] の 2 つの数（pt）で指定してください"


# deck.toml と実行環境の値


def _string(value: Any) -> str | None:
    return None if isinstance(value, str) else "は文字列で指定してください"


def _local(value: Any) -> str | None:
    """ローカルのパス。相対パスの基準と、ファイルの有無は読み込む側が扱う。"""
    if not isinstance(value, str) or not value.strip():
        return "は空でない文字列で指定してください"
    return "はローカルのパスで指定してください（外部 URL は使えません）" if _URL.match(value) else None


def _deck_path(value: Any) -> str | None:
    return None if _local(value) is None else "は deck.toml のパスで指定してください"


def _setting_file(value: Any) -> str | None:
    return None if _local(value) is None else "は設定名またはファイルのパスで指定してください"


def _entrypoint(value: Any) -> str | None:
    if isinstance(value, str) and ":" in value:
        file, function = value.rsplit(":", 1)
        if _local(file) is None and function.isidentifier():
            return None
    return "は slides.py:build の形式で指定してください"


def _page_size(value: Any) -> str | None:
    if (isinstance(value, str) and value in PAGE_SIZES) or (_numbers(value, 2) and min(value) > 0):
        return None
    return "は 16:9、4:3、または [幅, 高さ]（正の数、pt）で指定してください"


def _index(value: Any) -> str | None:
    return (
        None
        if isinstance(value, int) and not isinstance(value, bool) and value >= 0
        else "は 0 以上の整数で指定してください"
    )


def _count(value: Any) -> str | None:
    return (
        None
        if isinstance(value, int) and not isinstance(value, bool) and value >= 1
        else "は 1 以上の整数で指定してください"
    )


def _timeout(value: Any) -> str | None:
    return None if _is_number(value) and value > 0 else "は正の秒数で指定してください"


def _wait_seconds(value: Any) -> str | None:
    return None if _is_number(value) and value >= 0 else "は 0 以上の秒数で指定してください"


def _function_name(value: Any) -> str | None:
    # 数式の構文（math.syntax.function_names）と同じ規則。
    return None if isinstance(value, str) and re.fullmatch(r"[A-Za-z]+", value) else "は英字だけで指定してください"


def _fallback_renderer(value: Any) -> str | None:
    return None if value == "powerpoint" else "は powerpoint で指定してください（互換画像は PowerPoint で作ります）"


def _fallback_scale(value: Any) -> str | None:
    return None if _is_number(value) and 1 < value <= 8 else "は 1 より大きく 8 以下の数で指定してください"


def _raster_scale(value: Any) -> str | None:
    return None if _is_number(value) and 1 <= value <= 16 else "は 1 以上 16 以下の数で指定してください"


def toml_key(name: Any) -> str:
    """表のキーを TOML の書き方で示す（例：fonts."Yu Gothic UI"）。"""
    return name if isinstance(name, str) and _BARE_KEY.fullmatch(name) else json.dumps(str(name), ensure_ascii=False)


def check_keys(value: Any, allowed: Iterable[str], key: str, *, file: str) -> list[Problem]:
    """表の未知の設定を候補付きの問題にする（deck.toml 用）。key は表の位置で、最上位なら空文字列。"""
    if not isinstance(value, Mapping):
        return [_not_table(file, key, value)]
    allowed = frozenset(allowed)
    return [
        Problem(file, _key(key, name), "は不明な設定です", suggest(name, allowed))
        for name in value
        if name not in allowed
    ]


class _Checker:
    """1 つのファイルの設定を調べ、問題を集める。"""

    def __init__(self, file: str) -> None:
        self.file = file
        self.problems: list[Problem] = []

    def add(self, key: str, message: str, suggestions: Iterable[str] = ()) -> None:
        self.problems.append(Problem(self.file, key, message, tuple(suggestions)))

    def table(self, key: str, value: Any) -> bool:
        if isinstance(value, Mapping):
            return True
        self.problems.append(_not_table(self.file, key, value))
        return False

    def unknown(self, prefix: str, name: Any, allowed: Iterable[str]) -> None:
        self.add(_key(prefix, name), "は不明な設定です", suggest(name, allowed))

    def value(self, key: str, value: Any, check: Callable[[Any], str | None]) -> None:
        problem = check(value)
        if problem:
            self.add(key, problem + _value(value))

    def array(self, key: str, value: Any, check: Callable[[Any], str | None], *, show_value: bool = True) -> None:
        if not isinstance(value, (list, tuple)):
            self.add(key, "は配列で指定してください" + _value(value))
            return
        for number, item in enumerate(value, 1):
            problem = check(item)
            if problem:
                self.add(key, f"の {number} 番目{problem}" + (_value(item) if show_value else ""))

    def style_names(self, prefix: str, names: Iterable[str], kind: str | None, *, extra: Iterable[str] = ()) -> None:
        for name, found, reason in styles.unknown_names(kind, names, extra=extra):
            self.add(_key(prefix, name), reason or "は不明な設定です", found)

    def style_value(self, key: str, name: str, value: Any) -> None:
        problem = styles.check_style_value(name, value)
        if problem:
            self.add(key, problem + _value(value))

    def style_table(self, key: str, table: Mapping, kind: str | None) -> None:
        """書式の表。kind の書式（None なら全種類の書式）を受け付ける。"""
        self.style_names(key, table, kind)
        accepted = styles.ALL_STYLE_KEYS if kind is None else styles.STYLE_KEYS[kind]
        for name, item in table.items():
            if name in accepted:
                self.style_value(f"{key}.{name}", name, item)

    # テーマ

    def theme(self, data: Any) -> list[Problem]:
        if not self.table("", data):
            return self.problems
        for name, value in data.items():
            if name not in THEME_KEYS:
                self.unknown("", name, THEME_KEYS)
            elif name == "schema_version":
                self.value(name, value, _version)
            elif name == "fonts":
                self.fixed(name, value, FONT_KEYS, _text)
            elif name == "colors":
                self.colors(value)
            elif name == "margins":
                self.fixed(name, value, MARGIN_KEYS, _non_negative)
            elif name == "defaults":
                self.defaults(value)
            elif name == "roles":
                self.roles(value)
            else:
                self.page_kinds(value)
        return self.problems

    def fixed(self, key: str, value: Any, allowed: tuple[str, ...], check: Callable[[Any], str | None]) -> None:
        if not self.table(key, value):
            return
        for name, item in value.items():
            if name in allowed:
                self.value(f"{key}.{name}", item, check)
            else:
                self.unknown(key, name, allowed)

    def colors(self, value: Any) -> None:
        if not self.table("colors", value):
            return
        for name, item in value.items():
            key = f"colors.{name}"
            if name == "none":
                self.add(key, "は色の名前に使えません（none は塗らない・線なしの指定に使います）")
            elif not styles.COLOR_NAME.fullmatch(str(name)):
                self.add(key, "は色の名前に使えません（英字か _ で始め、英字・数字・_ だけを使います）")
            if not (isinstance(item, str) and styles.HEX_COLOR.fullmatch(item)):
                self.add(key, "は 6 桁の 16 進数で指定してください" + _value(item))

    def defaults(self, value: Any) -> None:
        if not self.table("defaults", value):
            return
        for name, item in value.items():
            key = f"defaults.{name}"
            if name == "tail_side":
                self.tail_default(key)
            elif name in styles.ALL_STYLE_KEYS:
                self.style_value(key, name, item)
            elif name in styles.DEFAULT_KINDS:
                if self.table(key, item):
                    if "tail_side" in item:
                        self.tail_default(f"{key}.tail_side")
                    self.style_table(key, {k: v for k, v in item.items() if k != "tail_side"}, name)
            elif isinstance(item, Mapping):
                self.unknown("defaults", name, styles.DEFAULT_KINDS)
            else:
                self.style_names("defaults", [name], None)

    def tail_default(self, key: str) -> None:
        # [defaults] と [defaults.box] は、すべての box と表のセルに届く。しっぽの寸法だけは既定に置くことができる。
        self.add(
            key,
            "は [defaults] と [defaults.<種類>] に置くことができません。しっぽの有無は役割と要素で決めます"
            "（tail_width_pt・tail_length_pt・tail_at は既定に置くことができます）",
        )

    def roles(self, value: Any) -> None:
        if not self.table("roles", value):
            return
        for name, role in value.items():
            key = f"roles.{name}"
            if not _ROLE_NAME.fullmatch(str(name)):
                self.add(key, "は役割の名前に使えません（英字か _ で始め、英字・数字・_ だけを使います）")
            if not self.table(key, role):
                continue
            self.style_names(key, role, None, extra=styles.ROLE_ONLY_KEYS)
            for item_name, item in role.items():
                item_key = f"{key}.{item_name}"
                if item_name == "inherits":
                    self.value(item_key, item, _text)
                elif item_name == "allowed_content":
                    self.value(item_key, item, _choice(ROLE_CONTENTS))
                elif item_name == "allowed_slide_ids":
                    self.array(item_key, item, _text)
                elif item_name == "restricted":
                    self.value(item_key, item, _boolean)
                elif item_name == "skip_checks":
                    self.array(item_key, item, _skippable_check)
                elif item_name in styles.ALL_STYLE_KEYS:
                    self.style_value(item_key, item_name, item)

    def page_kinds(self, value: Any) -> None:
        if not self.table("page_kinds", value):
            return
        for kind, settings in value.items():
            key = f"page_kinds.{kind}"
            if kind not in PAGE_KINDS:
                self.add(key, "は不明なページ種別です", suggest(kind, PAGE_KINDS))
                continue
            if not self.table(key, settings):
                continue
            for name, item in settings.items():
                item_key = f"{key}.{name}"
                if name not in PAGE_KIND_KEYS:
                    self.unknown(key, name, PAGE_KIND_KEYS)
                elif name in ("title_role", "default_role"):
                    self.value(item_key, item, _text)
                elif name == "page_number":
                    self.value(item_key, item, _boolean)
                elif name == "page_number_format":
                    self.value(item_key, item, _page_number_format)
                elif name == "background":
                    # 色か "none"（背景の図形を作らない）。
                    self.style_value(item_key, "fill", item)
                elif name in ("title_area", "content_area"):
                    self.value(item_key, item, _fraction_area)
                elif name.endswith("_area_pt"):
                    self.value(item_key, item, _pt_area)
                elif name in ("title_style", "footer_style"):
                    if self.table(item_key, item):
                        self.style_table(item_key, item, "text")
                else:
                    self.problems.extend(check_decorations(item, page_kind=kind, file=self.file))

    # 記述規則

    def rules(self, data: Any) -> list[Problem]:
        if not self.table("", data):
            return self.problems
        for name, value in data.items():
            if name not in RULES_KEYS:
                self.unknown("", name, RULES_KEYS)
            elif name == "schema_version":
                self.value(name, value, _version)
            elif name == "forbidden_characters":
                self.array(name, value, _character)
            elif name == "require_notes":
                self.value(name, value, _boolean)
            elif name == "notes_style":
                self.value(name, value, _choice(NOTES_STYLES))
            elif self.table(name, value):
                self.math_candidates(value)
        return self.problems

    def math_candidates(self, value: Mapping) -> None:
        for name, item in value.items():
            key = f"math_candidates.{name}"
            if name not in MATH_CANDIDATE_KEYS:
                self.unknown("math_candidates", name, MATH_CANDIDATE_KEYS)
            elif name in ("enabled", "numbers"):
                self.value(key, item, _boolean)
            elif name in ("variables", "operations"):
                self.array(key, item, _word)
            elif name == "exclude_patterns":
                self.array(key, item, _pattern, show_value=False)
            elif name == "exclude_roles":
                self.array(key, item, _text)
            elif not isinstance(item, (list, tuple)):
                self.add(key, "は配列で指定してください" + _value(item))
            else:
                for number, kind in enumerate(item, 1):
                    if kind not in PAGE_KINDS:
                        self.add(
                            key,
                            f"の {number} 番目は {'、'.join(PAGE_KINDS)} のどれかで指定してください{_value(kind)}",
                            suggest(kind, PAGE_KINDS),
                        )

    # deck.toml

    def deck(self, data: Any) -> list[Problem]:
        # 未知の名前は、最上位と各表で check_keys が候補を添えて報告する。
        self.problems.extend(check_keys(data, DECK_KEYS[""], "", file=self.file))
        if not isinstance(data, Mapping):
            return self.problems
        for name, value in data.items():
            if name not in DECK_KEYS[""]:
                continue
            if name not in DECK_KEYS:
                self.value(name, value, _DECK_VALUES[name])
            elif self.table(name, value):
                self.problems.extend(check_keys(value, DECK_KEYS[name], name, file=self.file))
                for key, item in value.items():
                    if key in DECK_KEYS[name]:
                        self.deck_value(name, key, item)
        return self.problems

    def deck_value(self, table: str, name: str, value: Any) -> None:
        key = f"{table}.{name}"
        if key == "inputs.documents":
            self.documents(value)
        elif key == "render.required":
            self.required_renderers(value)
        elif key in _DECK_ARRAYS:
            self.array(key, value, _DECK_ARRAYS[key])
        elif table == "metadata":
            self.value(key, value, _string)
        elif table == "deliverables":
            self.value(key, value, _local)
        else:
            self.value(key, value, _DECK_VALUES[key])

    def documents(self, value: Any) -> None:
        key = "inputs.documents"
        if not isinstance(value, (list, tuple)):
            self.add(key, "は配列で指定してください" + _value(value))
            return
        allowed = DECK_KEYS[key]
        seen: dict[str, int] = {}
        for number, item in enumerate(value, 1):
            at = f"の {number} 番目"
            if not isinstance(item, Mapping):
                self.add(key, f"{at}は表形式で指定してください" + _value(item))
                continue
            for name in item:
                if name not in allowed:
                    self.add(key, f"{at}の {name} は不明な設定です", suggest(name, allowed))
            for name in ("id", "path"):
                if name not in item:
                    self.add(key, f"{at}に {name} を指定してください")
            if "affects_content" in item and not isinstance(item["affects_content"], bool):
                self.add(key, f"{at}の affects_content は真偽値で指定してください" + _value(item["affects_content"]))
            if "id" in item:
                identifier = item["id"]
                if problem := _text(identifier):
                    self.add(key, f"{at}の id {problem}{_value(identifier)}")
                elif identifier in seen:
                    self.add(key, f"{at}の id {identifier} は {seen[identifier]} 番目と重複しています")
                else:
                    seen[identifier] = number
            if "path" in item and (problem := _local(item["path"])):
                self.add(key, f"{at}の path {problem}{_value(item['path'])}")

    def required_renderers(self, value: Any) -> None:
        key = "render.required"
        if not isinstance(value, (list, tuple)):
            self.add(key, "は配列で指定してください" + _value(value))
            return
        if not value:
            self.add(key, f"には {'、'.join(RENDERERS)} のうち 1 つ以上を指定してください")
            return
        seen = set()
        for number, item in enumerate(value, 1):
            if not (isinstance(item, str) and item in RENDERERS):
                self.add(
                    key,
                    f"の {number} 番目は {'、'.join(RENDERERS)} のどれかで指定してください{_value(item)}",
                    suggest(item, RENDERERS),
                )
            elif item in seen:
                self.add(key, f"の {number} 番目の {item} は重複しています")
            else:
                seen.add(item)

    # 実行環境

    def environment(self, data: Any) -> list[Problem]:
        if not self.table("", data):
            return self.problems
        top = ENVIRONMENT_KEYS[""]
        for name, value in data.items():
            if name not in top:
                self.unknown("", name, top)
            elif name == "schema_version":
                self.value(name, value, _version)
            elif name == "jobs":
                self.value(name, value, _count)
            elif name == "font_directories":
                self.font_directories(value)
            elif name == "fonts":
                self.font_files(value)
            elif self.table(name, value):
                self.renderers(value)
        return self.problems

    def font_directories(self, value: Any) -> None:
        key = "font_directories"
        if not isinstance(value, (list, tuple)) or any(not isinstance(item, str) or not item.strip() for item in value):
            self.add(key, "は文字列の配列で指定してください" + _value(value))
            return
        for number, item in enumerate(value, 1):
            if problem := _local(item):
                self.add(key, f"の {number} 番目{problem}{_value(item)}")

    def font_files(self, value: Any) -> None:
        if not self.table("fonts", value):
            return
        allowed = ENVIRONMENT_KEYS["fonts.*"]
        for family, font in value.items():
            key = f"fonts.{toml_key(family)}"
            if not self.table(key, font):
                continue
            for name, item in font.items():
                if name not in allowed:
                    self.unknown(key, name, allowed)
                else:
                    self.value(f"{key}.{name}", item, _local if name.endswith("path") else _index)

    def renderers(self, value: Mapping) -> None:
        for name, settings in value.items():
            key = _key("renderers", name)
            if name not in RENDERERS:
                self.unknown("renderers", name, RENDERERS)
                continue
            if not self.table(key, settings):
                continue
            allowed = ENVIRONMENT_KEYS[key]
            for item, setting in settings.items():
                if item not in allowed:
                    self.unknown(key, item, allowed)
                else:
                    self.value(
                        f"{key}.{item}", setting, {"timeout": _timeout, "wait_seconds": _wait_seconds}.get(item, _text)
                    )


# deck.toml の値の検査。表の中の項目は「表.名前」、配列は要素ごとの検査（_DECK_ARRAYS）。
_DECK_VALUES: dict[str, Callable[[Any], str | None]] = {
    "schema_version": _version,
    "extends": _deck_path,
    "entrypoint": _entrypoint,
    "size": _page_size,
    "theme": _setting_file,
    "rules": _setting_file,
    "output_dir": _local,
    "environment_override": _local,
    "render.primary": _choice(RENDERERS),
    "math.fallback": _choice(FALLBACK_MODES),
    "math.fallback_renderer": _fallback_renderer,
    "math.fallback_scale": _fallback_scale,
    "images.raster_scale": _raster_scale,
    "images.vector": _boolean,
    "template.path": _local,
    "template.layout_index": _index,
    "template.show_master_shapes": _boolean,
}
_DECK_ARRAYS: dict[str, Callable[[Any], str | None]] = {
    "inputs.read_roots": _local,
    "inputs.python": _local,
    "inputs.bibliography": _local,
    "inputs.asset_dirs": _local,
    "math.functions": _function_name,
}


def check_settings(
    kind: Literal["theme", "rules", "deck", "environment"], data: Mapping, *, file: str
) -> list[Problem]:
    """1 つのファイルの設定名と値を確かめる。

    kind はテーマ（theme）、記述規則（rules）、資料の設定（deck。deck.toml）、実行環境（environment。environment.toml）。
    参照と、ファイルをまたぐ必須の設定は合成した後に確かめる（check_theme_references・check_rules_references・
    check_deck_requirements）。パスは形だけを確かめ、相対パスの基準とファイルの有無は読み込む側が扱う。
    """
    checker = _Checker(file)
    if kind == "theme":
        return checker.theme(data)
    if kind == "rules":
        return checker.rules(data)
    if kind == "deck":
        return checker.deck(data)
    if kind == "environment":
        return checker.environment(data)
    raise ValueError(f"未対応の設定の種類です: {kind}")


def check_deck_requirements(config: Mapping, *, file: str, origins: Mapping[str, str] | None = None) -> list[Problem]:
    """合成した deck.toml の必須の設定（文書情報、雛形のパス）と、render.primary が render.required にあることを確かめる。

    origins は設定の位置（例：render.primary）から、その値を書いたファイルの表示名への対応で、extends で合成したときに渡す。
    位置にもその親の表にも対応がなければ file を示す。形の間違った値は check_settings が報告するので飛ばす。
    """
    origins = origins or {}

    def where(key: str) -> str:
        while key:
            if key in origins:
                return origins[key]
            key = key.rpartition(".")[0]
        return file

    problems: list[Problem] = []
    metadata = config.get("metadata", {})
    if isinstance(metadata, Mapping):
        for name in REQUIRED_METADATA:
            value = metadata.get(name)
            if value is None or (name == "title" and isinstance(value, str) and not value.strip()):
                key = f"metadata.{name}"
                problems.append(Problem(where(key), key, "を明示指定してください"))
    render = config.get("render", {})
    if isinstance(render, Mapping):
        primary, required = render.get("primary", RENDERERS[0]), render.get("required", [RENDERERS[0]])
        if (
            isinstance(primary, str)
            and primary in RENDERERS
            and isinstance(required, (list, tuple))
            and required
            and all(isinstance(item, str) and item in RENDERERS for item in required)
            and primary not in required
        ):
            problems.append(
                Problem(where("render.primary"), "render.primary", f"の {primary} を render.required に含めてください")
            )
    template = config.get("template")
    if isinstance(template, Mapping) and template and "path" not in template:
        problems.append(Problem(where("template"), "template.path", "を指定してください"))
    return problems


def check_decorations(decorations: Any, *, page_kind: str, file: str) -> list[Problem]:
    """ページ種別の飾り（`[[page_kinds.<種別>.decorations]]`）の設定を確かめる。色の参照は合成後に確かめる。

    飾りは線（line・arrow）か図形（box）で、文字は置かない。位置は box が area か area_pt、線が start・end か
    start_pt・end_pt の一方の組である。問題には飾りの番号（1 始まり）を添える。
    """
    key = f"page_kinds.{page_kind}.decorations"
    if not isinstance(decorations, (list, tuple)):
        return [Problem(file, key, "は配列で指定してください" + _value(decorations))]
    problems: list[Problem] = []
    seen: dict[str, int] = {}
    for number, item in enumerate(decorations, 1):

        def add(message: str, suggestions: Iterable[str] = (), number: int = number) -> None:
            problems.append(Problem(file, key, f"の {number} 番目{message}", tuple(suggestions)))

        if not isinstance(item, Mapping):
            add("は表形式で指定してください" + _value(item))
            continue
        kind = item.get("kind")
        if kind is None:
            add(f"に kind（{'、'.join(DECORATION_KINDS)}）を指定してください")
            continue
        if kind not in DECORATION_KINDS:
            add(
                f"の kind は {'、'.join(DECORATION_KINDS)} のどれかで指定してください{_value(kind)}",
                suggest(kind, DECORATION_KINDS),
            )
            continue
        identifier = item.get("id")
        if identifier is None:
            add("に id を指定してください")
        elif not (
            isinstance(identifier, str)
            and identifier.strip()
            and "/" not in identifier
            and not any(ord(character) < 32 for character in identifier)
        ):
            add("の id は空でなく、/ と制御文字を含まない文字列にしてください" + _value(identifier))
        elif identifier in seen:
            add(f"の id {identifier} は {seen[identifier]} 番目と重複しています")
        else:
            seen[identifier] = number
        allowed = _DECORATION_KEYS[kind]
        for name in item:
            if name in allowed:
                if name in styles.ALL_STYLE_KEYS and (problem := styles.check_style_value(name, item[name])):
                    add(f"の {name} {problem}{_value(item[name])}")
            elif name == "text":
                add("の text は使えません。飾りには文字を置きません")
            elif name in styles.UNSUPPORTED:
                add(f"の {name} {styles.UNSUPPORTED[name]}")
            elif name in styles.ALL_STYLE_KEYS or any(name in keys for keys in _DECORATION_KEYS.values()):
                add(f"の {name} は {kind} の飾りでは使えません")
            else:
                alias = styles.ALIASES.get(name)
                add(f"の {name} は不明な設定です", (alias,) if alias in allowed else suggest(name, allowed))
        if kind == "box":
            given = [name for name in ("area", "area_pt") if name in item]
            if len(given) != 1:
                add("には area か area_pt のどちらか一方を指定してください")
            elif problem := (_fraction_area if given[0] == "area" else _pt_area)(item[given[0]]):
                add(f"の {given[0]} {problem}{_value(item[given[0]])}")
            shape = item.get("shape", "rectangle")
            if shape not in styles.SHAPES:
                add(
                    f"の shape は {'、'.join(styles.SHAPES)} のどれかで指定してください{_value(shape)}",
                    suggest(shape, styles.SHAPES),
                )
            elif "corner_radius_pt" in item and shape != "rounded_rectangle":
                add("の corner_radius_pt は shape が rounded_rectangle のときだけ指定できます")
            continue
        fraction = any(name in item for name in ("start", "end"))
        points = any(name in item for name in ("start_pt", "end_pt"))
        if fraction == points:
            add("には start と end か、start_pt と end_pt のどちらか一方の組を指定してください")
            continue
        names, check = (("start", "end"), _fraction_point) if fraction else (("start_pt", "end_pt"), _pt_point)
        if any(name not in item for name in names):
            add(f"には {names[0]} と {names[1]} の両方を指定してください")
            continue
        found = [(name, check(item[name])) for name in names]
        for name, problem in found:
            if problem:
                add(f"の {name} {problem}{_value(item[name])}")
        if not any(problem for _, problem in found) and list(item[names[0]]) == list(item[names[1]]):
            add("の始点と終点を別にしてください")
    return problems


def _role_chain(roles: Mapping, name: str) -> tuple[list[str], list[str] | None, bool]:
    """name から継承をたどる。(たどった役割の列, 循環した役割の列か None, 途中で切れたか) を返す。"""
    path: list[str] = []
    current = name
    while True:
        if current in path:
            return path, path[path.index(current) :] + [current], False
        role = roles.get(current)
        if not isinstance(role, Mapping):
            return path, None, True
        path.append(current)
        parent = role.get("inherits")
        if parent is None:
            return path, None, False
        if not isinstance(parent, str) or not parent.strip():
            return path, None, True
        current = parent


def check_theme_references(theme: Mapping, *, file: str) -> list[Problem]:
    """合成したテーマの参照（色、役割、継承）と、継承を解決した役割の文字サイズを確かめる。

    色は `[colors]` の名前か 6 桁の 16 進数で、"none" は fill・line_color とページ種別の background だけに指定できる。
    形の間違った値は check_settings が報告するので、ここでは飛ばす。
    """
    problems: list[Problem] = []
    colors = _mapping(theme.get("colors"))
    roles = _mapping(theme.get("roles"))

    def color(key: str, value: Any, *, paint: bool, at: str = "") -> None:
        if not isinstance(value, str) or styles.check_style_value("fill", value):
            return
        if value == "none":
            if not paint:
                problems.append(
                    Problem(
                        file,
                        key,
                        f"{at}には none を指定できません（none は fill・line_color とページ種別の background だけに使えます）",
                    )
                )
        elif not (styles.HEX_COLOR.fullmatch(value) or value in colors):
            problems.append(Problem(file, key, f"{at}の色 {value} はテーマにありません", suggest(value, colors)))

    def style_table(key: str, table: Any, kind: str | None = None) -> None:
        # その種類が受け付けない書式名は check_settings が報告するので、色だけを確かめる。
        accepted = styles.ALL_STYLE_KEYS if kind is None else styles.STYLE_KEYS[kind]
        for name, value in _mapping(table).items():
            if name in styles.COLOR_KEYS and name in accepted:
                color(f"{key}.{name}", value, paint=name in styles.PAINT_KEYS)

    def role_reference(key: str, value: Any) -> None:
        if isinstance(value, str) and value.strip() and value not in roles:
            problems.append(Problem(file, key, f"の役割 {value} はテーマにありません", suggest(value, roles)))

    defaults = _mapping(theme.get("defaults"))
    style_table("defaults", {name: value for name, value in defaults.items() if name not in styles.DEFAULT_KINDS})
    for name, value in defaults.items():
        if name in styles.DEFAULT_KINDS:
            style_table(f"defaults.{name}", value, name)
    for name, role in roles.items():
        style_table(f"roles.{name}", role)
        if isinstance(role, Mapping):
            role_reference(f"roles.{name}.inherits", role.get("inherits"))
    for kind, settings in _mapping(theme.get("page_kinds")).items():
        if kind not in PAGE_KINDS or not isinstance(settings, Mapping):
            continue
        key = f"page_kinds.{kind}"
        for name in ("title_role", "default_role"):
            role_reference(f"{key}.{name}", settings.get(name))
        if "background" in settings:
            color(f"{key}.background", settings["background"], paint=True)
        for name in ("title_style", "footer_style"):
            style_table(f"{key}.{name}", settings.get(name), "text")
        decorations = settings.get("decorations")
        for number, item in enumerate(decorations if isinstance(decorations, (list, tuple)) else (), 1):
            shape_kind = _mapping(item).get("kind")
            accepted = _DECORATION_KEYS.get(shape_kind, frozenset()) if isinstance(shape_kind, str) else frozenset()
            for name, value in _mapping(item).items():
                if name in styles.PAINT_KEYS and name in accepted:
                    color(f"{key}.decorations", value, paint=True, at=f"の {number} 番目の {name} ")

    cycles: set[frozenset[str]] = set()
    for name in roles:
        path, cycle, broken = _role_chain(roles, name)
        if cycle:
            if frozenset(cycle) not in cycles:
                cycles.add(frozenset(cycle))
                problems.append(
                    Problem(file, f"roles.{cycle[0]}.inherits", f"の継承が循環しています（{' → '.join(cycle)}）")
                )
            continue
        if broken:
            continue
        resolved: dict[str, Any] = {}
        for role_name in reversed(path):
            resolved.update(roles[role_name])
        size, minimum = resolved.get("size_pt"), resolved.get("min_size_pt")
        if size is None or minimum is None:
            problems.append(
                Problem(file, f"roles.{name}", "に size_pt と min_size_pt を指定するか、inherits で継承してください")
            )
        elif _is_number(size) and _is_number(minimum) and 0 < size < minimum:
            problems.append(
                Problem(
                    file,
                    f"roles.{name}",
                    f"の size_pt は min_size_pt 以上にしてください（size_pt: {size}、min_size_pt: {minimum}）",
                )
            )
    return problems


def check_rules_references(rules: Mapping, theme: Mapping, *, file: str) -> list[Problem]:
    """合成した記述規則の参照（math_candidates.exclude_roles の役割がテーマにあること）を確かめる。"""
    roles = _mapping(theme.get("roles"))
    names = _mapping(rules.get("math_candidates")).get("exclude_roles")
    if not isinstance(names, (list, tuple)):
        return []
    return [
        Problem(
            file,
            "math_candidates.exclude_roles",
            f"の {number} 番目の役割 {name} はテーマにありません",
            suggest(name, roles),
        )
        for number, name in enumerate(names, 1)
        if isinstance(name, str) and name.strip() and name not in roles
    ]
