"""診断の並べ替えと件数、コマンドの結果の要約。

要約は利用者が読む既定の出力である。機械が読む用途には、結果の辞書を JSON で出す（`--json`）。
要約は結果の形を決め打ちしない。知らないコマンドや、形の違う結果でも例外を出さず、状態と文言だけを示す。
コマンドの最上位で捕まえた例外の結果（`exception` を持つ失敗）は、どのコマンドでも同じ形で示す。
"""

from __future__ import annotations

from contextvars import ContextVar
from functools import partial
import math
from pathlib import Path
import shlex
from typing import Any, Iterable

from .progress import RENDERER_LABELS, stage_label
from .selection import format_numbers

SEVERITY_ORDER = {"error": 0, "warning": 1, "info": 2}
SEVERITY_LABELS = {"error": "エラー", "warning": "警告", "info": "参考"}
COMMAND_LABELS = {
    "layout": "配置の検査",
    "doctor": "事前確認",
    "build": "生成",
    "check": "検査",
    "export": "反映",
    "clean": "整理",
    "renderer status": "表示ソフトの状態",
    "renderer clear": "残存記録の解除",
    "review status": "確認",
    "review record": "確認の記録",
    "review accept": "警告の許容",
    "review revoke": "記録の取り消し",
    "review carry": "確認の引き継ぎ",
    "review sheet": "比較画像",
    "measure": "計測",
}
STATUS_LABELS = {
    "passed": "合格",
    "failed": "失敗",
    "built": "完了",
    "stopped": "変換の前に停止",
    "written": "作成した",
    "exported": "完了",
    "preview": "削除の候補",
    "cleaned": "削除",
    "idle": "待機中",
    "busy": "使用中",
    "pending": "残存記録あり",
    "cleared": "解除",
    "nothing": "解除する記録なし",
    "refused": "解除を拒否",
    "recorded": "記録した",
    "accepted": "許容した",
    "revoked": "取り消した",
    "carried": "引き継いだ",
    "skipped": "省略",
    "complete": "完了",
    "incomplete": "未完了",
    "blocked": "止まっている",
}
# エラーは全件を出すが、この件数で打ち切る。
MAX_ERRORS = 50

_REVIEW_KINDS = {"content": "内容", "visual": "表示", "overview": "一覧画像"}
_VARIANT_NAMES = {"standard": "標準", "fallback": "互換"}
# 生成記録の照合で出る、ファイルごとの誤り。ページに属さないので、識別子ごとにまとめて示す。
_FILE_CODES = ("artifact_unrecorded", "artifact_changed", "input_changed")
# 所要時間の内訳で、段階をまとめる名前。最終形の PPTX の作成は互換画像の作成の一部に含める。
_TIMING_GROUPS = {
    "render": "変換",
    "pdf": "PDF の検査と画像化",
    "math_fallback": "互換画像",
    "final_form": "互換画像",
    "doctor": "事前確認",
    "close": "表示ソフトの終了",
}
# 確認を加えても通らない問題（終了コード 2）の識別子ごとの直し方。
_BLOCKED_OPERATIONS = (
    (
        ("artifact_unrecorded", "artifact_changed", "input_changed"),
        "slidekit build で生成し直す（入力または成果物が生成記録と一致しない）",
    ),
    (("build_incomplete",), "slidekit build で生成し直す（生成が完了していない）"),
    (("build_not_exportable",), "全ページを通常の slidekit build で生成し直す（下書き・部分生成は完成物にできない）"),
)
_REMAINING_NAMES = (("visual", "表示"), ("overview", "一覧画像"), ("content", "内容"), ("warnings", "警告"))
# 表示ソフトの使用中・時間切れ・残存記録で生成が失敗したときの次の操作（原因は資料でなく表示ソフトの状態にある）。
_RENDERER_OPERATIONS = {
    "renderer_timeout": (
        "変換の終了を待ち、slidekit renderer status で確かめる",
        "時間切れが続くなら、実行環境の timeout を延ばしてから slidekit build を再実行する",
    ),
    "renderer_busy": ("別の変換の終了を待ってから slidekit build を再実行する（待つ上限は --wait）",),
    "renderer_pending": (
        "slidekit renderer status で残存記録を確かめ、slidekit renderer clear で解除してから slidekit build を再実行する",
    ),
}


# ---------------------------------------------------------------- 並べ替えと件数


def _text(value: Any) -> str:
    return "" if value is None else str(value)


def _integer(value: Any) -> bool:
    return isinstance(value, int) and not isinstance(value, bool)


def _number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def _slide_numbers(slides: Iterable[dict] | None) -> dict:
    numbers: dict = {}
    for position, slide in enumerate(slides or (), 1):
        if isinstance(slide, dict) and slide.get("id") is not None:
            number = slide.get("number")
            numbers.setdefault(slide["id"], number if _integer(number) else position)
    return numbers


def _page(issue: dict, numbers: dict | None = None) -> int | None:
    number = issue.get("slide_number")
    if _integer(number):
        return number
    if numbers and issue.get("slide_id") is not None:
        return numbers.get(issue["slide_id"])
    return None


def sort_issues(issues: Iterable[dict], slides: Iterable[dict] | None = None) -> list[dict]:
    """ページ番号、重大度（error、warning、info）、code、element_id、related_element_id の順に並べた新しいリストを返す。

    ページ番号は slide_number、なければ slides（"id" と "number" を持つ）から slide_id で引く。ページのない診断は最後に置く。
    """
    numbers = _slide_numbers(slides)

    def key(issue: Any) -> tuple:
        if not isinstance(issue, dict):
            return (2, 0, len(SEVERITY_ORDER), "", "", "")
        number = _page(issue, numbers)
        return (
            0 if number is not None else 1,
            number or 0,
            SEVERITY_ORDER.get(issue.get("severity"), len(SEVERITY_ORDER)),
            _text(issue.get("code")),
            _text(issue.get("element_id")),
            _text(issue.get("related_element_id")),
        )

    return sorted(issues or (), key=key)


def count_issues(issues: Iterable[dict]) -> dict:
    """{"error": 件数, "warning": 件数, "info": 件数}。"""
    counts = {"error": 0, "warning": 0, "info": 0}
    for issue in issues or ():
        if isinstance(issue, dict) and issue.get("severity") in counts:
            counts[issue["severity"]] += 1
    return counts


# ---------------------------------------------------------------- 要約


# 要約の表示の選択（summarize の引数）。verbose は許容した警告と参考の一覧、wraps は配置の検査の行の文字列を出す。
_OPTIONS: ContextVar[dict] = ContextVar("slidekit_report_options", default={})


def _option(name: str) -> bool:
    return bool(_OPTIONS.get().get(name))


def summarize(command: str, result: dict, *, limit: int = 20, verbose: bool = False, wraps: bool = False) -> str:
    """コマンドの結果の要約。

    1 行目は状態と件数、2 行目以降に対象、診断（エラーは全件。ただし 50 件まで。警告と参考は合わせて limit 件まで）、
    所要時間、次の操作を示し、最後の行を合否と終了コードの行（「結果: …」）にする。
    許容した警告（宣言と記録）と生成の参考は件数だけを示し、verbose のときに一覧を出す。
    wraps は配置の検査で、自動で折り返した要素の行の文字列を示す。知らない command や形の違う result でも例外を出さない。
    """
    try:
        data = dict(result) if isinstance(result, dict) else {}
    except Exception:
        data = {}
    summary = _SUMMARIES.get(command) if isinstance(command, str) else None
    if data.get("exception") and data.get("status") == "failed":
        summary = partial(_command_failure, command=command)
    limit = limit if _integer(limit) and limit >= 0 else 20
    lines: list[str] = []
    token = _OPTIONS.set({"verbose": bool(verbose), "wraps": bool(wraps)})
    try:
        if summary is not None:
            try:
                lines = summary(data, limit)
            except Exception:
                # 要約が出せなくても、状態と文言は必ず示す（結果の辞書は --json で読むことができる）。
                lines = []
        if not lines:
            lines = _minimal(command, data)
        lines.append(_result_line(command, data))
    finally:
        _OPTIONS.reset(token)
    return "\n".join(lines)


# 終了コードの意味（cli の説明と同じ）。検査のコマンドは合格・不合格と示す。
_EXIT_LABELS = {0: "成功", 1: "未解決の事項あり", 2: "失敗"}
_CHECK_COMMANDS = frozenset({"layout", "doctor", "check"})


def _result_line(command: Any, data: dict) -> str:
    """最後の行。合否と、残っている事項の件数と、終了コード。末尾だけを読めば結果が分かるようにする。"""
    code = data.get("exit_code")
    if not _integer(code):
        code = 2 if data.get("status") == "failed" else 0
    command = command if isinstance(command, str) else ""
    if data.get("exception") == "KeyboardInterrupt" or data.get("message") == "処理を中断しました":
        return f"結果: 中断（終了コード {code}）"
    if command in _CHECK_COMMANDS or command == "export":
        label = {0: "合格" if command != "export" else "反映した", 1: "不合格", 2: "失敗"}.get(code, "失敗")
        if command == "export" and code != 0:
            label = "中止"
    else:
        label = _EXIT_LABELS.get(code, "失敗")
    parts = []
    counts = data.get("counts") if isinstance(data.get("counts"), dict) else None
    check = data if command == "check" else data.get("check") if isinstance(data.get("check"), dict) else None
    if counts is not None and _integer(counts.get("error")):
        parts.append(f"エラー {counts['error']}")
    elif check is not None and isinstance(check.get("issues"), list):
        issues = _issues(check["issues"])
        missing = sum(1 for issue in issues if issue.get("code") == "review_missing")
        unaccepted = sum(1 for issue in issues if issue.get("code") == "warning_unaccepted")
        parts.append(f"エラー {len(issues) - missing - unaccepted}")
        if check.get("final"):
            parts += [f"未完了の確認 {missing}", f"未対応の警告 {unaccepted}"]
    elif command == "review status":
        groups = _issues(data.get("reviews"))
        warnings = data.get("warnings") if isinstance(data.get("warnings"), dict) else {}
        pending = sum(len(group["pending"]) for group in groups if isinstance(group.get("pending"), list))
        parts += [
            f"エラー {len(_issues(data.get('errors')))}",
            f"未完了の確認 {pending}",
            f"未対応の警告 {len(_issues(warnings.get('pending')))}",
        ]
    remaining = data.get("remaining") if isinstance(data.get("remaining"), dict) else None
    if remaining is not None and (text := _remaining_text(remaining)):
        parts.append(f"残りの確認 {text}")
    parts.append(f"終了コード {code}")
    return f"結果: {label}（{'、'.join(parts)}）"


def _label(command: Any) -> str:
    return COMMAND_LABELS.get(command, _text(command) or "結果") if isinstance(command, str) else "結果"


def _minimal(command: Any, result: dict) -> list[str]:
    label = _label(command)
    status = result.get("status")
    lines = [f"{label}: {STATUS_LABELS.get(status, status) if isinstance(status, str) and status else '状態不明'}"]
    message = result.get("message")
    if message:
        exception = result.get("exception")
        lines.append(f"  {exception}: {message}" if exception else f"  {message}")
    return lines


def _declared(issue: dict) -> bool:
    """資料コードの宣言（slide.allow_warning）で許容した警告か。"""
    return issue.get("severity") == "warning" and isinstance(issue.get("declared"), dict)


def _counts_text(counts: dict, issues: Iterable[dict] = ()) -> str:
    text = f"エラー {counts.get('error', 0)} 件、警告 {counts.get('warning', 0)} 件"
    # 宣言で許容した警告は、件数に含めたうえで内数を示す（宣言のない警告と区別がつくようにする）。
    declared = sum(1 for issue in issues if _declared(issue))
    if declared:
        text += f"、うち宣言で許容 {declared} 件"
    if counts.get("info"):
        text += f"、参考 {counts['info']} 件"
    return text


def _counts(result: dict, issues: list[dict]) -> dict:
    counts = result.get("counts")
    if isinstance(counts, dict) and all(_integer(counts.get(key, 0)) for key in SEVERITY_ORDER):
        return counts
    return count_issues(issues)


def _issues(value: Any) -> list[dict]:
    return [issue for issue in value if isinstance(issue, dict)] if isinstance(value, list) else []


def _shown(path: Any) -> str:
    """カレントディレクトリからの相対パス。できなければそのまま。"""
    if path is None or path == "":
        return ""
    value = Path(str(path))
    if value.is_absolute():
        try:
            return str(value.relative_to(Path.cwd()))
        except (OSError, ValueError):
            return str(value)
    return str(value)


def _argument(path: Any) -> str:
    return shlex.quote(_shown(path))


def _next(operations: list[str]) -> list[str]:
    operations = [operation for operation in operations if operation][:3]
    if not operations:
        return []
    if len(operations) == 1:
        return [f"次の操作: {operations[0]}"]
    return ["次の操作:", *(f"  {operation}" for operation in operations)]


def _error_pages(issues: list[dict]) -> str:
    """エラーのあるページの --slides の指定。番号のないページ（選んだ章だけを構築した検査）は識別子で示す。"""
    errors = [issue for issue in issues if issue.get("severity") == "error"]
    numbers = format_numbers(number for issue in errors if (number := _page(issue)) is not None)
    names = []
    for issue in errors:
        slide_id = issue.get("slide_id")
        if _page(issue) is None and isinstance(slide_id, str) and slide_id:
            # 数字だけの識別子と all・pending は、番号や指定の語と区別するため id: を付ける。
            name = f"id:{slide_id}" if slide_id.isdigit() or slide_id in {"all", "pending"} else slide_id
            if name not in names and "," not in slide_id:
                names.append(name)
    return ",".join(part for part in (numbers, *names) if part)


def _failed_operation(command: str, issues: list[dict]) -> str:
    """検査・生成が失敗（終了コード 2）したときの次の操作。--slides の指定の誤りは指定を直す操作にする。"""
    if any(issue.get("severity") == "error" and issue.get("code") == "invalid_selection" for issue in issues):
        return f"--slides の指定を直して slidekit {command} を再実行する"
    return f"原因を直して slidekit {command} を再実行する"


def _fix_operation(project: Any, issues: list[dict]) -> str:
    """配置の誤りを直す操作。資料と、エラーのあるページを示す。"""
    if not project:
        return "編集元を直して slidekit layout で確かめる"
    pages = _error_pages(issues)
    return f"編集元を直して slidekit layout --project {_argument(project)}" + (
        f" --slides {shlex.quote(pages)}" if pages else ""
    )


# ---------------------------------------------------------------- 診断の 1 件


def _observations(issue: dict) -> list[dict]:
    value = issue.get("observations")
    if isinstance(value, dict):
        return [value]
    return [item for item in value if isinstance(item, dict)] if isinstance(value, list) else []


def _first(issue: dict, *keys: str) -> dict | None:
    return next((item for item in _observations(issue) if all(key in item for key in keys)), None)


def _gathered(issue: dict, key: str) -> list:
    values: list = []
    for item in _observations(issue):
        value = item.get(key)
        for entry in value if isinstance(value, list) else [value] if value is not None else []:
            if entry not in values:
                values.append(entry)
    return values


def _character(value: Any) -> str:
    if isinstance(value, dict):
        value = value.get("character", value.get("char", value.get("text")))
    value = _text(value)
    if len(value) == 1 and (not value.isprintable() or value.isspace()):
        return f"U+{ord(value):04X}"
    return value


def _listed(values: list, limit: int = 10) -> str:
    text = "、".join(_character(value) for value in values[:limit])
    return text + (f" ほか {len(values) - limit} 個" if len(values) > limit else "")


def _quoted(value: Any, limit: int = 30) -> str:
    text = " ".join(_text(value).split())
    return "「" + (text if len(text) <= limit else text[:limit] + "…") + "」"


def _overflow(area: Any, outer: Any) -> list[tuple[str, float]]:
    if not (
        isinstance(area, list)
        and isinstance(outer, list)
        and len(area) == len(outer) == 4
        and all(_number(value) for value in [*area, *outer])
    ):
        return []
    x, y, w, h = area
    left, top, width, height = outer
    sides = [("左", left - x), ("上", top - y), ("右", x + w - left - width), ("下", y + h - top - height)]
    return [(side, amount) for side, amount in sides if amount >= 0.05]


_SIDE_NAMES = {"left": "左", "top": "上", "right": "右", "bottom": "下"}


def _sizing_parts(item: dict) -> list[str]:
    """文字のはみ出しの観測値から、折り返した行数・高さの不足・1 行に収めるのに要る幅を示す文。"""
    parts = []
    lines, hard = item.get("lines"), item.get("hard_lines")
    single, shortfall = item.get("single_line_width_pt"), item.get("height_shortfall_pt")
    if _integer(lines) and _integer(hard) and lines > hard:
        parts.append(f"{lines} 行に折り返した" + (f"（明示改行で {hard} 行）" if hard > 1 else ""))
    if _number(shortfall):
        parts.append(f"高さが {shortfall:.1f} pt 足りない")
    if _number(single):
        parts.append(
            f"{'1 行' if not _integer(hard) or hard == 1 else '明示改行の行ごとに 1 行'}に収めるには内側の幅 {single:.1f} pt が要る"
        )
    return parts


def _observation_text(issue: dict) -> str:
    """要約で添える観測値の要点（括弧付き）。対応のない識別子は空にする。"""
    code = issue.get("code")
    if code == "text_overflow" and (item := _first(issue, "required_width_pt", "available_width_pt")):
        width = (item["required_width_pt"], item["available_width_pt"])
        height = (item.get("required_height_pt"), item.get("available_height_pt"))
        if all(_number(value) for value in (*width, *height)):
            parts = [f"必要な寸法 {width[0]:.1f}×{height[0]:.1f} pt、使える寸法 {width[1]:.1f}×{height[1]:.1f} pt"]
            return "（" + "。".join(parts + _sizing_parts(item)) + "）"
        if all(_number(value) for value in width):
            return f"（必要な幅 {width[0]:.1f} pt、使える幅 {width[1]:.1f} pt）"
    if code == "text_padding_overlap" and (item := _first(issue, "overflow_pt")):
        side = item["overflow_pt"].get("side") if isinstance(item["overflow_pt"], dict) else None
        amount = item["overflow_pt"].get("amount") if isinstance(item["overflow_pt"], dict) else None
        if side in _SIDE_NAMES and _number(amount):
            return "（" + "。".join([f"{_SIDE_NAMES[side]}の余白に {amount:.1f} pt 入る", *_sizing_parts(item)]) + "）"
    if code == "text_short_line" and (values := _gathered(issue, "short_texts")):
        return "（最後の行: " + "、".join(_quoted(value, 12) for value in values[:5]) + "）"
    if code in {"text_wrap_mid_phrase", "text_wrap_hyphen"} and (values := _gathered(issue, "breaks")):
        return "（" + "、".join(_quoted(value, 24) for value in values[:5]) + "）"
    if code in {"page_overflow", "content_area_overflow"} and (item := _first(issue, "area")):
        outer, name = (
            (item.get("page"), "ページ") if code == "page_overflow" else (item.get("content_area"), "本文領域")
        )
        sides = _overflow(item["area"], outer)
        if sides:
            return f"（{name}から" + "、".join(f"{side}へ {amount:.1f} pt" for side, amount in sides) + " はみ出し）"
    if code == "layout_bounds_overflow" and (item := _first(issue, "overflow_pt")):
        amount = item["overflow_pt"].get("amount") if isinstance(item["overflow_pt"], dict) else None
        side = item["overflow_pt"].get("side") if isinstance(item["overflow_pt"], dict) else None
        parts = [f"{_SIDE_NAMES.get(side, '')}へ {amount:.1f} pt はみ出し"] if _number(amount) else []
        total, members = item.get("stack_overflow_pt"), _issues(item.get("stack_elements"))
        if _number(total) and total > 0 and len(members) > 1:
            listed = "、".join(
                f"{_text(member.get('id'))} {member['height_pt']:.1f}"
                for member in members[:6]
                if _number(member.get("height_pt"))
            )
            parts.append(
                f"積み上げの全体で {total:.1f} pt 足りない（高さ pt: {listed}"
                + (" ほか" if len(members) > 6 else "")
                + "）"
            )
        if parts:
            return "（" + "。".join(parts) + "）"
    if code == "math_not_marked" and (values := _gathered(issue, "candidates")):
        return f"（候補: {_listed(values)}）"
    if code in {"font_glyph_missing", "forbidden_character"} and (values := _gathered(issue, "characters")):
        return f"（文字: {_listed(values)}）"
    if code == "pdf_element_text_overflow" and (item := _first(issue, "text")):
        return f"（文字: {_quoted(item['text'])}）"
    if code in _FILE_CODES and isinstance(issue.get("path"), str):
        return f"（{_shown(issue['path'])}）"
    return ""


def _page_label(issue: dict) -> str | None:
    """診断のページ。番号がなく識別子だけがあるとき（選んだ章だけを構築した検査）は識別子で示す。"""
    if (number := _page(issue)) is not None:
        return f"p.{number}"
    slide_id = issue.get("slide_id")
    return f"ページ {slide_id}" if isinstance(slide_id, str) and slide_id else None


def _issue_line(issue: dict) -> str:
    parts = []
    if (page := _page_label(issue)) is not None:
        parts.append(page)
    parts.append(SEVERITY_LABELS.get(issue.get("severity"), _text(issue.get("severity"))))
    parts.append(_text(issue.get("code")))
    element, related = issue.get("element_id"), issue.get("related_element_id")
    if element is not None and related is not None:
        parts.append(f"{element} / {related}")
    elif element is not None or related is not None:
        parts.append(_text(element if element is not None else related))
    parts.append(_text(issue.get("message")))
    return (
        " ".join(part for part in parts if part)
        + _observation_text(issue)
        + ("（宣言で許容）" if _declared(issue) else "")
    )


def _failure_lines(issue: dict, run: Any) -> list[str] | None:
    """例外の診断（errors.failure_diagnostic）。場所、その行のコード、トレースバックのファイルを示す。"""
    item = _first(issue, "location")
    if item is None or not isinstance(item["location"], dict):
        return None
    location = item["location"]
    parts = []
    if (page := _page_label(issue)) is not None:
        parts.append(page)
    parts.append(_text(issue.get("code")))
    where = location.get("relative_path") or location.get("path")
    if where:
        place = f"{where}:{location['line']}" if _integer(location.get("line")) else _text(where)
        notes = [_text(location.get("function"))] if location.get("function") else []
        if location.get("in_project") is False:
            notes.append("slidekit の内部で発生")
        parts.append(place + (f"（{'、'.join(notes)}）" if notes else ""))
    exception, message = item.get("exception"), item.get("message")
    if exception and message:
        parts.append(f"{exception}: {message}")
    else:
        parts.append(_text(exception or message or issue.get("message")))
    lines = ["  " + " ".join(part for part in parts if part)]
    if location.get("code"):
        lines.append(f"      {location['code']}")
    if item.get("traceback"):
        trace = Path(str(run)) / str(item["traceback"]) if run else item["traceback"]
        lines.append(f"  トレースバック: {_shown(trace)}")
    return lines


def _issue_lines(issue: dict, run: Any) -> list[str]:
    return _failure_lines(issue, run) or ["  " + _issue_line(issue)]


def _issue_sections(
    issues: list[dict],
    limit: int,
    *,
    run: Any = None,
    rest: str = "--json で全件を表示",
    collapse_info: bool = False,
) -> list[str]:
    """エラー、警告と参考の一覧。宣言で許容した警告（collapse_info なら参考も）は、verbose でなければ件数だけを示す。"""
    ordered = sort_issues(issues)
    errors = [issue for issue in ordered if issue.get("severity") == "error"]
    others = [issue for issue in ordered if issue.get("severity") != "error"]
    folded: list[str] = []
    if not _option("verbose"):
        declared = [issue for issue in others if _declared(issue)]
        info = [issue for issue in others if collapse_info and issue.get("severity") == "info"]
        others = [issue for issue in others if not _declared(issue) and not (collapse_info and issue in info)]
        if declared:
            folded.append(f"宣言で許容した警告 {len(declared)} 件")
        if info:
            folded.append(f"参考 {len(info)} 件")
    lines: list[str] = []
    if errors:
        lines.append("エラー:")
        for issue in errors[:MAX_ERRORS]:
            lines.extend(_issue_lines(issue, run))
        if len(errors) > MAX_ERRORS:
            lines.append(f"  ほか {len(errors) - MAX_ERRORS} 件（{rest}）")
    if others:
        severities = {issue.get("severity") for issue in others}
        lines.append("警告:" if severities == {"warning"} else "参考:" if severities == {"info"} else "警告と参考:")
        for issue in others[:limit]:
            lines.extend(_issue_lines(issue, run))
        if len(others) > limit:
            lines.append(f"  ほか {len(others) - limit} 件（{rest}）")
    if folded:
        lines.append("、".join(folded) + "（一覧は --verbose）")
    return lines


# ---------------------------------------------------------------- 共通の行


def _state(result: dict, issues: list[dict]) -> str:
    """配置の検査・事前確認の状態。エラーの診断は不合格、入力・環境・変換の失敗（終了コード 2）は失敗とする。"""
    status = result.get("status")
    counts = _counts_text(_counts(result, issues), issues)
    if status == "passed":
        return f"合格（{counts}）"
    if status == "failed":
        return f"{'失敗' if result.get('exit_code') == 2 else '不合格'}（{counts}）"
    return STATUS_LABELS.get(status, _text(status) or "状態不明") + f"（{counts}）"


def _pages(result: dict) -> str:
    selection = result.get("selection")
    if isinstance(selection, dict) and isinstance(selection.get("numbers"), list):
        numbers = [number for number in selection["numbers"] if _integer(number)]
        total = selection.get("total_slides")
        prefix = f"{total} ページ中 " if _integer(total) else ""
        return f"{prefix}{len(selection['numbers'])} ページ（{format_numbers(numbers)}）"
    if isinstance(selection, dict) and isinstance(selection.get("slide_ids"), list):
        # 選んだ章だけを構築したので、元のページ番号は分からない。識別子で示す。
        ids = [_text(value) for value in selection["slide_ids"]]
        listed = "、".join(ids[:5]) + (f" ほか {len(ids) - 5} ページ" if len(ids) > 5 else "")
        return f"{len(ids)} ページ（{listed}。選んだ章だけを構築）"
    return f"{result['slides']} ページ" if _integer(result.get("slides")) else ""


def _project_line(result: dict) -> str | None:
    if not result.get("project"):
        return None
    notes = [_pages(result)]
    if result.get("fallback_required"):
        notes.append("互換画像あり")
    notes = [note for note in notes if note]
    return f"資料: {_shown(result['project'])}" + (f"（{'、'.join(notes)}）" if notes else "")


def _timing_line(result: dict) -> str | None:
    timings = result.get("timings")
    groups: dict[str, float] = {}
    for item in timings if isinstance(timings, list) else []:
        if not isinstance(item, dict) or not _number(item.get("seconds")):
            continue
        name = _text(item.get("stage"))
        group = _TIMING_GROUPS.get(name.split(":")[0]) or _text(item.get("label")) or stage_label(name)
        groups[group] = groups.get(group, 0.0) + item["seconds"]
    elapsed = result.get("elapsed_seconds")
    if not _number(elapsed):
        if not groups:
            return None
        elapsed = sum(groups.values())
    shown = [
        (name, seconds) for name, seconds in sorted(groups.items(), key=lambda pair: -pair[1])[:4] if seconds >= 0.05
    ]
    parts = [f"{name} {seconds:.1f}" for name, seconds in shown]
    rest = elapsed - sum(seconds for _, seconds in shown)
    if shown and rest >= 0.05:
        parts.append(f"その他 {rest:.1f}")
    return f"所要時間: {elapsed:.1f} 秒" + (f"（{'、'.join(parts)}）" if parts else "")


def _compact(lines: list[str | None]) -> list[str]:
    return [line for line in lines if line]


# ---------------------------------------------------------------- コマンドごとの要約


def _layout(result: dict, limit: int) -> list[str]:
    issues = _issues(result.get("issues"))
    project = result.get("project")
    if result.get("exit_code") == 2:
        operation = _failed_operation("layout", issues)
    elif result.get("status") == "passed":
        operation = f"slidekit build --draft --project {_argument(project)}" if project else "slidekit build --draft"
    else:
        operation = _fix_operation(project, issues)
    return _compact(
        [
            f"配置の検査: {_state(result, issues)}",
            _project_line(result),
            *_issue_sections(issues, limit),
            *(_wrap_listing(result) if _option("wraps") else []),
            _timing_line(result),
            *_next([operation]),
        ]
    )


def _font_line(environment: dict) -> str | None:
    fonts = environment.get("fonts")
    fonts = fonts.get("fonts") if isinstance(fonts, dict) else None
    if not isinstance(fonts, dict) or not fonts:
        return None
    names, expected = [], set()
    for family, info in fonts.items():
        bold = isinstance(info, dict) and bool(info.get("bold"))
        names.append(f"{family}（通常・太字）" if bold else _text(family))
        expected.update({(family, "regular"), *([(family, "bold")] if bold else [])})
    found = set()
    renderers = environment.get("renderers")
    for info in renderers.values() if isinstance(renderers, dict) else []:
        for match in info.get("font_matches") or [] if isinstance(info, dict) else []:
            if isinstance(match, dict):
                found.add((match.get("family"), match.get("variant")))
    return "フォント: " + "、".join(names) + ("。Windows の登録書体と一致" if found and expected <= found else "")


def _renderer_line(result: dict, environment: dict) -> str | None:
    detected = environment.get("renderers") if isinstance(environment.get("renderers"), dict) else {}
    reported = result.get("renderers") if isinstance(result.get("renderers"), dict) else {}
    parts = []
    for name in [*reported, *(name for name in detected if name not in reported)]:
        info = {**(detected.get(name) or {}), **(reported.get(name) or {})}
        label = RENDERER_LABELS.get(name, _text(name))
        if info.get("status") == "failed" or info.get("ok") is False:
            parts.append(f"{label} を確認できない" + (f"（{info['message']}）" if info.get("message") else ""))
            continue
        text = " ".join(_text(part) for part in (label, info.get("version")) if part)
        parts.append(text + (f"（{info['executable']}）" if info.get("executable") else ""))
    return "表示ソフト: " + "、".join(parts) if parts else None


def _doctor(result: dict, limit: int) -> list[str]:
    issues = _issues(result.get("issues"))
    environment = result.get("environment") if isinstance(result.get("environment"), dict) else {}
    project = result.get("project")
    if result.get("status") == "passed":
        operation = f"slidekit build --project {_argument(project)}" if project else "slidekit build"
    elif result.get("exit_code") == 2:
        operation = _failed_operation("doctor", issues)
    else:
        operation = _fix_operation(project, issues)
    return _compact(
        [
            f"事前確認: {_state(result, issues)}",
            _project_line(result),
            _font_line(environment),
            _renderer_line(result, environment),
            *_issue_sections(issues, limit),
            _timing_line(result),
            *_next([operation]),
        ]
    )


def _short_stage(name: Any) -> str:
    """失敗した段階の短い名前（状態の行の括弧に入れるので、括弧を含めない）。"""
    name = _text(name)
    kind, _, rest = name.partition(":")
    if kind == "render" and rest:
        renderer = rest.partition(":")[0]
        return f"{RENDERER_LABELS.get(renderer, renderer)} の変換"
    if kind == "pdf":
        return "PDF の検査と画像化"
    if kind == "doctor":
        return "表示ソフトの事前確認"
    return stage_label(name)


def _build(result: dict, limit: int) -> list[str]:
    issues = _issues(result.get("issues"))
    run, project = result.get("build"), result.get("project")
    status, stage = result.get("status"), result.get("stage")
    counts = _counts_text(_counts(result, issues), issues)
    # 0.1 の結果は、エラーの診断がある完了を status "failed" で返していた（終了コード 1）。
    finished = status == "built" or (status == "failed" and stage == "complete" and result.get("exit_code") != 2)
    if status == "stopped":
        state = f"変換の前に停止（{counts}）"
    elif finished:
        state = f"完了（{counts}）"
    elif status == "failed":
        state = f"失敗（{_short_stage(stage)}）" if stage else "失敗"
    else:
        state = STATUS_LABELS.get(status, _text(status) or "状態不明") + f"（{counts}）"
    notes = [_pages(result)]
    if result.get("mode") == "draft":
        notes.append("下書き")
    if result.get("fallback"):
        notes.append("互換画像あり")
    notes = [note for note in notes if note]
    target = f"実行: {_shown(run)}" + (f"（{'、'.join(notes)}）" if notes else "") if run else None
    images = result.get("images")
    paths = (
        [path for values in images.values() if isinstance(values, list) for path in values if isinstance(path, str)]
        if isinstance(images, dict)
        else []
    )
    picture = None
    if finished and paths:
        picture = f"画像: {_shown(Path(str(run)) / paths[0] if run else paths[0])}" + (
            f" ほか {len(paths) - 1} 枚" if len(paths) > 1 else ""
        )
    carry = _carry_line(result.get("carry"))
    if status == "failed" and not finished:
        codes = [issue.get("code") for issue in issues if issue.get("severity") == "error"]
        operations = next(
            (list(_RENDERER_OPERATIONS[code]) for code in codes if code in _RENDERER_OPERATIONS),
            [_failed_operation("build", issues)],
        )
    elif status == "stopped" or any(issue.get("severity") == "error" for issue in issues):
        operations = [_fix_operation(project, issues)]
    elif result.get("mode") == "draft" or result.get("selection") or result.get("exportable") is False:
        operations = [
            "表示を確かめ、全ページを slidekit build"
            + (f" --project {_argument(project)}" if project else "")
            + " で生成する"
        ]
    else:
        operations = [f"slidekit review status --build {_argument(run)}" if run else "slidekit review status"]
    return _compact(
        [
            f"生成: {state}",
            target,
            *_issue_sections(issues, limit, run=run, rest="diagnostics.json", collapse_info=True),
            _timing_line(result),
            _comparison_line(result.get("comparison")),
            *carry,
            picture,
            *_next(operations),
        ]
    )


def _remaining_text(remaining: Any) -> str:
    if not isinstance(remaining, dict):
        return ""
    parts = [f"{label} {remaining[key]}" for key, label in _REMAINING_NAMES if _integer(remaining.get(key))]
    return "、".join(parts)


def _distinct(values: Any) -> list[str]:
    """ページの識別子の列から重複を除く（引き継ぎは記録ごとにページを挙げ、数式のページは標準表示と互換表示の 2 件になる）。"""
    return list(dict.fromkeys(_text(value) for value in values)) if isinstance(values, list) else []


def _run_names(values: Any) -> str:
    """{実行のパスか名前: 件数} を「0133 29 件、0120 2 件」の形にする。"""
    if not isinstance(values, dict):
        return ""
    return "、".join(f"{Path(_text(name)).name} {count} 件" for name, count in values.items() if _integer(count))


_KIND_NAMES = {"visual": "表示", "content": "内容"}


def _reason_text(reasons: Any) -> list[str]:
    """引き継げなかった確認の理由を、種類ごとに「表示: ページの表示の変更 2、表示環境の変更 1」の形で示す。"""
    grouped: dict[str, dict[str, int]] = {}
    documents: dict[str, list[str]] = {}
    for item in _issues(reasons):
        kind = _KIND_NAMES.get(item.get("kind"))
        reason = _text(item.get("reason"))
        if not kind or not reason:
            continue
        counts = grouped.setdefault(kind, {})
        counts[reason] = counts.get(reason, 0) + 1
        for name in item.get("documents") or []:
            if _text(name) not in documents.setdefault(kind, []):
                documents[kind].append(_text(name))
    lines = []
    for kind, counts in grouped.items():
        parts = [
            f"{reason} {count}"
            + (f"（{'、'.join(documents[kind])}）" if reason == "元資料の変更" and documents.get(kind) else "")
            for reason, count in counts.items()
        ]
        lines.append(f"{kind}: {'、'.join(parts)}")
    return lines


def _change_text(change: dict) -> str:
    before, after = change.get("before"), change.get("after")
    return f"{_text(change.get('item'))} {_text(before) or 'なし'} → {_text(after) or 'なし'}"


def _comparison_line(comparison: Any) -> str | None:
    """build の結果の comparison（前の実行とのページ画像の比較）の 1 行。"""
    if not isinstance(comparison, dict) or not comparison.get("run"):
        return None
    numbers = comparison.get("numbers") if isinstance(comparison.get("numbers"), dict) else {}

    def pages(key: str) -> str:
        values = comparison.get(key) if isinstance(comparison.get(key), list) else []
        known = [numbers[value] for value in values if _integer(numbers.get(value))]
        listed = format_numbers(known) if len(known) == len(values) else _listed_ids(values)
        return f"{len(values)}" + (
            f"（p.{listed}）" if values and len(known) == len(values) and key != "unchanged" else ""
        )

    parts = [
        f"変更なし {pages('unchanged')}",
        f"ページ番号だけ変わった {pages('renumbered')}",
        f"変わった {pages('changed')}",
    ]
    if comparison.get("added"):
        parts.append(f"新しいページ {pages('added')}")
    if comparison.get("removed"):
        parts.append(f"外したページ {len(comparison['removed'])}")
    return f"前の実行 {_text(comparison['run'])} との比較: " + "、".join(parts)


def _carry_line(carry: Any) -> list[str]:
    """生成の後の確認の引き継ぎ（build の結果の carry）。1 行目に件数、続く行に比べた実行・理由・表示環境の違いを示す。"""
    if not isinstance(carry, dict):
        return []
    status = carry.get("status")
    if status == "carried":
        count = carry.get("count")
        text = f"確認の引き継ぎ: {count if _integer(count) else 0} 件"
        notes = []
        if sources := _run_names(carry.get("sources")):
            notes.append(f"引き継ぎ元 {sources}")
        if renumbered := _distinct(carry.get("renumbered")):
            compared = carry.get("renumbered_from") if isinstance(carry.get("renumbered_from"), dict) else {}
            runs = "・".join(Path(_text(name)).name for name in compared)
            notes.append(
                f"ページ番号だけ変わったページ {len(renumbered)} 件" + (f"（実行 {runs} と比較）" if runs else "")
            )
        if remaining := _remaining_text(carry.get("remaining")):
            notes.append(f"未確認 {remaining}")
        # 候補にしなかった実行（形式 1 の実行など）。初めて生成した 0.1 の資料で、前の実行を飛ばしたことを示す。
        skipped_runs = carry.get("skipped_runs") if isinstance(carry.get("skipped_runs"), dict) else {}
        if skipped_runs:
            notes.append(
                "候補にしなかった実行 " + "・".join(f"{_text(name)} {count} 件" for name, count in skipped_runs.items())
            )
        lines = [text + (f"（{'、'.join(notes)}）" if notes else "")]
        reasons = _reason_text(carry.get("pending_reasons"))
        skipped = carry.get("skipped") if isinstance(carry.get("skipped"), dict) else {}
        if reasons or skipped:
            lines.append(
                "  引き継がなかった理由: "
                + "。".join(reasons + [f"{_text(reason)} {count}" for reason, count in skipped.items()])
            )
        changes = _issues(carry.get("environment_changes"))
        if changes:
            compared = carry.get("compared_run")
            shown = "、".join(_change_text(change) for change in changes[:6])
            more = f" ほか {len(changes) - 6} 件" if len(changes) > 6 else ""
            lines.append(
                f"  表示環境の変更" + (f"（実行 {_text(compared)} から）" if compared else "") + f": {shown}{more}"
            )
        return lines
    label = STATUS_LABELS.get(status, _text(status) or "状態不明")
    return [f"確認の引き継ぎ: {label}" + (f"（{carry['message']}）" if carry.get("message") else "")]


def _run_of(result: dict) -> Any:
    return result.get("build") or result.get("run") or result.get("run_id")


def _review_target(target: Any) -> str:
    target = target if isinstance(target, dict) else {}
    kind = _REVIEW_KINDS.get(target.get("kind"), _text(target.get("kind")) or "確認")
    details = [_text(target.get("renderer")), _VARIANT_NAMES.get(target.get("variant"), _text(target.get("variant")))]
    details = [detail for detail in details if detail]
    return kind + (f"（{'・'.join(details)}）" if details else "")


def _check_body(result: dict, limit: int, *, run: Any = None) -> tuple[list[str], dict]:
    """検査の結果の本体（状態の行を除く）と、件数。"""
    issues = _issues(result.get("issues"))
    warnings = _issues(result.get("warnings"))
    missing = [issue for issue in issues if issue.get("code") == "review_missing"]
    unaccepted = [issue for issue in issues if issue.get("code") == "warning_unaccepted"]
    errors = [issue for issue in issues if issue.get("code") not in {"review_missing", "warning_unaccepted"}]
    lines: list[str] = []
    # 生成記録の照合の誤りは識別子ごとにまとめる（入力の変更は多数のファイルにわたることがある）。
    grouped: dict[str, list[dict]] = {}
    for issue in errors:
        if issue.get("code") in _FILE_CODES:
            grouped.setdefault(issue["code"], []).append(issue)
    others = [issue for issue in errors if issue.get("code") not in _FILE_CODES]
    if grouped:
        lines.append("生成記録との不一致:")
        for value, group in grouped.items():
            paths = [_shown(issue["path"]) for issue in group if isinstance(issue.get("path"), str)]
            listed = (
                (paths[0] + (f" ほか {len(paths) - 1} 件" if len(paths) > 1 else "")) if paths else f"{len(group)} 件"
            )
            kinds: dict[str, list[str]] = {}
            for issue in group:
                if issue.get("input_kind") and isinstance(issue.get("path"), str):
                    kinds.setdefault(_text(issue["input_kind"]), []).append(_shown(issue["path"]))
            if kinds:
                listed = "、".join(
                    f"{kind}: {Path(found[0]).name}" + (f" ほか {len(found) - 1} 件" if len(found) > 1 else "")
                    for kind, found in kinds.items()
                )
            lines.append(f"  {value} {_text(group[0].get('message'))}（{listed}）")
    lines.extend(_issue_sections(others, limit, run=run))
    if missing:
        lines.append("未完了の確認:")
        targets: dict[str, list[dict]] = {}
        for issue in missing:
            targets.setdefault(_review_target(issue.get("target")), []).append(issue.get("target") or {})
        numbers: dict[str, list[int]] = {}
        for issue in missing:
            if _integer(issue.get("slide_number")):
                numbers.setdefault(_review_target(issue.get("target")), []).append(issue["slide_number"])
        for name, group in targets.items():
            first = next(
                (target.get("slide_id") for target in group if isinstance(target, dict) and target.get("slide_id")),
                None,
            )
            if numbers.get(name):
                lines.append(f"  {name} {len(group)} 件（p.{format_numbers(numbers[name])}）")
            else:
                lines.append(f"  {name} {len(group)} 件" + (f"（先頭: {first}）" if first else ""))
    if unaccepted:
        lines.append("未対応の警告:")
        known = {warning.get("diagnostic_id"): warning for warning in warnings}
        codes: dict[str, list[dict]] = {}
        for issue in unaccepted:
            warning = known.get(issue.get("diagnostic_id"), {})
            codes.setdefault(_text(warning.get("code")) or "不明な警告", []).append(warning)
        for name, group in codes.items():
            pages = format_numbers(number for warning in group if (number := _page(warning)) is not None)
            lines.append(f"  {name} {len(group)} 件" + (f"（p.{pages}）" if pages else ""))
    elif warnings and not result.get("final"):
        lines.extend(_issue_sections(warnings, limit, run=run))
    invalid = _issues(result.get("invalid_records"))
    if invalid:
        first = invalid[0]
        lines.append(
            f"無効な確認記録: {len(invalid)} 件（{_text(first.get('record_id'))}: {_text(first.get('reason'))}）"
        )
    counts = {"errors": len(errors), "warnings": len(warnings), "missing": len(missing), "unaccepted": len(unaccepted)}
    return lines, counts


def _declared_lines(entries: Any, limit: int, *, indent: str = "") -> list[str]:
    """資料コードの宣言で許容した警告を、ページ・種類・要素・理由・宣言の場所で並べる。"""
    declared = _issues(entries)
    if not declared:
        return []
    if not _option("verbose"):
        return [f"{indent}宣言で許容した警告: {len(declared)} 件（一覧は --verbose）"]
    lines = [f"{indent}宣言で許容した警告:"]
    for entry in declared[:limit]:
        parts = [page] if (page := _page_label(entry)) is not None else []
        parts.append(_text(entry.get("code")))
        element, related = entry.get("element_id"), entry.get("related_element_id")
        if element is not None or related is not None:
            parts.append(
                f"{element} / {related}"
                if element is not None and related is not None
                else _text(element if element is not None else related)
            )
        notes = [f"理由: {_text(entry.get('reason'))}"] if entry.get("reason") else []
        if entry.get("location"):
            notes.append(f"宣言: {_text(entry['location'])}")
        lines.append(
            f"{indent}  " + " ".join(part for part in parts if part) + (f"（{'、'.join(notes)}）" if notes else "")
        )
    if len(declared) > limit:
        lines.append(f"{indent}  ほか {len(declared) - limit} 件（--json で全件を表示）")
    return lines


def _blocked_operation(codes: set) -> str:
    """確認を加えても通らない問題の直し方。識別子で分ける。"""
    for names, operation in _BLOCKED_OPERATIONS:
        if codes & set(names):
            return operation
    return "slidekit build で生成し直す"


def _check(result: dict, limit: int) -> list[str]:
    run = _run_of(result)
    body, counts = _check_body(result, limit, run=run)
    status, final = result.get("status"), bool(result.get("final"))
    label = "最終検査" if final else "検査"
    numbers = [f"エラー {counts['errors']} 件", f"警告 {counts['warnings']} 件"]
    if final:
        numbers += [f"未完了の確認 {counts['missing']} 件", f"未対応の警告 {counts['unaccepted']} 件"]
    if status == "passed":
        state = "合格"
    elif status == "failed":
        state = "失敗" if result.get("exit_code") == 2 else "不合格"
    else:
        state = STATUS_LABELS.get(status, _text(status) or "状態不明")
    shown = _argument(run) if run else None
    if final:
        body.extend(_declared_lines(result.get("declared_warnings"), limit))
    if status == "passed":
        operation = (
            (f"slidekit export --build {shown}" if final else f"slidekit review status --build {shown}")
            if shown
            else ""
        )
    elif counts["errors"] and result.get("exit_code") == 2:
        operation = _blocked_operation({issue.get("code") for issue in _issues(result.get("issues"))})
    elif counts["errors"]:
        operation = "編集元を直して slidekit build で生成し直す"
    else:
        operation = f"slidekit review status --build {shown}" if shown else "slidekit review status"
    return _compact(
        [
            f"{label}: {state}（{'、'.join(numbers)}）",
            f"実行: {_shown(run)}" if run else None,
            *body,
            *_next([operation]),
        ]
    )


def _export(result: dict, limit: int) -> list[str]:
    status = result.get("status")
    check = result.get("check") if isinstance(result.get("check"), dict) else None
    run = _run_of(result) or (_run_of(check) if check else None)
    if status == "exported":
        deliverables = result.get("deliverables") if isinstance(result.get("deliverables"), dict) else {}
        files = [_shown(deliverables[key]) for key in ("pptx", "pdf", "notes") if deliverables.get(key)]
        record = result.get("record") or deliverables.get("record")
        return _compact(
            [
                "反映: 完了",
                f"実行: {_shown(run)}" if run else None,
                "完成物: " + "、".join(files) if files else None,
                f"採用記録: {_shown(record)}" if record else None,
                f"最終検査: {_shown(result['final_check'])}" if result.get("final_check") else None,
                f"旧版の退避先: {_shown(result['history'])}" if result.get("history") else None,
            ]
        )
    if check is not None:
        body, _ = _check_body(check, limit, run=run)
        body.extend(_declared_lines(check.get("declared_warnings"), limit))
        codes = {issue.get("code") for issue in _issues(check.get("issues"))}
        operations = [item for item in result.get("next") or [] if isinstance(item, str)]
        if result.get("exit_code") == 2 and codes & {"artifact_unrecorded", "artifact_changed", "input_changed"}:
            operations = [_blocked_operation(codes)]
        elif not operations:
            operations = [f"slidekit review status --build {_argument(run)}" if run else "slidekit review status"]
        lines = ["反映: 中止（最終検査に不合格）", f"実行: {_shown(run)}" if run else None, *body]
        # 足りない確認を記録するコマンドを並べる（review status と同じ。確認者と確認内容は画像を見てから書き換える）。
        if len(operations) == 1:
            lines.append(f"次の操作: {operations[0]}")
        else:
            lines.append("次の操作:")
            lines.extend(f"  {operation}" for operation in operations[: max(limit, 1)])
        return _compact(lines)
    return []


def _size(value: Any) -> str:
    if not _number(value):
        return ""
    for unit, scale in (("GB", 1e9), ("MB", 1e6), ("KB", 1e3)):
        if value >= scale:
            return f"{value / scale:.1f} {unit}"
    return f"{int(value)} B"


def _run_name(value: Any) -> str:
    return Path(_text(value)).name if value else ""


def _clean(result: dict, limit: int) -> list[str]:
    status = result.get("status")
    candidates = (
        [
            item if isinstance(item, dict) else {"run": item}
            for item in result.get("candidates") or []
            if isinstance(item, (dict, str))
        ]
        if isinstance(result.get("candidates"), list)
        else []
    )
    removed = result.get("removed") if isinstance(result.get("removed"), list) else []
    protected = _issues(result.get("protected"))
    total = result.get("bytes")
    if not _number(total):
        sizes = [item.get("bytes") for item in candidates if _number(item.get("bytes"))]
        total = sum(sizes) if sizes else None
    size = f"（{_size(total)}）" if _size(total) else ""
    if status == "cleaned":
        lines = [f"整理: {len(removed)} 件を削除{size}"]
    else:
        lines = [f"整理: 削除の候補 {len(candidates)} 件{size}。まだ削除していません"]
    for item in candidates[:limit]:
        # 生成記録のない実行（中断した生成など）は、状態と種類の代わりに「生成記録なし」と示す。
        states = [item.get("status"), item.get("mode")]
        if "status" in item and not any(states):
            states = ["生成記録なし"]
        fields = [
            item.get("run_id") or _run_name(item.get("run")),
            *states,
            _size(item.get("bytes")),
            item.get("reason"),
        ]
        lines.append("  " + "  ".join(_text(field) for field in fields if field))
    if len(candidates) > limit:
        lines.append(f"  ほか {len(candidates) - limit} 件（--json で全件を表示）")
    if protected:
        lines.append(f"保護: {len(protected)} 件")
        for item in protected[:limit]:
            reasons = item.get("reasons") if isinstance(item.get("reasons"), list) else []
            lines.append(f"  {_run_name(item.get('run'))}  {'、'.join(_text(reason) for reason in reasons)}".rstrip())
        if len(protected) > limit:
            lines.append(f"  ほか {len(protected) - limit} 件（--json で全件を表示）")
    if status != "cleaned" and candidates:
        lines.extend(_next(["同じ条件に --apply を付けると削除する"]))
    return lines


def _holder_text(holder: dict) -> str:
    inner = holder.get("holder") if isinstance(holder.get("holder"), dict) else {}
    values = {**holder, **inner}
    parts = []
    pid = inner.get("python_pid") or holder.get("pid")
    if pid:
        parts.append(f"PID {pid}")
    for key, label in (
        ("run", "実行"),
        ("project", "資料"),
        ("command", "コマンド"),
        ("host", "端末"),
        ("started_at", "開始"),
    ):
        if values.get(key):
            parts.append(f"{label} {_shown(values[key]) if key in {'run', 'project'} else _text(values[key])}")
    return "、".join(parts) or "情報なし"


_OWNED_RECORDS = {"pending": "残存記録", "active": "使用中の記録"}


def _dicts(value: Any) -> list[dict]:
    return [item for item in value if isinstance(item, dict)] if isinstance(value, list) else []


def _pids(items: list[dict]) -> str:
    return "、".join(_text(item.get("pid")) for item in items)


def _powerpoint_process_text(item: dict) -> str:
    """PowerPoint の処理の 1 行。PID・ウィンドウ・開始時刻と、slidekit が起動したと記録した PowerPoint と一致するか。"""
    title = _text(item.get("window_title"))
    window = f"題名「{title}」" if title else "題名のないウィンドウ" if item.get("has_window") else "ウィンドウなし"
    started = _text(item.get("started_at")) or "不明"
    if item.get("owned"):
        owner = f"slidekit が起動（{_OWNED_RECORDS.get(item.get('record'), '記録')}と一致）"
    else:
        owner = "slidekit が起動した記録と一致しない"
    return f"  PID {_text(item.get('pid'))}、{window}、開始 {started}、{owner}"


def _recorded_powerpoint(result: dict, pending: dict) -> tuple[str, bool]:
    """残存記録の変換に使った PowerPoint の説明と、それが slidekit の起動したもので、もう動いていないか。"""
    owned = pending.get("powerpoint")
    if not isinstance(owned, dict):
        return "記録なし（古い形式の記録か、PowerPoint を起動する前の記録）", False
    if owned.get("started") is not True:
        return "既に起動していた PowerPoint（slidekit は終了しない）", False
    record = next(
        (item for item in _dicts(result.get("powerpoint_records")) if item.get("source") == "pending"),
        None,
    )
    if record is None or record.get("pid") is None:
        return "slidekit が起動（PID の記録なし）", False
    if record.get("running"):
        return f"slidekit が起動（PID {_text(record['pid'])}、動作中）", False
    return f"slidekit が起動（PID {_text(record['pid'])}、終了済み）", True


def _renderer_status(result: dict, limit: int) -> list[str]:
    status = result.get("status")
    label = RENDERER_LABELS.get(result.get("renderer"), _text(result.get("renderer")) or "表示ソフト")
    lines = [f"表示ソフトの状態: {STATUS_LABELS.get(status, _text(status) or '状態不明')}（{label}）"]
    holder = result.get("holder")
    if isinstance(holder, dict) and holder:
        lines.append(f"使用中の処理: {_holder_text(holder)}")
    pending = result.get("pending")
    processes = _dicts(result.get("powerpoint_processes"))
    # renderer clear --stop-owned が終了する PowerPoint（slidekit が起動したと記録で確かめることができ、ウィンドウがない）。
    stoppable = [item for item in processes if item.get("stoppable")]
    stop_owned = f"slidekit renderer clear --stop-owned（slidekit が起動した PowerPoint（PID {_pids(stoppable)}）を終了"
    operations = []
    if isinstance(pending, dict) and pending:
        states = {"timeout": "時間切れ", "running": "変換中", "interrupted": "中断"}
        parts = [states.get(pending.get("status"), _text(pending.get("status")) or "状態不明")]
        if pending.get("run_id"):
            parts.append(f"記録 {pending['run_id']}")
        if pending.get("started_at"):
            parts.append(f"開始 {pending['started_at']}")
        lines.append(f"残存記録: {'、'.join(parts)}")
        alive, locked = result.get("pending_process_alive"), result.get("staged_input_locked")
        running = result.get("powerpoint_running")
        # 変換に使った PowerPoint が slidekit の起動したもので、もう動いていなければ、ほかの PowerPoint は複製を開いていない。
        recorded, gone = "", False
        if "powerpoint_records" in result:
            recorded, gone = _recorded_powerpoint(result, pending)
        if alive is not None:
            lines.append(f"  作業プロセス: {'動作中' if alive else '終了'}")
        if recorded:
            lines.append(f"  変換に使った PowerPoint: {recorded}")
        if locked is not None:
            # PowerPoint は読み取り専用で開いた複製のファイルを開いたままにしないので、排他で開くことができても
            # 開いていないとは言えない。
            shown = (
                "PowerPoint が開いている"
                if locked
                else "開いているかは確かめられない（PowerPoint が起動中）" if running and not gone else "開かれていない"
            )
            lines.append(f"  検査用の複製: {shown}")
        if alive:
            operations.append("変換の終了を待ち、slidekit renderer status で確かめる")
        elif stoppable:
            operations.append(stop_owned + "してから解除する）")
        elif locked:
            operations.append("PowerPoint で検査用の複製を閉じてから slidekit renderer clear を実行する")
        elif running and not gone:
            operations.append(
                "PowerPoint の終了を待つか、PowerPoint が検査用の複製を開いていないことを確かめてから"
                " slidekit renderer clear --force を実行する"
            )
        else:
            operations.append("slidekit renderer clear")
    elif "pending" in result:
        lines.append("残存記録: なし")
        if stoppable:
            # 前の処理が終了させずに残した PowerPoint。後続の変換は止めないが、残り続ける。
            lines.append(f"slidekit が起動した PowerPoint が残っている: PID {_pids(stoppable)}")
            operations.append(stop_owned + "する）")
    if result.get("powerpoint_running") is not None:
        count = f"（{len(processes)} 個）" if processes else ""
        lines.append(f"PowerPoint: {'起動中' + count if result['powerpoint_running'] else '起動していない'}")
        lines.extend(_powerpoint_process_text(item) for item in processes[:limit])
        if len(processes) > limit:
            lines.append(f"  ほか {len(processes) - limit} 個（--json で全件を表示）")
    if result.get("abandoned"):
        lines.append("排他: 前の処理が放さずに終了していた（slidekit が起動した PowerPoint が残っていることがある）")
    if result.get("state_dir"):
        lines.append(f"状態フォルダー: {result['state_dir']}")
    if status == "busy" and not operations:
        operations.append("変換の終了を待つ（slidekit renderer status で確かめる）")
    return lines + _next(operations)


def _renderer_clear(result: dict, limit: int) -> list[str]:
    status = result.get("status")
    head = f"残存記録の解除: {STATUS_LABELS.get(status, _text(status) or '状態不明')}"
    if result.get("reason"):
        head += f"（{result['reason']}）"
    lines = [head]
    for item in _dicts(result.get("stopped_powerpoint")):
        started = f"（開始 {_text(item['started_at'])}）" if item.get("started_at") else ""
        lines.append(f"終了した PowerPoint: PID {_text(item.get('pid'))}{started}")
    if result.get("archived"):
        lines.append(f"退避先: {result['archived']}")
    if result.get("removed_work_dir"):
        lines.append("作業フォルダー: 削除した")
    if status == "refused":
        lines.extend(_next(["slidekit renderer status で状態を確かめる"]))
    return lines


# ---------------------------------------------------------------- 確認記録の操作


def _group_name(kind: Any, renderer: Any = None, variant: Any = None) -> str:
    """確認の組（種類・表示ソフト・表示方法）の名前。内容の確認は表示ソフトによらない。"""
    name = _REVIEW_KINDS.get(kind, _text(kind) or "確認")
    if kind == "content":
        return name
    details = [_text(renderer), _VARIANT_NAMES.get(variant, _text(variant))]
    details = [detail for detail in details if detail]
    return name + (f"（{'・'.join(details)}）" if details else "")


def _paths_text(item: dict) -> str:
    """ファイルの一覧（paths の最初の数件と全体の件数 count）を「（最初のファイル ほか N 件）」にする。なければ空。"""
    paths = (
        [_shown(path) for path in item.get("paths") if isinstance(path, str)]
        if isinstance(item.get("paths"), list)
        else []
    )
    if not paths:
        return ""
    kinds = item.get("kinds") if isinstance(item.get("kinds"), dict) else None
    if kinds:
        # 変わった入力の種類ごとに、最初のファイルと件数を示す（どの入力が変わったかを推測しなくてよいように）。
        parts = []
        for kind, entry in kinds.items():
            entry = entry if isinstance(entry, dict) else {}
            first = Path(_text(entry.get("first"))).name if entry.get("first") else ""
            number = entry.get("count") if _integer(entry.get("count")) else 1
            parts.append(f"{_text(kind)}: {first}" + (f" ほか {number - 1} 件" if number > 1 else ""))
        return "（" + "、".join(parts) + "）"
    count = item.get("count") if _integer(item.get("count")) else len(paths)
    return f"（{paths[0]}" + (f" ほか {count - 1} 件" if count > 1 else "") + "）"


def _notice_lines(notices: Any) -> list[str]:
    """記録はできたが、その実行を反映できないことなどの知らせ。"""
    lines = []
    for notice in _issues(notices):
        where = _paths_text(notice) or (f"（{_text(notice['record_id'])}）" if notice.get("record_id") else "")
        lines.append(f"注意: {_text(notice.get('message'))}{where}")
    return lines


def _listed_ids(values: list, limit: int = 5) -> str:
    names = [_text(value) for value in values]
    return "、".join(names[:limit]) + (f" ほか {len(names) - limit} 件" if len(names) > limit else "")


def _status_operation(run: Any) -> str:
    return f"slidekit review status --build {_argument(run)}" if run else "slidekit review status"


def _pending_reasons(pending: list) -> str:
    """未確認のページを、引き継げなかった理由ごとに「。本文の変更 p.3,5、元資料の変更 p.6-7（proofs）」の形で示す。"""
    grouped: dict[str, list] = {}
    documents: dict[str, list[str]] = {}
    for item in pending:
        if not isinstance(item, dict) or not item.get("reason"):
            continue
        reason = _text(item["reason"])
        grouped.setdefault(reason, []).append(
            item.get("number") if _integer(item.get("number")) else item.get("slide_id")
        )
        for name in item.get("documents") or []:
            if _text(name) not in documents.setdefault(reason, []):
                documents[reason].append(_text(name))
    parts = []
    for reason, pages in grouped.items():
        numbers = [page for page in pages if _integer(page)]
        listed = f"p.{format_numbers(numbers)}" if len(numbers) == len(pages) else _listed_ids(pages)
        parts.append(f"{reason} {listed}" + (f"（{'、'.join(documents[reason])}）" if documents.get(reason) else ""))
    return "。" + "、".join(parts) if parts else ""


def _review_status(result: dict, limit: int) -> list[str]:
    status = result.get("status")
    notes = []
    if result.get("run_id"):
        notes.append(f"実行 {_text(result['run_id'])}")
    if _integer(result.get("slides")):
        notes.append(f"{result['slides']} ページ")
    if result.get("mode") == "draft":
        notes.append("下書き")
    elif result.get("exportable") is False:
        notes.append("部分生成")
    elif result.get("mode") == "full":
        notes.append("完全な生成")
    lines = [
        f"確認: {STATUS_LABELS.get(status, _text(status) or '状態不明')}" + (f"（{'、'.join(notes)}）" if notes else "")
    ]
    errors = _issues(result.get("errors"))
    warnings = result.get("warnings") if isinstance(result.get("warnings"), dict) else {}
    pending_warnings = _issues(warnings.get("pending"))
    counts = [
        f"{label} {warnings[key]}"
        for key, label in (("accepted", "許容"), ("declared", "宣言"))
        if _integer(warnings.get(key))
    ]
    counts.append(f"未対応 {len(pending_warnings)}")
    total = warnings.get("total") if _integer(warnings.get("total")) else len(pending_warnings)
    lines.append(f"  エラー {len(errors)} 件、警告 {total} 件（{'、'.join(counts)}）")
    groups = _issues(result.get("reviews"))
    numbers = {
        key: sum(group.get(key) for group in groups if _integer(group.get(key)))
        for key in ("required", "satisfied", "carried", "renumbered")
    }
    if groups:
        text = f"  確認 {numbers['required']} 件のうち {numbers['satisfied']} 件が有効"
        details = [f"引き継ぎ {numbers['carried']} 件"] if numbers["carried"] else []
        if numbers["renumbered"]:
            details.append(f"ページ番号だけ変わったもの {numbers['renumbered']} 件")
        lines.append(text + (f"（{'、'.join(details)}）" if details else ""))
    blocking = _issues(result.get("blocking"))
    if blocking:
        lines.append("  止めている問題:")
        for item in blocking:
            lines.append(f"    {_text(item.get('code'))} {_text(item.get('message'))}{_paths_text(item)}")
    if errors:
        lines.append("  エラー:")
        for issue in errors[:MAX_ERRORS]:
            lines.append("    " + _issue_line({**issue, "severity": None}))
        if len(errors) > MAX_ERRORS:
            lines.append(f"    ほか {len(errors) - MAX_ERRORS} 件（--json で全件を表示）")
    waiting = [group for group in groups if group.get("pending")]
    if waiting:
        lines.append("  未確認:")
        for group in waiting:
            name = _group_name(group.get("kind"), group.get("renderer"), group.get("variant"))
            pending = group["pending"] if isinstance(group["pending"], list) else []
            if group.get("kind") == "overview":
                lines.append(f"    {name} 未確認")
                continue
            pages = group.get("slides") or _listed_ids(
                [item.get("slide_id") for item in pending if isinstance(item, dict)]
            )
            lines.append(f"    {name} {pages}（{len(pending)} ページ{_pending_reasons(pending)}）")
    if pending_warnings:
        lines.append("  未対応の警告:")
        for warning in pending_warnings[:limit]:
            text = _issue_line({**warning, "severity": None, "message": None})
            lines.append(f"    {text}（{_text(warning.get('diagnostic_id'))}）")
        if len(pending_warnings) > limit:
            lines.append(f"    ほか {len(pending_warnings) - limit} 件（--json で全件を表示）")
    lines.extend(_declared_lines(result.get("declared_warnings"), limit, indent="  "))
    invalid = _issues(result.get("invalid_records"))
    if invalid:
        first = invalid[0]
        lines.append(
            f"  無効な確認記録: {len(invalid)} 件（{_text(first.get('record_id'))}: {_text(first.get('reason'))}）"
        )
    # 次の操作は確認の組ごとの記録と警告ごとの許容なので、3 件で打ち切らず limit 件まで示す。
    operations = (
        [item for item in result.get("next") or [] if isinstance(item, str)]
        if isinstance(result.get("next"), list)
        else []
    )
    if len(operations) == 1:
        lines.append(f"次の操作: {operations[0]}")
    elif operations:
        lines.append("次の操作:")
        lines.extend(f"  {operation}" for operation in operations[: max(limit, 1)])
        if len(operations) > max(limit, 1):
            lines.append(f"  ほか {len(operations) - max(limit, 1)} 件（--json で全件を表示）")
    return lines


def _review_record(result: dict, limit: int) -> list[str]:
    records = _issues(result.get("records"))
    run = _run_of(result)
    head = f"確認の記録: {STATUS_LABELS.get(result.get('status'), _text(result.get('status')) or '状態不明')}（{len(records)} 件"
    head += f"、{result['batch']}）" if result.get("batch") else "）"
    lines = [head]
    if run:
        lines.append(f"実行: {_shown(run)}")
    targets = result.get("targets") if isinstance(result.get("targets"), dict) else {}
    for key, values in targets.items():
        kind, _, rest = _text(key).partition(":")
        renderer, _, variant = rest.partition(":")
        values = values if isinstance(values, list) else []
        pages = [value for value in values if value is not None]
        lines.append(
            f"  {_group_name(kind, renderer or None, variant or None)} {len(values)} 件"
            + (f"（{_listed_ids(pages)}）" if pages else "")
        )
    remaining = result.get("remaining") if isinstance(result.get("remaining"), dict) else None
    operation = _status_operation(run)
    if remaining is not None:
        lines.append(f"残り: {_remaining_text(remaining) or 'なし'}")
        if not any(remaining.get(key) for key, _ in _REMAINING_NAMES):
            operation = f"slidekit export --build {_argument(run)}" if run else "slidekit export"
    return lines + _notice_lines(result.get("notices")) + _next([operation])


def _review_sheet(result: dict, limit: int) -> list[str]:
    sheets = [item for item in result.get("sheets") or [] if isinstance(item, str)]
    run = _run_of(result)
    changes = result.get("changes") if isinstance(result.get("changes"), dict) else {}
    counts: dict[str, int] = {}
    for value in changes.values():
        counts[_text(value)] = counts.get(_text(value), 0) + 1
    names = {"unchanged": "変更なし", "renumbered": "ページ番号だけ変わった", "changed": "変わった", "added": "新しい"}
    lines = [f"比較画像: {len(sheets)} 枚（{len(result.get('slides') or [])} ページ）"]
    if result.get("compared_run"):
        detail = "、".join(f"{names.get(key, key)} {count}" for key, count in counts.items())
        lines.append(f"  前の実行 {_text(result['compared_run'])} と比較" + (f"（{detail}）" if detail else ""))
    lines.extend(f"  {_shown(path)}" for path in sheets)
    lines.append("  比較画像は確認を助ける表示で、確認の記録ではない。ページ画像を開いて確かめてから記録する")
    return lines + _next([_status_operation(run)])


def _review_accept(result: dict, limit: int) -> list[str]:
    records = _issues(result.get("records"))
    run = _run_of(result)
    head = f"警告の許容: {STATUS_LABELS.get(result.get('status'), _text(result.get('status')) or '状態不明')}（{len(records)} 件"
    head += f"、{result['batch']}）" if result.get("batch") else "）"
    lines = [head]
    if run:
        lines.append(f"実行: {_shown(run)}")
    for record in records[:limit]:
        lines.append("  " + _issue_line({**record, "severity": None, "message": None}))
    if len(records) > limit:
        lines.append(f"  ほか {len(records) - limit} 件（--json で全件を表示）")
    skipped = _issues(result.get("skipped"))
    if skipped:
        reasons: dict[str, int] = {}
        for item in skipped:
            reasons[_text(item.get("reason"))] = reasons.get(_text(item.get("reason")), 0) + 1
        lines.append(
            f"記録しなかった警告: {len(skipped)} 件（"
            + "、".join(f"{name} {count}" for name, count in reasons.items())
            + "）"
        )
    return lines + _notice_lines(result.get("notices")) + _next([_status_operation(run)])


def _review_revoke(result: dict, limit: int) -> list[str]:
    revoked = (
        [value for value in result.get("revoked") or [] if value] if isinstance(result.get("revoked"), list) else []
    )
    run = _run_of(result)
    head = f"記録の取り消し: {STATUS_LABELS.get(result.get('status'), _text(result.get('status')) or '状態不明')}（{len(revoked)} 件"
    head += f"、{result['batch']}）" if result.get("batch") else "）"
    lines = [head]
    if run:
        lines.append(f"実行: {_shown(run)}")
    if revoked:
        lines.append(f"  取り消した記録: {_listed_ids(revoked)}")
    for item in _issues(result.get("cascaded"))[:limit]:
        lines.append(f"  元の記録も取り消した: {_shown(item.get('run'))} の {_text(item.get('record_id'))}")
    return lines + _notice_lines(result.get("notices")) + _next([_status_operation(run)])


def _review_carry(result: dict, limit: int) -> list[str]:
    status = result.get("status")
    run = _run_of(result)
    label = STATUS_LABELS.get(status, _text(status) or "状態不明")
    if status != "carried":
        lines = [f"確認の引き継ぎ: {label}"]
        if result.get("message"):
            lines.append(f"  {_text(result['message'])}")
        return lines
    records = _issues(result.get("records"))
    lines = [f"確認の引き継ぎ: {label}（{len(records)} 件）"]
    if run:
        lines.append(f"実行: {_shown(run)}")
    sources = result.get("sources") if isinstance(result.get("sources"), dict) else {}
    if sources:
        lines.append("  引き継ぎ元: " + "、".join(f"{_shown(name)} {count} 件" for name, count in sources.items()))
    if renumbered := _distinct(result.get("renumbered")):
        lines.append(f"  ページ番号だけ変わったページ: {_listed_ids(renumbered)}")
    skipped = result.get("skipped")
    reasons: dict[str, int] = {}
    if isinstance(skipped, dict):
        reasons = {_text(name): count for name, count in skipped.items() if _integer(count)}
    elif isinstance(skipped, list):
        for item in _issues(skipped):
            reasons[_text(item.get("reason"))] = reasons.get(_text(item.get("reason")), 0) + 1
    if reasons:
        lines.append("  引き継がなかった確認: " + "、".join(f"{name} {count} 件" for name, count in reasons.items()))
    skipped_runs = result.get("skipped_runs") if isinstance(result.get("skipped_runs"), dict) else {}
    if skipped_runs:
        lines.append(
            "  候補にしなかった実行: " + "、".join(f"{_text(name)} {count} 件" for name, count in skipped_runs.items())
        )
    if remaining := _remaining_text(result.get("remaining")):
        lines.append(f"  未確認: {remaining}")
    return lines + _notice_lines(result.get("notices")) + _next([_status_operation(run)])


# ---------------------------------------------------------------- 最上位の例外


def _command_failure(result: dict, limit: int, command: Any = None) -> list[str]:
    """コマンドの最上位で捕まえた例外の要約。文言は複数行（設定の問題の一覧など）をそのまま字下げして示す。

    場所（失敗の診断）は説明のための例外でないときだけ結果にあり、そのときだけ示す。
    """
    label = _label(command)
    exception = _text(result.get("exception"))
    if exception == "KeyboardInterrupt":
        return [f"{label}: 中断", f"  {_text(result.get('message')) or '処理を中断しました'}"]
    message = _text(result.get("message")) or exception
    first, *rest = message.splitlines() or [""]
    lines = [f"{label}: 失敗", "  " + (f"{exception}: {first}" if exception and exception != first else first)]
    lines.extend(f"    {line}" for line in rest)
    for issue in _issues(result.get("issues"))[:1]:
        item = _first(issue, "location")
        location = item.get("location") if item is not None else None
        if not isinstance(location, dict) or not (location.get("relative_path") or location.get("path")):
            continue
        where = location.get("relative_path") or location.get("path")
        place = f"{where}:{location['line']}" if _integer(location.get("line")) else _text(where)
        notes = [_text(location.get("function"))] if location.get("function") else []
        if location.get("in_project") is False:
            notes.append("slidekit の内部で発生")
        lines.append(f"  場所: {place}" + (f"（{'、'.join(notes)}）" if notes else ""))
        if location.get("code"):
            lines.append(f"      {location['code']}")
        lines.append("  トレースバック: 標準エラーに出力")
    return lines


def _measure(result: dict, limit: int) -> list[str]:
    results = _issues(result.get("results"))
    width = result.get("width_pt")
    notes = [f"幅 {width:.1f} pt" if _number(width) else "折り返さない"]
    if result.get("role"):
        notes.append(f"役割 {_text(result['role'])}")
    if _number(result.get("size_pt")):
        notes.append(f"{result['size_pt']:g} pt")
    lines = [f"計測: {len(results)} 件（{'、'.join(notes)}）"]
    for number, item in enumerate(results, 1):
        rows = _issues(item.get("lines"))
        height = item.get("frame_height_pt")
        head = f"  {number}: {len(rows)} 行" + (f"、枠の高さ {height:.1f} pt" if _number(height) else "")
        single = item.get("single_line_frame_width_pt")
        if len(rows) > 1 and _number(single):
            head += f"（1 行に収める枠の幅 {single:.1f} pt）"
        elif not _number(width) and _number(item.get("frame_width_pt")):
            head += f"、枠の幅 {item['frame_width_pt']:.1f} pt"
        lines.append(head)
        lines.extend(_wrap_rows(rows, "    "))
    return lines


_WRAP_MARKS = {"mid_phrase": " ← 文節の途中", "hyphen": " ← ハイフンの後", "split": " ← 語の途中"}


def _wrap_rows(rows: list[dict], indent: str) -> list[str]:
    """行の文字列と幅。自動の折り返しの位置の注意（文節の途中など）を行の末尾に添える。"""
    lines = []
    for row in rows:
        width = row.get("width_pt")
        text = f"{indent}「{_text(row.get('text')).rstrip()}」" + (f" {width:.1f} pt" if _number(width) else "")
        lines.append(text + _WRAP_MARKS.get(row.get("wrap"), ""))
    return lines


def _wrap_listing(result: dict) -> list[str]:
    """layout --wraps：自動で折り返した要素（テーマの要素を除く）の行の文字列と幅。"""
    layout = result.get("layout") if isinstance(result.get("layout"), dict) else {}
    lines = []
    for slide in _issues(layout.get("slides")):
        page = f"p.{slide['number']}" if _integer(slide.get("number")) else f"ページ {_text(slide.get('id'))}"
        for element in _issues(slide.get("elements")):
            rows = _issues(element.get("lines"))
            if str(element.get("id", "")).startswith("_theme/"):
                continue
            if not any(row.get("wrap") for row in rows):
                continue
            inner = element.get("inner_pt")
            width = (
                f"（内側の幅 {inner[2]:.1f} pt）"
                if isinstance(inner, list) and len(inner) == 4 and _number(inner[2])
                else ""
            )
            lines.append(f"  {page} {_text(element.get('id'))}{width}")
            lines.extend(_wrap_rows(rows, "    "))
    return ["自動で折り返した要素:", *lines] if lines else ["自動で折り返した要素: なし"]


_SUMMARIES = {
    "layout": _layout,
    "measure": _measure,
    "doctor": _doctor,
    "build": _build,
    "check": _check,
    "export": _export,
    "clean": _clean,
    "renderer status": _renderer_status,
    "renderer clear": _renderer_clear,
    "review status": _review_status,
    "review record": _review_record,
    "review accept": _review_accept,
    "review sheet": _review_sheet,
    "review revoke": _review_revoke,
    "review carry": _review_carry,
}
