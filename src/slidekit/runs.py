"""実行ディレクトリの割り当て、生成中の印、一覧と整理を扱う。

生成中の印 `.building` は JSON で、生成したプロセスの番号・端末名・開始時刻などを持つ。
`marker_state` はこの印から、生成が動いているか、中断したかを判定する。判定できないものは整理で保護する。
`records` は関数の中で読み込む（`records` もこのモジュールを使うので、読み込みを循環させない）。
"""

from __future__ import annotations

import contextlib
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import re
import shutil
import socket
import stat
import sys
import tomllib
from typing import Any

from .files import write_json

FINISHED_STATUSES = frozenset({"built", "stopped", "failed"})
MARKER_VERSION = 1
GITIGNORE = "# slidekit が作る実行記録と退避した完成物を追跡しない\n*\n"

_MARKER = ".building"
# この名前だけを含む出力先は slidekit が作ったものとみなし、.gitignore を置く。
_OUTPUT_ENTRIES = frozenset({"runs", "exports"})
# Windows の Python ではプロセスを確かめない（os.kill のシグナル 0 は Ctrl+C になる）。印は常に unknown として保護する。
_WINDOWS = os.name == "nt"
_RUN_NUMBER = re.compile(r"[0-9]+")
_MAX_PID = 2**31 - 1
_LEGACY_REASON = "旧形式の生成中の印です。生成が動いていないことを確かめてから .building を消してください"


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _is_number(name: str) -> bool:
    return _RUN_NUMBER.fullmatch(name) is not None


def _order(name: str) -> tuple:
    """実行番号の順。番号でない名前は、番号の後に名前の順で並べる。"""
    return (0, int(name), name) if _is_number(name) else (1, 0, name)


def _older(name: str, before: str) -> bool:
    if _is_number(name) and _is_number(before):
        return int(name) < int(before)
    return name < before


def _output_dir(project: Any) -> Path:
    """Project、または deck.toml（そのディレクトリでもよい）のパスから出力先を求める。

    パスのときは load_project を呼ばず、その deck.toml の output_dir だけを読む（既定 out。extends からは引き継がない）。
    設定の誤りやフォントの照合で、実行の一覧や指定が失敗しないようにするためである。
    """
    if not isinstance(project, (str, os.PathLike)):
        return Path(project.output_dir).resolve()
    path = Path(project).expanduser().resolve()
    if path.is_dir():
        path /= "deck.toml"
    with path.open("rb") as stream:
        value = tomllib.load(stream).get("output_dir", "out")
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"output_dir にはローカルのパスを指定してください: {path}")
    output = Path(value).expanduser()
    return (output if output.is_absolute() else path.parent / output).resolve()


# 生成中の印


def _boot_id() -> str | None:
    """Linux の起動ごとの識別子。ほかの OS では None。"""
    try:
        return Path("/proc/sys/kernel/random/boot_id").read_text(encoding="ascii").strip() or None
    except (OSError, UnicodeError):
        return None


def _pid_namespace() -> str | None:
    """Linux の PID 名前空間（例 `pid:[4026531836]`）。WSL の別のディストリビューションなどは番号の空間が違う。"""
    try:
        return os.readlink("/proc/self/ns/pid")
    except OSError:
        return None


def _process_stat(pid: int) -> tuple[str | None, int | None]:
    """Linux の /proc/<pid>/stat から、プロセスの状態（3 番目の欄）と開始時刻（22 番目の欄）を読む。読めなければ None。"""
    if _WINDOWS:
        return None, None
    try:
        data = Path(f"/proc/{pid}/stat").read_bytes()
    except OSError:
        return None, None
    # 2 番目の欄（実行ファイル名）は括弧で囲まれ、空白や括弧を含みうるので、最後の「)」より後を区切る。
    _, separator, tail = data.rpartition(b")")
    fields = tail.split()
    if not separator or len(fields) < 20:
        return None, None
    state = fields[0].decode("ascii", "replace")
    try:
        return state, int(fields[19])
    except ValueError:
        return state, None


def _write_gitignore(output: Path) -> None:
    """runs・exports だけを含む出力先に、中身を追跡しない .gitignore を置く（既存のものは変えない）。"""
    if any(entry.name not in _OUTPUT_ENTRIES for entry in output.iterdir()):
        return
    try:
        with (output / ".gitignore").open("x", encoding="utf-8") as stream:
            stream.write(GITIGNORE)
    except FileExistsError:
        # 同時に始めた別の生成が先に置いた。
        pass


def new_run(project: Any, *, command: str, options: dict) -> Path:
    """出力先の runs/NNNN を排他的に作り、生成中の印 .building を書く。

    出力先を新しく作ったときと、runs・exports だけを含む出力先には .gitignore を置く。
    options は JSON にできる値の辞書（例 {"draft": False, "slides": None}）。段階は construct から始まる。
    """
    pid = os.getpid()
    started = _now()
    marker = {
        "schema_version": MARKER_VERSION,
        "pid": pid,
        "host": socket.gethostname(),
        "platform": sys.platform,
        "boot_id": _boot_id(),
        "pid_namespace": _pid_namespace(),
        "process_start": _process_stat(pid)[1],
        "started_at": started,
        "updated_at": started,
        "project": str(project.path),
        # 実行保存先。資料を実行保存先ごと移しても、保存先から見た資料の相対位置で所属を判定できるようにする。
        "config": {"output_dir": str(Path(project.output_dir).resolve())},
        "command": command,
        "options": dict(options),
        "stage": "construct",
    }
    # JSON への変換に失敗する値は、実行ディレクトリを作る前に誤りにする（空の実行ディレクトリを残さない）。
    json.dumps(marker)
    output = Path(project.output_dir).resolve()
    runs = output / "runs"
    runs.mkdir(parents=True, exist_ok=True)
    _write_gitignore(output)
    # mkdir 自体を排他に使い、同時の build でも既存の実行を再利用しない。
    number = max((int(path.name) for path in runs.iterdir() if _is_number(path.name)), default=0) + 1
    while True:
        run = runs / f"{number:04d}"
        try:
            run.mkdir()
            break
        except FileExistsError:
            number += 1
    write_json(run / _MARKER, marker)
    return run


def update_marker(run: Path, *, stage: str) -> None:
    """生成中の印の stage と updated_at を書き換える（原子的に置き換える）。"""
    if not isinstance(stage, str) or not stage:
        raise ValueError("段階の名前を指定してください")
    path = Path(run) / _MARKER
    try:
        marker = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise FileNotFoundError(f"生成中の印がありません: {path}") from None
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ValueError(f"生成中の印を読めません: {path}: {error}") from error
    if not isinstance(marker, dict) or marker.get("schema_version") != MARKER_VERSION:
        raise ValueError(f"対応していない形式の生成中の印です: {path}")
    marker.update(stage=stage, updated_at=_now())
    write_json(path, marker)


def remove_marker(run: Path) -> None:
    (Path(run) / _MARKER).unlink(missing_ok=True)


def _same(recorded: Any, current: Any) -> bool:
    """片方が記録されていなければ比べない（仕様の形の印や、値の取得に対応していない環境のため）。"""
    return recorded is None or current is None or recorded == current


def marker_state(run: Path) -> dict:
    """生成中の印から実行の状態を判定する。

    戻り値：{"state": "none" | "running" | "interrupted" | "unknown", "pid", "host", "started_at", "stage", "reason"}。
    none は印がない、interrupted は生成のプロセスがもうない、unknown は判定できない（旧形式の印、別の端末・OS・PID 名前空間、
    Windows の Python）である。整理では running と unknown を保護し、interrupted を中断した生成として扱う。
    """
    return _marker_state(*_read_marker(Path(run)))


def _read_marker(run: Path) -> tuple[dict | None, str | None]:
    """生成中の印と、読めなかった理由。印がなければ両方 None。"""
    path = run / _MARKER
    if not os.path.lexists(path):
        return None, None

    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as error:
        return None, f"生成中の印を読めません: {error}"
    try:
        marker = json.loads(text)
    except json.JSONDecodeError:
        # 0.1 の印は資料のパスだけの文字列である。
        return None, _LEGACY_REASON
    if not isinstance(marker, dict) or marker.get("schema_version") != MARKER_VERSION:
        return None, "対応していない形式の生成中の印です"
    return marker, None


def _marker_state(marker: dict | None, problem: str | None) -> dict:
    """読み込み済みの印から生成の状態を判定する。"""
    result = {"state": "none", "pid": None, "host": None, "started_at": None, "stage": None, "reason": None}

    def judged(state: str, reason: str) -> dict:
        return {**result, "state": state, "reason": reason}

    if problem:
        return judged("unknown", problem)
    if marker is None:
        return result
    pid = marker.get("pid")
    result.update(pid=pid, host=marker.get("host"), started_at=marker.get("started_at"), stage=marker.get("stage"))
    if not isinstance(pid, int) or isinstance(pid, bool) or not 0 < pid <= _MAX_PID:
        return judged("unknown", "生成中の印のプロセス番号が不正です")
    if marker.get("host") != socket.gethostname():
        return judged(
            "unknown", f"別の端末（{marker.get('host')}）の生成です。その端末で生成が終わったことを確かめてください"
        )
    # WSL と Windows の Python は端末名が同じでも、プロセス番号の空間が違う。
    if not _same(marker.get("platform"), sys.platform):
        return judged(
            "unknown",
            f"別の OS（{marker.get('platform')}）の Python による生成です。その環境で生成が終わったことを確かめてください",
        )
    if _WINDOWS:
        return judged("unknown", "Windows の Python では生成のプロセスを確かめません")
    if not _same(marker.get("boot_id"), _boot_id()):
        return judged("interrupted", "端末を再起動する前の生成です")
    if not _same(marker.get("pid_namespace"), _pid_namespace()):
        return judged(
            "unknown",
            f"別の PID 名前空間（{marker.get('pid_namespace')}）の生成です。その環境で生成が終わったことを確かめてください",
        )
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return judged("interrupted", f"生成のプロセス（PID {pid}）が終了しています")
    except PermissionError:
        # 別の利用者のプロセスとして存在する。
        pass
    except OSError as error:
        return judged("unknown", f"生成のプロセス（PID {pid}）を確かめられません: {error}")
    state, start = _process_stat(pid)
    if state in {"Z", "X"}:
        # 終了して、親がまだ回収していないプロセス。
        return judged("interrupted", f"生成のプロセス（PID {pid}）が終了しています")
    if not _same(marker.get("process_start"), start):
        return judged("interrupted", f"PID {pid} は別のプロセスが使っています（番号の再利用）")
    return judged("running", f"PID {pid} が生成しています（段階 {marker.get('stage')}）")


# 一覧


def _read_manifest(run: Path) -> tuple[dict | None, str | None]:
    """生成記録を形式を問わずに読む。戻り値は（記録、読み込みに失敗した理由）。記録がなければ両方 None。"""
    path = run / "manifest.json"
    if not path.exists():
        return None, None
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        return None, f"生成記録を読めない（{error}）"
    if not isinstance(value, dict):
        return None, "生成記録を読めない（形式が不正）"
    return value, None


def _entries(runs_dir: Path) -> list[dict]:
    """実行ディレクトリごとに、生成記録と生成中の印を 1 回だけ読む（実行番号の順）。"""
    if not runs_dir.is_dir():
        return []
    entries = []
    for path in runs_dir.iterdir():
        if path.is_symlink() or not path.is_dir():
            continue
        manifest, problem = _read_manifest(path)
        marker_record, marker_problem = _read_marker(path)
        marker = _marker_state(marker_record, marker_problem)
        entry = {
            "path": path.resolve(),
            "run_id": path.name,
            "manifest": manifest is not None,
            "identity": manifest if manifest is not None else marker_record,
            "problem": problem,
            "status": None,
            "mode": None,
            "selection": None,
            "exportable": False,
            "marker": marker,
            "created_at": marker["started_at"],
        }
        if manifest is not None:
            # 形式 1 の生成記録には mode と exportable がない。全ページの通常の生成として示し、完成物にはしない。
            entry.update(
                status=manifest.get("status"),
                mode=manifest.get("mode") or "full",
                selection=manifest.get("selection"),
                exportable=manifest.get("exportable") is True,
                created_at=manifest.get("created_at") or marker["started_at"],
            )
        entries.append(entry)
    return sorted(entries, key=lambda entry: _order(entry["run_id"]))


def _size(path: Path) -> int:
    """ディレクトリの中の通常のファイルの合計の大きさ。リンクはたどらず、同じファイルへのハードリンクは 1 回だけ加算する。"""
    total, seen = 0, set()
    for directory, _, names in os.walk(path):
        for name in names:
            try:
                status = os.lstat(os.path.join(directory, name))
            except OSError:
                continue
            if not stat.S_ISREG(status.st_mode):
                continue
            if status.st_nlink > 1:
                key = (status.st_dev, status.st_ino)
                if key in seen:
                    continue
                seen.add(key)
            total += status.st_size
    return total


def list_runs(project: Any) -> list[dict]:
    """実行の一覧を実行番号の順に返す。project は Project か deck.toml のパス。

    各項目：{"run", "run_id", "status", "mode", "selection", "exportable", "state", "bytes", "created_at"}。
    生成記録がない実行の status・mode は None。state は marker_state の判定。
    """
    return [
        {
            "run": str(entry["path"]),
            "run_id": entry["run_id"],
            "status": entry["status"],
            "mode": entry["mode"],
            "selection": entry["selection"],
            "exportable": entry["exportable"],
            "state": entry["marker"]["state"],
            "bytes": _size(entry["path"]),
            "created_at": entry["created_at"],
        }
        for entry in _entries(_output_dir(project) / "runs")
    ]


def latest_run(project: Any, *, exportable_only: bool = True) -> Path | None:
    """生成記録があり生成中の印がない実行のうち、番号が最大のもの。project は Project か deck.toml のパス。

    exportable_only が真なら、status が built で exportable が真の実行（完成物にできる実行）だけを対象にする。
    exportable のない形式 1 の実行は対象にしない。該当する実行がなければ None。
    """
    runs_dir = _output_dir(project) / "runs"
    if not runs_dir.is_dir():
        return None
    names = sorted((path.name for path in runs_dir.iterdir() if _is_number(path.name)), key=_order, reverse=True)
    for name in names:
        run = runs_dir / name
        if run.is_symlink() or not run.is_dir() or os.path.lexists(run / _MARKER):
            continue
        manifest, _ = _read_manifest(run)
        if manifest is None:
            continue
        if exportable_only and (manifest.get("status") != "built" or manifest.get("exportable") is not True):
            continue
        return run.resolve()
    return None


# 整理


def protected_runs(project: Any, runs_dir: Path, output_dir: Path) -> dict[Path, list[str]]:
    """削除しない実行と、その理由を返す。

    保護するのは、採用中、反映した実行と退避した旧版、反映が未完了の実行と復旧に要る旧版、確認記録の引き継ぎ元、
    実行中（marker_state が running か unknown）、生成記録も生成中の印もない実行（割り当ての直後か、手で作ったもの）、
    生成記録の読み込みに失敗した実行、生成が終わっていない実行（status が FINISHED_STATUSES にない）、
    現在の資料の実行と確認できない実行である。理由は文言の列で返す（識別子と詳細は clean_runs の結果の protected）。
    """
    found = _protected(project, Path(output_dir), _entries(Path(runs_dir)))
    return {path: [reason["message"] for reason in reasons] for path, reasons in found.items()}


# 採用記録・反映の記録・確認記録の形式は records が持つ（形式 1 の記録も読む）。読めなければ records.RecordError になり、
# 保護すべき実行が分からないので整理を止める。


def _read_export_record(path: Path) -> dict:
    """採用記録（deliverables.json）と反映の記録（transaction.json・pending.json）を、形式 1 も含めて読む。"""
    from . import records

    return records.read_record(path, allow_legacy=True)


def _carry_sources(run: Path) -> list[Any]:
    """確認記録の写しが指す引き継ぎ元の実行（records.run_references の sources。生成記録は _entries で読み済み）。

    形式 2 の写しは人が記録した元の記録の実行だけを指すので、引き継ぎを重ねた途中の実行は保護しない。
    形式 1 の写しは 1 つ前の実行を指し、それを保護する。
    """
    from . import records

    return records._run_sources(run)


# 整理で実行を残す理由の識別子（clean --json の protected の reason_codes）。
PROTECTION_CODES = (
    "adopted",  # 現在採用している実行
    "exported",  # 完成物として反映した実行
    "previous_export",  # 旧版として退避した実行
    "export_pending",  # 反映が未完了の実行
    "export_recovery",  # 復旧に必要な旧版の実行
    "running",  # 実行中
    "running_unknown",  # 実行中か判定できない
    "unreadable",  # 生成記録を読むことができない
    "no_record",  # 生成記録も生成中の印もない
    "unfinished",  # 生成が終わっていない
    "other_project",  # 別の資料の実行
    "unknown_project",  # 所属を確認できない（記録に資料か実行保存先がない）
    "carry_source",  # 確認記録の引き継ぎ元
)


def _protected(project: Any, output_dir: Path, entries: list[dict]) -> dict[Path, list[dict]]:
    """保護する実行と理由（{"code", "message", ...}）の列。code は PROTECTION_CODES の値である。"""
    from . import records

    root = Path(project.root).resolve()
    protected: dict[Path, list[dict]] = {}
    identity = {
        "project": str(Path(getattr(project, "path", root / "deck.toml")).resolve()),
        "project_root": str(root),
        "config": {"output_dir": str(output_dir.resolve())},
    }

    def protect(path: Any, reason: str, code: str, **details: Any) -> None:
        if isinstance(path, (str, os.PathLike)) and str(path):
            # 相対パスは資料のディレクトリから読む（絶対パスはそのまま）。
            reasons = protected.setdefault((root / path).resolve(), [])
            if all(item["message"] != reason for item in reasons):
                reasons.append({"code": code, "message": reason, **details})

    record_name = project.config.get("deliverables", {}).get("record")
    if record_name:
        record_path = (root / record_name).resolve()
        if record_path.exists():
            protect(_read_export_record(record_path).get("run"), "現在採用している実行", "adopted")
    exports = output_dir / "exports"
    if exports.is_dir():
        for path in sorted(exports.glob("*/transaction.json")):
            transaction = _read_export_record(path)
            if transaction.get("status") == "committed":
                protect(transaction.get("run"), "完成物として反映した実行", "exported")
                protect(transaction.get("previous_run"), "旧版として退避した実行", "previous_export")
        pending = exports / "pending.json"
        if pending.exists():
            transaction = _read_export_record(pending)
            protect(transaction.get("run"), "完成物の反映が未完了の実行", "export_pending")
            protect(transaction.get("previous_run"), "復旧に必要な旧版の実行", "export_recovery")
    for entry in entries:
        run, marker = entry["path"], entry["marker"]
        if marker["state"] == "running":
            protect(run, f"実行中（PID {marker['pid']}、段階 {marker['stage']}）", "running")
        elif marker["state"] == "unknown":
            protect(run, f"実行中か判定できない: {marker['reason']}", "running_unknown")
        if entry["problem"]:
            protect(run, entry["problem"], "unreadable")
        elif not entry["manifest"]:
            if marker["state"] == "none":
                protect(run, "生成記録も生成中の印もない（割り当ての直後か、手で作った実行）", "no_record")
        elif not (isinstance(entry["status"], str) and entry["status"] in FINISHED_STATUSES):
            protect(run, f"生成が終わっていない（status: {entry['status']}）", "unfinished")
        # 実行保存先を共有する別の資料は整理しない。資料を実行保存先ごと移した場合は、確認の引き継ぎと同じく
        # 実行保存先から見た資料の相対位置で判定する。印だけの実行は記録された資料のパスが一致するときだけ対象にする。
        found = entry["identity"] or {}
        if run not in protected and not records._same_project(found, identity):
            recorded = found.get("project") if isinstance(found.get("project"), str) else None
            place = records._project_place(found)
            if recorded and place is not None:
                # 所属の情報はあるが、資料の場所も、実行保存先から見た資料の位置も違う（別の資料の実行）。
                protect(
                    run,
                    f"別の資料の実行（資料: {recorded}）",
                    "other_project",
                    recorded_project=recorded,
                    recorded_place=place,
                    current_place=records._project_place(identity),
                )
            else:
                # 生成記録か生成中の印に、資料の場所か実行保存先がない（古い形式の印など）。
                protect(
                    run,
                    "現在の資料の実行と確認できない（記録に資料か実行保存先がない）",
                    "unknown_project",
                    recorded_project=recorded,
                )
        for source in _carry_sources(run):
            protect(source, "確認記録の引き継ぎ元", "carry_source")
    return protected


def _failed_kind(entry: dict) -> list[str]:
    if entry["marker"]["state"] == "interrupted":
        return ["中断した生成"]
    if entry["status"] == "failed":
        return ["失敗した生成"]
    if entry["status"] == "stopped":
        return ["変換の前に止めた生成"]
    return []


def _draft_kind(entry: dict) -> list[str]:
    return [*(["下書き"] if entry["mode"] == "draft" else []), *(["部分生成"] if entry["selection"] else [])]


def clean_runs(
    project: Any,
    *,
    before: str | None = None,
    keep: int | None = None,
    failed: bool = False,
    drafts: bool = False,
    apply: bool = False,
) -> dict:
    """保護された実行を除く古い実行を、条件で絞って示し、apply が真のときだけ削除する。

    絞り込みの順：(1) 保護された実行を除く。(2) failed・drafts があれば、どちらかに当たる実行に絞る
    （failed は status が failed・stopped の実行と中断した実行、drafts は下書きと部分生成）。
    (3) before があれば、その実行より前に絞る。(4) keep があれば、残った実行の新しい keep 件を候補から外す。
    戻り値の candidates の各項目は {"run", "run_id", "status", "mode", "bytes", "reason"}、candidate_paths は候補のパスの列
    （0.1 の candidates と同じ形）、protected は {"run", "run_id", "reasons", "reason_codes", "details"} の列である（reason_codes は PROTECTION_CODES の値、details は理由ごとの code・message と、別資料の実行では記録された資料の場所と実行保存先から見た位置）。
    資料を実行保存先ごと移した場合も整理できる。実行保存先から見た資料の場所で判定し、別の資料や所属不明の実行は保護する。
    """
    if before is None and keep is None and not failed and not drafts:
        raise ValueError("--before、--keep、--failed、--drafts のいずれかを指定してください")
    if before is not None and (
        not isinstance(before, str) or not before or Path(before).name != before or before in {".", ".."}
    ):
        raise ValueError(f"--before には実行の識別子を指定してください: {before!r}")
    if keep is not None and (not isinstance(keep, int) or isinstance(keep, bool) or keep < 0):
        raise ValueError(f"--keep には 0 以上の整数を指定してください: {keep!r}")
    from . import records

    output = Path(project.output_dir).resolve()
    # 削除するときは、確認記録の操作と同じ、実行保存先に置く資料単位のロックを取る（資料のディレクトリにはロックのファイルを
    # 置かない）。候補の表示だけならロックを取らず、ファイルを何も書かない（削除しないので、表示の後に保護が増えても害はない）。
    with records._output_lock(output) if apply else contextlib.nullcontext():
        # ロックを持つ間は、確認・引き継ぎ・反映が新しい参照を作らないので、保護は増えない。
        # 生成記録と保護は最初に 1 回だけ読み、削除のたびに読み直さない。
        entries = _entries(output / "runs")
        protected = _protected(project, output, entries)
        selected = []
        for entry in entries:
            if entry["path"] in protected:
                continue
            reasons = []
            if failed or drafts:
                kinds = [*(_failed_kind(entry) if failed else []), *(_draft_kind(entry) if drafts else [])]
                if not kinds:
                    continue
                reasons.extend(kinds)
            if before is not None:
                if not _older(entry["run_id"], before):
                    continue
                reasons.append(f"{before} より前")
            selected.append((entry, reasons))
        if keep is not None:
            selected = selected[:-keep] if keep else selected
            for _, reasons in selected:
                reasons.append(f"新しい {keep} 件より古い" if keep else "保持しない（--keep 0）")
        candidates = [
            {
                "run": str(entry["path"]),
                "run_id": entry["run_id"],
                "status": entry["status"],
                "mode": entry["mode"],
                "bytes": _size(entry["path"]),
                "reason": "、".join(reasons),
            }
            for entry, reasons in selected
        ]
        removed = []
        if apply:
            for entry, _ in selected:
                shutil.rmtree(entry["path"])
                removed.append(str(entry["path"]))
    return {
        "command": "clean",
        "status": "cleaned" if apply else "preview",
        "conditions": {"before": before, "keep": keep, "failed": bool(failed), "drafts": bool(drafts)},
        "candidates": candidates,
        "candidate_paths": [item["run"] for item in candidates],
        "removed": removed,
        "protected": [
            {
                "run": str(entry["path"]),
                "run_id": entry["run_id"],
                "reasons": [reason["message"] for reason in protected[entry["path"]]],
                "reason_codes": [reason["code"] for reason in protected[entry["path"]]],
                "details": protected[entry["path"]],
            }
            for entry in entries
            if entry["path"] in protected
        ],
        "bytes": sum(item["bytes"] for item in candidates),
        "exit_code": 0,
    }
