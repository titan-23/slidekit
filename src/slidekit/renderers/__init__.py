"""PowerPoint によるローカルの PDF 変換。

`open_session()` で変換セッションを開き、事前確認と変換を同じセッションで行う。セッション（`powerpoint.py`）は
PowerShell の作業プロセスを 1 つだけ起動し、Windows 利用者ごとの排他を閉じるまで保持する。
`doctor_renderer()`・`render_pdf()` は、1 回限りのセッションを使う互換の関数である。
"""

from __future__ import annotations

import base64
from collections.abc import Callable, Iterator, Mapping
import contextlib
import functools
import hashlib
import json
import math
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
from typing import TYPE_CHECKING, Any
import uuid

if TYPE_CHECKING:
    from .powerpoint import PowerPointSession

# PowerShell の作業スクリプト。作業フォルダーへ BOM 付き UTF-8 で書いてから実行する。
SCRIPT = Path(__file__).with_name("powerpoint.ps1")
# renderer status・renderer clear の制限時間（秒）。
MAINTENANCE_TIMEOUT_SECONDS = 60


class RendererError(RuntimeError):
    """環境の不足または変換の失敗。生成の成功とは区別する。"""

    def __init__(self, message: str, *, code: str = "renderer_failed", **details: Any):
        super().__init__(message)
        self.code = code
        self.details = details


def _settings(environment: Mapping) -> dict:
    value = environment.get("powerpoint", {})
    if not isinstance(value, dict):
        raise RendererError("powerpoint の設定には表が必要です。", code="invalid_renderer")
    return value


def _executable(value: str) -> str:
    result = shutil.which(value)
    if result is None:
        raise RendererError(f"実行ファイルが見つかりません: {value}", code="renderer_missing")
    # WSL の wslpath は /init へのリンクであり、呼び出し名を保つ必要がある。
    return str(Path(result).absolute())


def _seconds(value: Any, label: str, *, positive: bool) -> float:
    """制限時間と待ち時間の値を確かめる。設定の読み込みでも確かめるが、引数で渡された値もここで拒否する。"""
    if (
        isinstance(value, bool)
        or not isinstance(value, (int, float))
        or not math.isfinite(value)
        or value < 0
        or (positive and value == 0)
    ):
        bound = "正" if positive else "0 以上"
        raise RendererError(f"{label}は{bound}の秒数で指定してください。", code="invalid_renderer")
    return float(value)


def _run(command: list[str], *, timeout: float = 30) -> subprocess.CompletedProcess:
    try:
        return subprocess.run(command, capture_output=True, timeout=timeout, check=False)
    except subprocess.TimeoutExpired as exc:
        raise RendererError("表示ソフトの処理が制限時間を超えました。", code="renderer_timeout") from exc
    except OSError as exc:
        raise RendererError(f"表示ソフトを実行できません: {exc}", code="renderer_unavailable") from exc


def _decode(value: bytes | str) -> str:
    if isinstance(value, str):
        return value.lstrip("﻿")
    for encoding in ("utf-8-sig", "cp932"):
        try:
            return value.decode(encoding)
        except UnicodeDecodeError:
            pass
    return value.decode("utf-8", errors="replace")


def _failure(value: dict) -> RendererError:
    """作業スクリプトの失敗の応答を、文言の接頭辞で分類した例外にする。"""
    message = str(value.get("message") or "PowerPoint による変換に失敗しました。")
    code = "renderer_failed"
    if "SLIDEKIT_PENDING:" in message:
        code = "renderer_pending"
    if "SLIDEKIT_BUSY:" in message:
        code = "renderer_busy"
    return RendererError(message, code=code, result=value)


def _result(process: subprocess.CompletedProcess) -> dict:
    output = _decode(process.stdout)
    lines = [line.removeprefix("SLIDEKIT_JSON:") for line in output.splitlines() if line.startswith("SLIDEKIT_JSON:")]
    if not lines:
        detail = (_decode(process.stderr) or output).strip()[-3000:]
        raise RendererError(f"PowerShell から実行結果を取得できません: {detail}")
    try:
        value = json.loads(lines[-1])
    except (ValueError, TypeError) as exc:
        raise RendererError("PowerShell の実行結果が不正です。") from exc
    if not isinstance(value, dict):
        raise RendererError("PowerShell の実行結果が不正です。")
    if process.returncode != 0 or not value.get("ok"):
        raise _failure(value)
    return value


def _powershell_command(executable: str, script: str) -> list[str]:
    # 動的なパスをコードに埋め込まず、固定した問い合わせだけを実行する。
    encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    return [executable, "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded]


def _powershell(environment: Mapping) -> str:
    """設定の PowerShell の実行ファイルを解決する。"""
    return _executable(_settings(environment).get("powershell", "powershell.exe"))


def _windows_context(powershell: str) -> dict:
    """Windows の一時フォルダー、LOCALAPPDATA、利用者の SID、フォントのフォルダーを PowerShell で問い合わせる。"""
    script = """
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
[Console]::OutputEncoding=New-Object System.Text.UTF8Encoding($false)
$data=@{ok=$true; temp=[IO.Path]::GetTempPath(); local_app_data=$env:LOCALAPPDATA;
sid=[System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value;
fonts=[Environment]::GetFolderPath('Fonts')}
Write-Output ('SLIDEKIT_JSON:' + ($data | ConvertTo-Json -Compress))
"""
    value = _result(_run(_powershell_command(powershell, script)))
    if not all(isinstance(value.get(key), str) and value[key] for key in ("temp", "local_app_data", "sid")):
        raise RendererError("Windows の一時フォルダーと利用者の情報を取得できません。")
    return {key: value.get(key) for key in ("temp", "local_app_data", "sid", "fonts")}


@functools.cache
def _cached_windows_context(powershell: str) -> dict:
    """`_windows_context()` の結果を、PowerShell の実行ファイルごとにプロセス内で覚える。失敗は覚えない。"""
    return _windows_context(powershell)


def windows_context(environment: Mapping) -> dict:
    """Windows の一時フォルダー、LOCALAPPDATA、利用者の SID、フォントのフォルダーを返す。

    キーは `temp`・`local_app_data`・`sid`・`fonts`（どれも Windows のパス）。
    PowerShell の実行ファイルごとに、プロセス内で 1 回だけ問い合わせる。
    """
    return dict(_cached_windows_context(_powershell(environment)))


def _unix_path(windows_path: str) -> Path:
    if os.name == "nt":
        return Path(windows_path)
    process = _run([_executable("wslpath"), "-u", windows_path])
    if process.returncode:
        raise RendererError(f"Windows パスを変換できません: {windows_path}")
    return Path(_decode(process.stdout).strip())


def _windows_path(path: Path) -> str:
    if os.name == "nt":
        return str(path.resolve())
    process = _run([_executable("wslpath"), "-w", str(path.resolve())])
    if process.returncode:
        raise RendererError(f"Windows パスへ変換できません: {path}")
    return _decode(process.stdout).strip()


def _work_folder(context: Mapping) -> Path:
    """Windows の一時フォルダーの下に、作業フォルダー `slidekit\\ppt-*` を新しく作る。"""
    root = _unix_path(context["temp"]) / "slidekit"
    root.mkdir(parents=True, exist_ok=True)
    return Path(tempfile.mkdtemp(prefix="ppt-", dir=root))


def _state_folder(context: Mapping) -> Path:
    """残存記録（pending.json）と使用中の記録（active.json）を置くフォルダー。Windows 利用者ごとに 1 つ。"""
    return _unix_path(context["local_app_data"]) / "slidekit" / "powerpoint"


def _write_script(work: Path) -> None:
    # Windows PowerShell 5 は BOM なしの日本語を ANSI として解釈する。
    (work / SCRIPT.name).write_text(SCRIPT.read_text(encoding="utf-8"), encoding="utf-8-sig")


def _script_command(powershell: str, work: str, mode: str, state: str, run_id: str) -> list[str]:
    """作業スクリプトを起動する引数。work と state は Windows のパスで、コードに埋め込まずに引数で渡す。"""
    return [
        powershell,
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        f"{work}\\{SCRIPT.name}",
        "-Mode",
        mode,
        "-StateDir",
        state,
        "-WorkDir",
        work,
        "-RunId",
        run_id,
    ]


def _holder(record: Any) -> dict | None:
    """使用中の記録（active.json）から、待つ側に見せる情報を作る。

    資料・実行・コマンドなどの情報（`holder`）に、作業プロセスの PID・開始時刻・セッションの識別子を加える。
    """
    if not isinstance(record, dict):
        return None
    holder = dict(record["holder"]) if isinstance(record.get("holder"), dict) else {}
    for key in ("pid", "started_at", "run_id"):
        if record.get(key) is not None:
            holder.setdefault(key, record[key])
    return holder


def _font_face(path: Path, index: int = 0) -> dict:
    from fontTools.ttLib import TTFont

    try:
        font = TTFont(path, fontNumber=index, lazy=True)
        try:
            digest = hashlib.sha256()
            for tag in sorted(font.reader.keys()):
                if tag == "DSIG":
                    continue
                data = font.reader[tag]
                if tag == "head":
                    data = data[:8] + b"\0\0\0\0" + data[12:]
                    # 集合からの抽出で更新される保存日時は、字形が同じかの判定に含めない。
                    data = data[:28] + b"\0" * 8 + data[36:]
                digest.update(tag.encode("ascii"))
                digest.update(len(data).to_bytes(8, "big"))
                digest.update(data)
            names = sorted({record.toUnicode() for record in font["name"].names if record.nameID in (1, 4, 6, 16)})
            versions = sorted({record.toUnicode() for record in font["name"].names if record.nameID == 5})
            return {"index": index, "table_sha256": digest.hexdigest(), "font_names": names, "version": versions}
        finally:
            font.close()
    except Exception as exc:
        raise RendererError(f"フォントを読み取れません: {path}（書体 {index}）", code="font_unreadable") from exc


def _face_count(path: Path) -> int:
    with path.open("rb") as stream:
        header = stream.read(12)
    return int.from_bytes(header[8:12], "big") if header[:4] == b"ttcf" else 1


def check_windows_fonts(environment: Mapping, installed_fonts: list[dict]) -> list[dict]:
    """寸法の計算に使う書体と Windows に登録された書体の実体を照合する。

    照合元と同じファイルの登録を先に調べる。同じファイルの同じ書体番号は、表のハッシュ値を計算し直さずに一致とする。
    """
    fonts = environment.get("fonts", {})
    matches = []
    path_cache: dict[str, Path] = {}
    face_cache: dict[tuple[Path, int], dict] = {}

    def local_path(value: str) -> Path:
        if value not in path_cache:
            path_cache[value] = _unix_path(value).resolve()
        return path_cache[value]

    def face(path: Path, index: int) -> dict:
        if (path, index) not in face_cache:
            face_cache[path, index] = _font_face(path, index)
        return face_cache[path, index]

    def file_name(entry: dict) -> str:
        return entry["path"].replace("\\", "/").split("/")[-1].casefold()

    def same_file(entry: dict, source_path: Path) -> bool:
        # ファイル名が違えば同じファイルではないので、パスの変換を省く。
        if file_name(entry) != source_path.name.casefold():
            return False
        try:
            destination = local_path(entry["path"])
            return destination == source_path or os.path.samefile(destination, source_path)
        except (OSError, RendererError):
            return False

    def find(candidates: list[dict], source_path: Path, index: int, source_face: dict) -> tuple[dict, int] | None:
        for entry in candidates:
            if same_file(entry, source_path):
                # 同じファイルの同じ書体なので、照合元の値をそのまま使う。
                return entry, index
        for entry in candidates:
            try:
                destination = local_path(entry["path"])
                count = _face_count(destination)
                for other in [index] + [other for other in range(count) if other != index]:
                    if other < count and face(destination, other)["table_sha256"] == source_face["table_sha256"]:
                        return entry, other
            except (OSError, RendererError):
                continue
        return None

    for family, config in fonts.items():
        if not isinstance(config, dict) or not config.get("path"):
            raise RendererError(f"フォントのファイル指定がありません: {family}", code="font_missing")
        variants = [("regular", config["path"], int(config.get("index", 0)))]
        if config.get("bold_path"):
            variants.append(("bold", config["bold_path"], int(config.get("bold_index", 0))))
        for variant, source, index in variants:
            source_path = Path(source).resolve()
            source_face = face(source_path, index)
            if family.casefold() not in {name.casefold() for name in source_face["font_names"]}:
                raise RendererError(
                    f"指定された書体名がフォントの内容と一致しません: {family}",
                    code="font_family_mismatch",
                    family=family,
                    path=str(source_path),
                )
            candidates = [
                entry
                for entry in installed_fonts
                if family.casefold() in entry.get("name", "").casefold()
                or source_path.name.casefold() == file_name(entry)
            ]
            found = find(candidates, source_path, index, source_face)
            if found is None:
                raise RendererError(
                    f"Windows に登録されたフォントと実体を照合できません: {family} / {variant}",
                    code="windows_font_mismatch",
                    family=family,
                    path=str(source_path),
                )
            entry, windows_index = found
            matches.append(
                {
                    "family": family,
                    "variant": variant,
                    "source": str(source_path),
                    "windows_path": entry["path"],
                    "windows_index": windows_index,
                    **source_face,
                }
            )
    return matches


def _render_paths(input_pptx: Path, output_pdf: Path) -> tuple[Path, Path]:
    """変換の入力があり、出力先がまだないことを確かめる。既存の PDF は採用しない。"""
    input_pptx, output_pdf = Path(input_pptx).resolve(), Path(output_pdf).resolve()
    if not input_pptx.is_file():
        raise RendererError(f"入力 PPTX がありません: {input_pptx}", code="input_missing")
    if output_pdf.exists():
        raise RendererError(f"PDF 出力先には新しいパスを指定してください: {output_pdf}", code="output_exists")
    return input_pptx, output_pdf


def _is_pdf(path: Path) -> bool:
    with path.open("rb") as stream:
        return stream.read(5) == b"%PDF-"


@contextlib.contextmanager
def open_session(
    environment: Mapping,
    *,
    wait_seconds: float | None = None,
    holder: dict | None = None,
    on_event: Callable[[dict], None] | None = None,
    state_dir: Path | None = None,
) -> Iterator[PowerPointSession]:
    """PowerPoint の変換セッションを開き、with を抜けるとき（例外を含む）に閉じる。

    - wait_seconds：別の変換が終わるのを待つ上限（秒）。None なら `environment["powerpoint"]["wait_seconds"]`（既定 600）。0 は待たない。
    - holder：待つ側に見せる情報（JSON にできる 2 KB までの辞書。`project`・`run`・`command`・`host`・`python_pid` など）。
    - on_event：途中経過の通知（`{"event": "waiting", "waited_seconds": …, "holder": …}` など）。呼び出し元のスレッドで、
      応答を待つ間に呼ぶ。
    - state_dir：試験専用。残存記録と使用中の記録の置き場所を差し替える。

    開くときに PowerShell の作業プロセスを起動して排他を取り、閉じるまで保持する。
    利用者の中断（Ctrl+C）は、準備完了の前なら作業プロセスを終了させ、後なら切り離して、`KeyboardInterrupt` を送り直す。
    """
    from .powerpoint import PowerPointSession

    session = PowerPointSession(
        environment, wait_seconds=wait_seconds, holder=holder, on_event=on_event, state_dir=state_dir
    )
    try:
        session.start()
        yield session
    except KeyboardInterrupt:
        session._interrupt()
        raise
    finally:
        session.close()


def doctor_renderer(environment: Mapping) -> dict:
    """PowerPoint の環境を確かめる（1 回限りのセッションを使う）。"""
    with open_session(environment) as session:
        return session.doctor()


def render_pdf(input_pptx: Path, output_pdf: Path, environment: Mapping) -> dict:
    """新しい PDF を出力する。失敗時は成功の結果を返さない（1 回限りのセッションを使う）。"""
    input_pptx, output_pdf = _render_paths(input_pptx, output_pdf)
    with open_session(environment) as session:
        return session.render(input_pptx, output_pdf)


def _maintenance(
    mode: str, environment: Mapping, *, state_dir: Path | None, force: bool = False, stop_owned: bool = False
) -> dict:
    """作業スクリプトを Status か Clear で 1 回実行し、結果の辞書を返す。排他は取り続けない。"""
    powershell = _powershell(environment)
    context = windows_context(environment)
    work = _work_folder(context)
    try:
        _write_script(work)
        windows_work = _windows_path(work)
        state = Path(state_dir) if state_dir is not None else _state_folder(context)
        command = _script_command(powershell, windows_work, mode, _windows_path(state), uuid.uuid4().hex)
        if force:
            command.append("-Force")
        if stop_owned:
            command.append("-StopOwned")
        return _result(_run(command, timeout=MAINTENANCE_TIMEOUT_SECONDS))
    finally:
        shutil.rmtree(work, ignore_errors=True)


def _items(value: Any) -> list[dict]:
    """作業スクリプトの結果の一覧を辞書の列にする。PowerShell が 1 件の配列を要素だけにした場合も列にする。"""
    if isinstance(value, dict):
        value = [value]
    return [item for item in value if isinstance(item, dict)] if isinstance(value, list) else []


def _whole(value: Any) -> int | None:
    return value if isinstance(value, int) and not isinstance(value, bool) else None


def _text_or_none(value: Any) -> str | None:
    return value if isinstance(value, str) and value else None


def _powerpoint_processes(value: Any, *, available: bool) -> list[dict]:
    """PowerPoint の処理の一覧（`renderer status` の `powerpoint_processes`）。

    `owned` は、slidekit が起動したと記録した PowerPoint（残存記録か使用中の記録の `powerpoint`）と PID・開始時刻の両方が一致するか、
    `record` は一致した記録（`pending`・`active`）である。`stoppable` は `renderer clear --stop-owned` が終了する対象か
    （一致し、ウィンドウがなく、available が真）。available は、排他を持つ処理も、動いている残存記録の作業プロセスもないこと。
    """
    processes = []
    for item in _items(value):
        owned = item.get("owned") is True
        has_window = bool(item.get("has_window"))
        record = item.get("record")
        processes.append(
            {
                "pid": _whole(item.get("pid")),
                "window_title": str(item.get("window_title") or ""),
                "has_window": has_window,
                "started_at": _text_or_none(item.get("started_at")),
                "owned": owned,
                "record": record if owned and record in ("pending", "active") else None,
                "stoppable": owned and not has_window and available,
            }
        )
    return processes


def _powerpoint_records(value: Any) -> list[dict]:
    """slidekit が起動したと記録した PowerPoint（`renderer status` の `powerpoint_records`）。running はその処理が動いているか。"""
    return [
        {
            "source": item.get("source") if item.get("source") in ("pending", "active") else None,
            "run_id": _text_or_none(item.get("run_id")),
            "pid": _whole(item.get("pid")),
            "started_at": _text_or_none(item.get("started_at")),
            "running": item.get("running") is True,
        }
        for item in _items(value)
    ]


def _stopped_powerpoint(value: Any) -> list[dict]:
    """`renderer clear --stop-owned` が終了した PowerPoint（PID・開始時刻・一致した記録）。"""
    return [
        {
            "pid": _whole(item.get("pid")),
            "started_at": _text_or_none(item.get("started_at")),
            "record": item.get("record") if item.get("record") in ("pending", "active") else None,
        }
        for item in _items(value)
    ]


def renderer_status(environment: Mapping, *, state_dir: Path | None = None) -> dict:
    """排他・残存記録・使用中の処理を調べる（`slidekit renderer status`）。

    排他は取ることができたら直ちに放す。`status` は `idle`・`busy`（別の処理が排他を持つ）・`pending`（残存記録がある）で、
    `idle` 以外なら `exit_code` は 1。`pending_process_alive` は残存記録の作業プロセスが動いているか、
    `staged_input_locked` は残存記録の検査用の複製を排他で開くことができないか（どちらも残存記録がなければ None）。
    PowerPoint は読み取り専用で開いた複製のファイルを開いたままにしないので、偽でも PowerPoint が複製を開いていることがある。
    `abandoned` は排他を持っていた処理が放さずに終わっていたか（この確認で取って放したので、後の処理には残らない）。
    `powerpoint_processes` は PowerPoint の処理ごとの PID・ウィンドウの題名・開始時刻と、slidekit が起動したと記録した PowerPoint と
    一致するか（`_powerpoint_processes`）、`powerpoint_records` はその記録と、記録の処理が動いているかである。
    記録は作業スクリプトが PowerPoint を起動したときに書く（残存記録と使用中の記録の `powerpoint`）。古い記録には記録がない。
    """
    value = _maintenance("Status", environment, state_dir=state_dir)
    busy = bool(value.get("busy"))
    pending = value.get("pending") if isinstance(value.get("pending"), dict) else None
    exists = pending is not None or bool(value.get("pending_exists"))
    status = "busy" if busy else "pending" if exists else "idle"
    alive = value.get("pending_process_alive")
    return {
        "command": "renderer status",
        "status": status,
        "busy": busy,
        "holder": _holder(value.get("holder")) if busy else None,
        "pending": pending,
        "pending_process_alive": alive,
        "staged_input_locked": value.get("staged_input_locked"),
        "powerpoint_running": bool(value.get("powerpoint_running")),
        "powerpoint_processes": _powerpoint_processes(
            value.get("powerpoint_processes"), available=not busy and not (exists and alive)
        ),
        "powerpoint_records": _powerpoint_records(value.get("powerpoint_records")),
        "abandoned": bool(value.get("abandoned")),
        "state_dir": value.get("state_dir"),
        "exit_code": 0 if status == "idle" else 1,
    }


def clear_pending(
    environment: Mapping,
    *,
    force: bool = False,
    stop_owned: bool = False,
    state_dir: Path | None = None,
) -> dict:
    """安全を確かめてから残存記録を解除する（`slidekit renderer clear`）。

    解除するのは、排他を取得したうえで、残存記録の作業プロセスが動いておらず、検査用の複製を排他で開くことに成功したときである。
    PowerPoint は読み取り専用で開いた複製のファイルを開いたままにしないので、排他で開くことに成功しても、PowerPoint が複製を開いていないとは
    言えない。PowerPoint が起動しているときは、複製の有無によらず force のときだけ解除する（複製を排他で開くことに失敗したときは force でも断る）。
    ただし、残存記録の変換に使った PowerPoint が slidekit の起動したものと記録されていて、もう動いていなければ、force を求めない。
    解除した記録は状態フォルダーの `cleared/` へ移し、作業フォルダーが Windows の一時フォルダーの `slidekit` の下にあれば消す。
    あわせて、24 時間より古く、どの記録も指していない作業フォルダー `ppt-*` を消す。
    `status` は `cleared`・`nothing`・`refused` で、断ったときは `exit_code` が 1。

    stop_owned が真なら、作業プロセスが動いていないことを確かめた後、slidekit が起動したと確かめることができる PowerPoint
    （記録と PID・開始時刻の両方が一致し、ウィンドウがないもの）だけを強制終了してから解除する。残存記録がなくても、前の処理が
    残した使用中の記録と一致する PowerPoint は終了する。一致する PowerPoint にウィンドウがあれば、終了せずに断る。
    終了した PowerPoint は `stopped_powerpoint`（`pid`・`started_at`・`record`）に入る。ほかの PowerPoint には触れない。
    """
    value = _maintenance("Clear", environment, state_dir=state_dir, force=force, stop_owned=stop_owned)
    status = value.get("status")
    if status not in ("cleared", "nothing", "refused"):
        raise RendererError(f"残存記録の解除の結果が不正です: {status}", code="renderer_protocol", result=value)
    return {
        "command": "renderer clear",
        "status": status,
        "reason": str(value.get("reason") or ""),
        "archived": value.get("archived"),
        "removed_work_dir": bool(value.get("removed_work_dir")),
        "removed_old_work_dirs": int(value.get("removed_old_work_dirs") or 0),
        "pending": value.get("pending") if isinstance(value.get("pending"), dict) else None,
        "holder": _holder(value.get("holder")),
        "stopped_powerpoint": _stopped_powerpoint(value.get("stopped")),
        "exit_code": 1 if status == "refused" else 0,
    }


__all__ = [
    "RendererError",
    "open_session",
    "doctor_renderer",
    "render_pdf",
    "renderer_status",
    "clear_pending",
    "windows_context",
    "check_windows_fonts",
]
