"""PowerPoint の変換セッション。

1 回の生成で PowerShell の作業プロセス（`powerpoint.ps1` の Session）を 1 つだけ起動し、Windows 利用者ごとの排他を
閉じるまで保持する。PowerPoint は最初の変換で起動し、セッションの間は同じものを使う。
リクエストは 1 行に 1 つの JSON（ASCII）で標準入力へ送り、作業プロセスは `SLIDEKIT_EVENT:`（途中経過）と
`SLIDEKIT_JSON:`（応答。起動時の失敗は番号 0）の行で答える。

作業プロセスを終了させるのは、準備完了の前（排他を待っているだけで、PowerPoint に触れていない）だけである。
準備完了の後は、時間切れと中断でも終了させず、標準入力を閉じて切り離す。切り離された作業プロセスは、
実行中の変換が終わると開いた複製を閉じ、残存記録を消し、排他を放して終わる。端末の Ctrl+C は作業プロセスに届かず、
Python が先に終わっても、作業プロセスは後片付けまで続く（`_start_worker`）。
"""

from __future__ import annotations

import base64
from collections.abc import Callable, Mapping
import contextlib
import copy
from datetime import datetime, timezone
import json
from pathlib import Path
import queue
import shutil
import subprocess
import threading
import time
from typing import Any
import uuid

from .. import renderers as base
from ..files import file_hash, write_json

# 別の変換が終わるのを待つ上限の既定値（秒）。
DEFAULT_WAIT_SECONDS = 600
# 準備完了を待つ上限は、待ち時間にこの秒数を加えたもの（PowerShell の起動の分）。
READY_MARGIN_SECONDS = 60
# 事前確認と後片付けの応答を待つ上限（秒）。変換は設定の timeout（既定 90 秒）。
DOCTOR_TIMEOUT_SECONDS = 60
CLOSE_TIMEOUT_SECONDS = 60
# 後片付けの応答の後、作業プロセスの終了を待つ上限（秒）。
EXIT_WAIT_SECONDS = 10
# 不正な終了のときに文言へ添える、プロトコル以外の出力の末尾の文字数。
TAIL_CHARACTERS = 3000
# 待つ側に見せる情報の上限（UTF-8 の JSON のバイト数）。
HOLDER_LIMIT_BYTES = 2048

_PREFIXES = (("SLIDEKIT_EVENT:", "event"), ("SLIDEKIT_JSON:", "response"))

# 切り離した作業プロセス。終了を回収するまで参照を残す。
_DETACHED: list = []


def _start_worker(command: list[str]) -> subprocess.Popen:
    """PowerShell の作業プロセスを起動する。標準入力・標準出力・標準エラーはパイプにする（試験はこの関数を差し替える）。

    切り離した作業プロセスが後片付けまで続くよう、次の 2 つを指定する（POSIX だけで働き、Windows の Python では無視される）。
    start_new_session：端末の Ctrl+C（処理グループ全体への SIGINT）を作業プロセスに届けない。中断は Python が受け、
    準備完了の前は終了、後は切り離しにする（`_interrupt`）。
    restore_signals=False：SIGPIPE を無視したまま起動する。WSL では powershell.exe の入出力を Linux 側の中継の処理が
    受け持つので、Python が終わって標準出力の読み口が閉じた後に作業プロセスが応答を書くと、既定の SIGPIPE で中継の処理が
    終わり、Windows 側の PowerShell も後片付けの前に終わらされる。
    """
    try:
        return subprocess.Popen(
            command,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            start_new_session=True,
            restore_signals=False,
        )
    except OSError as exc:
        raise base.RendererError(f"表示ソフトを実行できません: {exc}", code="renderer_unavailable") from exc


def _encode_holder(holder: Mapping | None) -> tuple[dict, str]:
    """待つ側に見せる情報を JSON にし、作業プロセスへ渡す base64 の文字列と合わせて返す。"""
    if holder is not None and not isinstance(holder, Mapping):
        raise TypeError("holder は辞書で指定してください")
    text = json.dumps(dict(holder or {}), ensure_ascii=False, default=str)
    data = text.encode("utf-8")
    if len(data) > HOLDER_LIMIT_BYTES:
        raise ValueError(f"holder は JSON で {HOLDER_LIMIT_BYTES} バイトまでにしてください（{len(data)} バイト）")
    return json.loads(text), base64.b64encode(data).decode("ascii")


def _close_quietly(stream: Any) -> None:
    if stream is not None:
        with contextlib.suppress(OSError, ValueError):
            stream.close()


def _reap_detached() -> None:
    """切り離した作業プロセスのうち、終わったものを回収する。"""
    _DETACHED[:] = [process for process in _DETACHED if process.poll() is None]


class _Tail:
    """出力の末尾を TAIL_CHARACTERS 文字まで残す。"""

    def __init__(self) -> None:
        self._text = ""
        self._lock = threading.Lock()

    def add(self, line: str) -> None:
        with self._lock:
            self._text = (self._text + line + "\n")[-TAIL_CHARACTERS:]

    def text(self) -> str:
        with self._lock:
            return self._text


def _read_output(stream: Any, lines: queue.Queue, tail: _Tail) -> None:
    """標準出力を 1 行ずつ読み、途中経過と応答を待ち行列へ入れる。それ以外の行は末尾だけを残す。"""
    try:
        for raw in iter(stream.readline, b""):
            text = base._decode(raw).rstrip("\r\n")
            kind = next((kind for prefix, kind in _PREFIXES if text.startswith(prefix)), None)
            if kind is None:
                tail.add(text)
                continue
            try:
                value = json.loads(text.split(":", 1)[1])
            except ValueError:
                value = None
            lines.put((kind, value) if isinstance(value, dict) else ("garbage", text))
    except (OSError, ValueError):
        pass
    finally:
        lines.put(("eof", None))
        _close_quietly(stream)


def _read_errors(stream: Any, tail: _Tail) -> None:
    """標準エラーを読み続け、末尾だけを残す（パイプが詰まって作業プロセスが止まらないようにする）。"""
    try:
        for raw in iter(stream.readline, b""):
            tail.add(base._decode(raw).rstrip("\r\n"))
    except (OSError, ValueError):
        pass
    finally:
        _close_quietly(stream)


class PowerPointSession(base.RendererSession):
    """PowerPoint の変換セッション。`renderers.open_session('powerpoint', …)` で開く。"""

    name = "powerpoint"

    def __init__(
        self,
        environment: Mapping,
        *,
        wait_seconds: float | None = None,
        holder: Mapping | None = None,
        on_event: Callable[[dict], None] | None = None,
        state_dir: Path | None = None,
    ) -> None:
        self.environment = environment
        self.run_id = uuid.uuid4().hex
        self.holder, self._holder_argument = _encode_holder(holder)
        self._wait_argument = wait_seconds
        self._on_event = on_event
        self._state_argument = None if state_dir is None else Path(state_dir)
        self.timeout: float | None = None
        self.wait_seconds: float | None = None
        self.work_dir: Path | None = None
        self.state_dir: Path | None = None
        self._windows_work = ""
        # new → starting → ready → closed。切り離したら detached、起動に失敗したら failed（理由は _reason）。
        self._state = "new"
        self._reason = ""
        self._ready = False
        self._process: Any = None
        self._lines: queue.Queue = queue.Queue()
        self._stdout_tail = _Tail()
        self._stderr_tail = _Tail()
        self._threads: list[threading.Thread] = []
        self._lock = threading.RLock()
        self._sequence = 0
        self._doctor: dict | None = None
        self._launches = 0
        self._renders = 0
        self._waited = 0.0
        self._abandoned = False
        self._powerpoint_started = False
        self._worker_pid: int | None = None
        self._keep_work = False
        self._closed: dict | None = None

    # 起動と準備完了

    def start(self) -> None:
        """作業プロセスを起動し、排他を取って準備完了になるまで待つ。待つ間は `waiting` の通知を on_event へ渡す。"""
        if self._state != "new":
            return
        settings = base._settings(self.name, self.environment)
        powershell = base._executable(settings.get("powershell", "powershell.exe"))
        self.timeout = base._seconds(settings.get("timeout", 90), "制限時間", positive=True)
        wait = (
            settings.get("wait_seconds", DEFAULT_WAIT_SECONDS) if self._wait_argument is None else self._wait_argument
        )
        self.wait_seconds = base._seconds(wait, "待ち時間", positive=False)
        _reap_detached()
        misses = base._cached_windows_context.cache_info().misses
        context = base.windows_context(self.environment)
        self._launches += base._cached_windows_context.cache_info().misses - misses
        self._state = "starting"
        self.work_dir = base._work_folder(context)
        self.state_dir = self._state_argument if self._state_argument is not None else base._state_folder(context)
        base._write_script(self.work_dir)
        # wslpath は作業フォルダーと状態フォルダーで 1 回ずつにし、変換ごとのファイル名は文字列でつなぐ。
        self._windows_work = base._windows_path(self.work_dir)
        command = base._script_command(
            powershell, self._windows_work, "Session", base._windows_path(self.state_dir), self.run_id
        )
        command += ["-WaitSeconds", repr(self.wait_seconds), "-Holder", self._holder_argument]
        self._process = _start_worker(command)
        self._launches += 1
        self._threads = [
            threading.Thread(
                target=_read_output,
                args=(self._process.stdout, self._lines, self._stdout_tail),
                name="slidekit-powerpoint-output",
                daemon=True,
            ),
            threading.Thread(
                target=_read_errors,
                args=(self._process.stderr, self._stderr_tail),
                name="slidekit-powerpoint-errors",
                daemon=True,
            ),
        ]
        for thread in self._threads:
            thread.start()
        try:
            self._wait_ready()
        except BaseException:
            if not self._ready and self._reason != "startup":
                self._kill()
            raise

    def _wait_ready(self) -> None:
        limit = self.wait_seconds + READY_MARGIN_SECONDS
        deadline = time.monotonic() + limit
        while True:
            kind, value = self._next(deadline)
            if kind == "event" and value.get("event") == "ready":
                self._ready = True
                self._state = "ready"
                pid = value.get("pid")
                self._worker_pid = pid if isinstance(pid, int) and not isinstance(pid, bool) else None
                waited = value.get("waited_seconds")
                self._waited = (
                    float(waited) if isinstance(waited, (int, float)) and not isinstance(waited, bool) else 0.0
                )
                self._abandoned = bool(value.get("abandoned"))
                return
            if kind == "event":
                self._notify(value)
            elif kind == "response" and value.get("id") == 0 and value.get("ok") is False:
                # 排他を待ち切れなかった・残存記録があった。作業プロセスは自分で終わる。
                self._state, self._reason = "failed", "startup"
                raise self._error(value)
            elif kind == "timeout":
                self._reason = "timeout"
                raise base.RendererError(
                    f"PowerShell の作業プロセスが {limit:g} 秒以内に準備を完了しませんでした。", code="renderer_timeout"
                )
            else:
                raise self._protocol(kind, value)

    # 公開の操作

    def doctor(self) -> dict:
        """事前確認（COM を起動しない）とフォントの照合。結果はセッションの中で使い回す。"""
        with self._lock:
            if self._doctor is None:
                self._require_ready()
                number = self._send("doctor")
                try:
                    response = self._receive(number, DOCTOR_TIMEOUT_SECONDS)
                except base.RendererError:
                    raise
                except BaseException:
                    self._detach("interrupted")
                    raise
                if response is None:
                    self._detach("timeout")
                    raise base.RendererError("PowerPoint の事前確認が制限時間を超えました。", code="renderer_timeout")
                if not response.get("ok"):
                    raise self._error(response)
                result = {key: value for key, value in response.items() if key != "id"}
                installed = result.pop("installed_fonts", None) or []
                result["font_matches"] = base.check_windows_fonts(self.environment, installed)
                self._doctor = result
            return copy.deepcopy(self._doctor)

    def render(self, input_pptx: Path, output_pdf: Path, *, label: str = "render") -> dict:
        """新しい PDF を出力する。既存の PDF は採用しない。時間切れでは残存記録を残し、作業プロセスを切り離す。"""
        input_pptx, output_pdf = base._render_paths(input_pptx, output_pdf)
        with self._lock:
            self._require_ready()
            self._renders += 1
            names = (f"input-{self._renders:04d}.pptx", f"output-{self._renders:04d}.pdf")
            staged_input, staged_output = (self.work_dir / name for name in names)
            fields = {
                "input": f"{self._windows_work}\\{names[0]}",
                "output": f"{self._windows_work}\\{names[1]}",
                "label": label,
            }
            started = time.monotonic()
            started_at = datetime.now(timezone.utc).isoformat()
            number = None
            try:
                shutil.copy2(input_pptx, staged_input)
                number = self._send("render", **fields)
                response = self._receive(number, self.timeout)
                if response is None:
                    response = self._late_response(number)
                if response is None:
                    self._leave_pending(number, fields, started_at, "timeout")
                    pending = self._pending_path
                    raise base.RendererError(
                        "PowerPoint が制限時間を超えました。利用者の資料は強制終了していません。変換が終わると記録は自動で解除されます。"
                        f"slidekit renderer status で確かめてください: {pending}",
                        code="renderer_timeout",
                        pending=str(pending),
                        work_dir=str(self.work_dir),
                    )
            except base.RendererError:
                raise
            except BaseException:
                # 変換中の中断は、時間切れと同じく残存記録を残して切り離す（状態は interrupted）。
                if number is not None and self._state == "ready":
                    self._leave_pending(number, fields, started_at, "interrupted")
                raise
            if response.get("cleanup") == "failed":
                # 開いた複製を閉じることができなかった。残存記録が後続の変換を止めるので、作業フォルダーも残す。
                self._keep_work = True
            if not response.get("ok"):
                raise self._error(response)
            self._powerpoint_started = self._powerpoint_started or bool(response.get("powerpoint_started"))
            result = {
                key: value
                for key, value in response.items()
                if key not in ("id", "cleanup", "powerpoint_started", "elapsed_seconds")
            }
            result["elapsed_seconds"] = round(time.monotonic() - started, 3)
            if not staged_output.is_file() or not base._is_pdf(staged_output):
                raise base.RendererError("PowerPoint の出力が PDF ではありません。", code="invalid_pdf")
            output_pdf.parent.mkdir(parents=True, exist_ok=True)
            # 既存の成果物には上書きせず、新しいファイルだけを採用する。
            with staged_output.open("rb") as source, output_pdf.open("xb") as target:
                shutil.copyfileobj(source, target)
            result.update(input_sha256=file_hash(staged_input), output_sha256=file_hash(output_pdf), label=label)
            staged_input.unlink(missing_ok=True)
            staged_output.unlink(missing_ok=True)
            return result

    def close(self) -> dict:
        """作業プロセスに後片付けを求め、終わるのを待つ。切り離した作業プロセスは待たない。2 回目以降は前の結果を返す。"""
        with self._lock:
            if self._closed is not None:
                return dict(self._closed)
            result: dict[str, Any] = {"renderer": self.name, "closed": True, "detached": False, "quit": None}
            try:
                if self._state == "ready":
                    try:
                        response = self._receive(self._send("close"), CLOSE_TIMEOUT_SECONDS)
                    except base.RendererError as error:
                        response = None
                        result["error"] = str(error)
                    except BaseException:
                        self._detach("interrupted")
                        raise
                    if response is None and self._state == "ready":
                        # 変換中ではないので、残存記録は作らずに切り離す。
                        self._detach("close")
                    elif response is not None:
                        self._state = "closed"
                        result["quit"] = bool(response.get("quit"))
                        if response.get("cleanup_warning"):
                            result["cleanup_warning"] = str(response["cleanup_warning"])
            finally:
                self._finish()
                result["detached"] = self._state == "detached"
                if self._reason:
                    result["reason"] = self._reason
                self._closed = result
            return dict(result)

    @property
    def operations(self) -> dict:
        """生成記録の `operations.sessions.powerpoint` の値。

        powershell_launches はセッションが起動した PowerShell の数（作業プロセスと、セッションの中で初めて行った Windows への
        問い合わせ）である。セッションの前に Windows へ問い合わせた処理（フォントの自動探索など）の起動は含めない。
        """
        return {
            "powershell_launches": self._launches,
            "renders": self._renders,
            "waited_seconds": self._waited,
            "powerpoint_started": self._powerpoint_started,
            "worker_pid": self._worker_pid,
            "abandoned": self._abandoned,
        }

    def _interrupt(self) -> None:
        """利用者の中断。準備完了の前なら作業プロセスを終了させ、後なら切り離す。"""
        with self._lock:
            if self._state == "ready":
                self._detach("interrupted")
            elif self._state == "starting":
                self._kill()

    # リクエストと応答

    def _require_ready(self) -> None:
        if self._state == "ready":
            return
        if self._state == "detached" and self._reason == "interrupted":
            raise base.RendererError(
                "変換は中断されました。このセッションでは続けられません。", code="renderer_interrupted"
            )
        if self._state == "detached" and self._reason == "timeout":
            raise base.RendererError(
                "前の処理が制限時間を超えたため、このセッションでは続けられません。", code="renderer_timeout"
            )
        if self._state in ("detached", "failed"):
            raise base.RendererError("変換の作業プロセスを使えません。", code="renderer_protocol")
        raise base.RendererError("変換セッションが開いていません。")

    def _send(self, op: str, **fields: Any) -> int:
        """リクエストを 1 行の JSON で送り、その番号を返す。"""
        self._sequence += 1
        number = self._sequence
        line = json.dumps({"id": number, "op": op, **fields}, ensure_ascii=True) + "\n"
        try:
            self._process.stdin.write(line.encode("ascii"))
            self._process.stdin.flush()
        except (OSError, ValueError):
            raise self._protocol("eof", None) from None
        return number

    def _receive(self, number: int, seconds: float) -> dict | None:
        """同じ番号の応答を返す。制限時間を過ぎたら None を返す（作業プロセスはまだ切り離さない）。"""
        deadline = time.monotonic() + seconds
        while True:
            kind, value = self._next(deadline)
            if kind == "timeout":
                return None
            if kind == "event" and value.get("event") != "ready":
                self._notify(value)
            elif kind == "response" and value.get("id") == number:
                return value
            else:
                raise self._protocol(kind, value)

    def _next(self, deadline: float) -> tuple[str, Any]:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            return "timeout", None
        try:
            return self._lines.get(timeout=remaining)
        except queue.Empty:
            return "timeout", None

    def _late_response(self, number: int) -> dict | None:
        """時間切れの直後までに届いた応答があれば返す（待たない）。"""
        while True:
            try:
                kind, value = self._lines.get_nowait()
            except queue.Empty:
                return None
            if kind == "response" and value.get("id") == number:
                return value
            if kind != "event":
                self._lines.put((kind, value))
                return None

    def _notify(self, value: dict) -> None:
        event = dict(value)
        if "holder" in event:
            event["holder"] = base._holder(event["holder"])
        if self._on_event is not None:
            self._on_event(event)

    def _error(self, value: dict) -> base.RendererError:
        """失敗の応答を例外にする。使用中なら待った時間と使用中の処理、残存記録ならその場所を詳細に入れる。"""
        error = base._failure(value)
        if error.code == "renderer_busy":
            error.details.update(holder=base._holder(value.get("holder")), waited_seconds=value.get("waited_seconds"))
        elif error.code == "renderer_pending":
            error.details.update(pending=str(self._pending_path), record=value.get("pending"))
        return error

    def _protocol(self, kind: str, value: Any) -> base.RendererError:
        """不正な出力か途中の終了。準備完了の後なら作業プロセスを切り離す（終了させない）。"""
        if kind == "eof":
            # 終了の直前に標準エラーへ出た文言も添える。
            for thread in self._threads[1:]:
                thread.join(timeout=0.5)
            message = "PowerShell の作業プロセスが途中で終了しました。"
        elif kind == "garbage":
            message = f"PowerShell の作業プロセスの出力が不正です: {str(value)[:200]}"
        elif kind == "event":
            message = "PowerShell の作業プロセスが想定外の通知を出しました。"
        else:
            message = "PowerShell の作業プロセスの応答が要求と合いません。"
            if isinstance(value, dict) and value.get("message"):
                message += f' {value["message"]}'
        if self._state == "ready":
            self._detach("protocol")
        else:
            self._reason = self._reason or "protocol"
        detail = (self._stdout_tail.text() + self._stderr_tail.text()).strip()[-TAIL_CHARACTERS:]
        return base.RendererError(
            f"{message} 出力: {detail}" if detail else message, code="renderer_protocol", output=detail
        )

    # 残存記録

    @property
    def _pending_path(self) -> Path:
        return self.state_dir / "pending.json"

    def _leave_pending(self, number: int, fields: dict, started_at: str, status: str) -> None:
        """変換中のリクエストを残したまま切り離す。残存記録が後続の変換を止め、作業プロセスが変換の後に消す。

        このセッションの記録なら status を書き換え、なければ作る。別の実行の記録は書き換えない。
        書いた直後に変換の終了の応答が届いていれば、作業プロセスは記録を消し終えているので、書いた記録を消す。
        """
        if self._late_response(number) is None:
            self._mark_pending(fields, started_at, status)
            if self._late_response(number) is not None:
                self._remove_own_pending(status)
        self._detach(status)

    def _mark_pending(self, fields: dict, started_at: str, status: str) -> None:
        path = self._pending_path
        try:
            current = json.loads(path.read_text(encoding="utf-8-sig"))
        except FileNotFoundError:
            current = None
        except (OSError, ValueError):
            # 読み込みに失敗した記録は、ほかの実行のものとして残す。
            return
        try:
            if current is None:
                record = {
                    "schema_version": 1,
                    "run_id": self.run_id,
                    "process_id": self._worker_pid,
                    "mode": "Render",
                    "label": fields["label"],
                    "input": fields["input"],
                    "output": fields["output"],
                    "work_dir": self._windows_work,
                    "started_at": started_at,
                    "status": status,
                    "holder": self.holder,
                }
                path.parent.mkdir(parents=True, exist_ok=True)
                with path.open("x", encoding="utf-8") as stream:
                    json.dump(record, stream, ensure_ascii=False)
            elif isinstance(current, dict) and current.get("run_id") == self.run_id:
                write_json(
                    path, {**current, "status": status, "process_id": current.get("process_id") or self._worker_pid}
                )
        except OSError:
            pass

    def _remove_own_pending(self, status: str) -> None:
        path = self._pending_path
        with contextlib.suppress(OSError, ValueError):
            current = json.loads(path.read_text(encoding="utf-8-sig"))
            if isinstance(current, dict) and current.get("run_id") == self.run_id and current.get("status") == status:
                path.unlink(missing_ok=True)

    # 切り離しと後片付け

    def _detach(self, reason: str) -> None:
        """作業プロセスを終了させずに切り離す。標準入力を閉じると、作業プロセスは実行中の処理の後に片付けて終わる。"""
        self._state, self._reason = "detached", reason
        self._keep_work = True
        if self._process is not None:
            _close_quietly(self._process.stdin)
            _DETACHED.append(self._process)

    def _kill(self) -> None:
        """準備完了の前の作業プロセスを終了させる。排他を待っているだけで、PowerPoint には触れていない。"""
        self._state = "failed"
        self._reason = self._reason or "interrupted"
        process = self._process
        if process is not None and process.poll() is None:
            with contextlib.suppress(OSError):
                process.kill()

    def _finish(self) -> None:
        """作業プロセスの終了を待ち（切り離したものは待たない）、作業フォルダーを片付ける。"""
        process = self._process
        if process is not None and self._state != "detached":
            _close_quietly(process.stdin)
            try:
                process.wait(timeout=EXIT_WAIT_SECONDS)
            except subprocess.TimeoutExpired:
                if self._ready:
                    self._detach(self._reason or "close")
                else:
                    self._kill()
                    with contextlib.suppress(subprocess.TimeoutExpired):
                        process.wait(timeout=EXIT_WAIT_SECONDS)
            if self._state != "detached":
                for thread in self._threads:
                    thread.join(timeout=1)
        if self.work_dir is not None and not self._keep_work:
            shutil.rmtree(self.work_dir, ignore_errors=True)
