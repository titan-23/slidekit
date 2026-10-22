"""変換の失敗・排他制御・変換セッション・フォント照合を、実ソフトとは別に確かめる。

PowerShell の作業プロセスは実際のパイプで同じ形式の行を返す偽物（FakeWorker）で置き換える。PowerPoint は使わない。
"""

import base64
import contextlib
import json
import os
from pathlib import Path
import subprocess
import threading
from types import SimpleNamespace

import pytest

from slidekit import renderers
from slidekit.files import file_hash
from slidekit.renderers import powerpoint

WORKER_PID = 4242

# 別の処理の使用中の記録（active.json）と、それから作る、待つ側に見せる情報。
ACTIVE = {
    "schema_version": 1,
    "pid": 1234,
    "run_id": "other-session",
    "started_at": "2026-10-12T00:00:00Z",
    "work_dir": "W:\\other",
    "holder": {
        "project": "/work/deck.toml",
        "run": "/work/out/runs/0011",
        "command": "build",
        "host": "example-pc",
        "python_pid": 99,
    },
}
HOLDER_VIEW = {
    "project": "/work/deck.toml",
    "run": "/work/out/runs/0011",
    "command": "build",
    "host": "example-pc",
    "python_pid": 99,
    "pid": 1234,
    "started_at": "2026-10-12T00:00:00Z",
    "run_id": "other-session",
}


@pytest.fixture(autouse=True)
def fresh_windows_context():
    """Windows への問い合わせの記憶を試験ごとに消す（模擬の結果を後の試験へ残さない）。"""
    renderers._cached_windows_context.cache_clear()
    yield
    renderers._cached_windows_context.cache_clear()


def result(value, returncode=0):
    return subprocess.CompletedProcess([], returncode, ("SLIDEKIT_JSON:" + json.dumps(value)).encode(), b"")


def windows(path):
    """試験用の Windows のパス。Unix のパスの区切りを \\ にして W: を付ける（unix() で元に戻すことができる）。"""
    return "W:" + str(path).replace("/", "\\")


def unix(text):
    assert text.startswith("W:\\"), text
    return Path(text[2:].replace("\\", "/"))


def synthetic_font(path, family, style="Regular", width=500):
    """字形 A だけを持つ小さな TrueType を作る。width を変えると表のハッシュ値が変わる。"""
    from fontTools.fontBuilder import FontBuilder
    from fontTools.pens.ttGlyphPen import TTGlyphPen

    builder = FontBuilder(1024, isTTF=True)
    builder.setupGlyphOrder([".notdef", "A"])
    builder.setupCharacterMap({65: "A"})
    pen = TTGlyphPen(None)
    builder.setupGlyf({".notdef": pen.glyph(), "A": pen.glyph()})
    builder.setupHorizontalMetrics({".notdef": (500, 0), "A": (width, 0)})
    builder.setupHorizontalHeader(ascent=800, descent=-200)
    name = family if style == "Regular" else f"{family} {style}"
    builder.setupNameTable(
        {
            "familyName": family,
            "styleName": style,
            "psName": name.replace(" ", "-"),
            "fullName": name,
            "uniqueFontIdentifier": name,
        }
    )
    builder.setupOS2()
    builder.setupPost()
    builder.save(path)
    return path


def ready(worker):
    worker.ready()


def respond_normally(worker, request):
    """作業スクリプトの正常な応答。変換では出力先へ PDF を書く。"""
    if request["op"] == "doctor":
        return {
            "id": request["id"],
            "ok": True,
            "renderer": "powerpoint",
            "version": "test-only",
            "executable": "C:\\Office\\POWERPNT.EXE",
            "windows_user_sid": "S-1-fake",
            "installed_fonts": [],
            "state_dir": worker.option("-StateDir"),
        }
    if request["op"] == "render":
        unix(request["output"]).write_bytes(b"%PDF-1.7\nfake PDF")
        started, worker.started_powerpoint = not worker.started_powerpoint, True
        return {
            "id": request["id"],
            "ok": True,
            "renderer": "powerpoint",
            "version": "16.0",
            "build": "test",
            "windows_user_sid": "S-1-fake",
            "run_id": worker.option("-RunId"),
            "powerpoint_started": started,
            "cleanup": "ok",
        }
    return {"id": request["id"], "ok": False, "message": f'未対応の要求です: {request["op"]}'}


class FakeWorker:
    """PowerShell の作業プロセスの代わり。実際のパイプを使い、作業スクリプトと同じ形式の行で応答する。

    startup(worker) は準備完了までの出力を書く（False を返すと起動で終わる）。respond(worker, request) はリクエストごとの
    応答の辞書を返す（None なら応答しない、'exit' なら出力を閉じて終わる）。close には後片付けの応答を返して終わる。
    """

    def __init__(self, command, *, startup=None, respond=None):
        self.command = command
        self.requests = []
        self.killed = False
        self.returncode = None
        self.pid = WORKER_PID
        self.started_powerpoint = False
        # 応答を止めた偽物を、試験の中か終わりに放す。
        self.release = threading.Event()
        self._startup = startup or ready
        self._respond = respond or respond_normally
        self._lock = threading.Lock()
        self._done = threading.Event()
        (requests, stdin), (stdout, output), (stderr, errors) = os.pipe(), os.pipe(), os.pipe()
        self.stdin, self._requests = os.fdopen(stdin, "wb"), os.fdopen(requests, "rb")
        self.stdout, self._output = os.fdopen(stdout, "rb"), os.fdopen(output, "wb")
        self.stderr, self._errors = os.fdopen(stderr, "rb"), os.fdopen(errors, "wb")
        self._thread = threading.Thread(target=self._serve, daemon=True)
        self._thread.start()

    def option(self, name):
        return self.command[self.command.index(name) + 1]

    def line(self, text, *, stream="out"):
        with self._lock:
            target = self._output if stream == "out" else self._errors
            with contextlib.suppress(OSError, ValueError):
                target.write((text + "\r\n").encode("utf-8"))
                target.flush()

    def event(self, **value):
        self.line("SLIDEKIT_EVENT:" + json.dumps(value, ensure_ascii=False))

    def send(self, **value):
        self.line("SLIDEKIT_JSON:" + json.dumps(value, ensure_ascii=False))

    def ready(self, **value):
        self.event(
            **{
                "event": "ready",
                "pid": self.pid,
                "sid": "S-1-fake",
                "state_dir": self.option("-StateDir"),
                "waited_seconds": 0.0,
                "abandoned": False,
                **value,
            }
        )

    def _serve(self):
        try:
            if self._startup(self) is False or self.killed:
                return
            for raw in self._requests:
                request = json.loads(raw)
                self.requests.append(request)
                if request["op"] == "close":
                    self.send(id=request["id"], ok=True, quit=self.started_powerpoint)
                    return
                response = self._respond(self, request)
                if response == "exit" or self.killed:
                    return
                if response is not None:
                    self.send(**response)
        except (OSError, ValueError):
            pass
        finally:
            if self.returncode is None:
                self.returncode = 0
            self._close_output()
            self._done.set()

    def _close_output(self):
        with self._lock:
            for stream in (self._output, self._errors):
                with contextlib.suppress(OSError, ValueError):
                    stream.close()

    def poll(self):
        return self.returncode if self._done.is_set() else None

    def wait(self, timeout=None):
        if not self._done.wait(timeout):
            raise subprocess.TimeoutExpired(self.command, timeout)
        return self.returncode

    def kill(self):
        self.killed = True
        self.returncode = -9
        self.release.set()
        self._close_output()

    def dispose(self):
        """試験の終わりに、止めた応答を放し、パイプをすべて閉じる。"""
        self.release.set()
        with contextlib.suppress(OSError, ValueError):
            self.stdin.close()
        self._thread.join(timeout=5)
        for stream in (self._requests, self.stdout, self.stderr):
            with contextlib.suppress(OSError, ValueError):
                stream.close()


@pytest.fixture
def fake_windows(tmp_path, monkeypatch):
    """Windows の一時領域と wslpath を模倣し、偽の作業プロセスを起動させる。PowerPoint の実行確認ではない。"""
    temporary, local = tmp_path / "日本語を含む TEMP", tmp_path / "Local App Data"
    temporary.mkdir()
    local.mkdir()
    state = SimpleNamespace(
        temporary=temporary,
        local=local,
        pending=local / "slidekit" / "powerpoint" / "pending.json",
        context_calls=0,
        unix_calls=[],
        windows_calls=[],
        workers=[],
        startup=None,
        respond=None,
    )

    def context(powershell):
        state.context_calls += 1
        return {
            "temp": str(temporary),
            "local_app_data": str(local),
            "sid": "S-1-fake",
            "fonts": str(tmp_path / "Fonts"),
        }

    def to_unix(value):
        state.unix_calls.append(value)
        return Path(value)

    def to_windows(path):
        state.windows_calls.append(Path(path))
        return windows(Path(path).resolve())

    def start(command):
        worker = FakeWorker(command, startup=state.startup, respond=state.respond)
        state.workers.append(worker)
        return worker

    monkeypatch.setattr(renderers, "_executable", lambda value: f"/fake/{value}")
    monkeypatch.setattr(renderers, "_windows_context", context)
    monkeypatch.setattr(renderers, "_unix_path", to_unix)
    monkeypatch.setattr(renderers, "_windows_path", to_windows)
    monkeypatch.setattr(powerpoint, "_start_worker", start)
    yield state
    for worker in state.workers:
        worker.dispose()


def blocking_render(*, write_record=True, write_pdf=False):
    """変換のリクエストに、放されるまで応答しない作業プロセス。放された後は変換を終えたものとして応答する。"""

    def respond(worker, request):
        if request["op"] != "render":
            return respond_normally(worker, request)
        pending = unix(worker.option("-StateDir")) / "pending.json"
        if write_record:
            # 作業スクリプトと同じく、開く前に残存記録を書く（BOM 付き UTF-8 でも読み込みに成功すること）。
            pending.parent.mkdir(parents=True, exist_ok=True)
            pending.write_text(
                json.dumps(
                    {
                        "schema_version": 1,
                        "run_id": worker.option("-RunId"),
                        "process_id": worker.pid,
                        "mode": "Render",
                        "label": request["label"],
                        "input": request["input"],
                        "output": request["output"],
                        "work_dir": worker.option("-WorkDir"),
                        "started_at": "2026-10-12T00:00:00Z",
                        "status": "running",
                        "holder": {},
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8-sig",
            )
        if write_pdf:
            unix(request["output"]).write_bytes(b"%PDF-1.7\nunfinished")
        worker.release.wait(10)
        if worker.killed:
            return None
        response = respond_normally(worker, request)
        pending.unlink(missing_ok=True)
        return response

    return respond


def test_missing_renderer_does_not_fall_back(monkeypatch):
    monkeypatch.setattr(renderers.shutil, "which", lambda _: None)
    with pytest.raises(renderers.RendererError, match="見つかりません") as error:
        renderers.doctor_renderer("powerpoint", {})
    assert error.value.code == "renderer_missing"


def test_executable_preserves_symlink_dispatch_name(tmp_path, monkeypatch):
    target = tmp_path / "init"
    target.write_text("dispatcher")
    link = tmp_path / "wslpath"
    link.symlink_to(target)
    monkeypatch.setattr(renderers.shutil, "which", lambda _: str(link))
    assert renderers._executable("wslpath") == str(link)


def test_existing_output_is_never_reused(tmp_path):
    source = tmp_path / "deck.pptx"
    source.write_bytes(b"pptx")
    output = tmp_path / "deck.pdf"
    output.write_bytes(b"old PDF")
    with pytest.raises(renderers.RendererError) as error:
        renderers.render_pdf(source, output, "powerpoint", {})
    assert error.value.code == "output_exists"
    assert output.read_bytes() == b"old PDF"


@pytest.mark.parametrize(
    "message,code",
    [
        ("SLIDEKIT_BUSY: busy", "renderer_busy"),
        ("SLIDEKIT_PENDING: pending", "renderer_pending"),
        ("conversion failed", "renderer_failed"),
    ],
)
def test_worker_errors_are_not_success(message, code):
    with pytest.raises(renderers.RendererError) as error:
        renderers._result(result({"ok": False, "message": message}, 2))
    assert error.value.code == code


def test_bad_worker_output_is_not_success():
    process = subprocess.CompletedProcess([], 0, b"PowerPoint started", b"")
    with pytest.raises(renderers.RendererError, match="実行結果を取得できません"):
        renderers._result(process)


def test_session_reuses_one_worker_for_doctor_and_renders(tmp_path, fake_windows):
    source = tmp_path / "deck.pptx"
    source.write_bytes(b"source")
    with renderers.open_session("powerpoint", {"fonts": {}}) as session:
        doctor = session.doctor()
        first = session.render(source, tmp_path / "standard.pdf", label="standard")
        second = session.render(source, tmp_path / "fallback.pdf", label="fallback")
        # 2 回目の事前確認は作業プロセスへ問い合わせない。
        assert session.doctor() == doctor
        operations = session.operations
    (worker,) = fake_windows.workers
    assert [request["op"] for request in worker.requests] == ["doctor", "render", "render", "close"]
    assert fake_windows.context_calls == 1
    assert operations == {
        "powershell_launches": 2,
        "renders": 2,
        "waited_seconds": 0.0,
        "powerpoint_started": True,
        "worker_pid": WORKER_PID,
        "abandoned": False,
    }
    assert worker.option("-Mode") == "Session"
    work = worker.option("-WorkDir")
    assert [request["input"] for request in worker.requests[1:3]] == [
        work + "\\input-0001.pptx",
        work + "\\input-0002.pptx",
    ]
    assert (first["label"], second["label"]) == ("standard", "fallback")
    assert first["run_id"] == second["run_id"] == worker.option("-RunId")
    assert (tmp_path / "standard.pdf").read_bytes().startswith(b"%PDF-")
    assert (first["input_sha256"], first["output_sha256"]) == (file_hash(source), file_hash(tmp_path / "standard.pdf"))
    # 閉じると作業プロセスの終了を待ち、作業フォルダーを消す。
    assert worker.poll() == 0 and not worker.killed
    assert not list((fake_windows.temporary / "slidekit").iterdir())


def test_windows_context_is_cached_per_process(fake_windows):
    with renderers.open_session("powerpoint", {}) as first:
        first_operations = first.operations
    with renderers.open_session("powerpoint", {}) as second:
        second_operations = second.operations
    assert fake_windows.context_calls == 1 and len(fake_windows.workers) == 2
    assert (first_operations["powershell_launches"], second_operations["powershell_launches"]) == (2, 1)
    # PowerShell の実行ファイルが違えば、別に問い合わせる。
    renderers.windows_context({"renderers": {"powerpoint": {"powershell": "pwsh.exe"}}})
    renderers.windows_context({})
    assert fake_windows.context_calls == 2


def test_windows_context_is_cached_and_reports_fonts(monkeypatch):
    launches = []

    def fake_run(command, **kwargs):
        launches.append(command)
        return result(
            {
                "ok": True,
                "temp": "C:\\Temp\\",
                "local_app_data": "C:\\Users\\example\\AppData\\Local",
                "sid": "S-1-5-21-example",
                "fonts": "C:\\Windows\\Fonts",
            }
        )

    monkeypatch.setattr(renderers, "_executable", lambda value: f"/fake/{value}")
    monkeypatch.setattr(renderers, "_run", fake_run)
    first = renderers.windows_context({})
    # 返した値を書き換えても、覚えた値は変わらない。
    first["fonts"] = "changed"
    assert renderers.windows_context({"renderers": {"powerpoint": {}}}) == {
        "temp": "C:\\Temp\\",
        "local_app_data": "C:\\Users\\example\\AppData\\Local",
        "sid": "S-1-5-21-example",
        "fonts": "C:\\Windows\\Fonts",
    }
    assert len(launches) == 1
    script = base64.b64decode(launches[0][launches[0].index("-EncodedCommand") + 1]).decode("utf-16-le")
    assert "[Environment]::GetFolderPath('Fonts')" in script


def test_session_paths_use_one_wslpath_call_per_folder(tmp_path, fake_windows):
    source = tmp_path / "deck.pptx"
    source.write_bytes(b"source")
    with renderers.open_session("powerpoint", {}) as session:
        for name in ("a.pdf", "b.pdf", "c.pdf"):
            session.render(source, tmp_path / name)
    (worker,) = fake_windows.workers
    assert fake_windows.unix_calls == [str(fake_windows.temporary), str(fake_windows.local)]
    assert [path.name for path in fake_windows.windows_calls] == [unix(worker.option("-WorkDir")).name, "powerpoint"]


def test_render_pdf_and_doctor_renderer_remain_compatible(tmp_path, fake_windows, monkeypatch):
    def respond(worker, request):
        response = respond_normally(worker, request)
        if request["op"] == "doctor":
            response["installed_fonts"] = [{"name": "Example Font", "path": "font.ttf"}]
        return response

    fake_windows.respond = respond
    monkeypatch.setattr(renderers, "check_windows_fonts", lambda environment, fonts: ["matched", fonts[0]["name"]])
    info = renderers.doctor_renderer("powerpoint", {})
    assert info["font_matches"] == ["matched", "Example Font"]
    assert set(info) == {"ok", "renderer", "version", "executable", "windows_user_sid", "state_dir", "font_matches"}
    source, target = tmp_path / "deck.pptx", tmp_path / "out" / "deck.pdf"
    source.write_bytes(b"source")
    rendered = renderers.render_pdf(source, target, "powerpoint", {})
    assert target.read_bytes().startswith(b"%PDF-")
    assert set(rendered) == {
        "ok",
        "renderer",
        "version",
        "build",
        "windows_user_sid",
        "run_id",
        "elapsed_seconds",
        "input_sha256",
        "output_sha256",
        "label",
    }
    assert (rendered["input_sha256"], rendered["label"]) == (file_hash(source), "render")
    # 1 回限りのセッションを 1 つずつ使い、閉じて作業フォルダーを消す。
    assert [[request["op"] for request in worker.requests] for worker in fake_windows.workers] == [
        ["doctor", "close"],
        ["render", "close"],
    ]
    assert not list((fake_windows.temporary / "slidekit").iterdir())
    # パスはコードに埋め込まず、引数で渡す。
    command = fake_windows.workers[1].command
    assert command[command.index("-File") + 1].endswith("\\powerpoint.ps1")
    assert len([part for part in command if "日本語を含む TEMP" in part]) == 2


@pytest.mark.parametrize(
    "response,code,keeps_work",
    [
        (
            {"ok": False, "message": "SLIDEKIT_PENDING: 前回の処理の終了を確認してください。", "cleanup": "ok"},
            "renderer_pending",
            False,
        ),
        ({"ok": False, "message": "PowerPoint で開けません。", "cleanup": "ok"}, "renderer_failed", False),
        # 開いた複製を閉じることができなかった。残存記録が残るので、作業フォルダーも残す。
        ({"ok": False, "message": "終了を確認できません。", "cleanup": "failed"}, "renderer_failed", True),
        # 成功の応答でも、出力が PDF でなければ採用しない。
        ({"ok": True, "renderer": "powerpoint", "cleanup": "ok"}, "invalid_pdf", False),
    ],
)
def test_session_render_failures_are_not_success(tmp_path, fake_windows, response, code, keeps_work):
    fake_windows.respond = lambda worker, request: {"id": request["id"], **response}
    source, target = tmp_path / "deck.pptx", tmp_path / "deck.pdf"
    source.write_bytes(b"source")
    with pytest.raises(renderers.RendererError) as error:
        renderers.render_pdf(source, target, "powerpoint", {})
    assert error.value.code == code
    if code == "renderer_pending":
        assert error.value.details["pending"] == str(fake_windows.pending)
    assert not target.exists()
    assert bool(list((fake_windows.temporary / "slidekit").iterdir())) is keeps_work


def garbage_output(worker, request):
    worker.line("PowerShell の警告の行")
    worker.line('SLIDEKIT_JSON:{"id": ')
    return None


def early_exit(worker, request):
    worker.line("終了の直前の誤り", stream="err")
    return "exit"


@pytest.mark.parametrize("respond,text", [(garbage_output, "PowerShell の警告の行"), (early_exit, "終了の直前の誤り")])
def test_worker_output_garbage_is_protocol_error(fake_windows, respond, text):
    fake_windows.respond = respond
    with pytest.raises(renderers.RendererError) as error:
        renderers.doctor_renderer("powerpoint", {})
    assert error.value.code == "renderer_protocol"
    assert text in str(error.value)
    # 準備完了の後なので終了させず、切り離す。
    (worker,) = fake_windows.workers
    assert worker.stdin.closed and not worker.killed


def libreoffice_run(calls, *, write_pdf=True):
    def fake_run(command, **kwargs):
        calls.append(command)
        if command[-1] == "--version":
            return subprocess.CompletedProcess(command, 0, b"LibreOffice 24.2\n", b"")
        if write_pdf:
            (Path(command[command.index("--outdir") + 1]) / "input.pdf").write_bytes(b"%PDF-1.7\nLibreOffice")
        return subprocess.CompletedProcess(command, 0, b"", b"")

    return fake_run


def test_libreoffice_session_reuses_doctor(tmp_path, monkeypatch):
    calls = []
    monkeypatch.setattr(renderers, "_executable", lambda value: f"/fake/{value}")
    monkeypatch.setattr(renderers, "_run", libreoffice_run(calls))
    source = tmp_path / "deck.pptx"
    source.write_bytes(b"source")
    with renderers.open_session("libreoffice", {}) as session:
        info = session.doctor()
        first = session.render(source, tmp_path / "a.pdf", label="standard")
        session.render(source, tmp_path / "b.pdf")
        operations = session.operations
    assert [command[-1] for command in calls].count("--version") == 1
    assert (info["version"], first["version"], first["label"]) == ("LibreOffice 24.2", "LibreOffice 24.2", "standard")
    assert operations == {"renders": 2, "doctor_runs": 1}
    assert (tmp_path / "b.pdf").read_bytes().startswith(b"%PDF-")


def test_libreoffice_conversion_without_output_fails(tmp_path, monkeypatch):
    source, target = tmp_path / "deck.pptx", tmp_path / "deck.pdf"
    source.write_bytes(b"source")
    monkeypatch.setattr(renderers, "_executable", lambda value: f"/fake/{value}")
    monkeypatch.setattr(renderers, "_run", libreoffice_run([], write_pdf=False))
    with pytest.raises(renderers.RendererError, match="変換に失敗"):
        renderers.render_pdf(source, target, "libreoffice", {})
    assert not target.exists()


def test_session_reports_waiting_with_holder(fake_windows):
    def startup(worker):
        worker.event(event="waiting", waited_seconds=0.0, holder=ACTIVE)
        worker.event(event="waiting", waited_seconds=5.0, holder=ACTIVE)
        worker.ready(waited_seconds=7.5)

    fake_windows.startup = startup
    events = []
    holder = {"project": "/mine/deck.toml", "run": "/mine/out/runs/0012", "command": "build"}
    with renderers.open_session("powerpoint", {}, wait_seconds=30, holder=holder, on_event=events.append) as session:
        operations = session.operations
    assert [(event["event"], event["waited_seconds"]) for event in events] == [("waiting", 0.0), ("waiting", 5.0)]
    assert events[1]["holder"] == HOLDER_VIEW
    assert operations["waited_seconds"] == 7.5
    # 待つ側に見せる自分の情報は、base64 の JSON で渡す。
    (worker,) = fake_windows.workers
    assert json.loads(base64.b64decode(worker.option("-Holder"))) == holder


def test_session_busy_after_wait_names_holder(fake_windows):
    def startup(worker):
        worker.send(
            id=0,
            ok=False,
            renderer="powerpoint",
            holder=ACTIVE,
            waited_seconds=10.0,
            message="SLIDEKIT_BUSY: 同じ Windows 利用者による変換が実行中です（PID 1234、資料 /work/deck.toml）。10.0 秒待ちました。",
        )
        return False

    fake_windows.startup = startup
    with pytest.raises(renderers.RendererError) as error:
        with renderers.open_session("powerpoint", {}, wait_seconds=10):
            pass
    assert error.value.code == "renderer_busy"
    assert "PID 1234" in str(error.value)
    assert (error.value.details["holder"], error.value.details["waited_seconds"]) == (HOLDER_VIEW, 10.0)
    # 作業プロセスは自分で終わる。終了させず、リクエストも送らない。
    (worker,) = fake_windows.workers
    assert not worker.killed and worker.requests == []
    assert not list((fake_windows.temporary / "slidekit").iterdir())


def test_wait_seconds_is_passed_to_worker_and_validated(fake_windows, isolated_environment):
    from slidekit.config import load_common_environment

    isolated_environment.write_text(
        "schema_version = 1\n[renderers.powerpoint]\nwait_seconds = 2.5\n", encoding="utf-8"
    )
    configured = load_common_environment()
    # 既定は 600 秒。設定の値を使い、引数があれば引数を優先する。
    for environment, argument in (({}, None), (configured, None), (configured, 0)):
        with renderers.open_session("powerpoint", environment, wait_seconds=argument):
            pass
    assert [float(worker.option("-WaitSeconds")) for worker in fake_windows.workers] == [600.0, 2.5, 0.0]
    # 0 以上の有限の秒数でなければ、作業プロセスを起動しない（設定ファイルの値は読み込みでも拒否する。test_environment.py）。
    for environment, argument in (
        ({}, -1),
        ({}, float("nan")),
        ({}, float("inf")),
        ({}, True),
        ({}, "10"),
        ({"renderers": {"powerpoint": {"wait_seconds": -1}}}, None),
    ):
        with pytest.raises(renderers.RendererError) as error:
            with renderers.open_session("powerpoint", environment, wait_seconds=argument):
                pass
        assert error.value.code == "invalid_renderer"
    assert len(fake_windows.workers) == 3


@pytest.fixture
def maintenance(tmp_path, monkeypatch):
    """renderer status・renderer clear の作業スクリプトの実行（_run）を模倣する。"""
    temporary = tmp_path / "TEMP"
    temporary.mkdir()
    state = SimpleNamespace(commands=[], value={"ok": True}, temporary=temporary)

    def fake_run(command, **kwargs):
        # 作業スクリプトを作業フォルダーへ書いてから実行する。
        assert unix(command[command.index("-File") + 1]).is_file()
        state.commands.append(command)
        return result(state.value)

    monkeypatch.setattr(renderers, "_executable", lambda value: f"/fake/{value}")
    monkeypatch.setattr(
        renderers,
        "_windows_context",
        lambda powershell: {
            "temp": str(temporary),
            "local_app_data": str(tmp_path / "Local"),
            "sid": "S-1-fake",
            "fonts": "",
        },
    )
    monkeypatch.setattr(renderers, "_unix_path", lambda value: Path(value))
    monkeypatch.setattr(renderers, "_windows_path", lambda path: windows(Path(path).resolve()))
    monkeypatch.setattr(renderers, "_run", fake_run)
    return state


def test_status_parses_worker_result(tmp_path, maintenance):
    record = {"run_id": "previous", "process_id": 77, "input": "W:\\work\\input-0001.pptx", "status": "timeout"}
    maintenance.value = {
        "ok": True,
        "busy": False,
        "holder": None,
        "pending": record,
        "pending_exists": True,
        "pending_process_alive": False,
        "staged_input_locked": False,
        "powerpoint_running": True,
        "abandoned": True,
        "state_dir": "W:\\state",
    }
    status = renderers.renderer_status("powerpoint", {}, state_dir=tmp_path / "state")
    assert status == {
        "command": "renderer status",
        "renderer": "powerpoint",
        "status": "pending",
        "busy": False,
        "holder": None,
        "pending": record,
        "pending_process_alive": False,
        "staged_input_locked": False,
        "powerpoint_running": True,
        "abandoned": True,
        "state_dir": "W:\\state",
        "exit_code": 1,
    }
    command = maintenance.commands[-1]
    assert command[command.index("-Mode") + 1] == "Status"
    assert command[command.index("-StateDir") + 1] == windows((tmp_path / "state").resolve())
    assert "-Force" not in command
    # 使用中なら busy で、使用中の処理の情報を示す。
    maintenance.value = {
        "ok": True,
        "busy": True,
        "holder": ACTIVE,
        "pending": None,
        "pending_exists": False,
        "powerpoint_running": True,
    }
    busy = renderers.renderer_status("powerpoint", {})
    assert (busy["status"], busy["holder"], busy["exit_code"]) == ("busy", HOLDER_VIEW, 1)
    # 読み込みに失敗した残存記録も残存記録として扱う。何もなければ idle。
    maintenance.value = {"ok": True, "busy": False, "pending": None, "pending_exists": True}
    assert renderers.renderer_status("powerpoint", {})["status"] == "pending"
    maintenance.value = {
        "ok": True,
        "busy": False,
        "pending": None,
        "pending_exists": False,
        "powerpoint_running": False,
    }
    idle = renderers.renderer_status("powerpoint", {})
    assert (idle["status"], idle["exit_code"], idle["abandoned"]) == ("idle", 0, False)
    # 作業フォルダーは毎回消す。LibreOffice は PowerShell を使わない。
    assert not list((maintenance.temporary / "slidekit").iterdir())
    count = len(maintenance.commands)
    assert renderers.renderer_status("libreoffice", {})["status"] == "idle"
    assert len(maintenance.commands) == count


def test_clear_refuses_when_worker_reports_busy_or_locked(maintenance):
    maintenance.value = {
        "ok": True,
        "status": "refused",
        "reason": "変換の実行中のため解除できません（PID 1234）。",
        "holder": ACTIVE,
        "archived": None,
        "removed_work_dir": False,
        "removed_old_work_dirs": 0,
    }
    busy = renderers.clear_pending("powerpoint", {})
    assert (busy["status"], busy["exit_code"], busy["holder"], busy["archived"]) == ("refused", 1, HOLDER_VIEW, None)
    assert "PID 1234" in busy["reason"]
    maintenance.value = {
        "ok": True,
        "status": "refused",
        "pending": {"run_id": "previous"},
        "reason": "PowerPoint が検査用の複製を開いています。PowerPoint で閉じてから再実行してください。",
    }
    locked = renderers.clear_pending("powerpoint", {}, force=True)
    assert (locked["status"], locked["exit_code"], locked["pending"]) == ("refused", 1, {"run_id": "previous"})
    assert "複製を開いています" in locked["reason"]
    # 結果の形が不正なら、解除できたとみなさない。作業スクリプトの失敗は例外にする。
    maintenance.value = {"ok": True, "status": "done"}
    with pytest.raises(renderers.RendererError) as error:
        renderers.clear_pending("powerpoint", {})
    assert error.value.code == "renderer_protocol"
    maintenance.value = {"ok": False, "message": "解除に失敗しました。"}
    with pytest.raises(renderers.RendererError) as error:
        renderers.clear_pending("powerpoint", {})
    assert error.value.code == "renderer_failed"


def test_clear_passes_force_flag(tmp_path, maintenance):
    archived = "W:\\state\\cleared\\20261012T000000000Z-pending.json"
    maintenance.value = {
        "ok": True,
        "status": "cleared",
        "reason": "残存記録を解除しました。",
        "archived": archived,
        "removed_work_dir": True,
        "removed_old_work_dirs": 2,
        "pending": {"run_id": "previous"},
    }
    cleared = renderers.clear_pending("powerpoint", {}, force=True, state_dir=tmp_path / "state")
    renderers.clear_pending("powerpoint", {})
    forced, plain = maintenance.commands
    assert forced[forced.index("-Mode") + 1] == plain[plain.index("-Mode") + 1] == "Clear"
    assert forced[-1] == "-Force" and "-Force" not in plain
    assert cleared == {
        "command": "renderer clear",
        "renderer": "powerpoint",
        "status": "cleared",
        "reason": "残存記録を解除しました。",
        "archived": archived,
        "removed_work_dir": True,
        "removed_old_work_dirs": 2,
        "pending": {"run_id": "previous"},
        "holder": None,
        "exit_code": 0,
    }
    maintenance.value = {"ok": True, "status": "nothing", "reason": "残存記録はありません。"}
    assert renderers.clear_pending("powerpoint", {})["exit_code"] == 0
    assert renderers.clear_pending("libreoffice", {})["status"] == "nothing"
    assert len(maintenance.commands) == 3


@pytest.mark.parametrize("worker_wrote_record", [True, False])
def test_timeout_pending_record_has_windows_paths(tmp_path, fake_windows, worker_wrote_record):
    fake_windows.respond = blocking_render(write_record=worker_wrote_record)
    source = tmp_path / "deck.pptx"
    source.write_bytes(b"source")
    # 偽物が記録を書き終えてから時間切れになるよう、記録を書く場合は制限時間に余裕を持たせる。
    timeout = 1.0 if worker_wrote_record else 0.2
    with pytest.raises(renderers.RendererError) as error:
        with renderers.open_session("powerpoint", {"renderers": {"powerpoint": {"timeout": timeout}}}) as session:
            session.render(source, tmp_path / "deck.pdf", label="standard")
    (worker,) = fake_windows.workers
    assert error.value.code == "renderer_timeout"
    assert "slidekit renderer status" in str(error.value)
    # このセッションの記録なら status を書き換え、なければ作る。解除に使う PID と Windows のパスを入れる。
    record = json.loads(fake_windows.pending.read_text(encoding="utf-8-sig"))
    work = worker.option("-WorkDir")
    assert (record["status"], record["run_id"], record["process_id"], record["label"]) == (
        "timeout",
        worker.option("-RunId"),
        WORKER_PID,
        "standard",
    )
    assert (record["input"], record["work_dir"]) == (work + "\\input-0001.pptx", work)
    assert error.value.details["pending"] == str(fake_windows.pending)
    assert Path(error.value.details["work_dir"]) == unix(work) and unix(work).is_dir()


def test_timeout_detaches_worker_without_killing(tmp_path, fake_windows):
    fake_windows.respond = blocking_render(write_record=False, write_pdf=True)
    source, target = tmp_path / "deck.pptx", tmp_path / "deck.pdf"
    source.write_bytes(b"source")
    sessions = []
    with pytest.raises(renderers.RendererError) as error:
        with renderers.open_session("powerpoint", {"renderers": {"powerpoint": {"timeout": 0.2}}}) as session:
            sessions.append(session)
            session.render(source, target)
    (worker,) = fake_windows.workers
    assert error.value.code == "renderer_timeout"
    assert worker.stdin.closed and not worker.killed
    assert sessions[0].close()["detached"] is True
    with pytest.raises(renderers.RendererError) as again:
        sessions[0].render(source, tmp_path / "again.pdf")
    assert again.value.code == "renderer_timeout"
    # 変換が後から終わっても、その出力は採用しない。
    worker.release.set()
    assert worker.wait(timeout=5) == 0
    assert unix(worker.requests[-1]["output"]).is_file()
    assert not target.exists()


def test_timeout_does_not_overwrite_previous_pending(tmp_path, fake_windows):
    fake_windows.pending.parent.mkdir(parents=True)
    fake_windows.pending.write_text('{"run_id":"previous"}')
    fake_windows.respond = blocking_render(write_record=False)
    source = tmp_path / "deck.pptx"
    source.write_bytes(b"source")
    with pytest.raises(renderers.RendererError) as error:
        with renderers.open_session("powerpoint", {"renderers": {"powerpoint": {"timeout": 0.2}}}) as session:
            session.render(source, tmp_path / "deck.pdf")
    assert error.value.code == "renderer_timeout"
    assert json.loads(fake_windows.pending.read_text()) == {"run_id": "previous"}


def raise_interrupt(event):
    raise KeyboardInterrupt


def test_interrupt_before_ready_kills_waiting_worker(fake_windows):
    def startup(worker):
        worker.event(event="waiting", waited_seconds=0.0, holder=ACTIVE)
        worker.release.wait(10)
        return False

    fake_windows.startup = startup
    with pytest.raises(KeyboardInterrupt):
        with renderers.open_session("powerpoint", {}, on_event=raise_interrupt):
            pass
    # 排他を待っているだけで PowerPoint に触れていないので、作業プロセスを終了させる。
    (worker,) = fake_windows.workers
    assert worker.killed
    assert not list((fake_windows.temporary / "slidekit").iterdir())


def test_interrupt_after_ready_detaches_worker(tmp_path, fake_windows):
    def respond(worker, request):
        if request["op"] == "render":
            # 変換中の中断を、通知を受けた on_event から起こす。
            worker.event(event="notice")
            worker.release.wait(10)
            return None
        return respond_normally(worker, request)

    fake_windows.respond = respond
    source = tmp_path / "deck.pptx"
    source.write_bytes(b"source")
    sessions = []
    with pytest.raises(KeyboardInterrupt):
        with renderers.open_session("powerpoint", {}, on_event=raise_interrupt) as session:
            sessions.append(session)
            session.render(source, tmp_path / "deck.pdf")
    (worker,) = fake_windows.workers
    assert worker.stdin.closed and not worker.killed
    assert json.loads(fake_windows.pending.read_text(encoding="utf-8"))["status"] == "interrupted"
    with pytest.raises(renderers.RendererError) as error:
        sessions[0].render(source, tmp_path / "again.pdf")
    assert error.value.code == "renderer_interrupted"
    # 変換していないときの中断は、後片付けを待たずに切り離す（残存記録は作らない）。
    with pytest.raises(KeyboardInterrupt):
        with renderers.open_session("powerpoint", {}) as session:
            session.doctor()
            raise KeyboardInterrupt
    idle = fake_windows.workers[1]
    assert [request["op"] for request in idle.requests] == ["doctor"]
    assert idle.stdin.closed and not idle.killed


@pytest.mark.skipif(not Path("/proc/self/stat").exists(), reason="Linux の /proc で処理の状態を読む試験である。")
def test_worker_has_own_session_and_ignores_sigpipe():
    # 切り離した作業プロセスが後片付けまで続くよう、端末の Ctrl+C が届かない別の session で、SIGPIPE を無視したまま起動する
    # （既定の SIGPIPE では、Python が終わった後の応答の書き込みで WSL の中継の処理が終わる）。
    import signal

    def read(path):
        process = powerpoint._start_worker(["cat", path])
        output, errors = process.communicate(timeout=30)
        assert process.returncode == 0, errors
        return output.decode("ascii")

    status = dict(line.split(":", 1) for line in read("/proc/self/status").splitlines() if ":" in line)
    assert int(status["SigIgn"].strip(), 16) & (1 << (signal.SIGPIPE - 1))
    # /proc/self/stat の 6 番目の欄が session。2 番目（実行ファイル名）は括弧の中なので、最後の「)」より後を区切る。
    session = int(read("/proc/self/stat").rpartition(")")[2].split()[3])
    assert session != os.getsid(0)


def test_windows_paths_are_arguments_not_interpolated_code(monkeypatch):
    calls = []

    def fake_run(command, **kwargs):
        calls.append(command)
        return subprocess.CompletedProcess(command, 0, b"/mnt/c/path\n", b"")

    monkeypatch.setattr(renderers, "_executable", lambda value: value)
    monkeypatch.setattr(renderers, "_run", fake_run)
    value = r"C:\a b\日本語 $(Write-Output private)\x"
    renderers._unix_path(value)
    assert calls[0] == ["wslpath", "-u", value]


def test_font_comparison_rejects_same_name_different_content(tmp_path, monkeypatch):
    font_path = tmp_path / "test.ttf"
    font_path.write_bytes(b"font")
    registered_path = tmp_path / "registered.ttf"
    registered_path.write_bytes(b"other font")
    monkeypatch.setattr(renderers, "_unix_path", lambda value: Path(value))
    monkeypatch.setattr(
        renderers,
        "_font_face",
        lambda path, index: {
            "table_sha256": path.read_bytes().hex(),
            "font_names": ["Example"],
            "version": ["1"],
            "index": index,
        },
    )
    with pytest.raises(renderers.RendererError) as error:
        renderers.check_windows_fonts(
            {"fonts": {"Example": {"path": str(font_path)}}},
            [{"name": "Example (TrueType)", "path": str(registered_path)}],
        )
    assert error.value.code == "windows_font_mismatch"


def test_collection_and_extracted_face_have_same_table_hash(tmp_path):
    from fontTools.ttLib import TTCollection, TTFont

    path = synthetic_font(tmp_path / "test.ttf", "Example")
    collection_path = tmp_path / "test.ttc"
    collection = TTCollection()
    collection.fonts = [TTFont(path)]
    collection.save(collection_path)
    collection.close()
    extracted_path = tmp_path / "extracted.ttf"
    font = TTFont(collection_path, fontNumber=0)
    font.recalcTimestamp = False
    font["head"].modified += 123
    font.save(extracted_path)
    font.close()
    assert renderers._font_face(collection_path)["table_sha256"] == renderers._font_face(extracted_path)["table_sha256"]


def test_font_comparison_rejects_wrong_face_name(tmp_path, monkeypatch):
    source = tmp_path / "wrong-face.ttf"
    monkeypatch.setattr(
        renderers,
        "_font_face",
        lambda path, index: {
            "table_sha256": "hash",
            "font_names": ["Another Family"],
            "version": ["1"],
            "index": index,
        },
    )
    with pytest.raises(renderers.RendererError) as error:
        renderers.check_windows_fonts({"fonts": {"Example": {"path": str(source)}}}, [])
    assert error.value.code == "font_family_mismatch"


def test_font_comparison_finds_same_faces_among_registered_fonts(tmp_path, monkeypatch):
    from fontTools.ttLib import TTCollection, TTFont

    regular = synthetic_font(tmp_path / "example.ttf", "Example")
    bold = synthetic_font(tmp_path / "example-bold.ttf", "Example", "Bold", width=600)
    # 同じ書体名で中身が違うフォントと、ファイルの読み込みに失敗した登録は照合先にしない。
    decoy = synthetic_font(tmp_path / "decoy.ttf", "Example", width=700)
    other = synthetic_font(tmp_path / "other.ttf", "Other", width=800)
    collection_path = tmp_path / "fonts.ttc"
    collection = TTCollection()
    collection.fonts = [TTFont(other), TTFont(regular), TTFont(bold)]
    collection.save(collection_path)
    collection.close()
    monkeypatch.setattr(renderers, "_unix_path", lambda value: Path(value))
    installed = [
        {"name": "Example (TrueType)", "path": str(decoy)},
        {"name": "Example Light (TrueType)", "path": str(tmp_path / "missing.ttf")},
        {"name": "Other & Example (TrueType)", "path": str(collection_path)},
    ]
    matches = renderers.check_windows_fonts(
        {"fonts": {"Example": {"path": str(regular), "bold_path": str(bold)}}}, installed
    )
    assert [
        (item["family"], item["variant"], item["source"], item["windows_path"], item["windows_index"])
        for item in matches
    ] == [
        ("Example", "regular", str(regular.resolve()), str(collection_path), 1),
        ("Example", "bold", str(bold.resolve()), str(collection_path), 2),
    ]
    # 照合元と同じファイルが登録されている場合も一致とする。
    same_file = renderers.check_windows_fonts(
        {"fonts": {"Example": {"path": str(regular)}}}, [{"name": "Example (TrueType)", "path": str(regular)}]
    )
    assert [(item["windows_path"], item["windows_index"]) for item in same_file] == [(str(regular), 0)]


def test_font_comparison_checks_same_file_first(tmp_path, monkeypatch):
    regular = synthetic_font(tmp_path / "example.ttf", "Example")
    bold = synthetic_font(tmp_path / "example-bold.ttf", "Example", "Bold", width=600)
    light = synthetic_font(tmp_path / "example-light.ttf", "Example", "Light", width=300)
    reads, converted = [], []
    real_face = renderers._font_face
    monkeypatch.setattr(
        renderers, "_font_face", lambda path, index=0: (reads.append((path.name, index)), real_face(path, index))[1]
    )
    monkeypatch.setattr(renderers, "_unix_path", lambda value: (converted.append(value), Path(value))[1])
    # 登録の順では、同じ書体名の別のファイルが先にある。
    installed = [
        {"name": "Example Bold (TrueType)", "path": str(bold)},
        {"name": "Example Light (TrueType)", "path": str(light)},
        {"name": "Example (TrueType)", "path": str(regular)},
    ]
    matches = renderers.check_windows_fonts({"fonts": {"Example": {"path": str(regular)}}}, installed)
    assert [(item["windows_path"], item["windows_index"]) for item in matches] == [(str(regular), 0)]
    # 照合元の表だけを読み、ほかの候補はパスの変換もしない。
    assert reads == [("example.ttf", 0)]
    assert converted == [str(regular)]


def test_worker_script_requires_force_while_powerpoint_runs():
    # PowerPoint は読み取り専用で開いた複製のファイルを開いたままにしないので、排他で開くことに成功しても、PowerPoint が複製を開いていないとは言えない
    # （PowerPoint 16.0、build 20430 で確かめた）。PowerPoint が起動していれば、複製の有無によらず --force を求める。
    script = Path(renderers.__file__).with_name("powerpoint.ps1").read_text(encoding="utf-8")
    clear = script[script.index("function Invoke-Clear") : script.index("function Remove-OwnWorkFolder")]
    assert "if ($running -and (-not $Force))" in clear
    assert "$null -eq $inputLocked" not in clear
    # 排他で開けない複製は、--force でも断る。
    assert clear.index("if ($inputLocked -eq $true)") < clear.index("if ($running -and (-not $Force))")


def test_worker_script_never_kills_powerpoint():
    script = Path(renderers.__file__).with_name("powerpoint.ps1").read_text(encoding="utf-8")
    assert "Stop-Process" not in script
    assert "taskkill" not in script.lower()
    assert ".Kill(" not in script
    assert "Global\\Slidekit.PowerPoint." in script
    assert "$presentation.Close()" in script
    assert "(-not $existing)" in script
    # 自分で起動したかは、PowerPoint の有無でなく、接続の前後の処理の番号で決める（終了中の PowerPoint を既存とみなさない）。
    new_app = script[script.index("function New-App") : script.index("function Reset-App")]
    assert new_app.index("$before = ") < new_app.index("New-Object -ComObject") < new_app.index("$after = ")
    assert "$before -notcontains" in new_app
    # 進捗の表示を止め（標準エラーに CLIXML を出さない）、排他は 1 秒ずつ待ち、使用中の記録を書く。
    assert "$ProgressPreference = 'SilentlyContinue'" in script
    assert "WaitOne(1000)" in script
    assert "active.json" in script
