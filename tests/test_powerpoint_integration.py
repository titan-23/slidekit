"""SLIDEKIT_TEST_POWERPOINT=1 を指定したときだけ、実際の Windows の PowerShell と PowerPoint を検査する。

直列に実行する（`-p no:xdist`）。利用者の PowerPoint と資料は終了させない。終了させるのは、検査用に起動した PowerShell だけである。
"""

import contextlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import threading
import time
from types import SimpleNamespace

import pytest

from slidekit import renderers

pytestmark = [
    pytest.mark.powerpoint,
    pytest.mark.skipif(
        os.environ.get("SLIDEKIT_TEST_POWERPOINT") != "1", reason="実機 PowerPoint を使う任意の検査である。"
    ),
]

# 排他を持つ検査用の PowerShell。標準入力の 1 行で放す（固定の待ち時間に頼らない）。
# ActivePath を渡すと、変換セッションと同じ形の使用中の記録を書き、放す前に消す。
HOLDER_SCRIPT = r"""param([string]$ActivePath = '')
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$mutex = New-Object System.Threading.Mutex($false, "Global\Slidekit.PowerPoint.$sid")
$locked = $false
try { $locked = $mutex.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $locked = $true }
if (-not $locked) {
    [Console]::Out.WriteLine('BUSY')
    exit 3
}
try {
    if ($ActivePath) {
        $record = @{schema_version=1; pid=$PID; run_id='holder-test'; started_at=[DateTime]::UtcNow.ToString('o');
                    holder=@{project='holder-test-deck.toml'; run='holder-test-run'; command='test'}}
        [IO.File]::WriteAllText($ActivePath, (ConvertTo-Json -InputObject $record -Compress -Depth 5),
                                (New-Object System.Text.UTF8Encoding($false)))
    }
    [Console]::Out.WriteLine("LOCKED $PID")
    [Console]::Out.Flush()
    [void][Console]::In.ReadLine()
} finally {
    if ($ActivePath) { Remove-Item -LiteralPath $ActivePath -Force -ErrorAction SilentlyContinue }
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
"""


def release(process):
    """検査用の PowerShell に 1 行を送り、排他を放させる。"""
    with contextlib.suppress(OSError, ValueError):
        process.stdin.write(b"release\n")
        process.stdin.flush()
        process.stdin.close()


@contextlib.contextmanager
def mutex_holder(active_path=None):
    """排他を持つ検査用の PowerShell を起動し、保持したところで制御を返す。抜けるときに放させる。"""
    context = renderers.windows_context({})
    powershell = renderers._executable("powershell.exe")
    with tempfile.TemporaryDirectory(
        prefix="slidekit-holder-test-", dir=renderers._unix_path(context["temp"])
    ) as folder:
        script = Path(folder) / "holder.ps1"
        script.write_text(HOLDER_SCRIPT, encoding="utf-8-sig")
        command = [
            powershell,
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            renderers._windows_path(script),
        ]
        if active_path is not None:
            command += ["-ActivePath", renderers._windows_path(active_path)]
        process = subprocess.Popen(command, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        try:
            line = process.stdout.readline().decode("utf-8", "replace").strip()
            assert line.startswith("LOCKED "), f"検査用の PowerShell が排他を取れない: {line}"
            yield SimpleNamespace(process=process, pid=int(line.split()[1]))
        finally:
            # 1 行で放させる。終わらない場合に備え、検査用の PowerShell だけを終了させる（PowerPoint には触れない）。
            release(process)
            try:
                process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                process.kill()
            process.communicate(timeout=10)


def make_pptx(path, text):
    from pptx import Presentation
    from pptx.util import Inches, Pt

    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
    textbox.text_frame.word_wrap = False
    run = textbox.text_frame.paragraphs[0].add_run()
    run.text = text
    run.font.name = "Yu Gothic UI"
    run.font.size = Pt(28)
    presentation.save(path)
    return path


def pdf_text(path):
    """PDF の各ページの文字（前後の空白を除く）。"""
    from pdf_tools import open_pdf, page_text

    with open_pdf(path) as document:
        return [page_text(page).strip() for page in document]


def test_real_powerpoint_renders_japanese(tmp_path):
    environment = {"fonts": {}}
    information = renderers.doctor_renderer("powerpoint", environment)
    assert information["version"]
    source = make_pptx(tmp_path / "input.pptx", "PowerPoint による変換確認")
    target = tmp_path / "output.pdf"
    result = renderers.render_pdf(source, target, "powerpoint", environment)
    assert result["renderer"] == "powerpoint"
    assert pdf_text(target) == ["PowerPoint による変換確認"]


def test_real_windows_mutex_blocks_another_call():
    with mutex_holder() as holder:
        with pytest.raises(renderers.RendererError) as error:
            renderers.doctor_renderer("powerpoint", {"fonts": {}, "renderers": {"powerpoint": {"wait_seconds": 0}}})
        assert error.value.code == "renderer_busy"
        # 確かめた直後に放させる。
        release(holder.process)
        assert holder.process.wait(timeout=10) == 0


def test_real_windows_pending_marker_blocks_without_deleting_record():
    context = renderers.windows_context({})
    with tempfile.TemporaryDirectory(
        prefix="slidekit-pending-test-", dir=renderers._unix_path(context["temp"])
    ) as folder:
        state = Path(folder) / "isolated-state"
        state.mkdir()
        marker = state / "pending.json"
        marker.write_text(json.dumps({"run_id": "test-only"}), encoding="utf-8")
        with pytest.raises(renderers.RendererError) as error:
            with renderers.open_session("powerpoint", {"fonts": {}}, wait_seconds=0, state_dir=state):
                pass
        assert error.value.code == "renderer_pending"
        assert json.loads(marker.read_text(encoding="utf-8")) == {"run_id": "test-only"}
        # 使用中の記録は排他を放す前に消す。
        assert not (state / "active.json").exists()


def test_real_session_renders_twice_with_one_worker(tmp_path):
    environment = {"fonts": {}}
    running_before = renderers.renderer_status("powerpoint", environment)["powerpoint_running"]
    # PowerShell の起動回数を確かめるため、Windows への問い合わせの記憶を消す。
    renderers._cached_windows_context.cache_clear()
    texts = ["1 回目の変換", "2 回目の変換"]
    sources = [make_pptx(tmp_path / f"input-{number}.pptx", text) for number, text in enumerate(texts, 1)]
    with renderers.open_session("powerpoint", environment) as session:
        session.doctor()
        results = [
            session.render(source, tmp_path / f"output-{number}.pdf", label=f"render-{number}")
            for number, source in enumerate(sources, 1)
        ]
        operations = session.operations
    assert operations["powershell_launches"] == 2 and operations["renders"] == 2
    assert isinstance(operations["worker_pid"], int)
    assert results[0]["run_id"] == results[1]["run_id"]
    assert [pdf_text(tmp_path / f"output-{number}.pdf") for number in (1, 2)] == [[text] for text in texts]
    # 閉じた後は排他を取ることができる。
    status = renderers.renderer_status("powerpoint", environment)
    assert (status["busy"], status["pending"]) == (False, None)
    if not running_before:
        # 起動した PowerPoint は終了している（終了には少し時間がかかる）。
        deadline = time.monotonic() + 30
        while status["powerpoint_running"] and time.monotonic() < deadline:
            time.sleep(1)
            status = renderers.renderer_status("powerpoint", environment)
        assert not status["powerpoint_running"]


def test_real_session_waits_for_another_holder():
    context = renderers.windows_context({})
    with tempfile.TemporaryDirectory(prefix="slidekit-wait-test-", dir=renderers._unix_path(context["temp"])) as folder:
        state = Path(folder) / "state"
        state.mkdir()
        with mutex_holder(state / "active.json") as holder:
            # 待たない設定では、使用中の処理の PID と資料を示して失敗する。
            with pytest.raises(renderers.RendererError) as error:
                with renderers.open_session("powerpoint", {"fonts": {}}, wait_seconds=0, state_dir=state):
                    pass
            assert error.value.code == "renderer_busy"
            assert f"PID {holder.pid}" in str(error.value) and "holder-test-deck.toml" in str(error.value)
            assert error.value.details["holder"]["pid"] == holder.pid
            # 待つ設定では、waiting を受け取り、保持側が放した後に成功する。
            # 放すのは最初の waiting を受け取ってから 2 秒後にする（PowerShell の起動の時間に左右されない）。
            events = []
            timers = []

            def on_event(event):
                events.append(event)
                if not timers:
                    timers.append(threading.Timer(2.0, release, args=(holder.process,)))
                    timers[0].start()

            try:
                with renderers.open_session(
                    "powerpoint",
                    {"fonts": {}},
                    wait_seconds=10,
                    state_dir=state,
                    holder={"project": "waiting-test", "command": "test"},
                    on_event=on_event,
                ) as session:
                    operations = session.operations
            finally:
                for timer in timers:
                    timer.cancel()
            assert events and events[0]["event"] == "waiting"
            assert events[0]["holder"]["project"] == "holder-test-deck.toml"
            assert 0 < operations["waited_seconds"] <= 10
        assert not (state / "active.json").exists()


def test_real_status_and_clear_in_isolated_state():
    environment = {"fonts": {}}
    context = renderers.windows_context(environment)
    powershell = renderers._executable("powershell.exe")
    # 終了した PowerShell の PID を持つ残存記録を作る（記録の処理が動いていない状態）。
    finished = subprocess.run(
        renderers._powershell_command(powershell, "Write-Output $PID"), capture_output=True, timeout=60, check=True
    )
    process_id = int(renderers._decode(finished.stdout).split()[-1])
    root = renderers._unix_path(context["temp"]) / "slidekit"
    root.mkdir(parents=True, exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix="ppt-", dir=root))
    try:
        with tempfile.TemporaryDirectory(
            prefix="slidekit-state-test-", dir=renderers._unix_path(context["temp"])
        ) as folder:
            state = Path(folder)
            staged = work / "input-0001.pptx"
            staged.write_bytes(b"not opened")
            record = {
                "schema_version": 1,
                "run_id": "status-test",
                "process_id": process_id,
                "mode": "Render",
                "label": "standard",
                "input": renderers._windows_path(staged),
                "output": renderers._windows_path(work / "output-0001.pdf"),
                "work_dir": renderers._windows_path(work),
                "started_at": "2026-10-12T00:00:00Z",
                "status": "timeout",
            }
            (state / "pending.json").write_text(json.dumps(record), encoding="utf-8")
            status = renderers.renderer_status("powerpoint", environment, state_dir=state)
            assert (status["status"], status["exit_code"]) == ("pending", 1)
            assert status["pending"]["run_id"] == "status-test"
            assert (status["pending_process_alive"], status["staged_input_locked"]) == (False, False)
            # PowerPoint が起動していれば、開いた複製を見分けることができないので、--force のときだけ解除する
            # （起動していなければ --force なしで解除する。確かめた後に PowerPoint が終了した場合も同じ）。
            cleared = renderers.clear_pending("powerpoint", environment, state_dir=state)
            if cleared["status"] == "refused":
                assert status["powerpoint_running"] and "--force" in cleared["reason"], cleared
                assert cleared["exit_code"] == 1 and (state / "pending.json").exists() and work.exists()
                cleared = renderers.clear_pending("powerpoint", environment, force=True, state_dir=state)
            assert (cleared["status"], cleared["exit_code"]) == ("cleared", 0), cleared
            archived = list((state / "cleared").glob("*-pending.json"))
            assert len(archived) == 1 and cleared["archived"].endswith(archived[0].name)
            assert json.loads(archived[0].read_text(encoding="utf-8-sig"))["run_id"] == "status-test"
            assert not (state / "pending.json").exists()
            assert cleared["removed_work_dir"] is True and not work.exists()
            nothing = renderers.clear_pending("powerpoint", environment, state_dir=state)
            assert (nothing["status"], nothing["exit_code"]) == ("nothing", 0)
    finally:
        shutil.rmtree(work, ignore_errors=True)


# 時間切れか中断で作業プロセスを切り離し、すぐに終わる子の Python。セッションの作業フォルダーと例外の識別子を出す。
DETACHING_CHILD = r"""
import json
from pathlib import Path
import sys

from slidekit import renderers

source, target, state, mode = sys.argv[1:5]
environment = {'fonts': {}, 'renderers': {'powerpoint': {'timeout': 1 if mode == 'timeout' else 300}}}
sessions = []
result = {'code': None}
try:
    with renderers.open_session('powerpoint', environment, state_dir=Path(state)) as session:
        sessions.append(session)
        session.render(Path(source), Path(target))
except renderers.RendererError as error:
    result['code'] = error.code
except KeyboardInterrupt:
    result['code'] = 'interrupted'
result['work_dir'] = str(sessions[0].work_dir) if sessions else None
print(json.dumps(result), flush=True)
"""


def make_long_pptx(path, count):
    """変換に数秒かかる、文字だけのページを count 枚持つ PPTX。"""
    from pptx import Presentation
    from pptx.util import Inches, Pt

    presentation = Presentation()
    for number in range(1, count + 1):
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(4))
        textbox.text_frame.word_wrap = True
        run = textbox.text_frame.paragraphs[0].add_run()
        run.text = f"{number} ページ目の変換確認。" * 20
        run.font.name = "Yu Gothic UI"
        run.font.size = Pt(20)
    presentation.save(path)
    return path


@pytest.mark.parametrize("mode", ["timeout", "interrupt"])
def test_real_detached_worker_cleans_up_after_python_exits(tmp_path, mode):
    # 時間切れと端末の Ctrl+C（処理グループ全体への SIGINT）の後に Python が終わっても、切り離した作業プロセスは
    # 変換を終えて残存記録・使用中の記録・作業フォルダーを消し、自分で起動した PowerPoint を終了する。
    import signal
    import sys

    environment = {"fonts": {}}
    running_before = renderers.renderer_status("powerpoint", environment)["powerpoint_running"]
    context = renderers.windows_context(environment)
    source = make_long_pptx(tmp_path / "long.pptx", 150)
    with tempfile.TemporaryDirectory(
        prefix="slidekit-detach-test-", dir=renderers._unix_path(context["temp"])
    ) as folder:
        state = Path(folder)
        # 子は自分の処理グループで起動する（端末の Ctrl+C と同じく、子の処理グループ全体へ SIGINT を送る）。
        child = subprocess.Popen(
            [sys.executable, "-c", DETACHING_CHILD, str(source), str(tmp_path / "out.pdf"), str(state), mode],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            start_new_session=True,
            env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
        )
        try:
            if mode == "interrupt":
                # 作業プロセスが変換を始めた（残存記録を running で書いた）後に送る。
                deadline = time.monotonic() + 120
                while time.monotonic() < deadline and child.poll() is None:
                    with contextlib.suppress(OSError, ValueError):
                        if json.loads((state / "pending.json").read_text(encoding="utf-8-sig"))["status"] == "running":
                            break
                    time.sleep(0.1)
                time.sleep(0.5)
                os.killpg(child.pid, signal.SIGINT)
            output, errors = child.communicate(timeout=180)
        finally:
            if child.poll() is None:
                child.kill()
        result = json.loads(output.decode("utf-8").strip().splitlines()[-1])
        assert result["code"] == ("renderer_timeout" if mode == "timeout" else "interrupted"), errors.decode(
            "utf-8", "replace"
        )
        # 子が終わった時点では、作業プロセスがまだ変換している。
        assert (state / "pending.json").exists()
        deadline = time.monotonic() + 120
        status = renderers.renderer_status("powerpoint", environment, state_dir=state)
        while status["status"] != "idle" and time.monotonic() < deadline:
            time.sleep(1)
            status = renderers.renderer_status("powerpoint", environment, state_dir=state)
        assert status["status"] == "idle", status
        assert not (state / "pending.json").exists() and not (state / "active.json").exists()
        # 作業フォルダーは、排他を放した後に作業プロセスが消す。
        work = Path(result["work_dir"])
        deadline = time.monotonic() + 30
        while work.exists() and time.monotonic() < deadline:
            time.sleep(0.5)
        assert not work.exists()
        if not running_before:
            deadline = time.monotonic() + 60
            while status["powerpoint_running"] and time.monotonic() < deadline:
                time.sleep(1)
                status = renderers.renderer_status("powerpoint", environment)
            assert not status["powerpoint_running"]


def test_real_windows_font_directories():
    context = renderers.windows_context({})
    assert (renderers._unix_path(context["fonts"]) / "arial.ttf").is_file()
    assert renderers._unix_path(context["local_app_data"]).is_dir()
