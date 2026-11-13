"""生成の間の終了の信号（SIGTERM・SIGHUP）を、Ctrl+C と同じ中断として扱うことを確かめる。

信号は試験の処理自身へ送る。差し替えが働かないと試験の処理が終わるので、主スレッドで、既定の処理のときだけ行う。
表示ソフトはフィクスチャ fake_session の偽の変換セッションに差し替える。
"""

import json
import signal
import threading

import pytest

from slidekit import pipeline

pytestmark = [
    pytest.mark.skipif(not hasattr(signal, "SIGHUP"), reason="SIGHUP のない環境（Windows）では試験しない"),
    pytest.mark.skipif(
        threading.current_thread() is not threading.main_thread(), reason="信号の処理は主スレッドでだけ設定できる"
    ),
]

DECK = """schema_version = 1
entrypoint = "slides.py:build"
size = "16:9"
[metadata]
title = "中断の試験"
author = ""
last_modified_by = ""
[inputs]
read_roots = ["."]
[render]
primary = "powerpoint"
required = ["powerpoint"]
[deliverables]
pptx = "presentation.pptx"
pdf = "presentation.pdf"
notes = "notes.md"
record = "deliverables.json"
"""

SLIDES = """from slidekit import Deck, Rect

def build(context):
    deck = Deck(context)
    slide = deck.slide(id="intro", kind="blank", title="Summary", notes="確認内容を説明します。", seconds=45)
    slide.text(id="body", text="Process checked", area=Rect(60, 110, 800, 90))
    return deck
"""


@pytest.fixture
def default_signals():
    """SIGTERM と SIGHUP を既定の処理にし、試験の後に元へ戻す。"""
    numbers = (signal.SIGTERM, signal.SIGHUP)
    saved = {number: signal.signal(number, signal.SIG_DFL) for number in numbers}
    yield numbers
    for number, handler in saved.items():
        signal.signal(number, handler)


@pytest.mark.parametrize("name", ["SIGTERM", "SIGHUP"])
def test_signal_becomes_interrupt_and_handlers_are_restored(default_signals, name):
    number = getattr(signal, name)
    with pytest.raises(KeyboardInterrupt):
        with pipeline._signals_as_interrupt():
            signal.raise_signal(number)
            pytest.fail("信号の後も続いた")
    assert all(signal.getsignal(number) is signal.SIG_DFL for number in default_signals)
    with pipeline._signals_as_interrupt():
        pass
    assert all(signal.getsignal(number) is signal.SIG_DFL for number in default_signals)


def test_second_signal_is_ignored_during_cleanup(default_signals):
    cleaned = []
    with pytest.raises(KeyboardInterrupt):
        with pipeline._signals_as_interrupt():
            try:
                signal.raise_signal(signal.SIGTERM)
            finally:
                signal.raise_signal(signal.SIGHUP)
                cleaned.append(True)
    assert cleaned == [True]


def test_caller_handler_is_kept(default_signals):
    calls = []

    def handler(number, frame):
        calls.append(number)

    signal.signal(signal.SIGTERM, handler)
    with pipeline._signals_as_interrupt():
        assert signal.getsignal(signal.SIGTERM) is handler
        signal.raise_signal(signal.SIGTERM)
    assert calls == [signal.SIGTERM] and signal.getsignal(signal.SIGTERM) is handler


def test_other_threads_do_not_change_handlers(default_signals):
    seen = []

    def work():
        with pipeline._signals_as_interrupt():
            seen.append(signal.getsignal(signal.SIGTERM))

    thread = threading.Thread(target=work)
    thread.start()
    thread.join()
    assert seen == [signal.SIG_DFL]


def test_sigterm_during_render_is_recorded_like_interrupt(
    tmp_path, write_deck, dejavu, dejavu_environment, fake_session
):
    deck = write_deck(tmp_path / "deck", deck=DECK, slides=SLIDES, theme=dejavu.theme)

    def terminate(name, label, source, output):
        signal.raise_signal(signal.SIGTERM)

    fake_session.on_render = terminate
    saved = signal.signal(signal.SIGTERM, signal.SIG_DFL)
    try:
        with pytest.raises(KeyboardInterrupt) as raised:
            pipeline.build(deck / "deck.toml")
        # 生成の後は元の処理に戻す。
        assert signal.getsignal(signal.SIGTERM) is signal.SIG_DFL
    finally:
        signal.signal(signal.SIGTERM, saved)
    (run,) = (deck / "out/runs").iterdir()
    result = raised.value.result
    assert (result["build"], result["status"], result["stage"]) == (str(run), "failed", "render:powerpoint:standard")
    issues = json.loads((run / "diagnostics.json").read_text(encoding="utf-8"))["issues"]
    assert [issue["code"] for issue in issues if issue["severity"] == "error"] == ["build_interrupted"]
    # 開いたセッションは閉じる（PowerPoint では作業プロセスを切り離し、作業プロセスが後片付けをする）。
    assert not (run / ".building").exists() and fake_session.closed == ["powerpoint"]
