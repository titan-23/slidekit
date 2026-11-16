"""進捗の表示と、段階ごとの所要時間の記録を確かめる。"""

import io
import threading

import pytest

from slidekit.progress import Progress, stage_label


class Clock:
    """試験用の時計。advance で進める。"""

    def __init__(self):
        self.now = 100.0

    def __call__(self):
        return self.now

    def advance(self, seconds):
        self.now += seconds


def test_stage_records_time_even_on_error():
    clock, stream = Clock(), io.StringIO()
    progress = Progress(stream, clock=clock)
    with progress.stage("construct", "構築"):
        clock.advance(0.5)
    clock.advance(1.0)
    with pytest.raises(RuntimeError, match="途中"):
        with progress.stage("layout"):
            clock.advance(2.25)
            progress.event("変換待ち: 別の処理が PowerPoint を使用中")
            raise RuntimeError("途中で失敗")
    assert progress.timings == [
        {"stage": "construct", "label": "構築", "seconds": 0.5},
        {"stage": "layout", "label": "配置の検査", "seconds": 2.25},
    ]
    assert progress.elapsed == 3.75
    assert stream.getvalue().splitlines() == [
        "[   0.0 秒] 構築",
        "[   1.5 秒] 配置の検査",
        "[   3.8 秒]   変換待ち: 別の処理が PowerPoint を使用中",
    ]
    # 実行中の段階はその時点までの秒数で示す。
    with progress.stage("render:powerpoint:standard"):
        clock.advance(4.0)
        assert progress.timings[-1] == {
            "stage": "render:powerpoint:standard",
            "label": "PowerPoint で変換（標準表示）",
            "seconds": 4.0,
        }


def test_null_stream_prints_nothing(capsys):
    clock = Clock()
    progress = Progress(None, clock=clock)
    with progress.stage("pdf:powerpoint:fallback"):
        clock.advance(1.0)
        progress.event("画像化")
    assert capsys.readouterr() == ("", "")
    assert progress.timings == [
        {"stage": "pdf:powerpoint:fallback", "label": "PDF の検査と画像化（互換表示）", "seconds": 1.0}
    ]
    assert stage_label("unknown_stage") == "unknown_stage"
    # 変換セッションを閉じる段階（自分で起動した PowerPoint の終了を待つ）。
    assert stage_label("close:powerpoint") == "PowerPoint の終了"


def test_error_prints_even_when_quiet(capsys):
    Progress(None).error("Traceback (most recent call last):\n  KeyError: 'b'")
    assert capsys.readouterr() == ("", "Traceback (most recent call last):\n  KeyError: 'b'\n")
    # 出力先を指定したときは、そこへ出す。
    stream = io.StringIO()
    Progress(stream).error("失敗")
    assert stream.getvalue() == "失敗\n"
    assert capsys.readouterr() == ("", "")


def test_lines_from_threads_are_not_mixed():
    stream = io.StringIO()
    progress = Progress(stream)

    def report(name):
        for index in range(200):
            progress.event(f"{name}-{index}")

    threads = [threading.Thread(target=report, args=(name,)) for name in "abcd"]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()
    lines = stream.getvalue().splitlines()
    assert len(lines) == 800
    assert all(line.startswith("[") and line.split("]   ", 1)[1][0] in "abcd" for line in lines)


def test_closed_stream_does_not_stop_processing():
    stream = io.StringIO()
    progress = Progress(stream)
    stream.close()
    with progress.stage("record"):
        progress.event("書けない")
    assert [item["stage"] for item in progress.timings] == ["record"]
