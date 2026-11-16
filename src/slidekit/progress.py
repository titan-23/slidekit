"""進捗の表示と、段階ごとの所要時間の記録。

進捗は 1 行ずつ「[  12.3 秒] ラベル」の形で出す。端末かどうかで形式を変えない。
書き込みはロックで守る（複数のスレッドから呼んでも行が混ざらない）。
"""

from __future__ import annotations

import contextlib
import sys
import threading
import time
from typing import Callable, Iterator, TextIO

# 段階の名前（機械向け）とラベル（日本語）。変換と PDF の検査は stage_label が作る。
STAGE_LABELS = {
    "construct": "構築",
    "assets": "素材の準備",
    "layout": "配置の検査",
    "pptx": "PPTX の書き出し",
    "environment": "実行環境の記録",
    "doctor": "表示ソフトの事前確認（変換待ちを含む）",
    "final_form": "最終形の PPTX の作成",
    "math_fallback": "互換画像の作成",
    "pptx_check": "PPTX の検査",
    "fallback_clone": "互換表示の複製",
    "collect": "結果の集約",
    "record": "生成記録の保存",
}
VARIANT_LABELS = {"standard": "標準表示", "fallback": "互換表示"}


def stage_label(name: str) -> str:
    """段階の名前のラベル。`render:powerpoint:{表示}`・`pdf:powerpoint:{表示}`・`close:powerpoint` も扱い、
    知らない名前はそのまま返す。"""
    if name in STAGE_LABELS:
        return STAGE_LABELS[name]
    kind, _, rest = str(name).partition(":")
    if kind == "close" and rest == "powerpoint":
        # 変換セッションを閉じる（自分で起動した PowerPoint の終了を待つ）。
        return "PowerPoint の終了"
    renderer, _, variant = rest.partition(":")
    if kind in {"render", "pdf"} and renderer == "powerpoint" and variant:
        variant = VARIANT_LABELS.get(variant, variant)
        return f"PowerPoint で変換（{variant}）" if kind == "render" else f"PDF の検査と画像化（{variant}）"
    return str(name)


class Progress:
    """進捗を stream へ出し、段階ごとの所要時間を記録する。

    stream が None なら何も出さず、所要時間だけを記録する（`--quiet`）。ただし `error()` は常に出す。
    """

    def __init__(self, stream: TextIO | None = None, *, clock: Callable[[], float] = time.monotonic) -> None:
        self._stream = stream
        self._clock = clock
        self._started = clock()
        self._lock = threading.Lock()
        self._stages: list[dict] = []

    @contextlib.contextmanager
    def stage(self, name: str, label: str | None = None) -> Iterator[None]:
        """開始時に 1 行出し、終了時に所要時間を記録する。例外で抜けたときも記録する。label を省くと stage_label(name)。"""
        label = stage_label(name) if label is None else label
        started = self._clock()
        record = {"stage": name, "label": label, "started": started, "seconds": None}
        with self._lock:
            self._stages.append(record)
        self._line(started, label)
        try:
            yield
        finally:
            ended = self._clock()
            with self._lock:
                record["seconds"] = ended - started

    def event(self, message: str) -> None:
        """段階の途中経過（「[  12.3 秒]   変換待ち: …」）。"""
        self._line(self._clock(), "  " + message)

    def error(self, text: str) -> None:
        """トレースバックなど。stream が None でも標準エラーへ出す。"""
        self._write(self._stream if self._stream is not None else sys.stderr, text)

    @property
    def timings(self) -> list[dict]:
        """[{"stage", "label", "seconds"}] を開始順に返す。実行中の段階はその時点までの秒数にする。"""
        now = self._clock()
        with self._lock:
            return [
                {
                    "stage": record["stage"],
                    "label": record["label"],
                    "seconds": round(now - record["started"] if record["seconds"] is None else record["seconds"], 3),
                }
                for record in self._stages
            ]

    @property
    def elapsed(self) -> float:
        """作ってからの秒数。"""
        return round(self._clock() - self._started, 3)

    def _line(self, moment: float, text: str) -> None:
        if self._stream is not None:
            self._write(self._stream, f"[{moment - self._started:6.1f} 秒] {text}")

    def _write(self, stream: TextIO, text: str) -> None:
        text = str(text)
        if not text.endswith("\n"):
            text += "\n"
        with self._lock:
            # 進捗の書き込みに失敗しても（閉じたパイプなど）、処理そのものは止めない。
            try:
                stream.write(text)
                stream.flush()
            except (OSError, ValueError):
                pass
