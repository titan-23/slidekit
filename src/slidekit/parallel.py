"""PowerPoint を使わない処理のプロセス並列（ページ画像の作成と互換画像の切り出し）。

子プロセスで動かす作業関数は、それぞれの領域のモジュールの最上位に置く（例：`evidence.render_page_image`、
`fallback.render_page_crops`）。引数と戻り値は pickle できる値だけにし、重いモジュールは関数の中で読み込む。
このモジュールはプロセス群の仕組みだけを持ち、重いモジュールを読み込まない。
"""

from __future__ import annotations

import concurrent.futures
import multiprocessing
import os
import signal
import threading
from typing import Any, Callable, Sequence

# 同時プロセス数を指定しないときの上限。
DEFAULT_MAX_JOBS = 8
# Windows の ProcessPoolExecutor が受け付ける上限。
_WINDOWS_MAX_JOBS = 61


def _cpu_count() -> int:
    counter = getattr(os, "process_cpu_count", None) or os.cpu_count
    return counter() or 1


def _positive_integer(value: Any) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value >= 1


def resolve_jobs(requested: int | None = None, environment: dict | None = None) -> int:
    """同時プロセス数を決める。

    優先順位は、引数（`--jobs`）、環境変数 SLIDEKIT_JOBS、実行環境の `jobs`、min(8, CPU の数) の順である。
    1 以上の整数でない値は ValueError にする。空の環境変数は指定なしとみなす。
    """
    if requested is not None:
        if not _positive_integer(requested):
            raise ValueError(f"--jobs は 1 以上の整数で指定してください: {requested!r}")
        return requested
    text = os.environ.get("SLIDEKIT_JOBS", "").strip()
    if text:
        if not text.isascii() or not text.isdecimal() or int(text) < 1:
            raise ValueError(f"環境変数 SLIDEKIT_JOBS は 1 以上の整数で指定してください: {text!r}")
        return int(text)
    value = (environment or {}).get("jobs")
    if value is not None:
        if not _positive_integer(value):
            raise ValueError(f"実行環境の jobs は 1 以上の整数で指定してください: {value!r}")
        return value
    return min(DEFAULT_MAX_JOBS, _cpu_count())


def _start_method() -> str:
    # 変換セッションの読み取りスレッドがあるため fork は使わない。
    if os.name != "nt" and "forkserver" in multiprocessing.get_all_start_methods():
        return "forkserver"
    return "spawn"


def _ignore_interrupt() -> None:
    """子プロセスでは Ctrl+C を無視し、中断の判断を親に任せる（親が待ちを打ち切り、プロセス群を閉じる）。"""
    signal.signal(signal.SIGINT, signal.SIG_IGN)


class Workers:
    """1 回の処理で使い回すプロセス群。

    プロセス群は、並列に動かす最初の `map` で作る。`with` を抜けるときか `close()` で閉じる。
    """

    def __init__(self, jobs: int) -> None:
        if not _positive_integer(jobs):
            raise ValueError(f"同時プロセス数は 1 以上の整数で指定してください: {jobs!r}")
        self.jobs = jobs
        self._lock = threading.Lock()
        self._processes: concurrent.futures.ProcessPoolExecutor | None = None
        self._closed = False

    def map(self, function: Callable, arguments: Sequence[tuple], *, minimum: int = 4) -> list:
        """function(*引数) を引数の組ごとに実行し、結果を arguments の順に返す。

        jobs が 1 以下か、組の数が minimum より少なければ、同じプロセスで順に実行する（プロセスの起動と
        重いモジュールの読み込みのコストを避ける）。子の例外は元のトレースバックを `__cause__` に付けて呼び出し元へ伝える。
        例外が起きたときは、arguments の順で最初に失敗した組の例外を出し、まだ始まっていない組を取り消す。
        """
        items = []
        for item in arguments:
            if not isinstance(item, (tuple, list)):
                raise TypeError(f"引数は組（tuple）の列で渡してください: {item!r}")
            items.append(tuple(item))
        self._check_open()
        if self.jobs <= 1 or len(items) < minimum:
            return [function(*item) for item in items]
        pool = self._process_pool()
        futures = [pool.submit(function, *item) for item in items]
        try:
            return [future.result() for future in futures]
        except BaseException:
            for future in futures:
                future.cancel()
            raise

    def close(self) -> None:
        """子の処理の終わりを待って閉じる。まだ始まっていない処理は取り消す。2 回目以降は何もしない。"""
        with self._lock:
            if self._closed:
                return
            self._closed = True
            processes, self._processes = self._processes, None
        if processes is not None:
            processes.shutdown(wait=True, cancel_futures=True)

    def __enter__(self) -> "Workers":
        return self

    def __exit__(self, *exc: Any) -> None:
        self.close()

    def _check_open(self) -> None:
        if self._closed:
            raise RuntimeError("閉じたプロセス群は使えない")

    def _process_pool(self) -> concurrent.futures.ProcessPoolExecutor:
        with self._lock:
            self._check_open()
            if self._processes is None:
                jobs = min(self.jobs, _WINDOWS_MAX_JOBS) if os.name == "nt" else self.jobs
                self._processes = concurrent.futures.ProcessPoolExecutor(
                    max_workers=jobs,
                    mp_context=multiprocessing.get_context(_start_method()),
                    initializer=_ignore_interrupt,
                )
            return self._processes
