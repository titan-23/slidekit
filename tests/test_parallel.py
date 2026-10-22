"""同時プロセス数の決め方と、プロセス群（Workers）の振る舞いを確かめる。"""

import os
import signal
import subprocess
import sys

import pytest

from slidekit import parallel, selection
from slidekit.parallel import Workers, resolve_jobs


def test_resolve_jobs_priority(monkeypatch):
    monkeypatch.setattr(parallel, "_cpu_count", lambda: 12)
    monkeypatch.delenv("SLIDEKIT_JOBS", raising=False)
    # 引数 > 環境変数 > 実行環境 > min(8, CPU の数) の順に使う。
    assert resolve_jobs(None) == 8
    assert resolve_jobs(None, {"jobs": 3}) == 3
    monkeypatch.setenv("SLIDEKIT_JOBS", "5")
    assert resolve_jobs(None, {"jobs": 3}) == 5
    assert resolve_jobs(2, {"jobs": 3}) == 2
    # 空の環境変数は指定なしとみなす。
    monkeypatch.setenv("SLIDEKIT_JOBS", " ")
    assert resolve_jobs(None, {"jobs": 3}) == 3
    monkeypatch.setattr(parallel, "_cpu_count", lambda: 2)
    assert resolve_jobs(None, {}) == 2
    for value in (0, -1, 2.0, True, "4"):
        with pytest.raises(ValueError, match="^--jobs は 1 以上の整数で指定してください"):
            resolve_jobs(value)
    for text in ("0", "two", "1.5", "-3", "４"):
        monkeypatch.setenv("SLIDEKIT_JOBS", text)
        with pytest.raises(ValueError, match="^環境変数 SLIDEKIT_JOBS は 1 以上の整数で指定してください"):
            resolve_jobs(None)
    monkeypatch.delenv("SLIDEKIT_JOBS")
    with pytest.raises(ValueError, match="^実行環境の jobs は 1 以上の整数で指定してください"):
        resolve_jobs(None, {"jobs": 0})


def test_map_keeps_order_and_propagates_errors():
    with Workers(3) as workers:
        assert workers.map(pow, [(2, exponent) for exponent in range(12)], minimum=1) == [2**n for n in range(12)]
        # 子の例外は arguments の順で最初に失敗した組のものを、子のトレースバック付きで出す。
        arguments = [("1",), ("2-3",), ("3-",), ("x-",), ("4",)]
        with pytest.raises(ValueError, match=r"^--slides の指定を解釈できません: '3-'$") as raised:
            workers.map(selection.parse, arguments, minimum=1)
        remote = str(raised.value.__cause__)
        assert "selection.py" in remote and "in parse" in remote
        # 失敗の後も同じプロセス群を使うことができる。同じプロセスで順に実行した場合と結果が同じになる。
        assert workers.map(selection.parse, arguments[:2], minimum=1) == [selection.parse("1"), selection.parse("2-3")]
    with pytest.raises(ValueError, match=r"^--slides の指定を解釈できません: '3-'$"):
        Workers(1).map(selection.parse, arguments)
    with pytest.raises(TypeError, match="組"):
        Workers(2).map(pow, [2, 3])
    with pytest.raises(ValueError):
        Workers(0)


def test_small_inputs_run_in_process():
    with Workers(4) as workers:
        # 組の数が minimum より少なければ、子プロセスを作らずに順に実行する。
        assert workers.map(os.getpid, [()] * 3) == [os.getpid()] * 3
        assert workers.map(os.getpid, [()] * 5, minimum=6) == [os.getpid()] * 5
        assert workers._processes is None
        assert Workers(1).map(os.getpid, [()] * 10, minimum=1) == [os.getpid()] * 10
        # minimum 以上なら子プロセスで実行する。
        assert os.getpid() not in workers.map(os.getpid, [()] * 4)
        assert workers._processes is not None


def test_children_leave_interrupts_to_parent():
    # 子プロセスは Ctrl+C を無視し、親が待ちを打ち切ってプロセス群を閉じる。親の扱いは変えない。
    with Workers(2) as workers:
        assert set(workers.map(signal.getsignal, [(signal.SIGINT,)] * 4)) == {signal.SIG_IGN}
    assert signal.getsignal(signal.SIGINT) is not signal.SIG_IGN


def test_close_is_final():
    workers = Workers(2)
    assert workers.map(pow, [(3, 2)] * 4) == [9, 9, 9, 9]
    workers.close()
    workers.close()
    assert workers._processes is None
    with pytest.raises(RuntimeError):
        workers.map(pow, [(2, 2)])


def test_common_modules_do_not_load_heavy_dependencies():
    # パッケージの読み込み（slidekit/__init__.py）で入るモジュールを基準にして、共通の部品が加えるものだけを調べる。
    script = (
        "import sys\n"
        "import slidekit\n"
        "before = set(sys.modules)\n"
        "import slidekit.errors, slidekit.parallel, slidekit.progress, slidekit.report, slidekit.selection\n"
        "added = {name.split('.')[0] for name in set(sys.modules) - before}\n"
        "heavy = ('pypdfium2', 'pypdf', 'PIL', 'matplotlib', 'numpy', 'lxml', 'pptx', 'fontTools', 'cairosvg')\n"
        "print(','.join(name for name in heavy if name in added))\n"
    )
    result = subprocess.run(
        [sys.executable, "-c", script],
        capture_output=True,
        text=True,
        check=True,
        env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
    )
    assert result.stdout.strip() == ""
