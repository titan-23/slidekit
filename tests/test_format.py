"""コードが Black で整形されていることを確かめる（設定は pyproject.toml の [tool.black]）。"""

from pathlib import Path
import subprocess
import sys

import pytest

ROOT = Path(__file__).resolve().parents[1]
TARGETS = ("src", "tests", "examples", "tools")


def test_code_is_formatted_with_black():
    # Black は追加依存 dev にある。導入していない環境では省略する。
    pytest.importorskip("black")
    command = [sys.executable, "-m", "black", "--check", "--quiet", *TARGETS]
    result = subprocess.run(command, cwd=ROOT, capture_output=True, text=True, check=False)
    assert result.returncode == 0, (
        "Black で整形されていないファイルがある。slidekit のディレクトリで `.venv/bin/python -m black src tests examples tools` を実行して整形する"
        f"（一覧は `--quiet` を外すと出る）\n{result.stderr}"
    )
