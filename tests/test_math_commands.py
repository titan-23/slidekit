"""数式の命令の表、構文解析との対応、命令の一覧の文書を確かめる。"""

import os
from pathlib import Path
import subprocess
import sys

from fontTools.ttLib import TTFont
import pytest

from slidekit.math import commands, syntax
from slidekit.math.syntax import MathSyntaxError, parse

DOCUMENT = Path(__file__).resolve().parents[1] / "docs" / "math.md"
CAMBRIA = Path("/mnt/c/Windows/Fonts/cambria.ttc")


def example(name):
    """命令名を使う最小の式。"""
    if name in commands.STRUCTURES:
        return commands.STRUCTURES[name].syntax
    if name in commands.FUNCTIONS:
        return "\\" + name + " x"
    token = "\\" + name
    if token in commands.SPACES:
        return "a" + token + " b"
    if name in commands.SYMBOLS:
        return token + " x" if commands.SYMBOLS[name].group == commands.OPERATOR else token
    if token in commands.DELIMITERS:
        return "\\left" + token + " x \\right."
    raise AssertionError(f"命令の使い方がわからない: {name}")


def test_symbols_exist_in_cambria_math():
    if not CAMBRIA.is_file():
        pytest.skip("Cambria Math が導入されていない")
    with TTFont(CAMBRIA, fontNumber=1) as face:
        assert "Cambria Math" in {item.toUnicode() for item in face["name"].names if item.nameID in (1, 16)}
        cmap = face.getBestCmap()
    characters = {symbol.char for symbol in (*commands.SYMBOLS.values(), *commands.CHARACTERS.values())}
    characters |= set(commands.NEGATIONS.values()) | set(commands.ACCENTS.values())
    characters |= {char for char in commands.DELIMITERS.values() if char}
    characters |= {char for char, _ in commands.SPACES.values() if char}
    missing = sorted(f"U+{ord(c):04X}" for text in characters for c in text if ord(c) not in cmap)
    assert not missing


def test_command_names_cover_parser():
    names = commands.command_names()
    assert len(names) > 250
    for name in sorted(names):
        assert parse(example(name)), name
    # 記号 1 文字の命令（\, \{ など）と、表の記号はすべて使うことができる
    for name, symbol in commands.SYMBOLS.items():
        assert parse(example(name) if name in names else "\\" + name)
    for token in commands.SPACES:
        parse("a" + token + " b")
    # 構造の表と、構文解析の処理が一致する
    handled = set(syntax._STRUCTURE_METHODS) | set(commands.FONTS) | set(commands.ACCENTS)
    assert handled | {"left", "right", "middle", "limits", "nolimits"} == set(commands.STRUCTURES)
    for name in ["foo", "alpah", "Alpha", "textbf", "mathbold", "begin", "big", "frak"]:
        with pytest.raises(MathSyntaxError) as error:
            parse("\\" + name + "{x}")
        assert error.value.reason == "未対応の数式命令"
        assert error.value.hint
    assert commands.UNSUPPORTED.keys().isdisjoint(names)


def test_math_reference_is_current():
    text = DOCUMENT.read_text(encoding="utf-8")
    assert text.count(commands.BEGIN) == 1 and text.count(commands.END) == 1
    region = text.split(commands.BEGIN + "\n", 1)[1].split(commands.END, 1)[0]
    assert region == commands.reference_markdown()


def test_reference_command_check_and_write(tmp_path, capsys):
    content = commands.reference_markdown()
    current = "# 題\n\n前の文\n\n" + commands.BEGIN + "\n" + content + commands.END + "\n\n後の文\n"
    good = tmp_path / "good.md"
    good.write_text(current, encoding="utf-8")
    assert commands.main(["--check", str(good)]) == 0
    stale = tmp_path / "stale.md"
    stale.write_text(current.replace("| `\\alpha` | α |", "| `\\alpha` | a |"), encoding="utf-8")
    assert commands.main(["--check", str(stale)]) == 1
    line = current.splitlines().index(next(item for item in current.splitlines() if item.startswith("| `\\alpha`"))) + 1
    assert f"{line} 行目" in capsys.readouterr().err
    assert commands.main(["--write", str(stale)]) == 0
    assert stale.read_text(encoding="utf-8") == current
    unmarked = tmp_path / "unmarked.md"
    unmarked.write_text("# 題\n", encoding="utf-8")
    assert commands.main(["--check", str(unmarked)]) == 2
    assert commands.main(["--write", str(unmarked)]) == 2
    assert unmarked.read_text(encoding="utf-8") == "# 題\n"
    capsys.readouterr()
    assert commands.main([]) == 0
    assert capsys.readouterr().out == content
    # python -m で動き、警告を出さない
    process = subprocess.run(
        [sys.executable, "-m", "slidekit.math.commands", "--check", str(good)],
        capture_output=True,
        text=True,
        env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
        check=False,
    )
    assert (process.returncode, process.stderr) == (0, "")
