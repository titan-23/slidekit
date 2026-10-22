"""導入の設定（pyproject.toml）、検証済みの依存（requirements-lock.txt）、版と公開する名前を確かめる。

宣言した依存とその推移的な依存（ロックファイルの行と比べるもの）は、導入済みの配布物の依存の宣言から求める。
ロックファイルを作るコマンド（Python 3.12、Linux）と同じ条件で環境の目印を評価する。導入済みの版がロックファイルと違う環境では、
推移的な依存が決まらないので省略する。
"""

from importlib import metadata
from pathlib import Path
import re
import shlex
import tomllib

import pytest

import slidekit
from slidekit import cli, tables

ROOT = Path(__file__).resolve().parents[1]
LOCK = ROOT / "requirements-lock.txt"
# 「名前==版」だけの行。環境の目印、URL、パス、-e、ハッシュ値の行は受け付けない。
PIN = re.compile(r"([A-Za-z0-9][A-Za-z0-9._-]*)==([A-Za-z0-9][A-Za-z0-9.+!_-]*)")
# ロックファイルを作るコマンド（--python-version 3.12、Linux）の環境の目印。
LOCK_ENVIRONMENT = {
    "python_version": "3.12",
    "python_full_version": "3.12.0",
    "sys_platform": "linux",
    "platform_system": "Linux",
    "os_name": "posix",
}


def _normalize(name: str) -> str:
    """配布物の名前を比べる形にそろえる（小文字、- . _ の並びは - にする）。"""
    return re.sub(r"[-_.]+", "-", name).lower()


def _project() -> dict:
    with (ROOT / "pyproject.toml").open("rb") as file:
        return tomllib.load(file)["project"]


def _declared() -> list[str]:
    """必須の依存と、すべての追加依存。"""
    project = _project()
    return list(project["dependencies"]) + [
        text for extra in project["optional-dependencies"].values() for text in extra
    ]


def _lock_lines() -> list[str]:
    return LOCK.read_text(encoding="utf-8").splitlines()


def _lock() -> dict[str, str]:
    """ロックファイルの名前（そろえた形）と版。"""
    return {_normalize(match[1]): match[2] for match in map(PIN.fullmatch, _lock_lines()) if match}


def test_lock_pins_every_declared_dependency():
    lock = _lock()
    declared = _declared()
    # 宣言した依存は、すべて版を 1 つに決める
    assert [text for text in declared if not PIN.fullmatch(text)] == []
    pins = {_normalize(match[1]): match[2] for match in map(PIN.fullmatch, declared)}
    assert {name: (version, lock.get(name)) for name, version in pins.items() if lock.get(name) != version} == {}


def test_lock_has_only_pinned_packages():
    lines = _lock_lines()
    assert [line for line in lines if line.strip() and not line.startswith("#") and not PIN.fullmatch(line)] == []
    names = [_normalize(match[1]) for match in map(PIN.fullmatch, lines) if match]
    assert names and len(names) == len(set(names))


def _note_command(note: list[str], start: str) -> list[str]:
    """注記の中の start で始まるコマンドを、行末までコピーしたときにシェルが渡す語に分ける。"""
    line = next(line for line in note if start in line)
    return shlex.split(line[line.index(start) :])


def test_lock_note_covers_every_extra():
    # 先頭の注記：検証した追加依存の組、導入のコマンド、作り直しのコマンドが、宣言したすべての追加依存を含む
    # コマンドは行末で終わり、行末までコピーして実行できる（コマンドの後に説明を続けない）
    extras = set(_project()["optional-dependencies"])
    lines = _lock_lines()
    note = lines[: next(index for index, line in enumerate(lines) if not line.startswith("#"))]
    assert note, "ロックファイルの先頭に注記がない"
    verified = re.search(r"\.\[([a-z,]+)\]", note[0])
    assert verified and set(verified[1].split(",")) == extras
    install = _note_command(note, "uv pip install -r")
    assert install == "uv pip install -r requirements-lock.txt && uv pip install --no-deps -e .".split()
    compile_words = _note_command(note, "uv pip compile")
    assert [word for word in compile_words if not word.isascii()] == []
    assert compile_words[:4] == ["uv", "pip", "compile", "pyproject.toml"]
    assert compile_words[-2:] == ["-o", "requirements-lock.txt"]
    pairs = list(zip(compile_words, compile_words[1:]))
    assert {value for option, value in pairs if option == "--extra"} == extras
    assert ("--python-version", "3.12") in pairs


def test_lock_matches_dependency_closure():
    # ロックファイルの行は、宣言した依存とその推移的な依存にちょうど一致する
    # （不要な行がない：追加依存 math を空にした後の matplotlib とその依存など。足りない行もない）
    from packaging.markers import default_environment
    from packaging.requirements import Requirement

    environment = {**default_environment(), **LOCK_ENVIRONMENT}
    lock = _lock()
    found: dict[str, str] = {}
    seen = set()
    pending = [Requirement(text) for text in _declared()]
    while pending:
        requirement = pending.pop()
        name = _normalize(requirement.name)
        if (name, frozenset(requirement.extras)) in seen:
            continue
        seen.add((name, frozenset(requirement.extras)))
        try:
            distribution = metadata.distribution(name)
        except metadata.PackageNotFoundError:
            pytest.skip(f"{name} が導入されていない（uv pip install -r requirements-lock.txt で導入する）")
        if lock.get(name, distribution.version) != distribution.version:
            pytest.skip(f"{name} の導入済みの版 {distribution.version} がロックファイルの {lock[name]} と違う")
        found[name] = distribution.version
        for text in distribution.requires or ():
            child = Requirement(text)
            extras = requirement.extras or {""}
            if child.marker is None or any(child.marker.evaluate({**environment, "extra": extra}) for extra in extras):
                pending.append(child)
    assert sorted(set(lock) - set(found)) == []
    assert sorted(set(found) - set(lock)) == []


def test_math_extra_installs_nothing():
    # 旧版の導入コマンド（.[math]）が動くように名前だけを残す。数式の計測は matplotlib を使わない
    assert _project()["optional-dependencies"]["math"] == []
    assert "matplotlib" not in _lock()


def test_version_matches_project_metadata(capsys):
    assert slidekit.__version__ == _project()["version"]
    with pytest.raises(SystemExit) as exit:
        cli.main(["--version"])
    assert exit.value.code == 0
    assert capsys.readouterr().out == f"slidekit {slidekit.__version__}\n"


def test_public_names_are_exported():
    # slide.table() の戻り値の型は、型の注釈で使うために公開する
    assert slidekit.Table is tables.Table
    assert [name for name in slidekit.__all__ if not hasattr(slidekit, name)] == []


def test_pymupdf_is_neither_declared_nor_imported():
    # PDF は slidekit.pdf（pypdfium2・pypdf・fontTools）で読む。AGPL の PyMuPDF を依存にも読み込みにも戻さない。
    declared = {_normalize(re.split(r"[\s<>=!~;\[]", text, maxsplit=1)[0]) for text in _declared()}
    assert "pymupdf" not in declared and "pymupdf" not in _lock()
    pattern = re.compile(r"^\s*(?:import|from)\s+(?:pymupdf|fitz)\b", re.MULTILINE)
    sources = [
        path for folder in ("src", "tests", "tools", "examples") for path in sorted((ROOT / folder).rglob("*.py"))
    ]
    assert [str(path.relative_to(ROOT)) for path in sources if pattern.search(path.read_text(encoding="utf-8"))] == []
