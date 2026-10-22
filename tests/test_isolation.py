"""試験が利用者の環境設定と Windows のフォントを読まないこと、PowerPoint の試験を並列に動かさないことを確かめる。"""

import os
from pathlib import Path
from types import SimpleNamespace

import pytest

from slidekit import fonts
from slidekit.config import load_project


def test_tests_do_not_read_user_settings_or_windows_fonts(tmp_path, tmp_path_factory, write_deck):
    environment = Path(os.environ["SLIDEKIT_ENVIRONMENT"])
    # 環境設定は試験ごとの空の一時ファイルで、tmp_path の外にある。
    assert tmp_path_factory.getbasetemp() in environment.parents
    assert tmp_path not in environment.parents
    assert environment.read_text(encoding="utf-8") == "schema_version = 1\n"
    assert fonts.font_directories({}) == ()
    assert os.environ["SLIDEKIT_JOBS"] == "1"
    # 同梱のテーマの書体（Yu Gothic UI など）を探さず、フォントを入力として追跡しない。
    project = load_project(write_deck(tmp_path / "deck"))
    assert project.environment["fonts"] == {}
    assert [record["path"] for record in project.environment["config_files"]] == [str(environment)]
    assert not [path for path in project.input_files if Path(path).suffix.lower() in {".ttf", ".ttc", ".otf"}]


def test_powerpoint_tests_refuse_parallel_workers(request, monkeypatch):
    conftest = next(
        plugin
        for plugin in request.config.pluginmanager.get_plugins()
        if Path(getattr(plugin, "__file__", "") or "") == Path(__file__).with_name("conftest.py")
    )

    def configure(**option):
        conftest.pytest_configure(SimpleNamespace(option=SimpleNamespace(**option)))

    monkeypatch.setenv("SLIDEKIT_TEST_POWERPOINT", "1")
    for workers in (2, "auto"):
        with pytest.raises(pytest.UsageError, match="SLIDEKIT_TEST_POWERPOINT=1 の試験は直列に実行する"):
            configure(numprocesses=workers)
    # -n を付けないか、-p no:xdist で xdist を読み込まなければ実行できる。
    configure(numprocesses=None)
    configure()
    # PowerPoint を使わない通常の試験は並列に実行できる。
    monkeypatch.delenv("SLIDEKIT_TEST_POWERPOINT")
    configure(numprocesses=2)
