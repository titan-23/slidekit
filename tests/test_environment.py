"""実行環境の設定の読み込みと検査、既定の書体の登録、資料の説明ファイルの場所を確かめる。"""

from pathlib import Path
import shutil
from types import SimpleNamespace

import pytest

from slidekit import config, fonts
from slidekit.config import load_common_environment, load_project
from slidekit.fonts import FontRegistry

# 候補表の書体のファイル。読み込みはファイルを開かないので、中身は何でもよい。
CANDIDATES = (
    "YuGothM.ttc",
    "YuGothB.ttc",
    "arial.ttf",
    "arialbd.ttf",
    "cambria.ttc",
    "calibri.ttf",
    "calibrib.ttf",
    "segoeui.ttf",
    "segoeuib.ttf",
    "meiryo.ttc",
    "meiryob.ttc",
)


def candidate_folder(path):
    path.mkdir(parents=True)
    for name in CANDIDATES:
        (path / name).write_bytes(name.encode())
    return path.resolve()


def font_files(project):
    return sorted(path for path in project.input_files if Path(path).suffix.lower() in {".ttf", ".ttc", ".otf"})


def test_font_directories_are_relative_to_environment_file(tmp_path, monkeypatch, write_deck):
    settings = tmp_path / "settings"
    (settings / "fonts").mkdir(parents=True)
    (tmp_path / "shared").mkdir()
    environment = settings / "environment.toml"
    environment.write_text('schema_version = 1\nfont_directories = ["fonts", "../shared"]\n', encoding="utf-8")
    monkeypatch.setenv("SLIDEKIT_ENVIRONMENT", str(environment))
    expected = [str((settings / "fonts").resolve()), str((tmp_path / "shared").resolve())]
    assert load_common_environment()["font_directories"] == expected
    assert load_project(write_deck(tmp_path / "deck")).environment["font_directories"] == expected
    assert fonts.font_directories(load_common_environment()) == tuple(map(Path, expected))


def test_missing_font_directory_is_rejected(tmp_path, isolated_environment, write_deck):
    root = write_deck(tmp_path / "deck")
    isolated_environment.write_text('schema_version = 1\nfont_directories = ["missing"]\n', encoding="utf-8")
    with pytest.raises(ValueError, match="font_directories のフォルダーがありません: .*missing"):
        load_project(root)
    for value in ('"fonts"', "[1]", '[""]'):
        isolated_environment.write_text(f"schema_version = 1\nfont_directories = {value}\n", encoding="utf-8")
        with pytest.raises(ValueError, match="font_directories は文字列の配列で指定してください"):
            load_project(root)


def test_default_font_table_is_shared(tmp_path, monkeypatch, isolated_environment, write_deck, dejavu):
    folder = (tmp_path / "fonts").resolve()
    folder.mkdir()
    shutil.copyfile(dejavu.path, folder / "DejaVuSans.ttf")
    shutil.copyfile(dejavu.bold_path, folder / "DejaVuSans-Bold.ttf")
    monkeypatch.setattr(fonts, "WINDOWS_FONT_FILES", {"DejaVu Sans": ("DejaVuSans.ttf", 0, "DejaVuSans-Bold.ttf", 0)})
    isolated_environment.write_text(f'schema_version = 1\nfont_directories = ["{folder}"]\n', encoding="utf-8")
    project = load_project(write_deck(tmp_path / "deck", theme=dejavu.theme))
    expected = {
        "path": str(folder / "DejaVuSans.ttf"),
        "index": 0,
        "bold_path": str(folder / "DejaVuSans-Bold.ttf"),
        "bold_index": 0,
    }
    # 読み込みの環境と寸法計算器が同じ候補表と同じフォルダーを使う。
    assert project.environment["fonts"] == {"DejaVu Sans": expected}
    with FontRegistry(project) as registry:
        assert registry.measure_text("Shared", 20)[0] > 0
        assert registry.inventory()["fonts"]["DejaVu Sans"]["path"] == expected["path"]
    alone = FontRegistry(
        SimpleNamespace(
            theme={"fonts": {"latin": "DejaVu Sans"}},
            environment={"font_directories": [str(folder)]},
            config={"math": {}},
        )
    )
    # 互換の measure_text は最後の行の下端に和文の書体を使うので、候補表にない Yu Gothic UI を探さないよう和文も指定する。
    alone.measure_text("Shared", 20, bold=True, japanese="DejaVu Sans")
    assert alone.inventory()["fonts"]["DejaVu Sans"]["bold"]["path"] == expected["bold_path"]


def test_empty_font_directories_disable_search(tmp_path, monkeypatch, isolated_environment, write_deck):
    monkeypatch.setattr(fonts, "_detect_font_directories", lambda environment: pytest.fail("自動探索を使った"))
    isolated_environment.write_text("schema_version = 1\nfont_directories = []\n", encoding="utf-8")
    project = load_project(write_deck(tmp_path / "deck"))
    assert project.environment["fonts"] == {}
    assert fonts.font_directories(project.environment) == ()
    with pytest.raises(
        FileNotFoundError,
        match=r'実フォントが見つからない: Arial。environment.toml の \[fonts."Arial"\] に path を指定する（探したフォルダー: なし）',
    ):
        project.measure_text("A", 20)


def test_only_used_fonts_are_added_and_tracked(tmp_path, isolated_environment, write_deck, dejavu):
    folder = candidate_folder(tmp_path / "fonts")
    # font_directories は表より前に書く（後ろに書くと表の項目になる）。
    isolated_environment.write_text(
        f'schema_version = 1\nfont_directories = ["{folder}"]\n'
        + dejavu.environment.removeprefix("schema_version = 1\n"),
        encoding="utf-8",
    )
    project = load_project(write_deck(tmp_path / "deck", theme=dejavu.theme))
    # テーマが DejaVu Sans だけなら、候補表の書体を環境にも入力の一覧にも加えない。
    assert set(project.environment["fonts"]) == {"DejaVu Sans"}
    assert font_files(project) == []
    theme = (
        dejavu.theme.replace('latin = "DejaVu Sans"', 'latin = "Arial"')
        + '[defaults]\nfont_japanese = "Meiryo"\n[roles.note]\nfont_latin = "Calibri"\n'
    )
    project = load_project(write_deck(tmp_path / "other", theme=theme))
    # テーマの fonts・defaults・roles に現れる書体だけを探す。
    assert set(project.environment["fonts"]) == {"DejaVu Sans", "Arial", "Meiryo", "Calibri"}
    assert project.environment["fonts"]["Arial"] == {
        "path": str(folder / "arial.ttf"),
        "index": 0,
        "bold_path": str(folder / "arialbd.ttf"),
        "bold_index": 0,
    }
    assert font_files(project) == []


def test_loading_does_not_hash_fonts(tmp_path, monkeypatch, isolated_environment, write_deck):
    folder = candidate_folder(tmp_path / "fonts")
    explicit = tmp_path / "explicit.ttf"
    explicit.write_bytes(b"explicit")
    isolated_environment.write_text(
        f'schema_version = 1\nfont_directories = ["{folder}"]\n[fonts.Example]\npath = "{explicit}"\n', encoding="utf-8"
    )
    hashed = []
    original = config.file_hash
    monkeypatch.setattr(config, "file_hash", lambda path: hashed.append(Path(path).resolve()) or original(path))
    project = load_project(
        write_deck(
            tmp_path / "deck", theme='[fonts]\njapanese = "Yu Gothic UI"\nlatin = "Example"\nmath = "Cambria Math"\n'
        )
    )
    assert set(project.environment["fonts"]) == {"Example", "Yu Gothic UI", "Cambria Math"}
    # 自動で見つけた書体も、明示した書体も、読み込みではハッシュ値を求めない。
    assert not [path for path in hashed if path == explicit.resolve() or folder in path.parents]
    assert font_files(project) == []
    # 実行環境の設定と同梱のテーマ・記述規則は追跡する。
    assert str(isolated_environment.resolve()) in project.input_files
    assert any(path.endswith("/themes/default.toml") for path in project.input_files)


def test_font_inside_output_is_rejected(tmp_path, isolated_environment, write_deck):
    root = write_deck(tmp_path / "deck")
    (root / "out").mkdir()
    (root / "out" / "font.ttf").write_bytes(b"font")
    isolated_environment.write_text(
        f'schema_version = 1\n[fonts.Example]\npath = "{root / "out" / "font.ttf"}"\n', encoding="utf-8"
    )
    with pytest.raises(ValueError, match="出力先が重複"):
        load_project(root)


@pytest.mark.parametrize("value", ["-1", '"10"', "true", "nan", "inf", "-0.5"])
def test_wait_seconds_must_be_non_negative_seconds(tmp_path, isolated_environment, write_deck, value):
    isolated_environment.write_text(f"schema_version = 1\n[powerpoint]\nwait_seconds = {value}\n", encoding="utf-8")
    with pytest.raises(ValueError, match="powerpoint.wait_seconds は 0 以上の秒数で指定してください"):
        load_project(write_deck(tmp_path / "deck"))


@pytest.mark.parametrize("value", ["0", "-2", "1.5", "true", '"4"'])
def test_jobs_must_be_positive_integer(tmp_path, isolated_environment, write_deck, value):
    isolated_environment.write_text(f"schema_version = 1\njobs = {value}\n", encoding="utf-8")
    with pytest.raises(ValueError, match="jobs は 1 以上の整数で指定してください"):
        load_project(write_deck(tmp_path / "deck"))


def test_valid_wait_seconds_and_jobs_are_kept(tmp_path, isolated_environment, write_deck):
    isolated_environment.write_text("schema_version = 1\njobs = 8\n[powerpoint]\nwait_seconds = 0\n", encoding="utf-8")
    environment = load_project(write_deck(tmp_path / "deck")).environment
    assert environment["jobs"] == 8
    assert environment["powerpoint"] == {"powershell": "powershell.exe", "timeout": 90, "wait_seconds": 0}
    isolated_environment.write_text("schema_version = 1\n[powerpoint]\nwait_seconds = 2.5\n", encoding="utf-8")
    assert load_common_environment()["powerpoint"]["wait_seconds"] == 2.5


def test_load_common_environment_reads_only_common_settings(tmp_path, monkeypatch, isolated_environment):
    folder = candidate_folder(tmp_path / "fonts")
    isolated_environment.write_text(
        f'schema_version = 1\njobs = 3\nfont_directories = ["{folder}"]\n[powerpoint]\nwait_seconds = 5\n',
        encoding="utf-8",
    )
    environment = load_common_environment()
    assert environment["jobs"] == 3
    assert environment["powerpoint"] == {"powershell": "powershell.exe", "timeout": 90, "wait_seconds": 5}
    # 資料がないので、書体は探さない。
    assert environment["fonts"] == {}
    assert [record["path"] for record in environment["config_files"]] == [str(isolated_environment)]
    assert environment["config_files"][0]["sha256"] == config.file_hash(isolated_environment)
    isolated_environment.write_text("schema_version = 1\njobs = 0\n", encoding="utf-8")
    with pytest.raises(ValueError, match="jobs"):
        load_common_environment()
    monkeypatch.setenv("SLIDEKIT_ENVIRONMENT", str(tmp_path / "missing.toml"))
    with pytest.raises(FileNotFoundError, match="実行環境の設定が見つかりません"):
        load_common_environment()
    # 指定がなく、利用者ごとの既定の場所にもなければ、既定値だけを返す。
    monkeypatch.delenv("SLIDEKIT_ENVIRONMENT")
    for name in ("HOME", "USERPROFILE", "APPDATA"):
        monkeypatch.setenv(name, str(tmp_path / "home"))
    assert load_common_environment() == {
        "fonts": {},
        "config_files": [],
        "powerpoint": {"powershell": "powershell.exe", "timeout": 90},
    }


def test_brief_path_is_project_brief_or_none(tmp_path, write_deck):
    root = write_deck(tmp_path / "deck")
    assert load_project(root).brief_path is None
    (root / "brief.md").write_text("発表の目的", encoding="utf-8")
    project = load_project(root)
    assert project.brief_path == (root / "brief.md").resolve()
    assert str(project.brief_path) in project.input_files
