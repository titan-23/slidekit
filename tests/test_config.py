"""資料の設定の読み込み（load_project）を確かめる。

設定名と値の問題を 1 つの SettingsError にまとめること、設定の継承（extends）、inputs.python の省略、
画像と互換画像の設定、実行環境の設定名、Project の書体（fonts）と文献（bibliography）を扱う。
関数の単位の設定の検査は test_schema.py にある。
"""

import json
from pathlib import Path
import sys

import pytest

from fake_fonts import FakeFonts
from slidekit.config import MAX_EXTENDS, load_common_environment, load_project
from slidekit.files import file_hash
from slidekit.fonts import FontRegistry
from slidekit.schema import SettingsError

DECK = """schema_version = 1
entrypoint = "slides.py:build"
[metadata]
title = "設定の試験"
author = "slidekit example"
last_modified_by = "slidekit example"
"""
BUILD = "def build(context):\n    pass\n"
# 公開されている一般的な文献だけを使う。
BIBLIOGRAPHY = """@book{knuth1984,
  author = {Donald E. Knuth},
  title = {The {\\TeX}book},
  publisher = {Addison-Wesley},
  year = {1984}
}
"""


def toml(value):
    """TOML の基本文字列（JSON の文字列の表記は TOML でも読むことができる）。"""
    return json.dumps(str(value))


def write(path, text):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return path


def settings_error(root):
    with pytest.raises(SettingsError) as caught:
        load_project(root)
    assert caught.value.code == "settings_invalid"
    return caught.value


def lines(error):
    return str(error).splitlines()


def make_template(path):
    from pptx import Presentation

    presentation = Presentation()
    presentation.slides.add_slide(presentation.slide_layouts[6])
    path.parent.mkdir(parents=True, exist_ok=True)
    presentation.save(path)


# 設定名と値の検査


@pytest.mark.parametrize(
    ("files", "expected"),
    [
        (
            {"rules.toml": 'forbidden_character = ["。"]\n'},
            "rules.toml: forbidden_character は不明な設定です（候補: forbidden_characters）",
        ),
        (
            {"rules.toml": "[math_candidates]\nexclude_pattern = []\n"},
            "rules.toml: math_candidates.exclude_pattern は不明な設定です（候補: exclude_patterns、exclude_page_kinds、exclude_roles）",
        ),
        ({"theme.toml": "[roles.note]\nsize = 14\n"}, "theme.toml: roles.note.size は不明な設定です（候補: size_pt）"),
        (
            {"theme.toml": '[page_kinds.contents]\ntitle_role = "slide_title"\n'},
            "theme.toml: page_kinds.contents は不明なページ種別です（候補: content）",
        ),
        (
            {"deck.toml": "\n[images]\nraster_scal = 4\n"},
            "deck.toml: images.raster_scal は不明な設定です（候補: raster_scale）",
        ),
    ],
)
def test_unknown_rules_and_theme_keys_name_file_key_and_candidates(tmp_path, write_deck, files, expected):
    root = write_deck(
        tmp_path / "deck", deck='theme = "theme.toml"\nrules = "rules.toml"\n' + DECK, extra=files.get("deck.toml", "")
    )
    write(root / "theme.toml", files.get("theme.toml", ""))
    write(root / "rules.toml", files.get("rules.toml", ""))
    error = settings_error(root)
    assert lines(error) == [expected]
    [problem] = error.details["problems"]
    assert expected.startswith(f"{problem['file']}: {problem['key']} ")


@pytest.mark.parametrize(
    ("extra", "expected"),
    [
        (
            '\n[math]\nfallbak="required"\n',
            "deck.toml: math.fallbak は不明な設定です（候補: fallback、fallback_scale）",
        ),
        ('\n[inputs]\npyton=["slides.py"]\n', "deck.toml: inputs.pyton は不明な設定です（候補: python）"),
        ("\n[images]\nraster_scal = 4\n", "deck.toml: images.raster_scal は不明な設定です（候補: raster_scale）"),
    ],
)
def test_unknown_deck_settings_have_candidates(tmp_path, write_deck, extra, expected):
    # test_model_config.py::test_unknown_settings_are_not_silently_ignored と同じ設定で、候補が示されることを確かめる。
    assert lines(settings_error(write_deck(tmp_path / "deck", deck=DECK, extra=extra))) == [expected]
    top = write_deck(tmp_path / "top", deck='entrypont = "slides.py:build"\n' + DECK)
    assert lines(settings_error(top)) == ["deck.toml: entrypont は不明な設定です（候補: entrypoint）"]


def test_settings_problems_are_reported_together(tmp_path, write_deck):
    # 関数の単位の同名の試験は test_schema.py にある。こちらは load_project を通す。
    root = write_deck(tmp_path / "deck", deck='rules = "rules.toml"\n' + DECK, theme="[roles.note]\nsize = 14\n")
    write(root / "rules.toml", 'forbidden_character = ["。"]\nnotes_style = "desu_masu"\n')
    error = settings_error(root)
    assert lines(error) == [
        "theme.toml: roles.note.size は不明な設定です（候補: size_pt）",
        "rules.toml: forbidden_character は不明な設定です（候補: forbidden_characters）",
        "rules.toml: notes_style は unrestricted、polite のどれかで指定してください（値: 'desu_masu'）",
    ]
    assert [(problem["file"], problem["key"]) for problem in error.details["problems"]] == [
        ("theme.toml", "roles.note.size"),
        ("rules.toml", "forbidden_character"),
        ("rules.toml", "notes_style"),
    ]


def test_problems_in_every_settings_file_are_reported_together(tmp_path, write_deck, isolated_environment):
    root = write_deck(
        tmp_path / "deck",
        deck='rules = "settings/rules.toml"\n' + DECK,
        theme='[colors]\nteal = "07F87"\n',
        extra="\n[images]\nraster_scal = 4\n",
    )
    write(root / "settings/rules.toml", 'notes_style = "desu_masu"\n')
    isolated_environment.write_text("schema_version = 1\njbos = 2\n", encoding="utf-8")
    assert lines(settings_error(root)) == [
        "deck.toml: images.raster_scal は不明な設定です（候補: raster_scale）",
        "theme.toml: colors.teal は 6 桁の 16 進数で指定してください（値: '07F87'）",
        "settings/rules.toml: notes_style は unrestricted、polite のどれかで指定してください（値: 'desu_masu'）",
        # 資料のディレクトリの外のファイルは絶対パスで示す。
        f"{isolated_environment}: jbos は不明な設定です（候補: jobs）",
    ]


def test_references_are_checked_after_merge_at_load(tmp_path, write_deck):
    root = write_deck(tmp_path / "deck", deck='rules = "rules.toml"\n' + DECK, theme='[roles.note]\ncolor = "mutd"\n')
    write(root / "rules.toml", '[math_candidates]\nexclude_roles = ["referenses"]\n')
    assert lines(settings_error(root)) == [
        "theme.toml: roles.note.color の色 mutd はテーマにありません（候補: muted）",
        "rules.toml: math_candidates.exclude_roles の 1 番目の役割 referenses はテーマにありません（候補: references）",
    ]
    # 資料のテーマで加えた役割は、記述規則から参照できる。
    write(root / "theme.toml", '[roles.caption]\ninherits = "note"\n')
    write(root / "rules.toml", '[math_candidates]\nexclude_roles = ["caption"]\n')
    assert load_project(root).theme["roles"]["caption"]["size_pt"] == 14.0


def test_toml_syntax_error_names_file(tmp_path, write_deck, isolated_environment):
    root = write_deck(
        tmp_path / "deck",
        deck='rules = "rules.toml"\n' + DECK,
        theme="[colors\nprimary = 1\n",
        extra="\n[images]\nraster_scal = 4\n",
    )
    write(root / "rules.toml", "notes_style = \n")
    found = lines(settings_error(root))
    # 読み込みに失敗したファイルがあっても、ほかのファイルの問題も合わせて報告する。
    assert found[0] == "deck.toml: images.raster_scal は不明な設定です（候補: raster_scale）"
    assert found[1].startswith("theme.toml: TOML として読めません（") and found[2].startswith(
        "rules.toml: TOML として読めません（"
    )
    assert len(found) == 3
    write(root / "deck.toml", DECK + "[images\n")
    [message] = lines(settings_error(root))
    assert message.startswith("deck.toml: TOML として読めません（")
    write(root / "deck.toml", DECK)
    isolated_environment.write_text("schema_version = 1\njobs = \n", encoding="utf-8")
    [message] = lines(settings_error(root))
    assert message.startswith(f"{isolated_environment}: TOML として読めません（")


def test_deck_values_and_requirements_are_checked_at_load(tmp_path, write_deck):
    root = write_deck(
        tmp_path / "deck",
        deck='schema_version = 1\nsize = { width = 960, height = 540 }\n[metadata]\ntitle = " "\n',
        extra='[[inputs.documents]]\nid = "outline"\npath = "outline.md"\n'
        '[[inputs.documents]]\nid = "outline"\npath = "notes.md"\n',
    )
    assert lines(settings_error(root)) == [
        # 使われていなかった {width, height} の形は受け付けない。
        "deck.toml: size は 16:9、4:3、または [幅, 高さ]（正の数、pt）で指定してください（値: {'width': 960, 'height': 540}）",
        "deck.toml: inputs.documents の 2 番目の id outline は 1 番目と重複しています",
        "deck.toml: metadata.title を明示指定してください",
        "deck.toml: metadata.author を明示指定してください",
        "deck.toml: metadata.last_modified_by を明示指定してください",
    ]


def test_problem_file_names_are_relative_inside_project(tmp_path, write_deck):
    shared = write(tmp_path / "shared/rules.toml", "require_notes = 1\n")
    root = write_deck(
        tmp_path / "deck",
        deck='theme = "settings/theme.toml"\nrules = "../shared/rules.toml"\n' + DECK,
        extra='[inputs]\nread_roots = [".", "../shared"]\n',
    )
    write(root / "settings/theme.toml", "[margins]\nleft = -1\n")
    assert lines(settings_error(root)) == [
        "settings/theme.toml: margins.left は 0 以上の有限の数で指定してください（値: -1）",
        f"{shared.resolve()}: require_notes は真偽値で指定してください（値: 1）",
    ]


def test_settings_outside_read_roots_are_not_read(tmp_path, write_deck):
    # 参照範囲の外の設定ファイルは中身を読まず、参照範囲の誤りとして報告する（0.1 と同じ）。
    outside = write(tmp_path / "outside/theme.toml", "[colors\n")
    write(tmp_path / "outside/environment.toml", "jobs = \n")
    root = write_deck(tmp_path / "deck", deck='theme = "../outside/theme.toml"\n' + DECK)
    # 入力の解決の誤りは、設定の場所を示す設定の誤りにする（トレースバックを出さない）。元の例外は __cause__ に残す。
    with pytest.raises(SettingsError, match="theme の指すファイルを使えません（宣言した参照範囲外です") as caught:
        load_project(root)
    assert isinstance(caught.value.__cause__, ValueError) and str(outside.resolve()) in str(caught.value)
    write_deck(root, deck='environment_override = "../outside/environment.toml"\n' + DECK)
    with pytest.raises(SettingsError, match="environment_override の指すファイルを使えません（宣言した参照範囲外です"):
        load_project(root)


def test_rules_references_are_not_checked_against_unread_theme(tmp_path, write_deck):
    # 資料のテーマを読み込むのに失敗したときは、記述規則の役割の参照を確かめず、失敗の原因を示す。
    # 同梱のテーマで確かめると、資料のテーマで加えた役割がないという誤検出になる。
    root = write_deck(tmp_path / "deck", deck='theme = "themes.toml"\nrules = "rules.toml"\n' + DECK)
    write(root / "theme.toml", '[roles.caption]\ninherits = "note"\n')
    write(root / "rules.toml", '[math_candidates]\nexclude_roles = ["caption"]\n')
    # パスの誤記は、ないファイルとして示す。
    with pytest.raises(SettingsError, match="theme の指すファイルを使えません（入力ファイルが見つかりません") as caught:
        load_project(root)
    assert str((root / "themes.toml").resolve()) in str(caught.value)
    assert isinstance(caught.value.__cause__, FileNotFoundError)
    # 参照範囲の外のテーマは参照範囲の誤りとして示す。
    write(tmp_path / "outside/theme.toml", '[roles.caption]\ninherits = "note"\n')
    write_deck(root, deck='theme = "../outside/theme.toml"\nrules = "rules.toml"\n' + DECK)
    with pytest.raises(SettingsError, match="theme の指すファイルを使えません（宣言した参照範囲外です"):
        load_project(root)
    # 構文の誤りと値の形の誤りは、その問題だけを示す。
    write(root / "theme.toml", '[roles.caption]\ninherits = "note"\n[colors\n')
    write_deck(root, deck='theme = "theme.toml"\nrules = "rules.toml"\n' + DECK)
    [message] = lines(settings_error(root))
    assert message.startswith("theme.toml: TOML として読めません（")
    write_deck(root, deck='theme = 3\nrules = "rules.toml"\n' + DECK)
    assert lines(settings_error(root)) == ["deck.toml: theme は設定名またはファイルのパスで指定してください（値: 3）"]


def test_page_background_none_is_accepted_at_load(tmp_path, write_deck):
    # 背景の図形を作らない "none" を、読み込みの検査が受け付ける。
    root = write_deck(
        tmp_path / "deck",
        theme='[page_kinds.content]\nbackground = "none"\n[page_kinds.cover]\nbackground = "surface"\n',
    )
    project = load_project(root)
    assert project.theme["page_kinds"]["content"]["background"] == "none"
    write(root / "theme.toml", '[page_kinds.content]\nbackground = "paper"\n')
    assert lines(settings_error(root)) == ["theme.toml: page_kinds.content.background の色 paper はテーマにありません"]


def test_decorations_are_checked_at_load(tmp_path, write_deck):
    # ページ種別の飾りの設定は、テーマの読み込みで確かめる。
    decoration = (
        '[[page_kinds.section.decorations]]\nkind = "line"\nid = "rule"\nstart = [0.1, 0.67]\nend = [0.9, 0.67]\n'
    )
    root = write_deck(tmp_path / "deck", theme=decoration + 'line_color = "primary"\n')
    [loaded] = load_project(root).theme["page_kinds"]["section"]["decorations"]
    assert (loaded["kind"], loaded["id"], loaded["line_color"]) == ("line", "rule", "primary")
    write(
        root / "theme.toml",
        decoration + 'line_color = "primery"\n[[page_kinds.section.decorations]]\nkind = "circle"\nid = "mark"\n',
    )
    assert lines(settings_error(root)) == [
        "theme.toml: page_kinds.section.decorations の 2 番目の kind は line、arrow、box のどれかで指定してください（値: 'circle'）",
        "theme.toml: page_kinds.section.decorations の 1 番目の line_color の色 primery はテーマにありません（候補: primary）",
    ]


# inputs.python


def test_inputs_python_is_optional_and_entry_is_tracked(tmp_path, write_deck):
    root = write_deck(tmp_path / "deck", deck=DECK)
    project = load_project(root)
    entry = str((root / "slides.py").resolve())
    assert project.config["inputs"]["python"] == [entry]
    assert project.input_files[entry] == file_hash(root / "slides.py")
    write(root / "main.py", BUILD)
    project = load_project(write_deck(root, deck=DECK.replace("slides.py:build", "main.py:build")))
    assert project.config["inputs"]["python"] == [str((root / "main.py").resolve())]


def test_listed_python_without_entry_gets_entry_added(tmp_path, write_deck):
    root = write_deck(tmp_path / "deck", deck=DECK, extra='[inputs]\npython = ["helper.py"]\n')
    write(root / "helper.py", "VALUE = 1\n")
    project = load_project(root)
    entry, helper = (str((root / name).resolve()) for name in ("slides.py", "helper.py"))
    assert project.config["inputs"]["python"] == [entry, helper]
    assert entry in project.input_files and helper in project.input_files
    # 入口を書いてあれば、書いた順のまま使う。
    write_deck(root, deck=DECK, extra='[inputs]\npython = ["helper.py", "slides.py"]\n')
    assert load_project(root).config["inputs"]["python"] == [helper, entry]
    write_deck(root, deck=DECK, extra='[inputs]\npython = ["missing.py"]\n')
    with pytest.raises(SettingsError, match="inputs.python の指すファイルを使えません（入力ファイルが見つかりません"):
        load_project(root)
    # 入力文書は識別子で示す。
    write_deck(root, deck=DECK, extra='[[inputs.documents]]\nid = "notes"\npath = "README.md"\n')
    with pytest.raises(SettingsError, match="deck.toml: inputs.documents の notes の指すファイルを使えません"):
        load_project(root)


# 設定の継承


def test_extends_resolves_paths_relative_to_each_file(tmp_path, write_deck):
    base = write_deck(
        tmp_path / "base",
        deck="""schema_version = 1
entrypoint = "slides.py:build"
environment_override = "environment.toml"
[metadata]
title = "共通の設定"
author = "slidekit example"
last_modified_by = "slidekit example"
[inputs]
python = ["helper.py"]
asset_dirs = ["figures"]
[[inputs.documents]]
id = "outline"
path = "outline.md"
[template]
path = "template.pptx"
""",
        theme='[colors]\nprimary = "112233"\n',
    )
    write(base / "helper.py", "VALUE = 1\n")
    write(base / "outline.md", "# 構成\n")
    write(base / "environment.toml", "schema_version = 1\n[powerpoint]\ntimeout = 45\n")
    (base / "figures").mkdir()
    make_template(base / "template.pptx")
    child = tmp_path / "child"
    write(child / "deck.toml", 'extends = "../base"\nrules = "rules.toml"\n')
    write(child / "rules.toml", "require_notes = true\n")
    project = load_project(child)
    base, child = base.resolve(), child.resolve()
    config = project.config
    assert project.root == child and project.path == child / "deck.toml"
    assert config["entrypoint"] == f"{base / 'slides.py'}:build"
    assert config["theme"] == str(base / "theme.toml") and project.theme["colors"]["primary"] == "112233"
    # 資料自身に書いたパスは、資料のディレクトリを基準にする。
    assert config["rules"] == str(child / "rules.toml") and project.rules["require_notes"] is True
    assert config["inputs"]["python"] == [str(base / "slides.py"), str(base / "helper.py")]
    assert config["inputs"]["asset_dirs"] == [str(base / "figures")]
    assert config["inputs"]["documents"] == [{"id": "outline", "path": str(base / "outline.md")}]
    assert config["template"]["path"] == str(base / "template.pptx")
    assert config["environment_override"] == str(base / "environment.toml")
    assert project.environment["powerpoint"]["timeout"] == 45
    # 継承元は単独でも読むことができる（同じ設定を、それぞれのディレクトリを基準に解決する）。
    assert load_project(base).config["inputs"]["documents"] == config["inputs"]["documents"]


def test_extends_inherits_rules_bibliography_and_bundled_names(tmp_path, write_deck):
    # 子が書かない記述規則と文献は、継承元のディレクトリを基準にする。継承元が書いた同梱の設定名は、そのまま使う。
    base = write_deck(
        tmp_path / "base",
        deck='theme = "default"\nrules = "rules.toml"\n' + DECK,
        extra='[inputs]\nbibliography = ["references.bib"]\n',
    )
    write(base / "rules.toml", 'notes_style = "polite"\n')
    write(base / "references.bib", BIBLIOGRAPHY)
    child = tmp_path / "child"
    write(child / "deck.toml", 'extends = "../base"\n')
    project = load_project(child)
    base = base.resolve()
    assert project.config["theme"] == "default"
    assert project.config["rules"] == str(base / "rules.toml") and project.rules["notes_style"] == "polite"
    assert project.config["inputs"]["bibliography"] == [str(base / "references.bib")]
    assert project.input_files[str(base / "references.bib")] == file_hash(base / "references.bib")
    write_deck(base, deck='theme = "default"\nrules = "general"\n' + DECK)
    assert load_project(child).config["rules"] == "general"


def test_extends_merges_tables_and_replaces_arrays(tmp_path, write_deck):
    write_deck(
        tmp_path / "base",
        deck="""schema_version = 1
[metadata]
title = "共通の題名"
author = "slidekit example"
last_modified_by = "slidekit example"
subject = "共通の主題"
[math]
fallback = "required"
functions = ["argmax"]
[[inputs.documents]]
id = "first"
path = "first.md"
[[inputs.documents]]
id = "second"
path = "second.md"
""",
    )
    child = write_deck(
        tmp_path / "child",
        deck="""extends = "../base/deck.toml"
[metadata]
title = "個別の題名"
[math]
functions = ["rank"]
[[inputs.documents]]
id = "own"
path = "own.md"
""",
    )
    write(child / "own.md", "# 個別\n")
    config = load_project(child).config
    assert config["metadata"] == {
        "title": "個別の題名",
        "author": "slidekit example",
        "last_modified_by": "slidekit example",
        "subject": "共通の主題",
    }
    assert config["math"]["fallback"] == "required" and config["math"]["functions"] == ["rank"]
    # 配列は配列ごと置き換えるので、継承元の元資料（ないファイル）は使わない。
    assert config["inputs"]["documents"] == [{"id": "own", "path": str((child / "own.md").resolve())}]
    # 入口を誰も書かなければ、資料のディレクトリの slides.py を使う。
    assert config["entrypoint"] == "slides.py:build"


def test_extends_does_not_inherit_output_or_deliverables(tmp_path, write_deck):
    base = write_deck(
        tmp_path / "base",
        deck=DECK.replace("[metadata]", 'output_dir = "build"\n[metadata]'),
        extra='[deliverables]\npptx = "shared.pptx"\nrecord = "shared.json"\n',
    )
    child = tmp_path / "child"
    write(child / "deck.toml", 'extends = "../base/deck.toml"\n')
    project = load_project(child)
    child = child.resolve()
    assert project.output_dir == child / "out" and project.config["output_dir"] == str(child / "out")
    assert project.config["deliverables"] == {
        "pptx": str(child / "presentation.pptx"),
        "pdf": str(child / "presentation.pdf"),
        "notes": str(child / "speaker_notes.md"),
        "record": str(child / "deliverables.json"),
    }
    parent = load_project(base)
    assert parent.output_dir == base.resolve() / "build"
    assert parent.config["deliverables"]["pptx"] == str(base.resolve() / "shared.pptx")


def test_extends_cycle_and_depth_are_rejected(tmp_path, write_deck):
    first, second = tmp_path / "first", tmp_path / "second"
    write_deck(first, deck='extends = "../second"\n' + DECK)
    write_deck(second, deck='extends = "../first/deck.toml"\n' + DECK)
    [message] = lines(settings_error(first))
    assert message == f"{(second / 'deck.toml').resolve()}: extends が循環しています: {(first / 'deck.toml').resolve()}"
    alone = write_deck(tmp_path / "alone", deck='extends = "deck.toml"\n' + DECK)
    assert lines(settings_error(alone)) == [f"deck.toml: extends が循環しています: {(alone / 'deck.toml').resolve()}"]
    # extends をたどるのは MAX_EXTENDS 回まで。
    levels = [tmp_path / f"level{number}" for number in range(MAX_EXTENDS + 2)]
    for number, level in enumerate(levels):
        parent = f'extends = "../level{number + 1}"\n' if number + 1 < len(levels) else ""
        write_deck(level, deck=parent + DECK)
    assert load_project(levels[1]).config["extends"] == [str((level / "deck.toml").resolve()) for level in levels[2:]]
    [message] = lines(settings_error(levels[0]))
    assert message == (
        f"{(levels[MAX_EXTENDS] / 'deck.toml').resolve()}: extends の継承は {MAX_EXTENDS} 段までにしてください: "
        f"{(levels[MAX_EXTENDS + 1] / 'deck.toml').resolve()}"
    )
    for value, expected in (
        ("3", "deck.toml: extends は deck.toml のパスで指定してください（値: 3）"),
        (
            '"https://example.invalid/deck.toml"',
            "deck.toml: extends は deck.toml のパスで指定してください（値: 'https://example.invalid/deck.toml'）",
        ),
        ('"../missing"', f"deck.toml: extends の deck.toml がありません: {(tmp_path / 'missing').resolve()}"),
    ):
        assert lines(settings_error(write_deck(tmp_path / "wrong", deck=f"extends = {value}\n" + DECK))) == [expected]


def test_extends_adds_parent_dirs_to_read_roots(tmp_path, write_deck):
    shared = write(tmp_path / "shared/notes.py", "NOTE = 1\n").parent.resolve()
    base = write_deck(tmp_path / "base", deck=DECK, extra='[inputs]\nread_roots = [".", "../shared"]\n')
    write(base / "parts.py", "VALUE = 1\n")
    child = tmp_path / "child"
    write(child / "deck.toml", 'extends = "../base"\n')
    project = load_project(child)
    base, child = base.resolve(), child.resolve()
    # 最も近いファイルが指定した列に、継承をたどった全ファイルのディレクトリを加える。
    assert project.read_roots == (base, shared, child)
    assert project.resolve_input(base / "parts.py") == base / "parts.py"
    assert project.resolve_input(shared / "notes.py") == shared / "notes.py"
    write(child / "deck.toml", 'extends = "../base"\n[inputs]\nread_roots = ["."]\n')
    assert load_project(child).read_roots == (child, base)
    # 継承がなければ、指定した列のまま。
    assert load_project(base).read_roots == (base, shared)


def test_extends_tracks_all_settings_files(tmp_path, write_deck):
    base = write_deck(tmp_path / "base", deck=DECK)
    write(base / "brief.md", "共通の説明\n")
    child = tmp_path / "child"
    write(child / "deck.toml", 'extends = "../base"\n')
    project = load_project(child)
    base, child = base.resolve(), child.resolve()
    for settings in (child / "deck.toml", base / "deck.toml"):
        assert project.input_files[str(settings)] == file_hash(settings)
    assert project.config["extends"] == [str(base / "deck.toml")]
    # 説明ファイルは資料のディレクトリから継承元の順に探す。
    assert project.brief_path == base / "brief.md" and str(base / "brief.md") in project.input_files
    write(child / "brief.md", "個別の説明\n")
    assert load_project(child).brief_path == child / "brief.md"
    assert load_project(base).config["extends"] == []
    # 継承元の設定ファイルも出力先と重ねることができない。
    write(child / "deck.toml", 'extends = "../base"\noutput_dir = "../base"\n')
    with pytest.raises(ValueError, match="資料設定と保存先が重複しています"):
        load_project(child)


def test_extends_problems_name_each_file(tmp_path, write_deck):
    base = write_deck(tmp_path / "base", deck="schema_version = 1\n[images]\nraster_scal = 4\n")
    child = tmp_path / "child"
    write(child / "deck.toml", 'extends = "../base"\nentrypont = "slides.py:build"\n[metadata]\ntitle = "個別の題名"\n')
    assert lines(settings_error(child)) == [
        "deck.toml: entrypont は不明な設定です（候補: entrypoint）",
        f"{(base / 'deck.toml').resolve()}: images.raster_scal は不明な設定です（候補: raster_scale）",
        "deck.toml: metadata.author を明示指定してください",
        "deck.toml: metadata.last_modified_by を明示指定してください",
    ]


def test_extends_errors_are_reported_with_other_problems(tmp_path, write_deck):
    # 継承をたどるのに失敗しても、読み込みに成功した deck.toml の設定名と値の問題を合わせて示す。合成した後の検査（必須の文書情報など）は行わない。
    root = write_deck(tmp_path / "deck", deck='extends = 3\nentrypont = "slides.py:build"\n' + DECK)
    typo = "deck.toml: entrypont は不明な設定です（候補: entrypoint）"
    assert lines(settings_error(root)) == [typo, "deck.toml: extends は deck.toml のパスで指定してください（値: 3）"]
    base = write(tmp_path / "base/deck.toml", 'extends = "../missing"\n[images]\nraster_scal = 4\n').resolve()
    write(root / "deck.toml", 'extends = "../base"\nentrypont = "slides.py:build"\n')
    read_problems = [typo, f"{base}: images.raster_scal は不明な設定です（候補: raster_scale）"]
    assert lines(settings_error(root)) == read_problems + [
        f"{base}: extends の deck.toml がありません: {(tmp_path / 'missing').resolve()}"
    ]
    write(base, 'extends = "../deck"\n[images]\nraster_scal = 4\n')
    assert lines(settings_error(root)) == read_problems + [
        f"{base}: extends が循環しています: {(root / 'deck.toml').resolve()}"
    ]
    write(base, "[images\n")
    found = lines(settings_error(root))
    assert found[0] == typo and found[1].startswith(f"{base}: TOML として読めません（") and len(found) == 2


# 画像と互換画像の設定


def test_fallback_scale_is_validated(tmp_path, write_deck):
    root = write_deck(tmp_path / "deck", deck=DECK)
    assert load_project(root).config["math"] == {"fallback": "off", "fallback_scale": 4, "functions": []}
    for value in ("2.5", "8"):
        write_deck(root, deck=DECK, extra=f"[math]\nfallback_scale = {value}\n")
        assert load_project(root).config["math"]["fallback_scale"] == float(value)
    for value in ("1", "0", "8.5", '"4"', "true", "nan"):
        write_deck(root, deck=DECK, extra=f"[math]\nfallback_scale = {value}\n")
        [message] = lines(settings_error(root))
        assert message.startswith("deck.toml: math.fallback_scale は 1 より大きく 8 以下の数で指定してください（値: ")


def test_images_settings_are_validated(tmp_path, write_deck):
    root = write_deck(tmp_path / "deck", deck=DECK)
    assert load_project(root).config["images"] == {"raster_scale": 2, "vector": False}
    write_deck(root, deck=DECK, extra="[images]\nraster_scale = 4.5\nvector = true\n")
    assert load_project(root).config["images"] == {"raster_scale": 4.5, "vector": True}
    for value in ("1", "16"):
        write_deck(root, deck=DECK, extra=f"[images]\nraster_scale = {value}\n")
        assert load_project(root).config["images"]["raster_scale"] == int(value)
    for value in ("0.5", "17", '"2"', "true"):
        write_deck(root, deck=DECK, extra=f"[images]\nraster_scale = {value}\n")
        [message] = lines(settings_error(root))
        assert message.startswith("deck.toml: images.raster_scale は 1 以上 16 以下の数で指定してください（値: ")
    # 表の中では未知の名前を値の問題より先に示す。
    write_deck(root, deck=DECK, extra='[images]\nvector = "yes"\nscale = 2\n')
    assert lines(settings_error(root)) == [
        "deck.toml: images.scale は不明な設定です",
        "deck.toml: images.vector は真偽値で指定してください（値: 'yes'）",
    ]


# 実行環境の設定名


@pytest.mark.parametrize(
    ("setting", "expected"),
    [
        ('temporary_dir = "tmp"\n', "temporary_dir は不明な設定です"),
        ('cache_dir = "cache"\n', "cache_dir は不明な設定です"),
        (
            '[fonts.Example]\npath = "a.ttf"\nwindows_path = "C:\\\\Windows\\\\Fonts\\\\a.ttf"\n',
            "fonts.Example.windows_path は不明な設定です",
        ),
        (
            '[fonts."Yu Gothic UI"]\npath = "YuGothM.ttc"\nwindows_bold_path = "C:\\\\Windows\\\\Fonts\\\\YuGothB.ttc"\n',
            'fonts."Yu Gothic UI".windows_bold_path は不明な設定です（候補: bold_path）',
        ),
        ("jbos = 2\n", "jbos は不明な設定です（候補: jobs）"),
        ("[powerpoint]\ntimout = 30\n", "powerpoint.timout は不明な設定です（候補: timeout）"),
        ("[powerpnt]\ntimeout = 30\n", "powerpnt は不明な設定です（候補: powerpoint）"),
    ],
)
def test_unknown_environment_settings_are_rejected(tmp_path, write_deck, isolated_environment, setting, expected):
    isolated_environment.write_text("schema_version = 1\n" + setting, encoding="utf-8")
    assert lines(settings_error(write_deck(tmp_path / "deck"))) == [f"{isolated_environment}: {expected}"]
    with pytest.raises(SettingsError, match="不明な設定です"):
        load_common_environment()


def test_environment_accepts_documented_settings(tmp_path, write_deck, isolated_environment):
    fonts = (tmp_path / "fonts").resolve()
    fonts.mkdir()
    isolated_environment.write_text(
        f"""schema_version = 1
jobs = 3
font_directories = [{toml(fonts)}]
[fonts."Yu Gothic UI"]
path = "YuGothM.ttc"
index = 1
bold_path = "YuGothB.ttc"
bold_index = 1
[powerpoint]
powershell = "powershell.exe"
timeout = 60
wait_seconds = 0
""",
        encoding="utf-8",
    )
    environment = load_project(write_deck(tmp_path / "deck")).environment
    folder = isolated_environment.parent
    assert environment["jobs"] == 3 and environment["font_directories"] == [str(fonts)]
    assert environment["fonts"]["Yu Gothic UI"] == {
        "path": str(folder / "YuGothM.ttc"),
        "index": 1,
        "bold_path": str(folder / "YuGothB.ttc"),
        "bold_index": 1,
    }
    assert environment["powerpoint"] == {"powershell": "powershell.exe", "timeout": 60, "wait_seconds": 0}


def test_environment_problems_are_collected(tmp_path, write_deck, isolated_environment):
    isolated_environment.write_text(
        'schema_version = 1\njobs = 0\nfont_directories = ["missing", "also_missing"]\n'
        '[fonts.Example]\nbold_path = "bold.ttf"\nindex = -1\n'
        "[powerpoint]\ntimeout = 0\n",
        encoding="utf-8",
    )
    folder = isolated_environment.parent
    expected = [
        f"{isolated_environment}: jobs は 1 以上の整数で指定してください（値: 0）",
        f"{isolated_environment}: fonts.Example.index は 0 以上の整数で指定してください（値: -1）",
        f"{isolated_environment}: powerpoint.timeout は正の秒数で指定してください（値: 0）",
        f"{isolated_environment}: font_directories のフォルダーがありません: {folder / 'missing'}",
        f"{isolated_environment}: font_directories のフォルダーがありません: {folder / 'also_missing'}",
        # 重ねた後の書体に path がなければ、その書体を書いたファイルを示す。
        f"{isolated_environment}: fonts.Example に path を指定してください",
    ]
    assert lines(settings_error(write_deck(tmp_path / "deck"))) == expected
    with pytest.raises(SettingsError) as caught:
        load_common_environment()
    assert lines(caught.value) == expected


def test_environment_override_completes_font_after_merge(tmp_path, write_deck, isolated_environment):
    isolated_environment.write_text('schema_version = 1\n[fonts.Example]\nbold_path = "bold.ttf"\n', encoding="utf-8")
    root = write_deck(tmp_path / "deck", deck='environment_override = "local.toml"\n' + DECK)
    write(root / "local.toml", 'schema_version = 1\n[fonts.Example]\npath = "regular.ttf"\n')
    font = load_project(root).environment["fonts"]["Example"]
    assert font == {
        "path": str(root.resolve() / "regular.ttf"),
        "bold_path": str(isolated_environment.parent / "bold.ttf"),
    }
    write(root / "local.toml", "schema_version = 1\n[fonts.Example]\nindex = 0\n")
    assert lines(settings_error(root)) == ["local.toml: fonts.Example に path を指定してください"]
    # 上書きの設定を読み込むのに失敗したときは、重ねた後の書体を確かめず、失敗の原因を示す。
    write(root / "local.toml", 'schema_version = 1\n[fonts.Example]\npath = "regular.ttf"\n[powerpoint\n')
    [message] = lines(settings_error(root))
    assert message.startswith("local.toml: TOML として読めません（")
    write_deck(root, deck='environment_override = "missing.toml"\n' + DECK)
    with pytest.raises(
        SettingsError, match="environment_override の指すファイルを使えません（入力ファイルが見つかりません"
    ):
        load_project(root)
    write(tmp_path / "outside/local.toml", 'schema_version = 1\n[fonts.Example]\npath = "regular.ttf"\n')
    write_deck(root, deck='environment_override = "../outside/local.toml"\n' + DECK)
    with pytest.raises(SettingsError, match="宣言した参照範囲外です"):
        load_project(root)


# Project の書体と文献


def test_project_fonts_are_shared_and_replaceable(tmp_path, write_deck):
    project = load_project(write_deck(tmp_path / "deck"))
    registry = project.fonts
    assert isinstance(registry, FontRegistry) and project.fonts is registry
    # 生成の流れはこの計測器を最後に閉じる。
    assert project._font_registry is registry
    fake = FakeFonts()
    project.fonts = fake
    assert project.fonts is fake and project._font_registry is fake
    # 互換の計測の関数も差し替えた書体を使う。
    assert project.measure_text("abcd", 20) == fake.measure_text("abcd", 20)
    assert project.measure_math("x+1", 20) == pytest.approx((30.0, 16.0))
    assert project.collect_errors is False


def test_project_bibliography_is_loaded_once(tmp_path, write_deck, monkeypatch):
    root = write_deck(tmp_path / "deck", deck=DECK, extra='[inputs]\nbibliography = ["references.bib"]\n')
    write(root / "references.bib", BIBLIOGRAPHY)
    project = load_project(root)
    bibliography = project.bibliography
    assert list(bibliography) == ["knuth1984"] and bibliography["knuth1984"].fields["year"] == "1984"
    assert project.bibliography is bibliography
    # 文献ファイルがなければ、bibtexparser を読み込まずに空の一覧を返す。
    monkeypatch.setitem(sys.modules, "bibtexparser", None)
    assert len(load_project(write_deck(tmp_path / "other", deck=DECK)).bibliography) == 0


def test_documents_without_content_effect_are_not_tracked(tmp_path, write_deck):
    """affects_content = false の入力文書は、存在と参照範囲を確かめ、生成記録の照合の対象から外す。"""
    root = write_deck(
        tmp_path / "deck",
        deck=DECK,
        extra='[[inputs.documents]]\nid = "readme"\npath = "README.md"\naffects_content = false\n'
        '[[inputs.documents]]\nid = "proof"\npath = "proof.md"\n',
    )
    write(root / "README.md", "参考\n")
    write(root / "proof.md", "証明\n")
    project = load_project(root)
    tracked = set(project.input_files)
    assert str((root / "README.md").resolve()) not in tracked and str((root / "proof.md").resolve()) in tracked
    documents = {document["id"]: document for document in project.config["inputs"]["documents"]}
    assert documents["readme"]["affects_content"] is False and Path(documents["readme"]["path"]).is_file()
    write_deck(root, deck=DECK, extra='[[inputs.documents]]\nid = "readme"\npath = "README.md"\naffects_content = 1\n')
    with pytest.raises(SettingsError, match="affects_content は真偽値で指定してください"):
        load_project(root)
