"""配置、入力保護、識別子、役割と出力直前の再計算を確かめる。"""

import json
import shutil

import pytest

from slidekit import pipeline
from slidekit.config import load_project
from slidekit.files import file_hash
from slidekit.fonts import FontRegistry
from slidekit.layout import Rect, align, distribute
from slidekit.model import Deck, Emphasis, Math, Text

BASE = """schema_version = 1
entrypoint = "slides.py:build"
size = "16:9"
[metadata]
title = "比較"
author = ""
last_modified_by = ""
"""


@pytest.fixture
def project_dir(tmp_path, write_deck):
    # 実行環境は conftest.py の空の設定を使う。
    return write_deck(tmp_path / "deck", deck=BASE)


def append(root, text):
    with (root / "deck.toml").open("a", encoding="utf-8") as stream:
        stream.write(text)


def test_inputs_are_resolved_from_project_and_symlinks_checked(project_dir, monkeypatch):
    (project_dir / "brief.md").write_text("発表の目的", encoding="utf-8")
    (project_dir.parent / "secret.txt").write_text("対象外", encoding="utf-8")
    (project_dir / "outside.txt").symlink_to(project_dir.parent / "secret.txt")
    monkeypatch.chdir(project_dir.parent)
    project = load_project(project_dir)
    assert project.size == (960, 540)
    assert str(project_dir / "brief.md") in project.input_files
    assert project.resolve_input("slides.py") == project_dir / "slides.py"
    with pytest.raises(ValueError, match="参照範囲外"):
        project.resolve_input("outside.txt")
    with pytest.raises(ValueError, match="URL"):
        project.resolve_input("https://example.invalid/image.png")


@pytest.mark.parametrize("value", ['title = ""\nauthor = ""\nlast_modified_by = ""', 'title = "比較"\nauthor = ""'])
def test_metadata_must_be_explicit(project_dir, value):
    (project_dir / "deck.toml").write_text("[metadata]\n" + value, encoding="utf-8")
    with pytest.raises(ValueError, match="metadata"):
        load_project(project_dir)


@pytest.mark.parametrize(
    "extra",
    [
        '\n[deliverables]\npptx = "slides.py"\n',
        '\n[deliverables]\npptx = "out/deck.pptx"\n',
        '\n[deliverables]\npptx = "same"\npdf = "same"\n',
    ],
)
def test_outputs_cannot_overwrite_inputs_or_share_paths(project_dir, extra):
    append(project_dir, extra)
    with pytest.raises(ValueError):
        load_project(project_dir)


def test_asset_directory_and_output_must_not_overlap(project_dir):
    (project_dir / "assets").mkdir()
    append(project_dir, '\n[inputs]\nasset_dirs = ["."]\n')
    with pytest.raises(ValueError, match="重複"):
        load_project(project_dir)


def test_environment_paths_relative_to_their_settings(project_dir, monkeypatch):
    shared = project_dir.parent / "environment.toml"
    shared.write_text(
        '[fonts.Example]\npath = "regular.ttf"\nindex = 1\n[renderers.powerpoint]\ntimeout = 60\n', encoding="utf-8"
    )
    monkeypatch.setenv("SLIDEKIT_ENVIRONMENT", str(shared))
    (project_dir / "local").mkdir()
    (project_dir / "local/env.toml").write_text(
        '[fonts.Example]\nbold_path = "bold.ttf"\nbold_index = 0\n', encoding="utf-8"
    )
    (project_dir / "deck.toml").write_text('environment_override = "local/env.toml"\n' + BASE, encoding="utf-8")
    project = load_project(project_dir)
    font = project.environment["fonts"]["Example"]
    assert font["path"] == str(project_dir.parent / "regular.ttf")
    assert font["bold_path"] == str(project_dir / "local/bold.ttf")
    assert project.environment["renderers"]["powerpoint"]["timeout"] == 60


def test_environment_and_bundled_resources_are_tracked_outside_read_roots(project_dir, monkeypatch, dejavu):
    shared = project_dir.parent / "environment.toml"
    font = (project_dir.parent / "font.ttf").resolve()
    bold = (project_dir.parent / "bold.ttf").resolve()
    shutil.copyfile(dejavu.path, font)
    shutil.copyfile(dejavu.bold_path, bold)
    shared.write_text('[fonts."DejaVu Sans"]\npath = "font.ttf"\nbold_path = "bold.ttf"\n', encoding="utf-8")
    monkeypatch.setenv("SLIDEKIT_ENVIRONMENT", str(shared))
    (project_dir / "theme.toml").write_text(dejavu.theme, encoding="utf-8")
    (project_dir / "deck.toml").write_text('theme = "theme.toml"\n' + BASE, encoding="utf-8")
    project = load_project(project_dir)
    assert str(shared) in project.input_files
    assert any(path.endswith("/themes/default.toml") for path in project.input_files)
    assert any(path.endswith("/rules/general.toml") for path in project.input_files)
    # 読み込みでは書体のハッシュ値を求めない。
    assert str(font) not in project.input_files and str(bold) not in project.input_files

    def recorded():
        # 計測に使った書体を、生成の記録（実行環境の記録）が追跡する。
        project = load_project(project_dir)
        with FontRegistry(project) as registry:
            registry.measure_text("Measured", 20)
            pipeline._environment(project, registry)
        return project.input_files

    first = recorded()
    assert first[str(font)] == file_hash(font) and first[str(bold)] == file_hash(bold)
    # 同じ書体名の別の実体に書き換えると、記録の値が変わる。
    shutil.copyfile(dejavu.bold_path, font)
    assert recorded()[str(font)] != first[str(font)]


@pytest.mark.parametrize(
    "extra",
    ['\n[render]\nrequird=["powerpoint"]\n', '\n[math]\nfallbak="required"\n', '\n[inputs]\npyton=["slides.py"]\n'],
)
def test_unknown_settings_are_not_silently_ignored(project_dir, extra):
    append(project_dir, extra)
    with pytest.raises(ValueError, match="不明な設定"):
        load_project(project_dir)


def test_custom_role_is_restricted_without_changing_body(project_dir):
    (project_dir / "theme.toml").write_text(
        """[roles.result_formula]
inherits = "body"
size_pt = 18
min_size_pt = 16
allowed_slide_ids = ["results"]
allowed_content = "math"
""",
        encoding="utf-8",
    )
    (project_dir / "deck.toml").write_text('theme = "theme.toml"\n' + BASE, encoding="utf-8")
    project = load_project(project_dir)
    deck = Deck(project)
    slide = deck.slide(id="results")
    formula = slide.equation(
        id="formula", expression="x^2", area=Rect(20, 100, 200, 50), role="result_formula", size_pt=16
    )
    assert formula.data["style"]["size_pt"] == 16
    assert project.theme["roles"]["body"]["min_size_pt"] == 18
    with pytest.raises(ValueError, match="数式だけ"):
        slide.text(id="bad", text="説明", area=slide.content, role="result_formula")
    with pytest.raises(ValueError, match="最小文字サイズ"):
        slide.text(id="small", text="説明", area=slide.content, min_size_pt=1)
    other = deck.slide(id="other")
    with pytest.raises(ValueError, match="使えません"):
        other.equation(id="formula", expression="x", area=other.content, role="result_formula")


def test_theme_inheritance_cycle_is_rejected(project_dir):
    (project_dir / "theme.toml").write_text('[roles.a]\ninherits = "b"\n[roles.b]\ninherits = "a"\n', encoding="utf-8")
    (project_dir / "deck.toml").write_text('theme = "theme.toml"\n' + BASE, encoding="utf-8")
    with pytest.raises(ValueError, match="継承"):
        load_project(project_dir)


def test_stable_ids_text_fragments_and_no_automatic_changes(project_dir):
    project = load_project(project_dir)
    deck = Deck(project)
    slide = deck.slide(id="comparison", title="方式の比較")
    slide.text(id="body", area=Rect(30, 80, 10, 10), text=[Text("A $x$\n"), Emphasis("重要"), Math("x^2")], size_pt=28)
    slide.line(id="separator", start=(20, 50), end=(200, 50))
    scene = deck.to_dict()
    body = next(e for e in scene["slides"][0]["elements"] if e["id"] == "body")
    assert body["style"]["size_pt"] == 28
    assert body["content"] == [
        {"kind": "text", "text": "A $x$\n"},
        {"kind": "emphasis", "text": "重要"},
        {"kind": "math", "text": "x^2"},
    ]
    assert json.loads(json.dumps(scene))["slides"][0]["id"] == "comparison"
    with pytest.raises(ValueError, match="重複"):
        slide.line(id="separator", start=(20, 60), end=(200, 60))
    with pytest.raises(ValueError):
        slide.box(id="box/child", area=Rect(20, 30, 100, 30))


def test_group_is_flattened_and_connections_follow_movement(project_dir):
    slide = Deck(load_project(project_dir)).slide(id="flow")
    group = slide.group(id="flow", area=Rect(40, 50, 200, 100))
    first = group.box(id="first", text="始点", area=Rect(0, 0, 50, 30))
    second = group.box(id="second", text="終点", area=Rect(100, 0, 50, 30))
    group.arrow(id="connection", start=first.anchor("right"), end=second.anchor("left"))
    group.line(id="rule", start=(0, 50), end=(100, 50))
    group.move(20, 10)
    scene = slide.to_dict(1)
    by_id = {e["id"]: e for e in scene["elements"]}
    assert by_id["flow/first"]["area"] == [60, 60, 50, 30]
    assert by_id["flow/connection"]["start"] == [110, 75]
    assert by_id["flow/connection"]["end_ref"] == {"element_id": "flow/second", "side": "left"}
    assert by_id["flow/rule"]["start"] == [60, 110]
    assert by_id["flow/first"]["parent_id"] == "flow"
    # 子の座標はページ上の座標に展開したまま、数式を含まず子が続くまとまりは PowerPoint のグループとして書き出す。
    assert by_id["flow"]["pptx_group"] is True and by_id["flow"]["area"] == [60, 60, 200, 100]


def test_math_rotation_in_ancestor_is_rejected(project_dir):
    # まとまりの回転・反転は作成時に拒否するので、数式の祖先が回転することはない。
    slide = Deck(load_project(project_dir)).slide(id="flow")
    for transform in ({"rotation": 20}, {"flip_h": True}, {"flip_v": True}):
        with pytest.raises(ValueError, match="回転・反転") as caught:
            slide.group(id="flow", area=Rect(50, 50, 200, 100), **transform)
        assert caught.value.code == "unsupported_argument" and caught.value.element_id == "flow"
    assert not slide.elements
    group = slide.group(id="flow", area=Rect(50, 50, 200, 100), rotation=0, flip_h=False)
    group.equation(id="label", expression="x", area=Rect(0, 0, 100, 30))
    with pytest.raises(ValueError, match="回転・反転") as caught:
        group.equation(id="turned", expression="y", area=Rect(0, 40, 100, 30), rotation=90)
    assert caught.value.code == "math_transform"
    assert not {"rotation", "flip_h"} & set(group.data)


def test_context_measurements_are_positive_and_scale_with_size(project_dir, dejavu, dejavu_environment):
    # 計測の方式は改修で変わるので、値ではなく性質で確かめる。書体は conftest の DejaVu Sans の設定を使う。
    (project_dir / "theme.toml").write_text(dejavu.theme, encoding="utf-8")
    (project_dir / "deck.toml").write_text('theme = "theme.toml"\n' + BASE, encoding="utf-8")
    context = load_project(project_dir)
    width, height = context.measure_text("Measured text", 20)
    assert width > 0 and height > 0
    larger = context.measure_text("Measured text", 40)
    assert larger[0] == pytest.approx(width * 2, rel=0.03)
    assert larger[1] == pytest.approx(height * 2, rel=0.03)
    assert context.measure_text("Measured text again", 20)[0] > width
    lines = context.measure_text("Measured text\nMeasured text", 20)
    assert lines[0] == pytest.approx(width) and lines[1] > height
    assert min(context.measure_text("Measured text", 20, bold=True)) > 0
    math_width, math_height = context.measure_math("x_i^2 + 1", 20)
    assert math_width > 0 and math_height > 0
    larger_math = context.measure_math("x_i^2 + 1", 40)
    assert larger_math[0] == pytest.approx(math_width * 2, rel=0.05)
    assert larger_math[1] == pytest.approx(math_height * 2, rel=0.05)


def test_page_kinds_and_aspect_ratios(project_dir):
    project = load_project(project_dir)
    project.config["size"] = [720, 540]
    deck = Deck(project)
    for kind in ("cover", "section", "content", "references", "blank"):
        deck.slide(id=kind, kind=kind, title=kind)
    for page in deck.to_dict()["slides"]:
        ids = {e["id"] for e in page["elements"]}
        assert ("_theme/footer" in ids) == (page["kind"] not in ("cover", "blank"))
        assert ("_theme/title" in ids) == (page["kind"] != "blank")
        assert page["content_area"][2] > 0


def test_page_number_format_shows_total(project_dir):
    """page_number_format の {page} と {total} を、ページ番号と全体のページ数に置き換える。"""
    project = load_project(project_dir)
    project.theme["page_kinds"]["content"]["page_number_format"] = "{page} / {total}"
    deck = Deck(project)
    for index in range(3):
        deck.slide(id=f"page{index}", kind="content", title="題名")
    labels = []
    for page in deck.to_dict()["slides"]:
        footer = next(e for e in page["elements"] if e["id"] == "_theme/footer")
        labels.append("".join(fragment["text"] for fragment in footer["content"]))
    assert labels == ["1 / 3", "2 / 3", "3 / 3"]


def test_template_uses_only_blank_layout_and_is_read_only(project_dir):
    from pptx import Presentation

    path = project_dir / "template.pptx"
    original = Presentation()
    original.slides.add_slide(original.slide_layouts[6])
    original.save(path)
    before = path.read_bytes()
    append(project_dir, '\n[template]\npath = "template.pptx"\nlayout_index = 6\n')
    project = load_project(project_dir)
    assert project.config["template"]["inspection"]["original_slide_count"] == 1
    assert path.read_bytes() == before
    (project_dir / "deck.toml").write_text(
        BASE + '\n[template]\npath = "template.pptx"\nlayout_index = 0\n', encoding="utf-8"
    )
    with pytest.raises(ValueError, match="空白"):
        load_project(project_dir)


def test_math_fallback_rejected_only_when_math_is_present(project_dir):
    append(project_dir, '\n[render]\nprimary="libreoffice"\nrequired=["libreoffice"]\n[math]\nfallback="off"\n')
    deck = Deck(load_project(project_dir))
    slide = deck.slide(id="content")
    slide.text(id="text", text="文章", area=slide.content)
    deck.to_dict()
    slide.equation(id="math", expression="x", area=slide.content)
    with pytest.raises(ValueError, match="互換画像"):
        deck.to_dict()


def test_rect_split_align_distribute_and_fit():
    area = Rect(20, 40, 400, 200)
    first, second = area.columns([1, 2], gap=10)
    assert first.w == 130 and second.x == 160 and second.w == 260
    assert area.inset(10) == Rect(30, 50, 380, 180)
    assert area.contain(100, 100) == Rect(120, 40, 200, 200)
    values = [Rect(0, 0, 30, 40), Rect(0, 0, 50, 20)]
    assert align(values, area, horizontal="right")[0].right == 420
    assert distribute(values, area)[1].right == 420
    with pytest.raises(ValueError):
        area.columns(2, gap=500)
    with pytest.raises(ValueError):
        Rect(0, 0, float("nan"), 1)


def test_rect_rows_split_height_in_order():
    area = Rect(20, 40, 400, 200)
    top, bottom = area.rows([1, 3], gap=10)
    assert top == Rect(20, 40, 400, 47.5)
    assert bottom == Rect(20, 97.5, 400, 142.5)
    assert [row.y for row in area.rows(4, gap=0)] == [40, 90, 140, 190]
    with pytest.raises(ValueError):
        area.rows(3, gap=150)
