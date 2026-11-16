"""まとまりの書き出しの判定（pptx_group）と、ページ種別の飾り・背景を確かめる。

PowerPoint のグループにするのは、描く子孫があり、子孫がまとまりの直後に続き、子孫に数式がないまとまりだけである。
"""

import itertools

import pytest

from fake_fonts import FakeFonts
from slidekit import Deck, Math, Rect
from slidekit.config import load_project
from slidekit.errors import InputError
from slidekit.pptx_validation import validate_pptx
from slidekit.schema import SettingsError
from slidekit.validators import validate_scene
from slidekit.writer import NS, write_pptx


@pytest.fixture
def make_deck(tmp_path, write_deck):
    """資料を書いて読み込み、計測を偽のフォントにした Deck を返す関数。"""
    numbers = itertools.count()

    def make(theme=None, extra=""):
        project = load_project(write_deck(tmp_path / f"deck{next(numbers)}", theme=theme, extra=extra))
        project.fonts = FakeFonts()
        return Deck(project)

    return make


def by_id(page):
    return {element["id"]: element for element in page["elements"]}


def flags(page):
    """まとまりの識別子から場面の pptx_group（なければ None）。"""
    return {element["id"]: element.get("pptx_group") for element in page["elements"] if element["kind"] == "group"}


def test_pptx_group_requires_contiguous_children(make_deck, tmp_path):
    deck = make_deck()
    slide = deck.slide(id="diagram", kind="blank")
    flow = slide.group(id="flow", area=Rect(40, 40, 300, 100), description="処理の流れ")
    first = flow.box(id="first", text="始点", area=Rect(0, 0, 80, 40))
    second = flow.box(id="second", text="終点", area=Rect(200, 0, 80, 40))
    flow.arrow(id="link", start=first.anchor("right"), end=second.anchor("left"))
    # 子の間にまとまりの外の要素がある（グループにすると重なり順が変わる）。
    split = slide.group(id="split", area=Rect(40, 200, 300, 60))
    split.box(id="a", text="前", area=Rect(0, 0, 80, 40))
    slide.text(id="between", text="外", area=Rect(400, 200, 60, 30))
    split.box(id="b", text="後", area=Rect(100, 0, 80, 40))
    # 描く子孫のないまとまり、子がまとまりだけのまとまり。
    slide.group(id="empty", area=Rect(400, 300, 10, 10))
    hollow = slide.group(id="hollow", area=Rect(400, 320, 10, 10))
    hollow.group(id="inner", area=Rect(0, 0, 5, 5))
    # 利用者が無効にしたまとまりと、入れ子のまとまり（外側を無効にしても内側は単独で判定する）。
    flat = slide.group(id="flat", area=Rect(40, 300, 200, 60), pptx_group=False)
    flat.box(id="c", text="展開", area=Rect(0, 0, 80, 40))
    outer = slide.group(id="outer", area=Rect(40, 380, 300, 100), pptx_group=True)
    outer.box(id="d", text="外", area=Rect(0, 0, 80, 40))
    nested = outer.group(id="nested", area=Rect(100, 0, 200, 60))
    nested.box(id="e", text="内", area=Rect(0, 0, 80, 40))
    loose = slide.group(id="loose", area=Rect(400, 380, 200, 100), pptx_group=False)
    tight = loose.group(id="tight", area=Rect(0, 0, 100, 50))
    tight.box(id="f", text="内", area=Rect(0, 0, 80, 40))
    scene = deck.to_dict()
    page = scene["slides"][0]
    assert flags(page) == {
        "flow": True,
        "split": None,
        "empty": None,
        "hollow": None,
        "hollow/inner": None,
        "flat": False,
        "outer": True,
        "outer/nested": True,
        "loose": False,
        "loose/tight": True,
    }
    # 書き出しと PPTX の検査は場面の pptx_group に従う。
    project = deck.project
    path = tmp_path / "deck.pptx"
    result = write_pptx(scene, project, path, fonts=project.fonts)
    assert result["groups"] == 4
    assert not [
        issue
        for issue in validate_pptx(path, scene, project, fonts=project.fonts)
        if issue["code"] in {"group_mismatch", "element_missing", "connection_mismatch"}
    ]
    with pytest.raises(InputError, match="pptx_group は真偽値か None") as caught:
        slide.group(id="wrong", pptx_group="yes")
    assert caught.value.code == "input_error" and caught.value.element_id == "wrong"


def test_math_group_is_flat(make_deck):
    # 数式を含むまとまりはグループにしない。数式を含まない内側のまとまりは、外側によらずグループにする。
    deck = make_deck()
    slide = deck.slide(id="formula", kind="blank")
    outer = slide.group(id="outer", area=Rect(40, 40, 400, 200))
    outer.equation(id="sum", expression="a+b", area=Rect(0, 0, 100, 40))
    plain = outer.group(id="plain", area=Rect(0, 60, 200, 60))
    plain.box(id="note", text="説明", area=Rect(0, 0, 80, 40))
    inline = slide.group(id="inline", area=Rect(40, 300, 300, 60))
    inline.text(id="label", text=["長さ ", Math("n")], area=Rect(0, 0, 200, 40))
    assert flags(deck.to_dict()["slides"][0]) == {"outer": None, "outer/plain": True, "inline": None}


def test_table_group_follows_pptx_group(make_deck):
    deck = make_deck()
    slide = deck.slide(id="tables", kind="blank")
    slide.table(id="plain", rows=[["1", "2"]], row_keys=["first"], column_keys=["a", "b"], area=Rect(40, 40, 200, 40))
    slide.table(
        id="flat",
        rows=[["1", "2"]],
        row_keys=["first"],
        column_keys=["a", "b"],
        area=Rect(40, 100, 200, 40),
        pptx_group=False,
    )
    formula = slide.table(
        id="formula",
        rows=[[Math("x"), "2"]],
        row_keys=["first"],
        column_keys=["a", "b"],
        area=Rect(40, 160, 200, 40),
        pptx_group=True,
    )
    assert flags(deck.to_dict()["slides"][0]) == {"plain": True, "flat": False, "formula": None}
    # 表には要素を追加できない（文献一覧も同じ）。
    with pytest.raises(AttributeError, match="表に要素は追加できません"):
        formula.references(id="list")


THEME = """[roles.body]
size_pt = 20
min_size_pt = 18
line_width_pt = 5

[defaults.line]
dashed = true

[[page_kinds.section.decorations]]
kind = "line"
id = "rule"
start = [0.1, 0.5]
end = [0.9, 0.5]
line_color = "primary"

[[page_kinds.section.decorations]]
kind = "box"
id = "mark"
shape = "diamond"
area_pt = [100, 356, 10, 10]
fill = "accent"

[[page_kinds.section.decorations]]
kind = "arrow"
id = "pointer"
start_pt = [500, 300]
end_pt = [450, 320]
line_width_pt = 2
"""


def test_theme_decorations_are_added_per_page_kind(make_deck):
    deck = make_deck(theme=THEME)
    part = deck.slide(id="part", kind="section", title="第 1 部")
    # 飾りに重ねた文字は、重なりの警告の対象にしない（背景と同じ）。
    part.text(id="lead", text="説明", area=Rect(96, 340, 300, 40))
    deck.slide(id="body", kind="content", title="本文")
    scene = deck.to_dict()
    section, content = scene["slides"]
    elements = by_id(section)
    rule, mark, pointer = (elements[f"_theme/decoration/{name}"] for name in ("rule", "mark", "pointer"))
    # 描く順は背景、飾り（設定の順）、題名、ページ番号。
    assert [element["id"] for element in section["elements"]] == [
        "_theme/background",
        "_theme/decoration/rule",
        "_theme/decoration/mark",
        "_theme/decoration/pointer",
        "_theme/title",
        "_theme/footer",
        "lead",
    ]
    assert {element["semantic"] for element in (rule, mark, pointer)} == {"decoration"}
    assert all(element["content"] == [] and element["role"] == "body" for element in (rule, mark, pointer))
    # 位置：比率はページの寸法（960×540）に掛け、pt はそのまま使う。
    assert (rule["start"], rule["end"], rule["area"]) == ([96.0, 270.0], [864.0, 270.0], [96.0, 270.0, 768.0, 0.0])
    assert (mark["area"], mark["shape"]) == ([100, 356, 10, 10], "diamond")
    assert (pointer["start"], pointer["end"], pointer["area"]) == (
        [500.0, 300.0],
        [450.0, 320.0],
        [450.0, 300.0, 50.0, 20.0],
    )
    # 書式：種類の既定（[defaults.line] の dashed、[defaults.box] の塗りと枠線）に飾りの指定を重ね、役割の書式は重ねない。
    colors = deck.project.theme["colors"]
    assert (rule["style"]["line_color"], rule["style"]["dashed"], rule["style"]["line_width_pt"]) == (
        colors["primary"],
        True,
        1.0,
    )
    assert (pointer["style"]["dashed"], pointer["style"]["line_width_pt"]) == (True, 2.0)
    assert (mark["style"]["fill"], mark["style"]["line_color"]) == (colors["accent"], colors["line"])
    assert "size_pt" not in mark["style"]
    # 飾りのないページ種別には置かない。
    assert not [element for element in content["elements"] if element.get("semantic") == "decoration"]
    assert validate_scene(scene, deck.project, deck.project.fonts) == []


def test_slide_can_omit_decorations(make_deck):
    deck = make_deck(theme=THEME)
    deck.slide(id="plain", kind="section", title="第 2 部", decorations=False)
    assert not [
        element for element in deck.to_dict()["slides"][0]["elements"] if element.get("semantic") == "decoration"
    ]
    with pytest.raises(InputError, match="decorations は真偽値") as caught:
        deck.slide(id="wrong", kind="section", decorations="no")
    assert caught.value.code == "input_error" and caught.value.slide_id == "wrong"


def test_invalid_decorations_are_rejected_at_load(tmp_path, write_deck):
    theme = """[[page_kinds.section.decorations]]
kind = "box"
id = "mark"
area = [0.1, 0.1, 0.1, 0.1]
area_pt = [10, 10, 10, 10]
text = "飾り"

[[page_kinds.section.decorations]]
kind = "line"
id = "mark"
start = [0.1, 0.5]
"""
    with pytest.raises(SettingsError) as caught:
        load_project(write_deck(tmp_path / "deck", theme=theme))
    message = str(caught.value)
    for expected in (
        "decorations の 1 番目の text は使えません",
        "1 番目には area か area_pt のどちらか一方",
        "2 番目の id mark は 1 番目と重複",
        "2 番目には start と end の両方",
    ):
        assert expected in message, expected


def test_background_none_omits_background_shape(make_deck, tmp_path):
    deck = make_deck(theme='[page_kinds.content]\nbackground = "none"\n[page_kinds.section]\nbackground = "accent"\n')
    deck.slide(id="body", kind="content", title="本文")
    deck.slide(id="part", kind="section", title="第 1 部")
    content, section = deck.to_dict()["slides"]
    assert "_theme/background" not in by_id(content)
    assert by_id(section)["_theme/background"]["style"]["fill"] == deck.project.theme["colors"]["accent"]
    path = tmp_path / "deck.pptx"
    scene = deck.to_dict()
    write_pptx(scene, deck.project, path)
    from zipfile import ZipFile
    from lxml import etree

    with ZipFile(path) as archive:
        names = [
            node.get("name")
            for node in etree.fromstring(archive.read("ppt/slides/slide1.xml")).iter(f"{{{NS['p']}}}cNvPr")
        ]
    assert "_theme/background" not in names and "_theme/title" in names
