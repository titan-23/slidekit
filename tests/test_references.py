"""文献一覧（slide.references）、出典の文献キー（deck.cited_keys）、出典の検査を確かめる。

試験の .bib は公開されている一般的な文献だけを試験の中で書く。計測は偽のフォント（1 文字 0.5 em）にする。
"""

import itertools

import pytest

from fake_fonts import FakeFonts
from slidekit import Deck, Rect, Text
from slidekit.bibliography import BibliographyError, format_reference
from slidekit.config import load_project
from slidekit.errors import ArgumentNameError, InputError
from slidekit.validators import validate_scene

BIB = r"""
@book{knuth1997,
  author = {Knuth, Donald E.},
  title = {The Art of Computer Programming, Volume 1: Fundamental Algorithms},
  publisher = {Addison-Wesley}, year = 1997
}
@article{dijkstra1959,
  author = {Dijkstra, Edsger W.},
  title = {A note on two problems in connexion with graphs},
  journal = {Numerische Mathematik}, volume = 1, pages = {269--271}, year = 1959,
  doi = {10.1007/BF01386390}
}
@book{cormen2009,
  author = {Cormen, Thomas H. and Leiserson, Charles E. and Rivest, Ronald L. and Stein, Clifford},
  title = {Introduction to Algorithms},
  publisher = {MIT Press}, year = 2009
}
@misc{web/guide,
  title = {A short guide},
  howpublished = {Online}, year = 2020
}
@misc{web_guide,
  title = {Another guide},
  howpublished = {Online}, year = 2021
}
@misc{broken,
  title = {Uses \unknowncommand{x}},
  year = 2000
}
"""


@pytest.fixture
def make_deck(tmp_path, write_deck):
    """文献ファイルを持つ資料を読み込み、計測を偽のフォントにした Deck を返す関数。"""
    numbers = itertools.count()

    def make(bib=BIB):
        root = tmp_path / f"deck{next(numbers)}"
        root.mkdir()
        (root / "references.bib").write_text(bib, encoding="utf-8")
        project = load_project(write_deck(root, extra='\n[inputs]\nbibliography = ["references.bib"]\n'))
        project.fonts = FakeFonts()
        return Deck(project)

    return make


def by_id(page):
    return {element["id"]: element for element in page["elements"]}


def test_references_lay_out_entries_with_sources(make_deck):
    deck = make_deck()
    slide = deck.slide(id="refs", kind="references", title="参考文献")
    group = slide.references(id="list", keys=["dijkstra1959", "knuth1997", "web/guide"], gap_pt=10, size_pt=16)
    page = deck.to_dict()["slides"][0]
    elements = by_id(page)
    items = [elements[f"list/{name}"] for name in ("dijkstra1959", "knuth1997", "web_guide")]
    # まとまりは既定で PowerPoint のグループにしない。項目は文献キーの順で、出典にその文献キーを付ける。
    assert elements["list"]["pptx_group"] is False
    assert [item["sources"] for item in items] == [
        [{"bib_key": "dijkstra1959"}],
        [{"bib_key": "knuth1997"}],
        [{"bib_key": "web/guide"}],
    ]
    assert all(
        item["parent_id"] == "list" and item["role"] == "references" and item["style"]["size_pt"] == 16
        for item in items
    )
    # 既定の領域はページの内容の領域。各項目は幅 area.w で計測した高さで、gap_pt の間隔で上から積む。
    content = Rect(*page["content_area"])
    fonts = deck.project.fonts
    tops = [content.y]
    for key, item in zip(("dijkstra1959", "knuth1997", "web/guide"), items):
        x, y, w, h = item["area"]
        assert (x, w) == (content.x, content.w) and y == pytest.approx(tops[-1])
        fragments = format_reference(deck.bibliography[key])
        assert h == pytest.approx(slide.measure(fragments, role="references", width=w, size_pt=16).height)
        assert item["layout_bounds"] == list(content)
        tops.append(y + h + 10)
    # 断片は既定の整形（format_reference）：著者、太字の題名、掲載情報、doi のリンク。
    first = items[0]["content"]
    assert first[0] == {"kind": "text", "text": "Edsger W. Dijkstra\n"}
    assert any(fragment.get("link") == "https://doi.org/10.1007/BF01386390" for fragment in first)
    # 戻り値のまとまりの範囲は実際に積んだ範囲。
    assert group.area == Rect(content.x, content.y, content.w, tops[-1] - 10 - content.y)
    assert elements["list"]["area"] == list(group.area)
    assert fonts.calls == []  # 数式はない


def test_references_area_formatter_and_group_proxy(make_deck):
    deck = make_deck()
    slide = deck.slide(id="refs", kind="references", title="参考文献")
    panel = slide.group(id="panel", area=Rect(100, 120, 400, 300))
    group = panel.references(
        id="list",
        keys=["knuth1997", "cormen2009"],
        area=Rect(10, 20, 300, None),
        gap_pt=0,
        formatter=lambda reference: [Text(reference.fields["title"], bold=True)],
        role="body",
    )
    page = deck.to_dict()["slides"][0]
    elements = by_id(page)
    first, second = elements["panel/list/knuth1997"], elements["panel/list/cormen2009"]
    # まとまりの中では、まとまりの左上を原点にする。高さを決めない領域では積み上げの範囲を記録しない。
    assert first["area"][:3] == [110, 140, 300] and "layout_bounds" not in first
    assert second["area"][1] == pytest.approx(first["area"][1] + first["area"][3])
    assert first["content"] == [
        {
            "kind": "text",
            "text": "The Art of Computer Programming, Volume 1: Fundamental Algorithms",
            "style": {"bold": True},
        }
    ]
    assert (first["role"], first["parent_id"], elements["panel/list"]["parent_id"]) == ("body", "panel/list", "panel")
    assert group.area.h == pytest.approx(second["area"][1] + second["area"][3] - 140)


def test_cited_keys_follow_first_appearance(make_deck):
    deck = make_deck()
    intro = deck.slide(id="intro", title="導入")
    intro.text(id="lead", text="最短経路", area=Rect(40, 100, 300, 40), sources=["dijkstra1959"])
    intro.text(id="more", text="教科書", area=Rect(40, 160, 300, 40), sources=[{"bib_key": "cormen2009"}, "knuth1997"])
    second = deck.slide(id="second", title="続き")
    second.text(id="again", text="再び", area=Rect(40, 100, 300, 40), sources=["knuth1997", "dijkstra1959"])
    assert deck.cited_keys() == ["dijkstra1959", "cormen2009", "knuth1997"]
    refs = deck.slide(id="refs", kind="references", title="参考文献")
    refs.references(id="list")
    page = deck.to_dict()["slides"][-1]
    assert [element["id"] for element in page["elements"] if element.get("parent_id") == "list"] == [
        "list/dijkstra1959",
        "list/cormen2009",
        "list/knuth1997",
    ]
    # 文献一覧の項目の出典は順を変えない。
    assert deck.cited_keys() == ["dijkstra1959", "cormen2009", "knuth1997"]


def test_reference_errors_leave_no_elements(make_deck):
    deck = make_deck()
    slide = deck.slide(id="refs", kind="references", title="参考文献")

    def rejected(match, code="invalid_reference", **arguments):
        with pytest.raises(InputError, match=match) as caught:
            slide.references(id="list", **arguments)
        assert caught.value.code == code and caught.value.slide_id == "refs"
        assert not slide.elements
        return caught.value

    rejected("文献一覧に載せる文献がありません")  # 出典がまだない
    rejected("keys に文献キーを 1 つ以上", keys=[])
    rejected("list か tuple", keys="knuth1997")
    rejected("重複しています: knuth1997", keys=["knuth1997", "knuth1997"])
    missing = rejected("未定義の文献キーです: knuth1979（候補: knuth1997）", keys=["knuth1979"])
    assert isinstance(missing, KeyError) and missing.element_id == "list"
    rejected("置き換えた項目の識別子が重複しています: web_guide", keys=["web/guide", "web_guide"])
    rejected("formatter は", code="input_error", keys=["knuth1997"], formatter="plain")
    rejected("gap_pt は 0 以上", code="invalid_geometry", keys=["knuth1997"], gap_pt=-1)
    rejected("幅を数値で", code="invalid_geometry", keys=["knuth1997"], area=Rect(0, 0, None, 100))
    reserved = rejected(
        "sources、text は references では使えません",
        code="unsupported_argument",
        keys=["knuth1997"],
        text="本文",
        sources=["knuth1997"],
    )
    assert isinstance(reserved, ArgumentNameError) and reserved.details["names"] == ["sources", "text"]
    # 対応しない LaTeX 命令は、項目を作るときに項目の識別子を付けて誤りにする（作りかけの項目も残さない）。
    broken = rejected("unknowncommand", keys=["knuth1997", "broken"])
    assert isinstance(broken, BibliographyError) and broken.element_id == "list/broken"


def test_sources_use_project_bibliography(make_deck):
    deck = make_deck()
    slide = deck.slide(id="intro")
    slide.text(id="known", text="既知", area=Rect(40, 100, 300, 40), sources=["knuth1997"])
    for source in ("knuth1979", {"bib_key": ["knuth1997"]}):
        with pytest.raises(InputError, match="未定義の文献キー") as caught:
            slide.text(id="unknown", text="未知", area=Rect(40, 160, 300, 40), sources=[source])
        assert caught.value.code == "source_missing"
    # 出典の判定は文献項目の LaTeX 表記を変換しない（変換できない項目も出典には使うことができる）。
    slide.text(id="raw", text="生", area=Rect(40, 220, 300, 40), sources=["broken"])
    assert deck.bibliography is deck.project.bibliography
    # 誤りを集めるときは、未定義の出典を除いて続ける。
    deck.project.collect_errors = True
    element = slide.text(id="collected", text="集める", area=Rect(40, 280, 300, 40), sources=["knuth1979", "knuth1997"])
    assert element.data["sources"] == [{"bib_key": "knuth1997"}]
    assert [issue.code for issue in deck.issues] == ["source_missing"]


def test_references_beyond_area_are_reported(make_deck):
    deck = make_deck()
    slide = deck.slide(id="refs", kind="references", title="参考文献")
    slide.references(id="list", keys=["knuth1997", "dijkstra1959", "cormen2009"], area=Rect(40, 100, 200, 60))
    scene = deck.to_dict()
    issues = [
        issue
        for issue in validate_scene(scene, deck.project, deck.project.fonts)
        if issue["code"] == "layout_bounds_overflow"
    ]
    assert issues and {issue["element_id"] for issue in issues} <= {
        "list/knuth1997",
        "list/dijkstra1959",
        "list/cormen2009",
    }
