"""同梱の使用例（examples/）を確かめる。

- binary_search が使用例の README の表の機能を使っていること（構築だけ。計測は偽のフォント）。
- 両方の使用例が配置と記述規則の検査（`pipeline.layout()`）で、宣言した許容のほかに診断を出さないこと
  （目印 examples。Windows の既定フォントを使い、なければ省略する）。
- 雛形と図のスクリプトが同梱のファイルと同じ内容を作ること。
- 使用例の README の相対リンクの先が、見出しを含めて実在すること（文書の試験と同じ確かめ方）。
- 使用例が自己完結し、作成者が `slidekit example` であること。
- slidekit の全体（.gitignore で除くものを除く）に、手元のパス（と、指定したときは非公開の語）がないこと。

使用例のスクリプトは、使用例のディレクトリに `__pycache__` を作らないように、ソースを直接実行して読み込む。
"""

import fnmatch
import json
import os
from pathlib import Path
import re
import tomllib
import types
import zipfile

import pypdf
import pytest

from fake_fonts import FakeFonts
from pdf_tools import Document, metadata as pdf_metadata
from slidekit import pipeline
from slidekit.assets import _check_svg
from slidekit.config import load_project
from slidekit.fallback import fallback_required
from slidekit.pdf import open_pdf

ROOT = Path(__file__).resolve().parents[1]
EXAMPLES = ROOT / "examples"
BINARY_SEARCH = EXAMPLES / "binary_search"
# 使用例は推奨の構成で、編集元と作業用のファイルは src/ に、完成物と README は直下にある。
BINARY_SEARCH_SOURCE = BINARY_SEARCH / "src"
EXAMPLE_NAMES = ("proposal", "binary_search")
AUTHOR = "slidekit example"

# binary_search のページの順（README の表と同じ）。
PAGES = ("cover", "idea", "branching", "trace", "invariant", "cost", "recurrence", "growth", "comparison", "references")
# binary_search の配置の検査で、資料のコードが許容を宣言した警告（ページ、種類、要素の組）。
DECLARED = {
    ("growth", "element_overlap", frozenset({"linear_value", "chart"})),
    ("growth", "element_overlap", frozenset({"binary_value", "chart"})),
}

# ソース・試験（試験用のデータを含む）・道具のディレクトリ。試験は C:\Users\example のような例の値を使うので、
# Windows の利用者のフォルダーの形は、これらの外（使用例・文書・設定・完成物）だけで調べる。
CODE_DIRECTORIES = {"src", "tests", "tools"}

# 手元のパス。この試験のファイル自身に一致しないように、文字の集合で書く。<…>・%…%・{…}・$… の置き場所は除く。
POSIX_HOME = re.compile(r"(?:/hom[e]/|/User[s]/)(?![<%{$])")
WINDOWS_HOME = re.compile(r"(?i:[a-z]:\\+user[s]\\+)(?![<%{$])")


def load_script(path: Path) -> types.ModuleType:
    """スクリプトを、バイトコードを書かずに読み込む（使用例のディレクトリに __pycache__ を作らない）。"""
    module = types.ModuleType(f"_example_{path.stem}")
    module.__file__ = str(path)
    exec(compile(path.read_text(encoding="utf-8"), str(path), "exec"), module.__dict__)
    return module


def read_toml(path: Path) -> dict:
    with path.open("rb") as stream:
        return tomllib.load(stream)


def construct_example(name: str, *, selection=None) -> tuple:
    """使用例を偽のフォントで構築し、(資料, 場面) を返す。"""
    project = load_project(EXAMPLES / name / "src" / "deck.toml")
    project.fonts = FakeFonts()
    project.selection = selection
    return project, pipeline.construct(project)


def elements_of(scene: dict, slide_id: str) -> dict:
    (slide,) = [slide for slide in scene["slides"] if slide["id"] == slide_id]
    return {element["id"]: element for element in slide["elements"]}


def has_math(element: dict) -> bool:
    return any(fragment["kind"] == "math" for fragment in element.get("content", ()))


def texts(element: dict) -> str:
    return "".join(fragment["text"] for fragment in element.get("content", ()))


def descendants(elements: dict, group_id: str) -> list[dict]:
    return [element for element in elements.values() if element["id"].startswith(group_id + "/")]


# ---- 機能 ----


def test_binary_search_example_covers_features():
    pytest.importorskip("bibtexparser")
    config = read_toml(BINARY_SEARCH_SOURCE / "deck.toml")
    assert config["size"] == "4:3" and config["template"]["layout_index"] == 6
    assert config["template"]["show_master_shapes"] is False and config["math"]["fallback"] == "required"
    assert "python" not in config["inputs"]
    assert [document["id"] for document in config["inputs"]["documents"]] == ["figure_script"]
    project, scene = construct_example("binary_search")
    assert "issues" not in scene
    slides = {slide["id"]: slide for slide in scene["slides"]}
    assert tuple(slides) == PAGES
    assert sum(slide["seconds"] for slide in scene["slides"]) == 360
    assert [slide["kind"] for slide in scene["slides"]].count("section") == 2
    assert (slides["cover"]["kind"], slides["references"]["kind"]) == ("cover", "references")
    assert fallback_required(scene, project)
    colors = {name: value.upper() for name, value in scene["theme"]["colors"].items()}

    # 1 表紙：題名の下の線と、折り返さない副題。
    cover = elements_of(scene, "cover")
    assert cover["title_rule"]["kind"] == "line"
    assert cover["subtitle"]["style"]["word_wrap"] is False and texts(cover["author"]) == AUTHOR

    # 2・6 章扉：ページ種別の飾り（ひし形 2 個と線）。
    for slide_id in ("idea", "cost"):
        decorations = {
            key: element
            for key, element in elements_of(scene, slide_id).items()
            if element.get("semantic") == "decoration"
        }
        assert sorted(decorations) == [
            "_theme/decoration/rule",
            "_theme/decoration/rule_end",
            "_theme/decoration/rule_start",
        ]
        assert [decorations[f"_theme/decoration/{name}"]["shape"] for name in ("rule_start", "rule_end")] == [
            "diamond"
        ] * 2

    # 3 分岐：記法で書いた文、内容から決める高さ、数式を含まないまとまり（PowerPoint のグループ）と間隔付きの矢印。
    branching = elements_of(scene, "branching")
    assert has_math(branching["premise"]) and branching["premise"]["sources"] == [{"bib_key": "knuth1998"}]
    assert branching["decision"]["pptx_group"] is True
    parts = descendants(branching, "decision")
    assert not any(has_math(element) for element in parts)
    assert any(element.get("shape") == "diamond" for element in parts)
    arrows = [element for element in parts if element["kind"] == "arrow"]
    assert len(arrows) == 2 and all(arrow["end_ref"]["gap"] == 6 for arrow in arrows)
    assert {"not_less", "less", "left_update", "right_update"} <= set(branching)
    assert all(has_math(branching[name]) for name in ("not_less", "less", "left_update", "right_update"))

    # 4 推移：見出しも数式の表（見出しの書式と交互の塗り）、key_id で作る状態図、半開区間の数式（数字の間は \, で区切る）。
    trace = elements_of(scene, "trace")
    headers = [element for key, element in trace.items() if key.startswith("values/_header/")]
    assert len(headers) == 9 and all(has_math(element) for element in headers)
    assert {element["style"]["fill"].upper() for element in headers} == {colors["primary"]}
    bands = [trace[f"values/{row}/i_0"]["style"]["fill"].upper() for row in ("value", "test")]
    assert bands == [colors["background"], colors["surface"]]
    assert trace["states"]["pptx_group"] is True
    states = descendants(trace, "states")
    assert not any(has_math(element) for element in states)
    assert "states/range_0_8" in trace and "states/range_0_8_4" in trace
    assert len([element for element in states if element["kind"] == "arrow"]) >= 3
    assert any(
        fragment["kind"] == "math" and fragment["text"] == r"[0,\,8)"
        for element in trace.values()
        for fragment in element.get("content", ())
    )

    # 5 区間：3 つの区間の数式の箱と、辺の上の位置（at=0）と間隔を指定した矢印 2 本。
    invariant = elements_of(scene, "invariant")
    assert [texts(invariant[name]) for name in ("below", "range", "above")] == ["[0, l)", "[l, r)", "[r, n)"]
    boundaries = [element for element in invariant.values() if element["kind"] == "arrow"]
    assert len(boundaries) == 2
    assert all(arrow["end_ref"]["at"] == 0 and arrow["end_ref"]["gap"] == 4 for arrow in boundaries)

    # 7 漸化式：積み上げ、役割 equation の左揃えと中央揃え、伸びる床関数、\text、出典、辺ごとの余白。
    recurrence = elements_of(scene, "recurrence")
    equations = [element for element in recurrence.values() if element["kind"] == "equation"]
    assert {element["role"] for element in equations} == {"equation"}
    assert {element["style"]["align"] for element in equations} == {"left", "center"}
    assert any(r"\left\lfloor" in texts(element) for element in equations)
    assert any(r"\text{回}" in texts(element) for element in equations)
    assert all("layout_bounds" in element for element in recurrence.values() if not element["id"].startswith("_theme"))
    assert recurrence["relation"]["sources"] == [{"bib_key": "knuth1998"}]
    padding = recurrence["meaning"]["style"]["padding_pt"]
    assert isinstance(padding, dict) and len(set(padding.values())) > 1

    # 8 図：SVG の図（ベクターの埋め込み、代替テキスト、元資料の出典）、使う場所を限った役割、凡例、許容の宣言。
    growth = elements_of(scene, "growth")
    chart = growth["chart"]
    assert Path(chart["path"]).name == "comparisons.svg" and chart.get("page") is None and chart["vector"] is True
    assert chart["description"] and chart["sources"] == [{"document_id": "figure_script"}]
    assert slides["growth"]["allow_roles"] == ["chart_label"]
    assert project.theme["roles"]["chart_label"]["restricted"] is True
    assert {growth[name]["role"] for name in ("x_label", "y_label")} == {"chart_label"}
    assert growth["legend"]["pptx_group"] is True
    assert "※" in texts(growth["computed"])
    declared = {
        ("growth", item["code"], frozenset({item["element_id"], item["related_element_id"]}))
        for item in slides["growth"]["allowed_warnings"]
    }
    assert declared == DECLARED

    # 9 比較：枠線付きの表（列と行の書式、数式だけのセルの役割）、強調、リンク、計測した寸法の箱。
    comparison = elements_of(scene, "comparison")
    cells = [element for key, element in comparison.items() if key.startswith("methods/")]
    assert all(element["style"]["line_color"] for element in cells)
    assert comparison["methods/worst/binary"]["role"] == "table_math"
    assert comparison["methods/premise/item"]["style"]["align"] == "left"
    assert comparison["methods/million/linear"]["style"]["fill"].upper() == colors["surface"]
    assert any(fragment["kind"] == "emphasis" for fragment in comparison["takeaway"]["content"])
    links = [fragment["link"] for fragment in comparison["library"]["content"] if fragment.get("link")]
    assert links == ["https://docs.python.org/3/library/bisect.html"]

    # 10 文献：文献キーごとの項目と出典。文献のページは数式の候補の検査から外す。
    references = elements_of(scene, "references")
    assert [element["sources"] for key, element in references.items() if key.startswith("entries/")] == [
        [{"bib_key": "knuth1998"}],
        [{"bib_key": "cormen2022"}],
    ]
    assert scene["rules"]["math_candidates"]["exclude_page_kinds"] == ["references"]
    cited = {
        source["bib_key"]
        for slide in scene["slides"]
        if slide["id"] != "references"
        for element in slide["elements"]
        for source in element.get("sources", ())
        if "bib_key" in source
    }
    assert cited == {"knuth1998", "cormen2022"}


def test_binary_search_builds_selected_chapters():
    """--slides で選んだページを含む章だけを作る（context.selected）。番号の指定では全ページを作る。"""
    pytest.importorskip("bibtexparser")
    _, scene = construct_example("binary_search", selection="trace")
    assert [slide["id"] for slide in scene["slides"]] == ["idea", "branching", "trace", "invariant"]
    _, scene = construct_example("binary_search", selection="cover,references")
    assert [slide["id"] for slide in scene["slides"]] == ["cover", "references"]
    _, scene = construct_example("binary_search", selection="growth")
    assert [slide["id"] for slide in scene["slides"]] == ["cost", "recurrence", "growth", "comparison"]
    _, scene = construct_example("binary_search", selection="3")
    assert tuple(slide["id"] for slide in scene["slides"]) == PAGES


@pytest.mark.parametrize("slides", ["trace,tracee", "tracee"])
def test_binary_search_reports_unknown_slide_id(slides, monkeypatch):
    """章を選ぶ資料でも、--slides の存在しない識別子は、識別子の誤りとして示す（選んだ章が 1 つもない場合を含む）。"""
    pytest.importorskip("bibtexparser")

    def load_with_fake_fonts(path):
        project = load_project(path)
        project.fonts = FakeFonts()
        return project

    monkeypatch.setattr(pipeline, "load_project", load_with_fake_fonts)
    result = pipeline.layout(BINARY_SEARCH_SOURCE / "deck.toml", slides=slides)
    assert result["exit_code"] == 2
    messages = [issue["message"] for issue in result["issues"] if issue["severity"] == "error"]
    # 章を 1 つも作らなかった場合も、作らなかった章のページの識別子を候補に示す。
    assert messages == ["ページの識別子がありません: tracee（候補: trace）"], messages


@pytest.mark.examples
@pytest.mark.parametrize("name", EXAMPLE_NAMES)
def test_examples_layout_is_clean(name, windows_fonts_environment):
    if name == "binary_search":
        pytest.importorskip("bibtexparser")
    result = pipeline.layout(EXAMPLES / name / "src" / "deck.toml")
    assert result["exit_code"] == 0, result["issues"]
    undeclared = [issue for issue in result["issues"] if "declared" not in issue]
    assert undeclared == []
    declared = {
        (issue["slide_id"], issue["code"], frozenset({issue["element_id"], issue["related_element_id"]}))
        for issue in result["issues"]
    }
    assert declared == (DECLARED if name == "binary_search" else set())


# ---- 雛形と図 ----


def test_example_template_and_figure_are_reproducible(tmp_path):
    template = load_script(BINARY_SEARCH_SOURCE / "make_template.py").build_template(tmp_path / "template.pptx")
    with zipfile.ZipFile(BINARY_SEARCH_SOURCE / "template.pptx") as bundled, zipfile.ZipFile(template) as made:
        assert bundled.namelist() == made.namelist()
        for name in bundled.namelist():
            assert bundled.read(name) == made.read(name), name
    # 同梱の図は自己完結した SVG で、文字を持たない（目盛りの数字は輪郭）。
    data = (BINARY_SEARCH_SOURCE / "figures/comparisons.svg").read_bytes()
    _check_svg(data)
    assert b"<text" not in data and b"<title>" in data
    script = load_script(BINARY_SEARCH_SOURCE / "make_figure.py")
    if not script.TICK_FONT.exists():
        pytest.skip(f"図の目盛りの数字の書体がない: {script.TICK_FONT}")
    assert script.build_figure(tmp_path / "comparisons.svg").read_bytes() == data


def documents_module() -> types.ModuleType:
    """文書の試験（test_documents.py）のモジュール。使用例の README のリンクも文書と同じ確かめ方で照らす。

    関数の中で読み込む（文書の試験がこの試験のファイルでもう一度収集されないように）。
    """
    import test_documents

    return test_documents


def test_example_readme_links_resolve():
    """使用例の README の相対リンクが、ファイル・ディレクトリと見出し（#…）までたどることができる。

    文書の試験（test_documents.py）は使用例の README を対象にしないので、ここで同じ確かめ方を使う。
    """
    check_links = documents_module().test_relative_links_resolve
    readmes = sorted(EXAMPLES.glob("*/README.md"))
    assert BINARY_SEARCH / "README.md" in readmes
    for readme in readmes:
        check_links(readme.relative_to(ROOT).as_posix())


def test_example_readme_link_check_follows_headings(tmp_path, monkeypatch):
    """使用例の README のリンクの確かめ方は、ファイルの有無だけでなく、指した文書と README 自身の見出しまで照らす。"""
    documents = documents_module()
    monkeypatch.setattr(documents, "ROOT", tmp_path)
    guide = tmp_path / "docs" / "usage.md"
    guide.parent.mkdir()
    guide.write_text("# 利用ガイド\n\n## 部分の生成\n", encoding="utf-8")
    readme = tmp_path / "examples" / "demo" / "README.md"
    readme.parent.mkdir(parents=True)
    name = readme.relative_to(tmp_path).as_posix()

    readme.write_text(
        "# 使用例\n\n## 作り方\n\n[ガイド](../../docs/usage.md#部分の生成)、[手順](#作り方)\n", encoding="utf-8"
    )
    documents.test_relative_links_resolve(name)
    for target, problem in (
        ("../../docs/usage.md#ない見出し", "見出し #ない見出し がない"),
        ("#ない見出し", "見出し #ない見出し がない"),
        ("../../docs/missing.md", "がない"),
    ):
        readme.write_text(f"# 使用例\n\n[リンク]({target})\n", encoding="utf-8")
        with pytest.raises(AssertionError, match=problem):
            documents.test_relative_links_resolve(name)


# ---- 自己完結と作成者 ----


def example_inputs(directory: Path, config: dict) -> list[str]:
    """deck.toml が参照する入力のパス（入口、テーマ、記述規則、文献、素材のフォルダー、元資料、雛形）。"""
    inputs = config.get("inputs", {})
    paths = [
        config["entrypoint"].rsplit(":", 1)[0],
        *inputs.get("python", ()),
        *inputs.get("bibliography", ()),
        *inputs.get("asset_dirs", ()),
        *(document["path"] for document in inputs.get("documents", ())),
    ]
    for key in ("theme", "rules"):
        if key in config:
            paths.append(config[key])
    if "template" in config:
        paths.append(config["template"]["path"])
    return paths


def test_examples_are_self_contained():
    from pptx import Presentation

    for name in EXAMPLE_NAMES:
        directory = EXAMPLES / name / "src"
        config = read_toml(directory / "deck.toml")
        assert "extends" not in config, name
        assert config["inputs"]["read_roots"] == ["."], name
        for value in example_inputs(directory, config):
            path = (directory / value).resolve()
            assert path.is_relative_to(directory) and path.exists(), (name, value)
        # 完成物は使用例のディレクトリの直下に、採用記録は src/ に保存する（推奨の構成）。
        for key, value in config.get("deliverables", {}).items():
            expected = directory if key == "record" else directory.parent
            assert (directory / value).resolve().parent == expected, (name, key, value)
        metadata = config["metadata"]
        assert (metadata["author"], metadata["last_modified_by"]) == (AUTHOR, AUTHOR), name
        if "template" in config:
            properties = Presentation(directory / config["template"]["path"]).core_properties
            assert (properties.author, properties.last_modified_by) == (AUTHOR, AUTHOR), name
        # 反映した完成物があれば、その文書情報の作成者も確かめる。
        pptx = directory / config["deliverables"]["pptx"]
        if pptx.is_file():
            properties = Presentation(pptx).core_properties
            assert (properties.author, properties.last_modified_by) == (AUTHOR, AUTHOR), name
        pdf = directory / config["deliverables"]["pdf"]
        if pdf.is_file():
            assert pdf_metadata(pdf)["author"] == AUTHOR, name


# ---- 手元のパスと非公開の語 ----


def _ignored(root: Path):
    """公開しないものかを判定する関数を返す。.gitignore で除くものと、.git・.venv・out/（.gitignore がなくても除く）。

    .gitignore の行は名前の照合（`*`・`?`・`[…]`。先頭の `**/` は除き、末尾の `/` はディレクトリだけ）として読む。
    否定（`!`）と場所を限る行は、この照合で読むことができないので失敗にする。
    """
    # .git は git の記録で、公開するファイルではない（git は .gitignore になくても除く）。
    patterns = [(".git", True), (".venv", True), ("out", True)]
    ignore = root / ".gitignore"
    for line in ignore.read_text(encoding="utf-8").splitlines() if ignore.is_file() else ():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        name = line.removeprefix("**/").rstrip("/")
        assert not line.startswith("!") and "/" not in name, f".gitignore の行を名前の照合で読めない: {line}"
        patterns.append((name, line.endswith("/")))

    def ignored(path: Path) -> bool:
        return any(
            fnmatch.fnmatchcase(path.name, name) and (path.is_dir() or not directory) for name, directory in patterns
        )

    return ignored


def _walk(path: Path, ignored):
    for child in sorted(path.iterdir()):
        if ignored(child):
            continue
        if child.is_dir():
            yield from _walk(child, ignored)
        else:
            yield child


def _texts(path: Path):
    """ファイルの（場所、文字）。文字にできない形式なら、文字を None にする。

    PPTX は XML の部品（文書情報、スライド、ノート、関係）、PDF は文書情報・XMP・各ページの文字とリンク先を調べる。
    ほかは形式によらず UTF-8 の文字として読む。
    """
    if path.suffix in (".pptx", ".potx"):
        with zipfile.ZipFile(path) as package:
            for name in package.namelist():
                if name.endswith((".xml", ".rels")):
                    yield f"{path}:{name}", package.read(name).decode("utf-8", errors="replace")
    elif path.suffix == ".pdf":
        reader = pypdf.PdfReader(path)
        yield f"{path}:metadata", "\n".join(value for value in pdf_metadata(path).values() if value)
        stream = reader.trailer["/Root"].get("/Metadata")
        yield f"{path}:xmp", (
            stream.get_object().get_data().decode("utf-8", errors="replace") if stream is not None else ""
        )
        with open_pdf(path) as document:
            for page in document:
                yield f"{path}:{page.index + 1}", "\n".join([_page_text(page), *_link_targets(reader, page.index)])
    else:
        try:
            yield str(path), path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            yield str(path), None


def _page_text(page) -> str:
    """ページの文字（描画順。見えない文字を含む）。

    行が変わる所（基線が文字サイズの半分より大きく変わる所と、文字サイズの半分より大きく左へ戻る所）で改行し、同じ行で
    字の間が空いている所に空白を入れる（別の枠の語や数式の字が続いて 1 語にならず、英数字の境界で照らす語も見つかるようにする）。
    """
    text, last = [], None
    for span in page.spans():
        for char in span.chars:
            if last is not None:
                previous, size = last[0], max(last[1], span.size)
                gap = char.bbox[0] - previous.bbox[2]
                if abs(char.origin[1] - previous.origin[1]) > size / 2 or gap < -size / 2:
                    text.append("\n")
                elif gap > size * 0.15:
                    text.append(" ")
            text.append(chr(char.unicode))
            last = char, span.size
    return "".join(text)


def _link_targets(reader, index: int) -> list[str]:
    """ページのリンクの宛先（URI と、ファイルを開くリンクのファイル名）。reader は pypdf の PdfReader。"""
    targets = []
    for annotation in reader.pages[index].get("/Annots") or []:
        action = annotation.get_object().get("/A")
        action = action.get_object() if action is not None else {}
        for key in ("/URI", "/F"):
            value = action.get(key)
            value = value.get_object() if value is not None else None
            # ファイルの指定は、文字列か、ファイル名（/UF・/F）を持つ辞書である。
            if isinstance(value, dict):
                value = value.get("/UF") or value.get("/F")
            if value is not None:
                targets.append(str(value))
    return targets


def scanned_texts(root: Path = ROOT):
    """調べる文字（場所、内容、ソース・試験・道具か）。内容が None なら文字にできない。

    範囲は slidekit の全体で、.gitignore で除くもの（.venv、out/、__pycache__ など）だけを除く。
    使用例の雛形と図、完成物と採用記録、直下の設定（pyproject.toml・requirements-lock.txt・.gitignore）も、形式によらず調べる。
    """
    for path in _walk(root, _ignored(root)):
        code = path.relative_to(root).parts[0] in CODE_DIRECTORIES
        for place, text in _texts(path):
            yield place, text, code


def _scan(words_of) -> None:
    """調べる範囲の文字ごとに words_of(文字, ソース・試験・道具か) で見つけた語を集め、ないことを確かめる。

    文字にできないファイルも、調べていないものとして報告する（調べ方を加えるまで失敗にする）。
    """
    found = []
    for place, text, code in scanned_texts():
        words = ["文字にできない形式（調べ方を加える）"] if text is None else words_of(text, code)
        found.extend((place, word) for word in dict.fromkeys(words))
    assert found == [], found


def test_scan_covers_every_published_file(tmp_path):
    """手元のパスと非公開の語の検査は、.gitignore で除くもののほかのすべてのファイルを、形式によらず調べる。

    完成物と採用記録、完成物でない PPTX・PDF（雛形と図）、直下の設定、文書の .md でないファイルも調べる。
    文字にできない形式は None で返す。
    """

    def write(relative: str, content) -> Path:
        path = tmp_path / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        if isinstance(content, bytes):
            path.write_bytes(content)
        else:
            path.write_text(content, encoding="utf-8")
        return path

    def pptx(relative: str, mark: str) -> None:
        with zipfile.ZipFile(write(relative, b""), "w") as package:
            package.writestr("ppt/slides/slide1.xml", f"<p:sld><a:t>{mark}</a:t></p:sld>")
            package.writestr("ppt/slides/_rels/slide1.xml.rels", "<Relationships/>")
            package.writestr("ppt/media/image1.png", b"\x89PNG\r\n\x1a\n")

    # 実際の .gitignore で除くものを確かめる。
    write(".gitignore", (ROOT / ".gitignore").read_text(encoding="utf-8") + "# mark-gitignore\n")
    for relative in ("pyproject.toml", "requirements-lock.txt", "README.md", "docs/notes.txt", "workflows/steps.md"):
        write(relative, f"mark-{relative}\n")
    for relative in ("src/pkg/module.py", "tests/data/sample.json", "tools/tool.py"):
        write(relative, f"code-{relative}\n")
    for relative in (
        ".git/config",
        ".venv/lib/site.py",
        "out/runs/0001/manifest.json",
        "examples/demo/out/runs/0001/deck.json",
        "src/pkg/__pycache__/module.cpython-314.pyc",
        "src/pkg/module.pyc",
        "src/pkg.egg-info/PKG-INFO",
        "build/lib/module.py",
        "environment.local.toml",
    ):
        write(relative, "mark-ignored\n")
    write("examples/demo/deck.toml", '[deliverables]\npptx = "demo.pptx"\nrecord = "deliverables.json"\n')
    write("examples/demo/deliverables.json", json.dumps({"schema_version": 2, "run": "mark-record"}))
    write("examples/demo/slides.py", "mark-slides\n")
    pptx("examples/demo/template.pptx", "mark-template")
    pptx("examples/demo/demo.pptx", "mark-deliverable")
    figure = write("examples/demo/figures/figure.pdf", b"")
    with Document() as document:
        page = document.new_page()
        page.insert_text((20, 40), "mark-pdf-text")
        page.insert_link({"from": (0, 0, 20, 20), "uri": "https://mark-pdf-link/"})
        document.set_metadata({"author": "mark-pdf-author"})
        document.save(figure)
    # XMP は pdf_tools で書くことができないので、pypdf で加える。
    writer = pypdf.PdfWriter(clone_from=figure)
    writer.xmp_metadata = b"<x:xmpmeta>mark-pdf-xmp</x:xmpmeta>"
    writer.write(figure)
    unreadable = write("examples/demo/data.bin", b"\xff\xfe\x00\x81")

    found, missing = {}, []
    for place, text, code in scanned_texts(tmp_path):
        if text is None:
            missing.append(place)
            continue
        for mark in re.findall(r"(?:mark|code)-[\w./-]*\w", text):
            found.setdefault(mark, set()).add(code)
    assert missing == [str(unreadable)]
    published = ["gitignore", "pyproject.toml", "requirements-lock.txt", "README.md", "docs/notes.txt"]
    published += ["workflows/steps.md", "slides", "template", "record", "deliverable"]
    published += ["pdf-text", "pdf-link", "pdf-author", "pdf-xmp"]
    code = ["src/pkg/module.py", "tests/data/sample.json", "tools/tool.py"]
    assert found == {
        **{f"mark-{name}": {False} for name in published},
        **{f"code-{name}": {True} for name in code},
    }


def test_examples_and_documents_have_no_local_paths():
    home = str(Path.home())
    user = re.escape(Path.home().name)
    # 実際の利用者のフォルダー（Windows・macOS の形）は、範囲のすべてで調べる。
    own = re.compile(rf"(?i:[\\/]users[\\/]+{user})(?![A-Za-z0-9_])")

    def words_of(text: str, code: bool) -> list[str]:
        patterns = [POSIX_HOME, own] + ([] if code else [WINDOWS_HOME])
        words = [match.group() for pattern in patterns for match in pattern.finditer(text)]
        return words + [value for value in (home, str(ROOT)) if value in text]

    _scan(words_of)


def test_private_terms_are_absent():
    """SLIDEKIT_PRIVATE_TERMS が 1 行 1 語の UTF-8 のファイルを指すときだけ、その語がないことを確かめる。

    ASCII の語は大文字小文字を区別し、英数字の境界で照らす（一般的な英単語の一部に一致しないように）。
    ほかの語は部分文字列で照らす。語の一覧はリポジトリに置かない。
    """
    source = os.environ.get("SLIDEKIT_PRIVATE_TERMS")
    if not source:
        pytest.skip("SLIDEKIT_PRIVATE_TERMS で語の一覧を指定したときだけ調べる")
    terms = [
        (number, line.strip())
        for number, line in enumerate(Path(source).read_text(encoding="utf-8").splitlines(), 1)
        if line.strip()
    ]
    assert terms, "語の一覧が空である"
    matchers = []
    for number, term in terms:
        if term.isascii():
            matches = re.compile(rf"(?<![A-Za-z0-9]){re.escape(term)}(?![A-Za-z0-9])").search
        else:
            matches = lambda text, term=term: term in text  # noqa: E731
        matchers.append((number, matches))
    # 語そのものは報告に出さない（一覧の行番号だけを示す）。
    _scan(lambda text, _code: [f"{number} 行目の語" for number, matches in matchers if matches(text)])
