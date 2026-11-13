"""照合用の資料の道具（tools/calibrate.py）と、予測と実際を重ねる確認画像の道具（tools/layout_overlay.py）の試験。

PowerPoint は使わない。slidekit build の代わりに、資料を偽のフォントで構築して標準 14 書体で描いた PDF を持つ実行を作る。
"""

from __future__ import annotations

import functools
import json
from pathlib import Path
import subprocess

from PIL import Image
import pytest

from fake_fonts import FakeFonts
from pdf_tools import Document, text_length
from slidekit import measure
from slidekit.config import load_project
from slidekit.pipeline import construct
from test_layout_fixture import RGB, InkFonts, load_tool, synthetic_run

ROOT = Path(__file__).resolve().parents[1]
CALIBRATE = ROOT / "tools" / "calibrate.py"
OVERLAY = ROOT / "tools" / "layout_overlay.py"

CASES = """schema_version = 1
deck = "cases"

[defaults]
latin = "Helvetica"
japanese = "Helvetica"
math = "Times-Roman"

[[cases]]
name = "plain"
expressions = ["x+1"]
mode = ["inline", "display"]
tags = ["text_baseline"]
"""


def calibrate():
    return load_tool(CALIBRATE)


def overlay_tool():
    return load_tool(OVERLAY)


def write_config(tmp_path, text=CASES, name="calibration.toml") -> Path:
    path = tmp_path / name
    path.write_text(text, encoding="utf-8")
    return path


# ---- 設定と資料 ----


def test_calibrate_expands_cases_into_a_generic_deck(tmp_path):
    tool = calibrate()
    config = write_config(
        tmp_path,
        """schema_version = 1

[defaults]
line_spacing = [1.0, 1.2]

[[cases]]
name = "script"
expressions = ["m^{1/4}", 'x^{\\frac{a}{b}}']
size_pt = [12, 20, 32]
tags = ["text_baseline"]

[[cases]]
expressions = ['\\frac{a}{\\frac{b}{c}}']
mode = ["inline", "display"]
valign = ["top", "bottom"]
before = "値は ["
after = "] です。"
""",
    )
    conditions = tool.prepare(config, tmp_path / "work")
    work = tmp_path / "work"
    assert sorted(path.name for path in work.iterdir()) == [
        "calibration.toml",
        "conditions.json",
        "deck.toml",
        "slides.py",
        "theme.toml",
    ]
    assert json.loads((work / "conditions.json").read_text(encoding="utf-8")) == conditions
    assert conditions["deck"] == "cases" and conditions["versions"]["measure"] == measure.MEASURE_VERSION
    items = [item for page in conditions["pages"] for item in page["items"]]
    assert len(items) == 2 * 3 * 2 + 1 * 2 * 2 * 2
    ids = [item["id"] for item in items]
    assert len(set(ids)) == len(ids) and "script_01_inline_s12_sp1.2_left_top" in ids
    assert "case2_01_display_s20_sp1_left_bottom" in ids
    for page in conditions["pages"]:
        boxes = [item["area"] for item in page["items"]]
        assert all(
            tool.LEFT <= x and x + w <= tool.RIGHT and tool.TOP <= y and y + h <= tool.BOTTOM for x, y, w, h in boxes
        )
        # 数式の照合データの計測は枠を上下に 4 pt、左右に 2 pt 広げて見るので、枠の間を GAP 空ける。
        assert not any(
            a[0] < b[0] + b[2] + tool.GAP
            and b[0] < a[0] + a[2] + tool.GAP
            and a[1] < b[1] + b[3] + tool.GAP
            and b[1] < a[1] + a[3] + tool.GAP
            for i, a in enumerate(boxes)
            for b in boxes[i + 1 :]
        )
    (large,) = [item for item in items if item["id"] == "script_01_inline_s32_sp1.2_left_top"]
    assert large["area"][2:] == [pytest.approx(14.4 * 32), pytest.approx(5.6 * 32 * 1.2)]
    project = load_project(work / "deck.toml")
    project.fonts = FakeFonts()
    scene = construct(project)
    assert "issues" not in scene and len(scene["slides"]) == len(conditions["pages"])
    elements = {element["id"]: element for slide in scene["slides"] for element in slide["elements"]}
    small = elements["script_01_inline_s12_sp1.2_left_top"]
    assert small["description"] == "tags: text_baseline"
    assert [fragment["text"] for fragment in small["content"]] == ["[", "m^{1/4}", "]"]
    assert {key: small["style"][key] for key in ("size_pt", "line_spacing", "word_wrap", "padding_pt")} == {
        "size_pt": 12.0,
        "line_spacing": 1.2,
        "word_wrap": False,
        "padding_pt": 0.0,
    }
    display = elements["case2_01_display_s20_sp1_left_bottom"]
    assert display["kind"] == "equation" and display["style"]["valign"] == "bottom"
    assert [fragment["text"] for fragment in elements["case2_01_inline_s20_sp1_left_top"]["content"]] == [
        "値は [",
        r"\frac{a}{\frac{b}{c}}",
        "] です。",
    ]
    load_tool(ROOT / "tools" / "make_layout_fixture.py").check_public(scene)


def test_calibrate_writes_layout_and_math_decks(tmp_path):
    tool = calibrate()
    layout = tool.prepare(write_config(tmp_path, 'schema_version = 1\ndeck = "layout"\n'), tmp_path / "layout")
    text = (tmp_path / "layout" / "deck.toml").read_text(encoding="utf-8")
    assert layout["extends"] == "tests/data/layout/deck/deck.toml" and "extends = " in text
    project = load_project(tmp_path / "layout" / "deck.toml")
    assert project.output_dir == (tmp_path / "layout" / "out").resolve()
    config = write_config(tmp_path, 'schema_version = 1\ndeck = "math"\nexpressions = ["x^2", "\\\\unknownmacro"]\n')
    math = tool.prepare(config, tmp_path / "math")
    assert [item["expression"] for page in math["pages"] for item in page["items"]] == ["x^2", "x^2"]
    assert list(math["rejected"]) == [r"\unknownmacro"] and math["size_pt"] == 20
    project = load_project(tmp_path / "math" / "deck.toml")
    project.fonts = FakeFonts()
    assert len(construct(project)["slides"]) == 2


@pytest.mark.parametrize(
    "text",
    [
        "deck = 'cases'\n",  # schema_version がない
        "schema_version = 1\ndeck = 'other'\n",
        "schema_version = 1\n",  # cases がない
        "schema_version = 1\n[[cases]]\nexpressions = []\n",
        "schema_version = 1\n[[cases]]\nexpressions = ['x']\nmode = 'block'\n",
        "schema_version = 1\n[[cases]]\nexpressions = ['x']\nsize_pt = 0\n",
        "schema_version = 1\n[[cases]]\nexpressions = ['x']\ncolor = 'red'\n",
        "schema_version = 1\n[[cases]]\nexpressions = ['x']\ntags = ['a b']\n",
        "schema_version = 1\n[[cases]]\nname = 'a'\nexpressions = ['x']\n[[cases]]\nname = 'a'\nexpressions = ['y']\n",
        "schema_version = 1\n[[cases]]\nexpressions = ['x']\nsize_pt = 100\n",  # 枠がページに入らない
        "schema_version = 1\ndeck = 'layout'\nexpressions = ['x']\n",
    ],
)
def test_calibrate_rejects_malformed_config(tmp_path, text):
    tool = calibrate()
    with pytest.raises(tool.CalibrationError):
        tool.prepare(write_config(tmp_path, text), tmp_path / "work")
    assert not (tmp_path / "work").exists()  # 誤りがあれば作業ディレクトリを作らない


def test_calibrate_needs_new_work_directory_outside_slidekit(tmp_path):
    tool = calibrate()
    config = write_config(tmp_path)
    with pytest.raises(tool.CalibrationError, match="slidekit の外"):
        tool.prepare(config, ROOT / "tests" / "calibration-work")
    assert not (ROOT / "tests" / "calibration-work").exists()
    (tmp_path / "used").mkdir()
    (tmp_path / "used" / "file.txt").write_text("x", encoding="utf-8")
    with pytest.raises(tool.CalibrationError, match="新しいディレクトリ"):
        tool.prepare(config, tmp_path / "used")


# ---- 変換（slidekit build の代わり）と計測 ----


def fake_build(work: Path, number: int = 1) -> Path:
    """資料を偽のフォントで構築し、要素の枠の左上から標準 14 書体で描いた PDF と生成記録を持つ実行を作る。"""
    project = load_project(work / "deck.toml")
    project.fonts = FakeFonts()
    scene = construct(project)
    run = work / "out" / "runs" / f"{number:04d}"
    standard = run / "renders" / "powerpoint" / "standard"
    standard.mkdir(parents=True)
    (run / "deck.json").write_text(json.dumps(scene, ensure_ascii=False), encoding="utf-8")
    families = {
        element["style"][key]
        for slide in scene["slides"]
        for element in slide["elements"]
        for key in ("font_latin", "font_japanese", "font_math")
        if element.get("style", {}).get(key)
    }
    inventory = {
        family: {"path": f"{family}.ttf", "index": 0, "sha256": f"{n:064d}", "version": "Version 1.0"}
        for n, family in enumerate(sorted(families), 1)
    }
    manifest = {
        "project": str(work / "deck.toml"),
        "environment": {
            "fonts": {"fonts": inventory},
            "packages": {"pypdfium2": "5.14.0"},
            "renderers": {"powerpoint": {"version": "16.0.1.2"}},
        },
    }
    (run / "manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
    with Document() as document:
        for slide in scene["slides"]:
            page = document.new_page(width=960, height=540)
            for element in slide["elements"]:
                if element["id"].startswith("_theme/"):
                    continue
                x, y = element["area"][:2]
                baseline = y + 18 + number  # 実行ごとに 1 pt ずつ下げる
                for fragment in element["content"]:
                    font = "tiro" if fragment["kind"] == "math" else "helv"
                    page.insert_text((x, baseline), fragment["text"], fontname=font, fontsize=20, color=RGB)
                    x += text_length(fragment["text"], fontname=font, fontsize=20)
        document.save(standard / "deck.pdf")
    return run


def runner_for(work: Path):
    """slidekit build の代わりに fake_build を呼ぶ runner と、呼ばれたコマンドの一覧。"""
    calls = []

    def run(command, **options):
        calls.append(command)
        assert options["capture_output"] and command[1:4] == ["-m", "slidekit", "build"] and "--json" in command
        built = fake_build(work, len(calls))
        result = {"status": "built", "build": str(built), "counts": {"error": 0}}
        return subprocess.CompletedProcess(command, 0, stdout=json.dumps(result), stderr="")

    return run, calls


def test_calibrate_converts_with_slidekit_build(tmp_path):
    tool = calibrate()
    work = tmp_path / "work"
    tool.prepare(write_config(tmp_path), work)
    runner, calls = runner_for(work)
    run = tool.convert(work, runner=runner)
    assert run == work / "out" / "runs" / "0001" and calls[0][-6:-3] == ["--project", str(work / "deck.toml"), "--json"]
    assert tool.find_run(work) == run.resolve() and tool.find_run(work, "1") == run.resolve()

    def failed(command, **options):
        return subprocess.CompletedProcess(command, 2, stdout=json.dumps({"status": "failed", "build": ""}), stderr="")

    with pytest.raises(tool.CalibrationError, match="変換を終えていない"):
        tool.convert(work, runner=failed)

    def broken(command, **options):
        return subprocess.CompletedProcess(command, 2, stdout="", stderr="Traceback\nRuntimeError: x")

    with pytest.raises(tool.CalibrationError, match="RuntimeError: x"):
        tool.convert(work, runner=broken)


def test_calibrate_run_measures_and_compares(tmp_path, monkeypatch, capsys):
    tool = calibrate()
    work = tmp_path / "work"
    runner, calls = runner_for(work)
    convert = functools.partial(tool.convert, runner=runner)
    monkeypatch.setattr(tool, "convert", convert)
    fonts = InkFonts({})
    monkeypatch.setattr(tool, "measure_run", functools.partial(tool.measure_run, fonts=fonts))
    missing = tmp_path / "none.json"
    arguments = ["--layout-fixture", str(missing), "--math-fixture", str(missing)]
    code = tool.main(["run", str(write_config(tmp_path)), "--work", str(work), *arguments])
    output = capsys.readouterr().out
    assert code == 1 and len(calls) == 1  # 偽のフォントの計測は標準 14 書体の PDF と合わない
    assert "文字の配置: 要素 2 個" in output and "数式: 項目 2 件、計測できない項目 0 件" in output
    results = work / "results" / "0001"
    assert sorted(path.name for path in results.iterdir()) == [
        "comparison.json",
        "layout_fixture.json",
        "math_fixture.json",
    ]
    result = json.loads((results / "comparison.json").read_text(encoding="utf-8"))
    layout = result["layout"]
    assert layout["cases"] == 2 and layout["unread"] == [] and layout["fixture"] == {"available": False}
    rows = {row["key"]: row for row in layout["differences"]}
    inline = rows["p01/plain_01_inline_s20_sp1_left_top"]
    assert inline["tags"][:2] == ["p01", "text_baseline"] and inline["lines"] == [1, 1]
    predicted = measure.layout_element(
        json.loads((results / "layout_fixture.json").read_text(encoding="utf-8"))["cases"][0]["element"], fonts
    )
    assert inline["baseline_pt"] == [
        pytest.approx(19 - predicted.block_top + 14 - predicted.text.lines[0].baseline, abs=0.01)
    ]
    # 許容差の規則（text_baseline）は計測の比較にも当てはまる。
    assert any("規則 text_baseline" in message for message in layout["problems"])
    math = result["math"]
    (group,) = math["measurement"]["groups"]
    assert (group["size_pt"], group["math"], group["count"]) == (20.0, "Times-Roman", 2)
    assert {row["id"] for row in group["items"]} == {
        "plain_01_inline_s20_sp1_left_top",
        "plain_01_display_s20_sp1_left_top",
    }
    items = json.loads((results / "math_fixture.json").read_text(encoding="utf-8"))["items"]
    advance = text_length("x+1", fontname="tiro", fontsize=20)
    (inline_item,) = [item for item in items if item["mode"] == "inline"]
    assert inline_item["advance_pt"] == pytest.approx(advance, abs=0.01) and inline_item["size_pt"] == 20

    # 変換せずに計測し直す。既存の照合データとの差も求める。
    layout_data = json.loads((results / "layout_fixture.json").read_text(encoding="utf-8"))
    moved = json.loads(json.dumps(layout_data))
    for case in moved["cases"]:  # 位置が違っても、枠の左上からの位置が同じなら同じ
        case["element"]["area"][0] += 100
        case["actual"]["line_extents_pt"] = [[a + 100, b + 100] for a, b in case["actual"]["line_extents_pt"]]
        case["actual"]["ink_pt"][0] += 100
        case["actual"]["ink_pt"][2] += 100
    moved["cases"][1]["actual"]["baselines_pt"][0] += 0.5
    (tmp_path / "layout.json").write_text(json.dumps(moved), encoding="utf-8")
    math_data = {
        "size_pt": 20,
        "font": {"family": "Times-Roman"},
        "text_fonts": {"japanese": "Helvetica", "latin": "Helvetica"},
        "items": [{**inline_item, "advance_pt": inline_item["advance_pt"] - 0.2}],
    }
    (tmp_path / "math.json").write_text(json.dumps(math_data), encoding="utf-8")
    monkeypatch.setattr(tool, "convert", lambda *args, **kwargs: pytest.fail("変換しない"))
    fixtures = ["--layout-fixture", str(tmp_path / "layout.json"), "--math-fixture", str(tmp_path / "math.json")]
    assert tool.main(["run", "--work", str(work), "--measure-only", *fixtures]) == 1
    result = json.loads((results / "comparison.json").read_text(encoding="utf-8"))
    fixture = result["layout"]["fixture"]
    assert (fixture["same"], [row["key"] for row in fixture["changed"]], fixture["new"]) == (
        1,
        ["p01/plain_01_display_s20_sp1_left_top"],
        [],
    )
    assert fixture["changed"][0]["baseline_pt"] == pytest.approx(0.5)
    fixture = result["math"]["fixture"]
    assert fixture["new"] == ["display:x+1"] and fixture["changed"][0]["largest_pt"] == pytest.approx(0.2, abs=0.001)
    assert fixture["changed"][0]["largest_em"] == pytest.approx(0.01, abs=0.0001)
    assert "文字の照合データとの差: 同じ 1 件、変わった 1 件" in capsys.readouterr().out
    # 設定ファイルを書かずに run すると、作業ディレクトリの資料を新しい実行に変換し直す。
    monkeypatch.setattr(tool, "convert", convert)
    assert tool.main(["run", "--work", str(work), *fixtures]) == 1
    assert len(calls) == 2 and (work / "results" / "0002" / "comparison.json").is_file()
    assert tool.find_run(work).name == "0002"
    # --measure-only は設定ファイルを受け付けない。生成条件のない作業ディレクトリは変換しない。
    with pytest.raises(SystemExit):
        tool.main(["run", str(write_config(tmp_path)), "--work", str(work), "--measure-only"])
    assert tool.main(["run", "--work", str(tmp_path / "empty")]) == 2
    assert "生成条件がありません" in capsys.readouterr().err


def test_calibrate_measures_math_deck_like_math_tool(tmp_path, monkeypatch):
    tool = calibrate()
    work = tmp_path / "work"
    tool.prepare(write_config(tmp_path, 'schema_version = 1\ndeck = "math"\nexpressions = ["x+1"]\n'), work)
    run = fake_build(work)
    record = {"family": "Cambria Math", "file": "cambria.ttc", "index": 1, "version": "Version 1.0", "sha256": "0" * 64}
    monkeypatch.setattr(tool.math_tool(), "font_record", lambda project: record)
    theme = load_project(work / "deck.toml").theme["fonts"]
    existing = {
        "size_pt": 20,
        "font": {"family": theme["math"]},
        "text_fonts": {"japanese": theme["japanese"], "latin": theme["latin"]},
        "items": [
            {"expression": "x+1", "mode": "display", "ink_width_pt": 1.0, "ink_height_pt": 1.0},
            {"expression": "y", "mode": "inline", "advance_pt": 10.0, "ink_pt": None},
        ],
    }
    (tmp_path / "math.json").write_text(json.dumps(existing), encoding="utf-8")
    result = tool.measure_run(work, run, fonts=InkFonts({}), layout_fixture=None, math_fixture=tmp_path / "math.json")
    data = json.loads((work / "results" / "0001" / "math_fixture.json").read_text(encoding="utf-8"))
    assert (data["powerpoint"], data["font"], data["size_pt"]) == ("16.0.1.2", record, 20.0)
    assert data["text_fonts"] == {"japanese": theme["japanese"], "latin": theme["latin"]}
    assert [(item["expression"], item["mode"]) for item in data["items"]] == [("x+1", "inline"), ("x+1", "display")]
    advance = text_length("x+1", fontname="tiro", fontsize=20)
    assert data["items"][0]["advance_pt"] == pytest.approx(advance, abs=0.01)
    math = result["math"]
    assert math["failed"] == [] and math["measurement"]["groups"][0]["count"] == 2
    # 式の一覧の一部の資料では、照合データだけにある式を示さない。
    fixture = math["fixture"]
    assert (fixture["new"], [row["key"] for row in fixture["changed"]], fixture["missing"]) == (
        ["inline:x+1"],
        ["display:x+1"],
        [],
    )


def test_calibrate_reports_unusable_input(tmp_path, capsys):
    tool = calibrate()
    assert tool.main(["run", "--work", str(tmp_path / "none"), "--measure-only"]) == 2
    assert "PDF のある実行がありません" in capsys.readouterr().err
    with pytest.raises(SystemExit) as raised:
        tool.main(["--help"])
    assert raised.value.code == 0
    assert "run" in capsys.readouterr().out and "adopt" in tool.__doc__


# ---- 採用 ----


def results_of(tmp_path, deck: str, files: dict) -> Path:
    work = tmp_path / "work"
    results = work / "results" / "0001"
    results.mkdir(parents=True)
    (work / "conditions.json").write_text(json.dumps({"schema_version": 1, "deck": deck}), encoding="utf-8")
    for name, value in files.items():
        (results / name).write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
    return work


def test_calibrate_adopts_layout_fixture_explicitly(tmp_path):
    tool = calibrate()
    candidate = load_tool(ROOT / "tools" / "make_layout_fixture.py").make_fixture(synthetic_run(tmp_path / "source"))
    work = results_of(tmp_path, "layout", {"layout_fixture.json": candidate, "comparison.json": {"run": "0001"}})
    output = tmp_path / "powerpoint_layout.json"
    assert tool.adopt(work, output=output) == (output, 3)
    assert json.loads(output.read_text(encoding="utf-8")) == candidate
    # 一部だけ置き換える。ほかの要素は既存の値のまま。
    existing = json.loads(json.dumps(candidate))
    for case in existing["cases"]:
        case["actual"]["baselines_pt"] = [value + 1 for value in case["actual"]["baselines_pt"]]
    output.write_text(json.dumps(existing), encoding="utf-8")
    assert tool.adopt(work, output=output, only=["first/lead"]) == (output, 1)
    adopted = {f"{case['slide_id']}/{case['element_id']}": case for case in json.loads(output.read_text())["cases"]}
    assert adopted["first/lead"] == candidate["cases"][0]
    assert adopted["first/note"]["actual"]["baselines_pt"] == [
        value + 1 for value in candidate["cases"][1]["actual"]["baselines_pt"]
    ]
    with pytest.raises(tool.CalibrationError, match="両方にない要素"):
        tool.adopt(work, output=output, only=["first/none"])
    existing["environment"]["powerpoint"] = "16.0.0.1"
    output.write_text(json.dumps(existing), encoding="utf-8")
    with pytest.raises(tool.CalibrationError, match="environment"):
        tool.adopt(work, output=output, only=["first/lead"])
    stale = {**candidate, "measure_version": f"{measure.MEASURE_VERSION}-old"}
    (work / "results" / "0001" / "layout_fixture.json").write_text(json.dumps(stale), encoding="utf-8")
    with pytest.raises(tool.CalibrationError, match="--measure-only"):
        tool.adopt(work, output=output)


def test_calibrate_adopts_math_fixture_explicitly(tmp_path, capsys):
    tool = calibrate()
    head = {
        "schema_version": 1,
        "font": {"family": "Cambria Math", "sha256": "0" * 64},
        "powerpoint": "16",
        "size_pt": 20,
    }
    items = [
        {"expression": "x", "mode": "inline", "advance_pt": 10.0, "ink_pt": [0.5, 9.5, 9.0, 0.2]},
        {"expression": "x", "mode": "display", "ink_width_pt": 9.0, "ink_height_pt": 9.2},
        {"expression": "a:b", "mode": "inline", "advance_pt": 30.0, "ink_pt": [0.5, 29.5, 14.0, 0.2]},
    ]
    candidate = {**head, "items": items}
    work = results_of(tmp_path, "math", {"math_fixture.json": candidate, "comparison.json": {"math": {"failed": []}}})
    output = tmp_path / "powerpoint_math.json"
    assert tool.main(["adopt", "--work", str(work), "--output", str(output)]) == 0
    assert "採用した照合データ" in capsys.readouterr().out
    assert json.loads(output.read_text(encoding="utf-8")) == candidate
    existing = {**head, "items": [{**item, "excluded": "old"} for item in items]}
    output.write_text(json.dumps(existing), encoding="utf-8")
    assert tool.adopt(work, output=output, only=["x", "inline:a:b"]) == (output, 3)
    assert json.loads(output.read_text(encoding="utf-8"))["items"] == items
    with pytest.raises(tool.CalibrationError, match="両方にない式"):
        tool.adopt(work, output=output, only=["display:a:b"])
    # 式の一覧の一部の計測値は、全体を置き換えない。
    (work / "conditions.json").write_text(json.dumps({"deck": "math", "subset": True}), encoding="utf-8")
    with pytest.raises(tool.CalibrationError, match="--only"):
        tool.adopt(work, output=output)
    assert tool.adopt(work, output=output, only=["x"]) == (output, 2)
    (work / "conditions.json").write_text(json.dumps({"deck": "math"}), encoding="utf-8")
    (work / "results" / "0001" / "comparison.json").write_text(json.dumps({"math": {"failed": ["x"]}}))
    with pytest.raises(tool.CalibrationError, match="計測できない式"):
        tool.adopt(work, output=output)


def test_calibrate_does_not_adopt_custom_cases(tmp_path):
    tool = calibrate()
    work = results_of(tmp_path, "cases", {"comparison.json": {}})
    with pytest.raises(tool.CalibrationError, match="採用しない"):
        tool.adopt(work, output=tmp_path / "out.json")
    assert not (tmp_path / "out.json").exists()


# ---- 予測と実際を重ねる確認画像 ----


def overlay_run(root, **options) -> Path:
    """synthetic_run に、偽のフォントで計測した layout.json を加えた実行。"""
    run = synthetic_run(root, **options)
    scene = json.loads((run / "deck.json").read_text(encoding="utf-8"))
    fonts = FakeFonts()
    slides = [
        {
            "id": slide["id"],
            "elements": [
                measure.layout_record(measure.layout_element(element, fonts))
                for element in slide["elements"]
                if element["kind"] in ("text", "equation") and element["content"]
            ],
        }
        for slide in scene["slides"]
    ]
    (run / "layout.json").write_text(json.dumps({"schema_version": 1, "slides": slides}), encoding="utf-8")
    return run


def test_layout_overlay_draws_prediction_and_pdf(tmp_path, capsys):
    tool = overlay_tool()
    run = overlay_run(tmp_path / "a")
    output = tmp_path / "overlay"
    assert (
        tool.main(["--run", str(run), "--element", "first/lead", "--element", "formula", "--output", str(output)]) == 0
    )
    printed = capsys.readouterr().out
    data = json.loads((output / "overlay.json").read_text(encoding="utf-8"))
    lead, formula = data["rows"]
    assert (lead["slide_id"], lead["element_id"], formula["slide_id"]) == ("first", "lead", "second")
    # 予測の基線は枠の上端 40 pt から 0.9 em（偽のフォント）、PDF の基線は 64 pt。
    (line,) = lead["differences"]["line_differences"]
    assert line["baseline"] == {"pt": 6.0, "em": 0.3}
    assert line["left"]["pt"] == pytest.approx(4.0)
    assert set(lead["differences"]["ink"]) == {"left", "top", "right", "bottom"}
    assert "1 行目: 基線 +6.00 pt（+0.300 em）" in printed and "インク: 左端" in printed
    for name in ("overlay.png", lead["image"], formula["image"]):
        assert (output / name).is_file()
    with Image.open(output / lead["image"]) as image:
        colors = {color for _count, color in image.convert("RGB").getcolors(maxcolors=1 << 24)}
        size = image.size
    assert any(r > 200 and g < 80 and b < 80 for r, g, b in colors)  # 実際（赤）
    assert any(b > 200 and r < 80 for r, g, b in colors)  # 予測（青）
    assert size[0] >= (300 + 2 * 20) * 8 and size[1] > tool.HEADER + 40 * 8  # 枠と余白 1 em を 8 px/pt で


def test_layout_overlay_remeasures_and_compares_runs(tmp_path, capsys):
    tool = overlay_tool()
    first, second = overlay_run(tmp_path / "a"), overlay_run(tmp_path / "b", first_text="Hello")
    rows = tool.overlay([first, second], ["lead", "note"], tmp_path / "out", remeasure=True, fonts=FakeFonts(), scale=2)
    assert [(row["run_dir"], row["element_id"], row["source"]) for row in rows] == [
        (str(first.resolve()), "lead", "remeasure"),
        (str(first.resolve()), "note", "remeasure"),
        (str(second.resolve()), "lead", "remeasure"),
        (str(second.resolve()), "note", "remeasure"),
    ]
    assert len({row["image"] for row in rows}) == 4
    note = rows[1]["differences"]
    assert note["lines"] == [2, 2] and [value["baseline"]["pt"] for value in note["line_differences"]] == [6.0, 6.0]
    with Image.open(tmp_path / "out" / "overlay.png") as sheet:
        assert sheet.height > 4 * tool.HEADER
    # 実行ディレクトリの中には書かない。見つからない要素だけなら失敗にする。
    assert tool.main(["--run", str(first), "--element", "lead", "--output", str(first / "overlay")]) == 2
    assert "実行ディレクトリの外" in capsys.readouterr().err
    assert tool.main(["--run", str(first), "--element", "none", "--output", str(tmp_path / "x")]) == 2
    assert "要素 none がない" in capsys.readouterr().err


def test_layout_overlay_requires_unique_element(tmp_path):
    tool = overlay_tool()
    scene = {"slides": [{"id": "a", "elements": [{"id": "x"}]}, {"id": "b", "elements": [{"id": "x"}, {"id": "g/y"}]}]}
    with pytest.raises(tool.OverlayError, match="複数のページ"):
        tool.find_element(scene, "x")
    assert tool.find_element(scene, "b/x")[1]["id"] == "b"
    assert tool.find_element(scene, "g/y")[2]["id"] == "g/y" and tool.find_element(scene, "b/g/y")[0] == 1
    assert tool.find_element(scene, "c/x") is None
