"""SLIDEKIT_TEST_POWERPOINT=1 を指定したときだけ、生成の流れ（`pipeline.build`）を実際の PowerPoint で確かめる。

次のことを確かめる。
- PowerPoint の変換は 1 つのセッションで、互換画像を付ける資料でも標準表示 1 回と互換表示 1 回だけであること。
  標準表示は最終形の PPTX（互換画像だけが仮の画像）を変換し、最終 PPTX と、スライドの XML と関係が一致すること。
- 数式のページだけの確認用複製（章 `p14:sectionLst` のある資料を含む）を、PowerPoint が開いて変換できること。
- 互換画像の切り出しの範囲が数式のインク（伸びた括弧などのはみ出しを含む）と枠線を含むこと。
  はみ出しが互換画像の図形の範囲を超えるときは、切り詰めて `fallback_crop_clipped` で知らせること。
- 題名の数式にも互換画像が付き、互換表示の PDF で欠けないこと。

修復が要らないことは、ウィンドウを表示しない変換では PowerPoint が修復を求めないため、次の 4 つの組で判定する。
1. 変換が制限時間内に成功する。
2. 変換の後に排他と残存記録がない（`renderer status` が idle）。PowerPoint を起動したセッションなら、PowerPoint も終了している。
3. PDF に内容が描かれる。互換画像は代わりの枠でなく PPTX の画像の図形と同じ範囲に描かれる。
4. PPTX の検査（`validate_pptx`）に、部品と関係の欠落などの診断がない。
資料は作成 API で作り、フィクスチャ windows_fonts_environment で Windows の既定フォントを使う。直列に実行する（`-p no:xdist`）。
"""

import json
import os
from pathlib import Path
import time
import uuid

from lxml import etree
import pytest

from pdf_tools import image_boxes, open_pdf, render_region

pytestmark = [
    pytest.mark.powerpoint,
    pytest.mark.skipif(
        os.environ.get("SLIDEKIT_TEST_POWERPOINT") != "1", reason="実機 PowerPoint を使う任意の検査である。"
    ),
]

DECK = """schema_version = 1
entrypoint = "slides.py:build"
size = "16:9"
[metadata]
title = "生成の流れの検査"
author = ""
last_modified_by = ""
[math]
fallback = "required"
"""

# 3 ページのうち 2 ページ目だけに数式がある資料。
THREE_PAGES = """from slidekit import Deck, Math, Rect

def build(context):
    deck = Deck(context)
    for number, page in enumerate(("first", "second", "third"), 1):
        slide = deck.slide(id=page, kind="content", title=f"Page {number}", notes="説明します。")
        slide.text(id="body", text="Process checked", area=Rect(60, 110, 400, 60))
        if page == "second":
            slide.text(id="equation", text=[Math("x_i^2 + \\\\frac{a}{b}")], area=Rect(60, 200, 300, 80))
    return deck
"""

# 枠線付きの低い文字枠に伸びる括弧の式（下へのはみ出しが MAX_MARGIN_PT を超える low と、超えない mid）、
# 枠線付きの楕円に数式を入れた資料。計測の予測では、low は 24.8 pt、mid は 8.8 pt 下へはみ出す。
CROPS = """from slidekit import Deck, Math, Rect

def build(context):
    deck = Deck(context)
    slide = deck.slide(id="crops", kind="content", title="Crops", notes="説明します。")
    slide.box(id="low", text=[Math("\\\\left(\\\\frac{\\\\sqrt{x}}{y}\\\\right)")], area=Rect(60, 200, 200, 24),
              line_color="1F2937", line_width_pt=2, padding_pt=0, word_wrap=False)
    slide.box(id="mid", text=[Math("\\\\left(\\\\frac{\\\\sqrt{x}}{y}\\\\right)")], area=Rect(60, 320, 200, 40),
              line_color="1F2937", line_width_pt=2, padding_pt=0, word_wrap=False)
    slide.box(id="circle", text=[Math("x^2+y^2")], shape="ellipse", area=Rect(400, 160, 200, 120),
              line_color="1F2937", line_width_pt=3)
    return deck
"""

# 題名に数式がある資料。
TITLE_MATH = """from slidekit import Deck, Math, Rect, Text

def build(context):
    deck = Deck(context)
    slide = deck.slide(id="area", kind="content", title=[Text("面積 "), Math("\\\\pi r^2")], notes="説明します。")
    slide.text(id="body", text="Process checked", area=Rect(60, 160, 400, 60))
    return deck
"""

P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p14": P14,
}
SECTION_EXTENSION = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
# PPTX の検査の診断のうち、部品と関係の欠落・重複と XML の誤り（修復が要る PPTX の兆候。判定 4）。
PACKAGE_CODES = frozenset(
    {
        "package_duplicate",
        "invalid_xml",
        "relationship_missing",
        "package_reference_missing",
        "presentation_missing",
        "slide_missing",
    }
)


def build_deck(root, write_deck, slides, *, allowed=(), **options):
    """資料を書いて生成し、（結果、実行ディレクトリ、生成記録）を返す。options は `pipeline.build` に渡す。

    生成は完了し、allowed（意図したはみ出しなど）のほかにエラーの診断がないことを確かめる。
    """
    from slidekit import pipeline

    write_deck(root, deck=DECK, slides=slides)
    result = pipeline.build(root / "deck.toml", **options)
    assert result["status"] == "built", result
    errors = [issue for issue in result["issues"] if issue["severity"] == "error" and issue["code"] not in allowed]
    assert errors == [], errors
    run = Path(result["build"])
    return result, run, json.loads((run / "manifest.json").read_text(encoding="utf-8"))


def contains(outer, inner, tolerance=0.3):
    """矩形 inner が outer に tolerance（pt）の余裕で含まれるか。矩形は（x0, y0, x1, y1）の列か `Box`。"""
    (left, top, right, bottom), (x0, y0, x1, y1) = outer, inner
    return x0 >= left - tolerance and y0 >= top - tolerance and x1 <= right + tolerance and y1 <= bottom + tolerance


def dark_pixels(page, rectangle):
    """PDF のページ（`slidekit.pdf` のページ）の範囲を 2 倍で描き、暗い画素（明度 128 未満）の数を返す。

    範囲はページの中に切り詰める（ページの外の画素は含めない）。
    """
    from slidekit.pdf import Box

    area = Box.of(rectangle) & page.rect
    if area.is_empty:
        return 0
    samples = render_region(page, 2, area).tobytes()
    return sum(1 for index in range(0, len(samples), 3) if sum(samples[index : index + 3]) < 3 * 128)


def package_problems(path, scene, project):
    """PPTX の検査のうち、部品と関係の欠落などの診断（判定 4）。scene はその PPTX に残したページだけの場面にする。"""
    from slidekit.pptx_validation import validate_pptx

    return [
        issue
        for issue in validate_pptx(path, scene, project, fallback_expected=False)
        if issue["code"] in PACKAGE_CODES
    ]


def assert_idle(project, operations):
    """変換の後に排他と残存記録がなく、PowerPoint を起動したセッションなら PowerPoint も終了している（判定 2）。

    operations はセッションの記録（`operations.sessions.powerpoint` か、セッションの `operations`）。
    PowerPoint は `Quit` の後、終了するまでに数秒かかる（実際に計測すると 0.5〜2.6 秒）ので、最大 30 秒待つ。
    """
    from slidekit import renderers

    status = renderers.renderer_status("powerpoint", project.environment)
    assert status["status"] == "idle", status
    if operations["powerpoint_started"]:
        deadline = time.monotonic() + 30
        while status["powerpoint_running"] and time.monotonic() < deadline:
            time.sleep(1)
            status = renderers.renderer_status("powerpoint", project.environment)
        assert status["powerpoint_running"] is False, status


def drawn_at_frame(boxes, crop):
    """互換画像が代わりの枠でなく、PPTX の画像の図形と同じ範囲（frame_pt）に描かれ、切り出しの範囲を覆うか（判定 3）。

    boxes は PDF のページに描かれた画像の範囲（`image_boxes`）。
    """
    return any(
        max(abs(a - b) for a, b in zip(box, crop["frame_pt"])) <= 0.6 and contains(box, crop["crop_pt"], 0.5)
        for box in boxes
    )


def test_real_build_uses_one_session_and_final_form(tmp_path, write_deck, windows_fonts_environment):
    from slidekit import renderers
    from slidekit.config import load_project

    queries = renderers._cached_windows_context.cache_info().misses
    result, run, record = build_deck(tmp_path / "deck", write_deck, THREE_PAGES)
    queried = renderers._cached_windows_context.cache_info().misses - queries
    operations = record["operations"]
    session = operations["sessions"]["powerpoint"]
    # PowerShell の起動は作業プロセスの 1 回と、Windows への問い合わせ（プロセス内で最初の 1 回だけで、2 回目以降は記憶を使う）。
    # 問い合わせはセッションの中でだけ起こる（セッションの外で PowerShell を起動しない）。変換は標準表示と互換表示の 2 回。
    assert queried <= 1 and session["powershell_launches"] == 1 + queried and session["renders"] == 2
    assert [operations[f"render:powerpoint:{variant}"]["label"] for variant in ("standard", "fallback")] == [
        "standard",
        "fallback",
    ]
    assert operations["render:powerpoint:standard"]["input_sha256"] == operations["final_form"]["sha256"]
    assert operations["final_form"]["structure"]["matched"] is True
    assert not (run / "intermediate/native.pdf").exists()
    # 閉じた後に排他を取ることができ、PowerPoint を起動していたら終了している。
    assert_idle(load_project(tmp_path / "deck"), session)
    # 互換画像は標準表示の PDF から切り出した画像で、その PDF の数式の範囲に暗い画素がある（数式が描かれている）。
    (crop,) = operations["math_fallback"]["crops"]
    with open_pdf(run / "renders/powerpoint/standard/deck.pdf") as document:
        assert dark_pixels(document[1], crop["crop_pt"]) > 20


def test_real_fallback_clone_with_math_slides_only_opens(tmp_path, write_deck, windows_fonts_environment):
    from slidekit import fallback, renderers
    from slidekit.config import load_project
    from slidekit.ooxml import read_package, write_package

    result, run, record = build_deck(tmp_path / "deck", write_deck, THREE_PAGES, keep_intermediate=True)
    project = load_project(tmp_path / "deck")
    scene = json.loads((run / "deck.json").read_text(encoding="utf-8"))
    # 確認用複製に残る、数式のページだけの場面。
    math_scene = {**scene, "slides": [page for page in scene["slides"] if page["id"] == "second"]}
    operations = record["operations"]
    assert operations["fallback_clone"]["kept_slides"] == [{"number": 1, "id": "second"}]
    (crop,) = operations["math_fallback"]["crops"]
    # 判定 2：生成の変換の後に排他と残存記録がなく、PowerPoint を起動していたら終了している。
    assert_idle(project, operations["sessions"]["powerpoint"])
    # 判定 4：最終 PPTX（生成の中の検査）と確認用複製に、部品と関係の欠落がない。
    assert not [issue for issue in result["issues"] if issue["code"] in PACKAGE_CODES | {"pdf_page_count"}]
    assert package_problems(run / "intermediate/fallback.pptx", math_scene, project) == []
    # 判定 1・3：数式のページだけが描かれ、互換画像は PPTX の画像の図形と同じ範囲に描かれて、数式の範囲に暗い画素がある。
    fallback_pdf = run / "renders/powerpoint/fallback/deck.pdf"
    with open_pdf(fallback_pdf) as document:
        assert len(document) == 1
        assert drawn_at_frame(image_boxes(fallback_pdf, 0), crop), (image_boxes(fallback_pdf, 0), crop)
        assert dark_pixels(document[0], crop["crop_pt"]) > 20
    # 章のある資料：最終 PPTX に 2 つの章（1・2 ページと 3 ページ）を加え、数式のページだけの複製を作って変換する。
    parts = read_package(run / "deck.pptx")
    presentation = etree.fromstring(parts["ppt/presentation.xml"])
    identifiers = [node.get("id") for node in presentation.findall("p:sldIdLst/p:sldId", NS)]
    extensions = presentation.find("p:extLst", NS)
    if extensions is None:
        extensions = etree.SubElement(presentation, f"{{{NS['p']}}}extLst")
    extension = etree.SubElement(extensions, f"{{{NS['p']}}}ext", uri=SECTION_EXTENSION)
    sections = etree.SubElement(extension, f"{{{P14}}}sectionLst", nsmap={"p14": P14})
    for name, members in (("前半", identifiers[:2]), ("後半", identifiers[2:])):
        section = etree.SubElement(sections, f"{{{P14}}}section", name=name, id="{" + str(uuid.uuid4()).upper() + "}")
        listing = etree.SubElement(section, f"{{{P14}}}sldIdLst")
        for identifier in members:
            etree.SubElement(listing, f"{{{P14}}}sldId", id=identifier)
    parts["ppt/presentation.xml"] = etree.tostring(
        presentation, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    sectioned = tmp_path / "sectioned.pptx"
    write_package(parts, sectioned)
    assert package_problems(sectioned, scene, project) == []
    clone = tmp_path / "sectioned-fallback.pptx"
    prepared = fallback.prepare_fallback(sectioned, clone, math_only=True)
    assert prepared["kept_slides"] == [{"number": 1, "id": "second"}]
    kept = etree.fromstring(read_package(clone)["ppt/presentation.xml"])
    assert [node.get("id") for node in kept.iter(f"{{{P14}}}sldId")] == [identifiers[1]]
    assert package_problems(clone, math_scene, project) == []
    pdf = tmp_path / "sectioned-fallback.pdf"
    with renderers.open_session("powerpoint", {**project.environment, "fonts": {}}) as session:
        session.render(clone, pdf, label="fallback")
    assert_idle(project, session.operations)
    with open_pdf(pdf) as document:
        assert len(document) == 1
        assert drawn_at_frame(image_boxes(pdf, 0), crop), (image_boxes(pdf, 0), crop)
        assert dark_pixels(document[0], crop["crop_pt"]) > 20


def test_real_fallback_crop_keeps_overflow_and_outline(tmp_path, write_deck, windows_fonts_environment):
    from slidekit import fallback
    from slidekit.pdf_layout import read_slide

    # 低い文字枠の式は枠からはみ出す（意図した配置のエラー）。
    result, run, record = build_deck(
        tmp_path / "deck", write_deck, CROPS, allowed=("text_overflow", "pdf_element_text_overflow")
    )
    crops = {crop["element_id"]: crop for crop in record["operations"]["math_fallback"]["crops"]}
    assert set(crops) == {"low", "mid", "circle"}
    scene = json.loads((run / "deck.json").read_text(encoding="utf-8"))
    with open_pdf(run / "renders/powerpoint/standard/deck.pdf") as document:
        reading = read_slide(document[0], scene["slides"][0])
    step = 1 / record["operations"]["math_fallback"]["image_scale"]
    for element_id, crop in crops.items():
        assert crop["outline_pt"] > 0
        x0, y0, x1, y1 = crop["rectangle_pt"]
        half = crop["outline_pt"] / 2
        # 枠線の外側の半分を含む。
        assert contains(crop["crop_pt"], [x0 - half, y0 - half, x1 + half, y1 + half], tolerance=step)
        # 数式のインク（字形と、伸びた括弧・根号・分数の線）を、互換画像の図形の範囲の中で含む。
        ink, frame = reading.elements[element_id].ink, crop["frame_pt"]
        assert ink is not None
        inside = [max(ink[0], frame[0]), max(ink[1], frame[1]), min(ink[2], frame[2]), min(ink[3], frame[3])]
        assert contains(crop["crop_pt"], inside, tolerance=step), (element_id, ink, crop)
        # 互換画像の図形（切り出しの範囲を含む）は、数式の図形の矩形から MAX_MARGIN_PT までに収まる。
        assert contains(
            [
                x0 - fallback.MAX_MARGIN_PT,
                y0 - fallback.MAX_MARGIN_PT,
                x1 + fallback.MAX_MARGIN_PT,
                y1 + fallback.MAX_MARGIN_PT,
            ],
            crop["frame_pt"],
            tolerance=step,
        )
    # 伸びた括弧は枠の外へはみ出す。図形の矩形どおりの切り出しでは欠ける。
    for element_id in ("low", "mid"):
        ink, rectangle = reading.elements[element_id].ink, crops[element_id]["rectangle_pt"]
        assert ink[1] < rectangle[1] - 1 or ink[3] > rectangle[3] + 1, (element_id, ink)
    # はみ出しが MAX_MARGIN_PT を超える low だけを切り詰め、警告で知らせる（知らせずに欠けることはない）。
    assert [crops[key]["clipped"] for key in ("low", "mid", "circle")] == [True, False, False]
    warned = {issue.get("element_id") for issue in result["issues"] if issue["code"] == "fallback_crop_clipped"}
    assert warned == {"low"}
    # 互換表示では、はみ出した部分も画像に含まれて描かれる。mid は標準表示と同じだけの暗い画素がある。
    # 枠線の帯は除く（互換表示では枠線も画像の一部として描かれ、縁の明暗が少し違う）。
    for element_id in ("low", "mid"):
        ink, rectangle = reading.elements[element_id].ink, crops[element_id]["rectangle_pt"]
        band = crops[element_id]["outline_pt"] / 2 + 1
        outside = [[ink[0], ink[1], ink[2], rectangle[1] - band], [ink[0], rectangle[3] + band, ink[2], ink[3]]]
        areas = [area for area in outside if area[3] - area[1] >= 1]
        with open_pdf(run / "renders/powerpoint/standard/deck.pdf") as document:
            standard = sum(dark_pixels(document[0], area) for area in areas)
        with open_pdf(run / "renders/powerpoint/fallback/deck.pdf") as document:
            shown = sum(dark_pixels(document[0], area) for area in areas)
        assert shown > 0, element_id
        if element_id == "mid":
            assert standard > 20 and abs(shown - standard) <= max(20, standard // 10), (standard, shown)


def test_real_powerpoint_title_math_has_fallback(tmp_path, write_deck, windows_fonts_environment):
    result, run, record = build_deck(tmp_path / "deck", write_deck, TITLE_MATH)
    crops = [crop for crop in record["operations"]["math_fallback"]["crops"] if crop["element_id"] == "_theme/title"]
    assert len(crops) == 1
    (crop,) = crops
    assert crop["clipped"] is False
    with open_pdf(run / "renders/powerpoint/standard/deck.pdf") as document:
        standard = dark_pixels(document[0], crop["crop_pt"])
    fallback_pdf = run / "renders/powerpoint/fallback/deck.pdf"
    images = image_boxes(fallback_pdf, 0)
    # 題名の互換画像が、切り出しの範囲を覆う位置に描かれる。
    assert any(contains(bbox, crop["crop_pt"], tolerance=0.5) for bbox in images), (images, crop)
    with open_pdf(fallback_pdf) as document:
        shown = dark_pixels(document[0], crop["crop_pt"])
    # 互換表示でも題名の文字と数式が欠けない（暗い画素の数が標準表示と大きく違わない）。
    assert standard > 20 and abs(shown - standard) <= max(20, standard // 10), (standard, shown)
    assert not [
        issue
        for issue in result["issues"]
        if issue["code"] == "pdf_text_missing" and issue["scope"] == "pdf:powerpoint:standard"
    ]
