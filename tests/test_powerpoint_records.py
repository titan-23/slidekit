"""SLIDEKIT_TEST_POWERPOINT=1 を指定したときだけ、表示確認の証拠と引き継ぎを実際の PowerPoint で確かめる。

次のことを確かめる。
- 同じ部品のページを 2 回変換すると、部品の証拠が一致し、画像の差（`evidence.render_difference`）が雑音の範囲に収まること。
- 標準表示の証拠に使う数式の本体（`mc:Choice`）が、互換画像を付けない中間 PPTX の数式と同じに描かれること。
- 先頭にページを加えて再生成すると、番号の桁数が同じページの表示の確認が、番号の領域を除いた比較で自動で引き継がれること。

資料は作成 API で作り、フィクスチャ windows_fonts_environment で Windows の既定フォントを使う。直列に実行する（`-p no:xdist`）。
"""

import json
import os
from pathlib import Path

from PIL import Image
import pytest

from pdf_tools import open_pdf

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
title = "表示確認の証拠の検査"
author = ""
last_modified_by = ""
[math]
fallback = "required"
"""

# 文字・数式・図形・画像を含む 3 ページ（content 種別で、ページ番号の飾りがある）。題名はページの位置によらない。
PAGES = """from slidekit import Deck, Math, Rect

PAGES = {PAGES}


def build(context):
    deck = Deck(context)
    for page in PAGES:
        slide = deck.slide(id=page, kind="content", title=page.title(), notes="説明します。")
        if page == "text":
            slide.text(id="body", text="Process checked line by line", area=Rect(60, 110, 500, 60))
        elif page == "formula":
            slide.text(id="body", text="The sum is", area=Rect(60, 110, 300, 40))
            slide.text(id="equation", text=[Math("\\\\sum_{{i=1}}^{{n}} x_i^2 + \\\\frac{{a}}{{b}}")], area=Rect(60, 170, 400, 80))
        elif page == "figure":
            slide.box(id="frame", text="Step", area=Rect(60, 110, 200, 80), line_color="1F2937", line_width_pt=2)
            slide.box(id="circle", text="Goal", shape="ellipse", area=Rect(400, 110, 200, 80), line_color="1F2937")
            slide.image(id="picture", path="picture.png", area=Rect(60, 260, 160, 120))
        else:
            slide.text(id="body", text="Inserted page", area=Rect(60, 110, 400, 60))
    return deck
"""


def write(root, write_deck, pages):
    root = write_deck(root, deck=DECK, slides=PAGES.format(PAGES=repr(tuple(pages))))
    picture = root / "picture.png"
    if not picture.exists():
        image = Image.new("RGB", (160, 120), "white")
        for x in range(160):
            for y in range(40, 80):
                image.putpixel((x, y), (30 + x, 90, 160))
        image.save(picture)
    return root


def build(root, **options):
    from slidekit import pipeline

    result = pipeline.build(root / "deck.toml", **options)
    assert result["status"] == "built" and result["counts"]["error"] == 0, result
    run = Path(result["build"])
    return result, run, json.loads((run / "manifest.json").read_text(encoding="utf-8"))


def complete(run):
    """全ページの表示・一覧画像・内容の確認と、警告の許容を記録する（試験の記録）。"""
    from slidekit import records

    records.record_review(run, "visual,content", "all", "検査担当", "表示と本文を確認", variant="all")
    records.record_review(run, "overview", None, "検査担当", "一覧を確認")
    for warning in records.check_build(run)["warnings"]:
        if not warning.get("declared"):
            records.accept_warning(run, warning["diagnostic_id"], "検査担当", "試験の資料の表示条件による")


def test_real_renders_of_same_parts_are_within_noise(tmp_path, write_deck, windows_fonts_environment):
    from slidekit.evidence import NOISE_PIXELS, render_difference

    root = write(tmp_path / "deck", write_deck, ("text", "formula", "figure"))
    _, first, before = build(root, carry_from="none")
    _, second, after = build(root, carry_from="none")
    assert [page["page_parts"] for page in before["slides"]] == [page["page_parts"] for page in after["slides"]]
    compared = 0
    for old, new in zip(before["slides"], after["slides"]):
        assert sorted(old["visuals"]) == sorted(new["visuals"])
        for key, visual in old["visuals"].items():
            difference = render_difference(first / visual["path"], second / new["visuals"][key]["path"])
            assert difference["strong_pixels"] is not None and difference["strong_pixels"] <= NOISE_PIXELS, (
                old["id"],
                key,
                difference,
            )
            compared += 1
    # 標準表示 3 ページと、数式のページの互換表示。
    assert compared == 4


def test_real_choice_equals_unwrapped_math(tmp_path, write_deck, windows_fonts_environment):
    """標準表示（最終形の PPTX の mc:Choice）と、互換画像を付けない中間 PPTX の数式の描画が同じであること。

    改修後の生成は中間 PPTX を変換しないので、試験の中で中間 PPTX を同じ PowerPoint で変換して比べる。
    """
    from slidekit import renderers
    from slidekit.config import load_project
    from slidekit.evidence import NOISE_PIXELS, RENDER_SCALE, render_difference, render_page_image

    root = write(tmp_path / "deck", write_deck, ("text", "formula", "figure"))
    _, run, record = build(root, carry_from="none", keep_intermediate=True)
    native = run / "intermediate/native.pptx"
    assert native.is_file()
    output = tmp_path / "native" / "deck.pdf"
    output.parent.mkdir()
    with renderers.open_session(load_project(root).environment) as session:
        session.render(native, output, label="native")
    with open_pdf(output) as document:
        assert len(document) == len(record["slides"])
    for index, page in enumerate(record["slides"]):
        picture = render_page_image(
            str(output),
            index,
            str(tmp_path / "native" / f"slide-{page['number']:03d}.png"),
            RENDER_SCALE,
            thumbnail=False,
        )
        standard = run / page["visuals"]["powerpoint:standard"]["path"]
        difference = render_difference(standard, picture["path"])
        assert difference["strong_pixels"] is not None and difference["strong_pixels"] <= NOISE_PIXELS, (
            page["id"],
            difference,
        )


def test_real_page_insert_carries_with_footer_mask(tmp_path, write_deck, windows_fonts_environment):
    from slidekit import records

    root = write(tmp_path / "deck", write_deck, ("text", "formula", "figure"))
    _, first, _ = build(root, carry_from="none")
    complete(first)
    write(root, write_deck, ("inserted", "text", "formula", "figure"))
    result, second, _ = build(root)
    carry = result["carry"]
    assert carry["status"] == "carried" and sorted(carry["renumbered"]) == ["figure", "formula", "text"]
    copies = json.loads((second / "review.json").read_text(encoding="utf-8"))["records"]
    visuals = [record for record in copies if record["kind"] == "visual"]
    # 標準表示の 3 ページと、数式のページの互換表示。どれも番号の領域を除いて比べ、差は雑音の範囲である。
    assert sorted((record["slide_id"], record["variant"]) for record in visuals) == [
        ("figure", "standard"),
        ("formula", "fallback"),
        ("formula", "standard"),
        ("text", "standard"),
    ]
    assert all(record["render_check"]["masked"] for record in visuals if record["variant"] == "standard")
    status = records.review_status(second)
    pending = {
        (group["kind"], group["variant"], tuple(item["slide_id"] for item in group["pending"]))
        for group in status["reviews"]
        if group["pending"]
    }
    assert pending == {
        ("visual", "standard", ("inserted",)),
        ("content", None, ("inserted",)),
        ("overview", "standard", (None,)),
    }
