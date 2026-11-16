"""SLIDEKIT_TEST_POWERPOINT=1 のときだけ、校正用の資料の折り返しのページを実際の PowerPoint で変換し、計測と比べる。

直列に実行する（`-p no:xdist`）。資料は一時ディレクトリへ複製してから構築し、PPTX と PDF も一時ディレクトリに作る。
比べ方は test_layout_fixture.compare_case と同じで、ここでは行数・インク・文字ごとの書体を確かめる（照合データの試験は全項目を比べる）。
書き出しと計測に同じ FontSource を渡すので、幅の曖昧な記号と欧文の書体にない文字は、計測と同じ書体の区切り（forced）で書かれる。
"""

import os
import shutil

import pytest

from slidekit import renderers
from slidekit.config import load_project
from slidekit.pdf_layout import read_pdf
from slidekit.pipeline import construct
from slidekit.writer import write_pptx
from test_layout_fixture import DECK, WRAP_PAGES, compare_case, load_tool

pytestmark = [
    pytest.mark.powerpoint,
    pytest.mark.skipif(
        os.environ.get("SLIDEKIT_TEST_POWERPOINT") != "1", reason="実機 PowerPoint を使う任意の検査である。"
    ),
]


def test_real_powerpoint_wraps_like_measurement(windows_fonts_environment, tmp_path):
    tool = load_tool()
    deck = shutil.copytree(DECK, tmp_path / "deck")
    project = load_project(deck / "deck.toml")
    fonts = project.fonts
    try:
        scene = construct(project)
        scene["slides"] = [slide for slide in scene["slides"] if slide["id"] in WRAP_PAGES]
        assert tuple(slide["id"] for slide in scene["slides"]) == WRAP_PAGES
        pptx, pdf = tmp_path / "wrap.pptx", tmp_path / "wrap.pdf"
        write_pptx(scene, project, pptx, fonts=fonts)
        renderers.render_pdf(pptx, pdf, project.environment)
        failures, count = [], 0
        for slide, reading in zip(scene["slides"], read_pdf(pdf, scene)):
            for element in tool.text_elements(slide):
                rendered = reading.elements.get(element["id"])
                if rendered is None:
                    failures.append(f"{slide['id']}/{element['id']}: PDF から読めない")
                    continue
                count += 1
                case = {
                    "slide_id": slide["id"],
                    "element_id": element["id"],
                    "tags": tool.case_tags(slide, element),
                    "element": element,
                    "actual": tool.actual_record(rendered),
                }
                failures += compare_case(case, fonts, checks=("lines", "ink", "fonts"))
    finally:
        fonts.close()
    assert count > 40
    assert not failures, f"PowerPoint の折り返しが計測と違う（{len(failures)} 件）:\n" + "\n".join(failures)
