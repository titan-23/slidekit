"""この例の雛形 template.pptx を、python-pptx の既定の雛形から作る。

slidekit は雛形のテーマ・マスター・レイアウトを引き継ぎ、雛形のページは外す。
ここでは配色とフォントをテーマに書き込み、外れることを示す見本のページを 1 枚入れる。
文書情報の日時と ZIP の時刻を固定するので、同じ版の python-pptx からは同じ内容のファイルになる（ZIP の各部品の中身が同じ）。

    python make_template.py            # template.pptx を作り直す
    python make_template.py OUTPUT     # 別の場所へ書き出す
"""

from __future__ import annotations

from datetime import datetime, timezone
import io
from pathlib import Path
import sys
import tomllib
import zipfile

from lxml import etree
from pptx import Presentation
from pptx.util import Pt

HERE = Path(__file__).resolve().parent
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
FIXED_TIME = datetime(2026, 1, 1, tzinfo=timezone.utc)
THEME_NAME = "slidekit example"
# テーマの色の枠と、theme.toml の色の名前の対応。
COLOR_SLOTS = {
    "dk1": "text",
    "lt1": "background",
    "dk2": "primary",
    "lt2": "surface",
    "accent1": "primary",
    "accent2": "accent",
    "accent3": "muted",
    "accent4": "line",
}


def _theme(xml: bytes, colors: dict) -> bytes:
    """テーマの名前・配色・フォントを、この例の theme.toml の値へ書き換える。"""
    root = etree.fromstring(xml)
    root.set("name", THEME_NAME)
    scheme = root.find(f".//{{{A}}}clrScheme")
    scheme.set("name", THEME_NAME)
    for slot, name in COLOR_SLOTS.items():
        node = scheme.find(f"{{{A}}}{slot}")
        for child in list(node):
            node.remove(child)
        etree.SubElement(node, f"{{{A}}}srgbClr", val=colors[name])
    fonts = root.find(f".//{{{A}}}fontScheme")
    fonts.set("name", THEME_NAME)
    for kind in ("majorFont", "minorFont"):
        font = fonts.find(f"{{{A}}}{kind}")
        font.find(f"{{{A}}}latin").set("typeface", "Arial")
        font.find(f"{{{A}}}ea").set("typeface", "Yu Gothic UI")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def build_template(output: Path) -> Path:
    """雛形を output に書き、そのパスを返す。"""
    colors = tomllib.loads((HERE / "theme.toml").read_text(encoding="utf-8"))["colors"]
    presentation = Presentation()
    presentation.slide_width, presentation.slide_height = Pt(720), Pt(540)
    sample = presentation.slides.add_slide(presentation.slide_layouts[0])
    sample.shapes.title.text = "雛形の見本ページ"
    sample.placeholders[1].text = "slidekit は生成時にこのページを外す"
    properties = presentation.core_properties
    properties.title = "slidekit example template"
    properties.author = properties.last_modified_by = "slidekit example"
    properties.created = properties.modified = FIXED_TIME
    properties.revision = 1
    buffer = io.BytesIO()
    presentation.save(buffer)
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(buffer) as source, zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "ppt/theme/theme1.xml":
                data = _theme(data, colors)
            fixed = zipfile.ZipInfo(info.filename, date_time=FIXED_TIME.timetuple()[:6])
            fixed.compress_type = zipfile.ZIP_DEFLATED
            target.writestr(fixed, data)
    return output


if __name__ == "__main__":
    print(build_template(Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "template.pptx"))
