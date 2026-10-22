"""数式の計測の照合データ（tests/data/powerpoint_math.json）を PowerPoint で作る（開発用。試験からは呼ばない）。

使い方（PowerPoint を使う。PowerPoint を使う処理は一度に一つだけ実行する）:
  python tests/tools/powerpoint_math_data.py --output PATH [--work DIR]
  python tests/tools/powerpoint_math_data.py --output PATH --work DIR --measure-only   # 変換済みの PDF を計測し直す

- 照合データの式（EXPRESSIONS）を、20 pt の文中の数式（Arial の [ と ] で挟む）と独立した数式（equation()、左揃え）にした
  資料を作業ディレクトリに作る（deck.toml・slides.py）。slidekit の書き出しで PPTX を作り、PowerPoint で PDF にする。
- 文中の数式の送り幅は、PDF の [ の送り幅の右端と ] の原点の差。描画範囲は目印の間を 8 px/pt で画像化して求め、
  [ の右端と基線を原点にした (左, 右, 基線より上, 基線より下) にする。枠の上端より上の描画は、枠の中の描画に続く分だけを
  含める（上の行の式が枠の下へはみ出した描画を拾わない）。
- 独立した数式は、枠の範囲を 8 px/pt で画像化した描画範囲の幅と高さ。
- 数式の書体は資料の実行環境で解決し、ファイル名・書体番号・版・sha256 を書く。照合データにパスは書かない。
- 書き出しに失敗する式（未対応の命令など）は、照合データに入れず、標準エラーに示す。

作業ディレクトリの既定は一時フォルダーの slidekit-powerpoint-math/（--measure-only で使い回す。終わったら消してよい）。
作り直す時期：PowerPoint・数式の書体の版を変えたとき、数式の組み方を変えたとき。差分を確かめてから入れ替える。
"""

from __future__ import annotations

import argparse
from datetime import date
import json
from pathlib import Path
import sys
import tempfile

SCHEMA_VERSION = 1
SIZE_PT = 20.0
SCALE = 8  # 描画範囲を求める画像の倍率（px/pt）
DEFAULT_WORK = Path(tempfile.gettempdir()) / "slidekit-powerpoint-math"
MARK_FONT = "Arial"
COLUMNS = (24.0, 336.0, 648.0)  # 枠の左端（16:9 の資料、幅 300 pt）
INLINE_ROWS, INLINE_HEIGHT = 8, 56.0
DISPLAY_ROWS, DISPLAY_HEIGHT = 6, 80.0
FRAME_WIDTH = 288.0
ROW_GAP = 8.0  # 上下の枠の間（pt）

# 一般的な記法だけを使う。
EXPRESSIONS = (
    # 文字
    "x",
    "abc",
    "0123",
    "3.14",
    r"\alpha",
    r"\Theta",
    r"\epsilon",
    r"\phi",
    r"\ell",
    # 演算
    "a+b",
    "a-b",
    "-a",
    r"a\times b",
    r"a\cdot b",
    "a=b",
    r"a\le b",
    r"a\to b",
    "a,b",
    "a;b",
    "a:b",
    "n!",
    "f'",
    "f''",
    # 添字
    "x_i",
    "x^2",
    "x_i^2",
    "x_{ij}",
    "x^{n+1}",
    "e^{-x}",
    r"T_{\text{max}}",
    # 分数
    r"\frac{a}{b}",
    r"\frac{1}{2}",
    r"\frac{n(n+1)}{2}",
    "a/b",
    r"\binom{n}{k}",
    # 根号
    r"\sqrt{n}",
    r"\sqrt{x+y}",
    r"\sqrt[3]{x}",
    r"\sqrt{\frac{a}{b}}",
    # 括弧
    "(x)",
    "f(x)",
    "O(n)",
    r"O(\sqrt{n})",
    r"O(n\log n)",
    r"\left(\frac{a}{b}\right)",
    "[0,n)",
    r"\{x\}",
    "|x|",
    r"\lfloor x\rfloor",
    r"\left\{x \middle| x>0\right\}",
    r"\langle u,v\rangle",
    # 関数
    r"\log n",
    r"\log(n)",
    r"\log_2 n",
    r"\sin x",
    r"\max(a,b)",
    r"\min_i x_i",
    r"\max\limits_{i} x_i",
    r"\lim_{n\to\infty} a_n",
    r"\operatorname{rank}(A)",
    # 大型演算子
    r"\sum_{i=1}^{n} a_i",
    r"\sum\limits_{i=1}^{n} a_i",
    r"\prod_i x_i",
    r"\int_0^1 f(x)\,dx",
    r"\bigcup_i S_i",
    # 記号付き
    r"\hat{x}",
    r"\tilde{x}",
    r"\bar{x}",
    r"\overline{xy}",
    r"\vec{v}",
    r"\dot{x}",
    r"\widehat{xy}",
    # 書体
    r"\mathrm{abc}",
    r"\mathbf{v}",
    r"\boldsymbol{\alpha}",
    r"\mathcal{O}",
    r"\mathbb{R}",
    r"\mathfrak{g}",
    r"\mathsf{A}",
    r"\mathtt{A}",
    # 文章
    r"\text{abc}",
    r"\text{最大}",
    r"x\text{ 個}",
    # 空白
    r"a\,b",
    r"a\:b",
    r"a\;b",
    r"a\quad b",
    r"a\qquad b",
    r"a\ b",
    # その他
    r"\not\in",
    r"x\in S",
    r"A\subseteq B",
    r"\overset{?}{=}",
    r"\cdots",
    r"\ldots",
    r"a \bmod b",
    r"a \pmod{n}",
    # 前後の空きが記号の種類で決まる式（… と ⋯ の前後、文字の後のピリオド）
    r"a_1, \ldots, a_n",
    r"a_1 + \cdots + a_n",
    "p.q",
    # 添字の中の文章、大型演算子の引数と添字の位置、伸びる括弧の大きさ（照合の規則を確かめた式）
    r"x_{\text{最大}}",
    r"\sum_i a_i b_i",
    r"\int f\,dx",
    r"\sum^n a_i",
    r"\prod_{i=1}^{n} x_i",
    r"\oint_C f",
    r"\left(x^2\right)",
    r"\left(\frac{a^2}{b}\right)",
    # 添字の土台が並びのとき（位置は最後の原子で決まる）、関数名の後の \{、数字の間のカンマ（太字の数字は数を作らない）、
    # 関係の記号の \bot、根号と関数の適用の後の句読点（照合データにない式で PowerPoint との差が見つかった規則。2026-10-20）
    r"(\frac{a}{b})^2",
    r"{\frac{a}{b}}^2",
    r"\left(\frac{a}{b}\right)^2",
    r"(\sqrt{x})^2",
    r"(a+b)_i",
    r"\max\{1,2\}",
    r"\min\{w,p-j\}",
    "(0,0)",
    "f(3,2)",
    r"\mathbf{(0,0)}",
    r"\text{/}\bot",
    r"x\bot y",
    r"\sqrt{x},y",
    r"\log x,y",
    # 関数の適用の後の普通の文字・開き括弧・分数・大型演算子、中身の末尾が関数の適用である根号の後の句読点、横書きの分数の前後、
    # 本体のない関数名を土台にした \underset の前、\text の中の引用符と ×（PowerPoint が lang="ja-JP" で和文の書体で描く記号。
    # 生成した資料と汎用の式の PDF で PowerPoint との差が見つかった規則。2026-10-20）
    r"\log x\,y",
    r"\log(x)y",
    r"\max(a,b)\,c",
    r"\log x(y)",
    r"\log x\frac{a}{b}",
    r"\log x\sum_i a_i",
    r"\sqrt{x\log x},y",
    "x a/b",
    "2a/b",
    "a/b,x",
    "x+a/b",
    "x(a/b)",
    "a/b x",
    "a/b(x)",
    "(x)a/b",
    r"x\underset{i}{\max} y",
    r"\text{“ab”}",
    r"\text{a×b}",
)

DECK_TOML = """# 数式の計測の照合データを作る資料（tests/tools/powerpoint_math_data.py が作る。内容は汎用の式だけ）。
schema_version = 1
entrypoint = "slides.py:build"
size = "16:9"
output_dir = "out"

[metadata]
title = "数式の照合データ"
author = "slidekit example"
last_modified_by = "slidekit example"

[math]
fallback = "off"
"""

SLIDES_TEMPLATE = '''"""数式の計測の照合データを作る資料。tests/tools/powerpoint_math_data.py が作る。"""

from slidekit import Deck, Math, Rect, Text

ITEMS = {items}


def build(context):
    deck = Deck(context)
    for page in ITEMS:
        slide = deck.slide(id=page["id"], kind="blank", title=page["id"])
        for item in page["items"]:
            area = Rect(*item["area"])
            if item["mode"] == "inline":
                slide.text(id=item["id"], text=[Text("["), Math(item["expression"]), Text("]")], area=area,
                           size_pt={size}, word_wrap=False, font_latin="{mark}")
            else:
                slide.equation(id=item["id"], expression=item["expression"], area=area, size_pt={size},
                               word_wrap=False, align="left", valign="top")
    return deck
'''


def supported(expressions=EXPRESSIONS) -> tuple[list[str], dict[str, str]]:
    """書き出しに成功する式と、失敗する式の理由。"""
    from slidekit.math.office_math import math_element

    accepted, rejected = [], {}
    for expression in expressions:
        try:
            math_element(expression, size=SIZE_PT)
        except (ValueError, KeyError) as error:
            rejected[expression] = str(error)
        else:
            accepted.append(expression)
    return accepted, rejected


def layout_items(expressions) -> list[dict]:
    """ページごとの枠。文中の数式のページの後に、独立した数式のページを置く。"""
    pages = []
    for mode, rows, height in (("inline", INLINE_ROWS, INLINE_HEIGHT), ("display", DISPLAY_ROWS, DISPLAY_HEIGHT)):
        per_page = rows * len(COLUMNS)
        for start in range(0, len(expressions), per_page):
            items = []
            for offset, expression in enumerate(expressions[start : start + per_page]):
                row, column = divmod(offset, len(COLUMNS))
                number = start + offset
                items.append(
                    {
                        "id": f"{mode}_{number:03d}",
                        "mode": mode,
                        "expression": expression,
                        "area": [COLUMNS[column], 16.0 + row * (height + ROW_GAP), FRAME_WIDTH, height],
                    }
                )
            pages.append({"id": f"{mode}_{start // per_page + 1:02d}", "items": items})
    return pages


def write_deck(work: Path, pages: list[dict]) -> Path:
    work.mkdir(parents=True, exist_ok=True)
    (work / "deck.toml").write_text(DECK_TOML, encoding="utf-8")
    source = SLIDES_TEMPLATE.format(items=json.dumps(pages, ensure_ascii=False, indent=1), size=SIZE_PT, mark=MARK_FONT)
    (work / "slides.py").write_text(source, encoding="utf-8")
    return work / "deck.toml"


def _font_key(name: str) -> str:
    from slidekit.pdf_layout import _font_name

    return _font_name(name)


def _chars(page, rect) -> list[dict]:
    """rect（pt）に掛かる文字（空白を除く）。page は `slidekit.pdf` のページ。書体名は部分集合の接頭辞を除く。

    以前使っていた PyMuPDF の rawdict（clip 付き）と同じ条件で選ぶ：字形の範囲（`Char.ink`）が rect に接するか重なり、
    文字の範囲（bbox。横は送り幅、縦はアセンダー・ディセンダー）が rect と重なる（辺が接するだけのものを除く）文字。
    塗りと輪郭の両方で描く文字は 1 字にする。PyMuPDF は字形の範囲の横に輪郭の範囲を、文字の範囲の縦に 1 em に縮めない
    アセンダー・ディセンダーを使うので、rect の辺に近い文字は選ぶ結果が違うことがある（照合データの枠では同じ）。
    bbox は文字の範囲、size は文字サイズ（pt）。
    """
    from slidekit.pdf import Box

    rect = Box.of(rect)
    spans = page.spans()
    filled = {span.seqno for span in spans if span.kind == 0}
    result = []
    for span in spans:
        if span.kind == 1 and span.seqno - 1 in filled:
            continue
        for char in span.chars:
            text = chr(char.unicode)
            ink = char.ink
            if not text.strip() or ink[2] < rect.x0 or ink[3] < rect.y0 or ink[0] > rect.x1 or ink[1] > rect.y1:
                continue
            bbox = (ink[0], char.bbox[1], ink[2], char.bbox[3])
            # 幅 0 の文字も、rect の中にあれば選ぶ（PyMuPDF と同じ比べ方）。
            if bbox[0] < rect.x1 and bbox[1] < rect.y1 and bbox[2] > rect.x0 and bbox[3] > rect.y0:
                result.append(
                    {"c": text, "font": span.font, "size": span.size, "origin": tuple(char.origin), "bbox": bbox}
                )
    return result


# 最後に画像化したページ（ページ、倍率、画像）。同じページの枠を続けて調べるので、1 ページ分を使い回す。
_RENDERED: list = []


def _page_image(page, scale):
    """ページを scale 倍で画像化した RGB の画像。"""
    if not (_RENDERED and _RENDERED[0] is page and _RENDERED[1] == scale):
        _RENDERED[:] = [page, scale, page.render(scale)]
    return _RENDERED[2]


def ink_mask(page, rect, scale=SCALE):
    """rect（pt）の中の描画の画素（描画を 255、背景を 0 にした白黒の画像）と、画像の原点 (x, y)（px）。

    画像は rect を整数の画素に広げ、ページの範囲で切り詰めて切り出す（PyMuPDF の get_pixmap(clip=...) と同じ範囲）ので、
    画素の位置は画像の原点から pt に戻す。page は `slidekit.pdf` のページ、rect は `Box` か 4 つの数。
    """
    from PIL import Image, ImageChops

    from slidekit.pdf import Box

    left, top, right, bottom = Box.of(rect).pixels(scale)
    page_left, page_top, page_right, page_bottom = page.rect.pixels(scale)
    left, top = max(left, page_left), max(top, page_top)
    right, bottom = max(left, min(right, page_right)), max(top, min(bottom, page_bottom))
    image = _page_image(page, scale).crop((left, top, right, bottom))
    background = max(image.getcolors(maxcolors=1 << 24))[1]
    difference = ImageChops.difference(image, Image.new("RGB", image.size, background)).convert("L")
    return difference.point(lambda value: 255 if value > 40 else 0), (left, top)


def _box_pt(box, origin, scale):
    """画像の中の範囲 (左, 上, 右, 下)（px）を pt にする。"""
    left, top, right, bottom = box
    x, y = origin
    return ((x + left) / scale, (y + top) / scale, (x + right) / scale, (y + bottom) / scale)


def ink_box(page, rect, scale=SCALE):
    """rect（pt）の中の描画範囲 (左, 上, 右, 下)（pt）。何も描かれていなければ None。"""
    mask, origin = ink_mask(page, rect, scale)
    box = mask.getbbox()
    return None if box is None else _box_pt(box, origin, scale)


def connected_ink_box(page, rect, top, scale=SCALE):
    """rect（pt）の中で、上端 top（pt）より下の描画と、それに続く上の描画の範囲 (左, 上, 右, 下)（pt）。なければ None。

    top より上の描画は、top の行から上へ、描画のない行が来るまでの分だけを含める。間に描画のない行がある上の描画
    （上の行の高い式が枠の下へはみ出した分）は含めない。
    """
    mask, origin = ink_mask(page, rect, scale)
    rows = mask.getprojection()[1]
    first = min(max(0, round(top * scale) - origin[1]), len(rows))
    if first < len(rows) and rows[first]:
        while first > 0 and rows[first - 1]:
            first -= 1
    box = mask.crop((0, first, mask.width, mask.height)).getbbox()
    if box is None:
        return None
    left, upper, right, lower = box
    return _box_pt((left, upper + first, right, lower + first), origin, scale)


def measure_inline(page, area, *, math_font="Cambria Math", scale=SCALE) -> dict | None:
    """目印の [ と ] で挟んだ文中の数式の送り幅と描画範囲。目印が見つからなければ None。"""
    from slidekit.pdf import Box

    x, y, width, height = area
    # 高い式の上端が枠の上へ出ても計測できるように、上を上の行の枠との間（ROW_GAP）の半分まで広げる。上の行の高い式の
    # 描画は枠の下端より下へ出ることがある（文中の \left(\frac{\frac{a}{b}}{c}\right) など）ので、枠の上端より上の描画は、
    # 枠の中の描画に描画のない行を挟まずに続く分だけを含める（connected_ink_box）。
    rect = Box(x - 2, y - ROW_GAP / 2, x + width + 2, y + height + 2)
    key = _font_key(math_font)
    marks = [char for char in _chars(page, rect) if _font_key(char["font"]) != key]
    opening = next((char for char in marks if char["c"] == "["), None)
    closing = next((char for char in reversed(marks) if char["c"] == "]"), None)
    if opening is None or closing is None or closing["origin"][0] <= opening["bbox"][2]:
        return None
    start, end, baseline = opening["bbox"][2], closing["origin"][0], opening["origin"][1]
    box = connected_ink_box(page, Box(start + 0.05, rect.y0, end - 0.05, rect.y1), y, scale)
    ink = (
        None
        if box is None
        else [
            round(box[0] - start, 3),
            round(box[2] - start, 3),
            round(baseline - box[1], 3),
            round(box[3] - baseline, 3),
        ]
    )
    return {"advance_pt": round(end - start, 3), "ink_pt": ink}


def measure_display(page, area, *, scale=SCALE) -> dict | None:
    """独立した数式の枠の描画範囲の幅と高さ。何も描かれていなければ None。"""
    x, y, width, height = area
    box = ink_box(page, (x - 2, y - 2, x + width + 2, y + height + 2), scale)
    if box is None:
        return None
    return {"ink_width_pt": round(box[2] - box[0], 3), "ink_height_pt": round(box[3] - box[1], 3)}


def measure_pdf(pdf: Path, pages: list[dict], *, math_font="Cambria Math") -> tuple[list[dict], list[str]]:
    """PDF の各枠を計測する。戻り値は（照合データの items、計測できなかった式）。"""
    from slidekit.pdf import open_pdf

    items, failed = [], []
    try:
        with open_pdf(pdf) as document:
            if len(document) != len(pages):
                raise ValueError(f"PDF と資料のページ数が一致しない: PDF={len(document)}, 資料={len(pages)}")
            for page, layout in zip(document, pages):
                for item in layout["items"]:
                    if item["mode"] == "inline":
                        value = measure_inline(page, item["area"], math_font=math_font)
                    else:
                        value = measure_display(page, item["area"])
                    if value is None:
                        failed.append(f"{item['mode']}: {item['expression']}")
                        continue
                    items.append({"expression": item["expression"], "mode": item["mode"], **value})
    finally:
        _RENDERED.clear()
    return items, failed


def font_record(project) -> dict:
    """数式の書体のファイル名・書体番号・版・sha256（パスは書かない）。"""
    from fontTools.ttLib import TTFont

    from slidekit.files import file_hash
    from slidekit.fonts import default_font

    family = project.theme["fonts"]["math"]
    specification = (project.environment.get("fonts") or {}).get(family) or default_font(family, project.environment)
    if specification is None:
        raise FileNotFoundError(f"数式の書体が見つかりません: {family}")
    path, index = Path(specification["path"]), int(specification.get("index", 0))
    with TTFont(path, fontNumber=index if path.suffix.lower() in (".ttc", ".otc") else -1, lazy=True) as font:
        version = "; ".join(sorted({entry.toUnicode() for entry in font["name"].names if entry.nameID == 5}))
    return {"family": family, "file": path.name, "index": index, "version": version, "sha256": file_hash(path)}


def render(work: Path) -> tuple[Path, dict, object]:
    """資料を構築して PPTX にし、PowerPoint で PDF にする。戻り値は（PDF、PowerPoint の検査結果、資料）。"""
    from slidekit import renderers
    from slidekit.config import load_project
    from slidekit.pipeline import construct
    from slidekit.writer import write_pptx

    project = load_project(work / "deck.toml")
    scene = construct(project)
    (work / "deck.json").write_text(json.dumps(scene, ensure_ascii=False, indent=1), encoding="utf-8")
    pptx, pdf = work / "math.pptx", work / "math.pdf"
    write_pptx(scene, project, pptx)
    pdf.unlink(missing_ok=True)
    with renderers.open_session(
        "powerpoint", project.environment, holder={"project": "powerpoint_math_data", "command": "math data"}
    ) as session:
        doctor = session.doctor()
        session.render(pptx, pdf, label="math data")
    return pdf, doctor, project


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="数式の計測の照合データを PowerPoint で作る（PowerPoint を使う）")
    parser.add_argument("--output", type=Path, required=True, help="書き出す照合データの JSON")
    parser.add_argument("--work", type=Path, default=DEFAULT_WORK, help="作業ディレクトリ（資料・PPTX・PDF を置く）")
    parser.add_argument(
        "--measure-only", action="store_true", help="変換せず、作業ディレクトリの math.pdf と doctor.json を測り直す"
    )
    arguments = parser.parse_args(argv)
    expressions, rejected = supported()
    for expression, reason in rejected.items():
        print(f"書き出せない式を除く: {expression}: {reason}", file=sys.stderr)
    pages = layout_items(expressions)
    work = arguments.work.resolve()
    write_deck(work, pages)
    from slidekit.config import load_project

    if arguments.measure_only:
        pdf, project = work / "math.pdf", load_project(work / "deck.toml")
        doctor = json.loads((work / "doctor.json").read_text(encoding="utf-8"))
    else:
        pdf, doctor, project = render(work)
        (work / "doctor.json").write_text(
            json.dumps({"version": doctor.get("version")}, ensure_ascii=False), encoding="utf-8"
        )
    font = font_record(project)
    items, failed = measure_pdf(pdf, pages, math_font=font["family"])
    for name in failed:
        print(f"測れない式: {name}", file=sys.stderr)
    data = {
        "schema_version": SCHEMA_VERSION,
        "created": date.today().isoformat(),
        "powerpoint": doctor.get("version"),
        "font": font,
        "text_fonts": {"japanese": project.theme["fonts"]["japanese"], "latin": project.theme["fonts"]["latin"]},
        "size_pt": SIZE_PT,
        "items": items,
    }
    from slidekit.files import write_json

    write_json(arguments.output, data)
    print(
        f"書いた照合データ: {arguments.output}（式 {len(items)} 件、除いた式 {len(rejected)} 件、測れない式 {len(failed)} 件）"
    )
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
