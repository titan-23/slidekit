"""実行の PDF の画像に、要素の予測と実際の基線・行の範囲・インクを重ねた確認画像を作る（開発用。パッケージに含めない）。

使い方:
  python tools/layout_overlay.py --run RUN [--run RUN ...] --element ID [--element ID ...] --output DIR
      [--scale 8] [--margin-em 1] [--remeasure [--project deck.toml]] [--pdf PDF]

- 画像は実行の標準表示の PDF（renders/powerpoint/standard/deck.pdf。最終 PPTX を PowerPoint で変換したもの）の、
  要素の枠の周りを pypdfium2 で scale 倍（1 pt あたりの画素数）にしたもの。
- 予測は実行の layout.json（生成したときの計測）。--remeasure では今のコードで計測し直す（tests/test_layout_fixture.py の
  predicted_lines と同じく、行頭と行末の空白と、線で描く数式の端を補正する）。実際は pdf_layout.read_slide で PDF から読む。
- 予測は青（太い線）、実際は赤（細い線。予測の上に描く）、文字領域は灰色で描く。行の基線は横線、行の範囲は基線の両端の
  縦線、インクは矩形で示す。予測と実際が重なると、赤い線の両側に青い線が見える。
- 行ごとの基線・左端・右端とインクの 4 辺の差（実際 − 予測）を pt と em（要素の最大の文字サイズ）で標準出力に示し、
  出力先に overlay.json と画像（実行と要素ごとの画像と、それを縦に並べた overlay.png）を書く。
- 要素は「ページの識別子/要素の識別子」か要素の識別子で指定する。文字サイズや行間を変えた複数の実行を並べて比べることが
  でき、実行にない要素は省く。
- 配置を直すときの確認の補助であり、表示確認の記録（review）には使わない。出力先は実行ディレクトリの外にする。
"""

from __future__ import annotations

import argparse
from collections.abc import Mapping
import importlib
import json
from pathlib import Path
import re
import sys

ROOT = Path(__file__).resolve().parents[1]
TESTS = ROOT / "tests"
PDF = Path("renders") / "powerpoint" / "standard" / "deck.pdf"
PREDICTED, ACTUAL, FRAME = (0, 90, 255, 220), (230, 0, 0, 220), (150, 150, 150, 255)
EDGES = ("left", "top", "right", "bottom")
EDGE_NAMES = {"left": "左端", "top": "上端", "right": "右端", "bottom": "下端"}
HEADER = 22  # 画像の上の見出しの高さ（px）
_UNSAFE = re.compile(r"[^A-Za-z0-9_.-]")


class OverlayError(ValueError):
    """確認画像を作れない入力。"""


def comparison():
    """文字の配置の比べ方（tests/test_layout_fixture.py の predicted_lines と _size）。"""
    if str(TESTS) not in sys.path:
        sys.path.insert(0, str(TESTS))
    return importlib.import_module("test_layout_fixture")


def _read_json(path: Path) -> dict:
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise OverlayError(f"ファイルがありません: {path}") from None


def find_element(scene: Mapping, spec: str) -> tuple[int, Mapping, Mapping] | None:
    """指定に合う（ページの位置、ページ、要素）。なければ None、複数のページに合えば OverlayError。

    指定は要素の識別子か「ページの識別子/要素の識別子」（まとまりの子の識別子は / を含むので、両方の読み方を試す）。
    """
    found = []
    for index, slide in enumerate(scene.get("slides", [])):
        prefix = f"{slide.get('id')}/"
        for element in slide.get("elements", []):
            identifier = str(element.get("id"))
            if spec == identifier or (spec.startswith(prefix) and spec[len(prefix) :] == identifier):
                found.append((index, slide, element))
    if len(found) > 1:
        pages = "、".join(str(slide.get("id")) for _index, slide, _element in found)
        raise OverlayError(f"要素 {spec} が複数のページにある（{pages}）。「ページの識別子/要素の識別子」で指定する")
    return found[0] if found else None


def recorded_prediction(layout: Mapping, slide_id, element_id) -> dict | None:
    """実行の layout.json の予測（内容のある行の基線・左端・右端、インク、文字領域）。"""
    for slide in layout.get("slides", []):
        if slide.get("id") != slide_id:
            continue
        for record in slide.get("elements", []):
            if record.get("id") == element_id:
                lines = [
                    {
                        "text": line["text"],
                        "baseline": line["baseline_pt"],
                        "left": line["left_pt"],
                        "right": line["left_pt"] + line["width_pt"],
                    }
                    for line in record["lines"]
                    if line["text"].strip()
                ]
                x, y, w, h = record["text_area_pt"]
                return {"lines": lines, "ink": record.get("ink_pt"), "area": [x, y, x + w, y + h]}
    return None


def measured_prediction(element: Mapping, fonts) -> dict:
    """今のコードで計測し直した予測（試験と同じ補正をした行の範囲）。"""
    from slidekit import measure

    layout = measure.layout_element(element, fonts)
    lines = comparison().predicted_lines(layout, element, fonts)
    ink, area = layout.ink, layout.text_area
    return {
        "lines": [{key: line[key] for key in ("text", "baseline", "left", "right")} for line in lines],
        "ink": None if ink is None else [ink.x, ink.y, ink.x + ink.w, ink.y + ink.h],
        "area": [area.x, area.y, area.x + area.w, area.y + area.h],
    }


def actual_values(rendered) -> dict:
    return {
        "lines": [{"baseline": line.baseline, "left": line.left, "right": line.right} for line in rendered.lines],
        "ink": None if rendered.ink is None else list(rendered.ink),
    }


def differences(predicted: Mapping, actual: Mapping, size: float) -> dict:
    """実際 − 予測（pt と em）。行数が違えば行ごとの差を求めない。"""

    def value(difference: float) -> dict:
        return {"pt": round(difference, 3), "em": round(difference / size, 4)}

    result: dict = {"lines": [len(predicted["lines"]), len(actual["lines"])], "line_differences": None, "ink": None}
    if len(predicted["lines"]) == len(actual["lines"]):
        result["line_differences"] = [
            {key: value(a[key] - p[key]) for key in ("baseline", "left", "right")}
            for p, a in zip(predicted["lines"], actual["lines"])
        ]
    if predicted["ink"] is not None and actual["ink"] is not None:
        result["ink"] = {edge: value(a - p) for edge, p, a in zip(EDGES, predicted["ink"], actual["ink"])}
    return result


def _render(pdf: Path, index: int, box, scale: float):
    """PDF の index ページの box（pt の (左, 上, 右, 下)）を scale 倍にした RGB の画像と、その左上の画素の位置。

    slidekit.pdf の Page.render と同じく、画素数を決めた画像に行列で描く（倍率を正確に保つ）。
    """
    import pypdfium2 as pdfium
    import pypdfium2.raw as raw

    from slidekit.pdf import Box

    document = pdfium.PdfDocument(str(pdf))
    try:
        page = document[index]
        page_box = Box(0.0, 0.0, float(page.get_width()), float(page.get_height()))
        left, top, right, bottom = (Box(*box) & page_box).pixels(scale)
        width, height = max(1, right - left), max(1, bottom - top)
        bitmap = pdfium.PdfBitmap.new_native(width, height, raw.FPDFBitmap_BGR)
        try:
            bitmap.fill_rect((255, 255, 255, 255), 0, 0, width, height)
            matrix = raw.FS_MATRIX(scale, 0, 0, scale, -left, -top)
            clip = raw.FS_RECTF(0, 0, width, height)
            raw.FPDF_RenderPageBitmapWithMatrix(bitmap, page.raw, matrix, clip, raw.FPDF_ANNOT)
            image = bitmap.to_pil().convert("RGB")
        finally:
            bitmap.close()
        page.close()
    finally:
        document.close()
    return image, (left, top)


def draw(image, origin, scale: float, predicted: Mapping, actual: Mapping, size: float, title: str):
    """画像に予測と実際を重ね、上に見出しを付けた画像を返す。"""
    from PIL import Image, ImageDraw

    overlay = Image.new("RGBA", image.size, (0, 0, 0, 0))
    pen = ImageDraw.Draw(overlay)
    width = max(1, round(scale / 8))
    tick = 0.25 * size * scale

    def point(x, y):
        return (x * scale - origin[0], y * scale - origin[1])

    left, top, right, bottom = predicted["area"]
    pen.rectangle([point(left, top), point(right, bottom)], outline=FRAME, width=width)
    for values, color, size_px in ((predicted, PREDICTED, 3 * width), (actual, ACTUAL, width)):
        if values["ink"] is not None:
            x0, y0, x1, y1 = values["ink"]
            pen.rectangle([point(x0, y0), point(x1, y1)], outline=color, width=size_px)
        for line in values["lines"]:
            (x0, y), (x1, _) = point(line["left"], line["baseline"]), point(line["right"], line["baseline"])
            pen.line([(x0, y), (x1, y)], fill=color, width=size_px)
            for x in (x0, x1):
                pen.line([(x, y - tick), (x, y + tick)], fill=color, width=size_px)
    body = Image.alpha_composite(image.convert("RGBA"), overlay).convert("RGB")
    sheet = Image.new("RGB", (max(body.width, 360), body.height + HEADER), "white")
    sheet.paste(body, (0, HEADER))
    label = f"{title}  blue: predicted  red: actual  gray: text area"
    ImageDraw.Draw(sheet).text((4, 4), label.encode("ascii", "replace").decode(), fill=(0, 0, 0))
    return sheet


def report_lines(row: Mapping) -> list[str]:
    """1 要素の差の文言。"""
    head = f"{row['run']} {row['slide_id']}/{row['element_id']}（{row['size_pt']:g} pt、予測 {row['source']}）"
    lines = [head]
    difference = row["differences"]
    predicted, actual = difference["lines"]
    if difference["line_differences"] is None:
        lines.append(f"  行数: 予測 {predicted}、実際 {actual}（行ごとの差は求めない）")
    else:
        names = {"baseline": "基線", "left": "左端", "right": "右端"}
        for number, values in enumerate(difference["line_differences"], 1):
            parts = [f"{names[key]} {value['pt']:+.2f} pt（{value['em']:+.3f} em）" for key, value in values.items()]
            lines.append(f"  {number} 行目: " + "、".join(parts))
    if difference["ink"] is not None:
        parts = [
            f"{EDGE_NAMES[edge]} {value['pt']:+.2f} pt（{value['em']:+.3f} em）"
            for edge, value in difference["ink"].items()
        ]
        lines.append("  インク: " + "、".join(parts))
    return lines


def overlay(
    runs,
    elements,
    output: Path,
    *,
    scale: float = 8.0,
    margin_em: float = 1.0,
    remeasure: bool = False,
    project: Path | None = None,
    pdf: Path | None = None,
    fonts=None,
) -> list[dict]:
    """実行と要素ごとに確認画像を書き、差の一覧（overlay.json の rows）を返す。

    fonts は --remeasure で使う FontSource で、省略すると資料（project か、生成記録の資料）の書体を使う。
    """
    from PIL import Image

    from slidekit.pdf import open_pdf
    from slidekit.pdf_layout import read_slide

    runs = [Path(run).resolve() for run in runs]
    output = Path(output).resolve()
    for run in runs:
        if output == run or run in output.parents:
            raise OverlayError(f"出力先は実行ディレクトリの外にする（表示確認の記録と分ける）: {output}")
    if not 0 < scale <= 32 or margin_em < 0:
        raise OverlayError("scale は 0 より大きく 32 以下、margin-em は 0 以上にする")
    if remeasure and fonts is None:
        from slidekit.config import load_project
        from slidekit.fonts import FontRegistry

        path = project or _read_json(runs[0] / "manifest.json").get("project")
        if not path or not Path(path).exists():
            raise OverlayError("--remeasure では資料（deck.toml）を --project で指定する")
        with FontRegistry(load_project(path)) as registry:
            return overlay(
                runs, elements, output, scale=scale, margin_em=margin_em, remeasure=True, pdf=pdf, fonts=registry
            )
    compare = comparison()
    rows, panels = [], []
    for number, run in enumerate(runs, 1):
        scene = _read_json(run / "deck.json")
        recorded = None if remeasure else _read_json(run / "layout.json")
        source = Path(pdf) if pdf is not None and len(runs) == 1 else run / PDF
        if not source.is_file():
            raise OverlayError(f"標準表示の PDF がありません: {source}")
        targets = [found for spec in elements if (found := find_element(scene, spec)) is not None]
        for spec in elements:
            if find_element(scene, spec) is None:
                print(f"実行 {run.name} に要素 {spec} がない（省く）", file=sys.stderr)
        with open_pdf(source) as document:
            if len(document) != len(scene.get("slides", [])):
                raise OverlayError(f"PDF と場面のページ数が一致しない: {source}")
            readings = {index: read_slide(document[index], slide) for index, slide, _element in targets}
        for index, slide, element in targets:
            slide_id, element_id = slide.get("id"), str(element.get("id"))
            rendered = readings[index].elements.get(element_id)
            if rendered is None:
                print(f"実行 {run.name} の {slide_id}/{element_id} を PDF から読めない（省く）", file=sys.stderr)
                continue
            if remeasure:
                predicted = measured_prediction(element, fonts)
            else:
                predicted = recorded_prediction(recorded, slide_id, element_id)
                if predicted is None:
                    print(f"実行 {run.name} の layout.json に {slide_id}/{element_id} がない（省く）", file=sys.stderr)
                    continue
            actual = actual_values(rendered)
            size = compare._size(element)
            row = {
                "run": run.name,
                "run_dir": str(run),
                "slide_id": slide_id,
                "element_id": element_id,
                "size_pt": size,
                "source": "remeasure" if remeasure else "layout.json",
                "predicted": predicted,
                "actual": actual,
                "differences": differences(predicted, actual, size),
            }
            rows.append(row)
            boxes = [predicted["area"], *(value for value in (predicted["ink"], actual["ink"]) if value is not None)]
            margin = margin_em * size
            box = (
                min(b[0] for b in boxes) - margin,
                min(b[1] for b in boxes) - margin,
                max(b[2] for b in boxes) + margin,
                max(b[3] for b in boxes) + margin,
            )
            image, origin = _render(source, index, box, scale)
            title = f"{run.name} {slide_id}/{element_id} {size:g}pt"
            panel = draw(image, origin, scale, predicted, actual, size, title)
            # 同じ名前の実行（別の資料の runs/0001 など）を並べても重ならないように、指定した順の番号を付ける。
            name = _UNSAFE.sub("_", f"{number:02d}_{run.name}_{slide_id}_{element_id}") + ".png"
            output.mkdir(parents=True, exist_ok=True)
            panel.save(output / name)
            row["image"] = name
            panels.append(panel)
    if not rows:
        raise OverlayError("指定した要素を、どの実行でも読めない")
    gap = 12
    sheet = Image.new(
        "RGB", (max(panel.width for panel in panels), sum(panel.height + gap for panel in panels) - gap), "white"
    )
    y = 0
    for panel in panels:
        sheet.paste(panel, (0, y))
        y += panel.height + gap
    sheet.save(output / "overlay.png")
    from slidekit.files import write_json

    write_json(output / "overlay.json", {"schema_version": 1, "scale": scale, "rows": rows})
    return rows


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="実行の PDF の画像に、要素の予測と実際の基線・行の範囲・インクを重ねる（開発用。PowerPoint は使わない）"
    )
    parser.add_argument("--run", type=Path, action="append", required=True, help="実行ディレクトリ（複数指定できる）")
    parser.add_argument(
        "--element", action="append", required=True, help="要素（ページの識別子/要素の識別子か、要素の識別子）"
    )
    parser.add_argument("--output", type=Path, required=True, help="画像と overlay.json を書くディレクトリ")
    parser.add_argument("--scale", type=float, default=8.0, help="1 pt あたりの画素数（既定 8）")
    parser.add_argument("--margin-em", type=float, default=1.0, help="切り出す範囲の余白（文字サイズの em。既定 1）")
    parser.add_argument("--remeasure", action="store_true", help="layout.json の代わりに、今のコードで計測し直す")
    parser.add_argument("--project", type=Path, help="--remeasure で書体を使う資料（既定は生成記録の資料）")
    parser.add_argument("--pdf", type=Path, help="読む PDF（実行が 1 つのときだけ。既定は標準表示の PDF）")
    arguments = parser.parse_args(argv)
    try:
        rows = overlay(
            arguments.run,
            arguments.element,
            arguments.output,
            scale=arguments.scale,
            margin_em=arguments.margin_em,
            remeasure=arguments.remeasure,
            project=arguments.project,
            pdf=arguments.pdf,
        )
    except OverlayError as error:
        print(f"確認画像を作れません: {error}", file=sys.stderr)
        return 2
    for row in rows:
        print("\n".join(report_lines(row)))
    print(f"画像: {arguments.output / 'overlay.png'}（要素ごとの画像と overlay.json も同じ場所）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
