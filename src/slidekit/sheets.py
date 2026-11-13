"""確認を助ける比較画像（`slidekit review sheet`）。ページごとに、前の実行の画像・今の画像・互換表示の画像を横に並べる。

画像の組み立てだけを行い、実行と確認記録は読まない（並べるページと画像のパスは `records.review_sheet` が決める）。
比較画像は確認を助ける表示であり、それを見たことを確認の記録として扱わない。
"""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path

from .evidence import RENDER_SCALE

CELL_WIDTH = 480  # 1 枚の画像の幅（画素）
LABEL_HEIGHT = 22  # 列と行の見出しの高さ（画素）
ROWS_PER_SHEET = 6
_GAP = 8
_BACKGROUND = (255, 255, 255)
_TEXT = (32, 37, 43)
_MISSING = (235, 238, 242)
_BOX = (220, 30, 30)
_FRAME = (160, 168, 178)


def _font(size: int):
    from PIL import ImageFont

    try:
        return ImageFont.load_default(size=size)
    except TypeError:  # 古い Pillow
        return ImageFont.load_default()


def _cell(path: str | None, boxes: Sequence[Sequence[float]], width: int):
    """1 枚の画像を幅 width に縮め、差のある領域（pt）に枠を描く。画像がなければ灰色の枠にする。"""
    from PIL import Image, ImageDraw

    if path is None:
        image = Image.new("RGB", (width, round(width * 9 / 16)), _MISSING)
        return image
    with Image.open(path) as source:
        image = source.convert("RGB")
    ratio = width / image.width
    image = image.resize((width, max(1, round(image.height * ratio))))
    draw = ImageDraw.Draw(image)
    factor = RENDER_SCALE * ratio
    for x, y, w, h in boxes:
        draw.rectangle(
            [x * factor - 2, y * factor - 2, (x + w) * factor + 2, (y + h) * factor + 2], outline=_BOX, width=3
        )
    return image


def write_sheets(rows: Sequence[dict], columns: Sequence[str], directory: Path, prefix: str) -> list[Path]:
    """rows の各項目（label と、列の名前から {"path", "boxes"} への cells）を、ROWS_PER_SHEET 行ずつの PNG にする。"""
    from PIL import Image, ImageDraw

    directory.mkdir(parents=True, exist_ok=True)
    font = _font(14)
    paths = []
    for start in range(0, len(rows), ROWS_PER_SHEET):
        chunk = rows[start : start + ROWS_PER_SHEET]
        rendered = [
            [
                _cell(
                    (row["cells"].get(column) or {}).get("path"),
                    (row["cells"].get(column) or {}).get("boxes") or [],
                    CELL_WIDTH,
                )
                for column in columns
            ]
            for row in chunk
        ]
        heights = [max(image.height for image in images) for images in rendered]
        width = len(columns) * (CELL_WIDTH + _GAP) + _GAP
        height = LABEL_HEIGHT + sum(height + LABEL_HEIGHT + _GAP for height in heights) + _GAP
        canvas = Image.new("RGB", (width, height), _BACKGROUND)
        draw = ImageDraw.Draw(canvas)
        for index, column in enumerate(columns):
            draw.text((_GAP + index * (CELL_WIDTH + _GAP), 4), column, fill=_TEXT, font=font)
        y = LABEL_HEIGHT
        for row, images, row_height in zip(chunk, rendered, heights):
            draw.text((_GAP, y + 2), row["label"], fill=_TEXT, font=font)
            y += LABEL_HEIGHT
            for index, image in enumerate(images):
                left = _GAP + index * (CELL_WIDTH + _GAP)
                canvas.paste(image, (left, y))
                # 白いページが背景と区別できるように、画像の周りに細い枠を描く。
                draw.rectangle([left - 1, y - 1, left + image.width, y + image.height], outline=_FRAME)
            y += row_height + _GAP
        path = directory / f"{prefix}-{start // ROWS_PER_SHEET + 1:03d}.png"
        canvas.save(path, "PNG")
        paths.append(path)
    return paths
