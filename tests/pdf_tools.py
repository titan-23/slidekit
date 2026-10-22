"""試験用の PDF を作る補助と読む補助。

作る側は reportlab で PDF を書く。呼び方は以前使っていた PyMuPDF に近く、座標は pt、原点はページの左上、y は下向きである
（`slidekit.pdf` と同じ）。描く命令は記録しておき、`save()` のときにまとめて書く。
読む側は `slidekit.pdf`（文字・描画・画像化）と pypdf（文書情報・リンク）を使う。画像の範囲と線の区間は、`slidekit.pdf` を
通さずに pypdfium2 で直接読む。
"""

from __future__ import annotations

from dataclasses import dataclass, field
import io
from pathlib import Path
from typing import Any, Callable

from slidekit.pdf import Box, open_pdf

# PyMuPDF の短い書体名と、PDF の標準 14 書体の名前。
BASE_FONTS = {
    "helv": "Helvetica",
    "heit": "Helvetica-Oblique",
    "hebo": "Helvetica-Bold",
    "tiro": "Times-Roman",
    "tiit": "Times-Italic",
    "tibo": "Times-Bold",
    "cour": "Courier",
    "cobo": "Courier-Bold",
    "symb": "Symbol",
}
LINK_URI = "uri"
_registered: dict[str, str] = {}


def _font(fontname: str, fontfile: str | Path | None) -> str:
    """reportlab に登録した書体名。fontfile があれば TrueType の書体を埋め込む。"""
    if fontfile is None:
        return BASE_FONTS.get(fontname, fontname)
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    key = str(Path(fontfile).resolve())
    if key not in _registered:
        name = f"{fontname}-{len(_registered)}"
        pdfmetrics.registerFont(TTFont(name, key))
        _registered[key] = name
    return _registered[key]


def text_length(text: str, fontname: str = "helv", fontsize: float = 11, fontfile: str | Path | None = None) -> float:
    """文字列の送り幅の合計（pt）。PyMuPDF の get_text_length と同じ値（標準 14 書体は AFM の幅）。"""
    from reportlab.pdfbase import pdfmetrics

    return pdfmetrics.stringWidth(text, _font(fontname, fontfile), fontsize)


def _rgb(color) -> tuple[float, float, float] | None:
    if color is None:
        return None
    if isinstance(color, (int, float)):
        return (float(color),) * 3
    if len(color) == 1:
        return (float(color[0]),) * 3
    return tuple(float(value) for value in color[:3])


@dataclass
class Shape:
    """1 つの経路にまとめて描く図形（PyMuPDF の Shape に近い）。`finish()` で書式を決め、`commit()` でページに加える。"""

    page: "Page"
    items: list[tuple] = field(default_factory=list)
    styles: list[tuple[list[tuple], dict]] = field(default_factory=list)

    def draw_rect(self, rect) -> None:
        self.items.append(("rect", tuple(Box.of(rect))))

    def draw_line(self, start, end) -> None:
        self.items.append(("line", tuple(start), tuple(end)))

    def draw_polyline(self, points) -> None:
        self.items.append(("polyline", [tuple(point) for point in points]))

    def finish(
        self,
        *,
        color=None,
        fill=None,
        width: float = 1,
        fill_opacity: float = 1,
        stroke_opacity: float = 1,
        closePath: bool = False,
    ) -> None:  # noqa: N803
        """記録した線を 1 つの図形として描く設定にする。closePath の既定は偽にする（PyMuPDF は真）。"""
        self.styles.append(
            (
                self.items,
                {
                    "color": _rgb(color),
                    "fill": _rgb(fill),
                    "width": width,
                    "fill_opacity": fill_opacity,
                    "stroke_opacity": stroke_opacity,
                    "close": closePath,
                },
            )
        )
        self.items = []

    def commit(self) -> None:
        for items, style in self.styles:
            self.page._operations.append(("path", items, style))
        self.styles = []


class Page:
    """作る PDF の 1 ページ。"""

    def __init__(self, width: float, height: float):
        self.width = float(width)
        self.height = float(height)
        self._operations: list[tuple] = []

    @property
    def rect(self) -> Box:
        return Box(0.0, 0.0, self.width, self.height)

    def insert_text(
        self,
        point,
        text: str,
        *,
        fontname: str = "helv",
        fontsize: float = 11,
        color=(0, 0, 0),
        fill=None,
        render_mode: int = 0,
        border_width: float = 0.05,
        fill_opacity: float = 1,
        stroke_opacity: float = 1,
        fontfile: str | Path | None = None,
        rotate: float = 0,
    ) -> None:
        """point（左上が原点）を最初の行の基点として文字を書く。改行 \\n で行を分け、行送りは文字サイズの 1.2 倍にする
        （PyMuPDF は、書体のアセンダーとディセンダーの合計が 1 em を超えるとその合計を使う）。render_mode は PDF の文字の
        描画モード（0 塗り、1 輪郭、2 塗りと輪郭、3 見えない文字）。fill を省くと描画モードによらず color で塗る（PyMuPDF は
        描画モード 0 のときだけ）。rotate は point を中心に左回りに回す角度（度。PyMuPDF と同じ向き）。"""
        self._operations.append(
            (
                "text",
                (float(point[0]), float(point[1])),
                text,
                {
                    "font": _font(fontname, fontfile),
                    "size": float(fontsize),
                    "color": _rgb(color),
                    "fill": _rgb(fill) if fill is not None else _rgb(color),
                    "mode": render_mode,
                    "border": border_width,
                    "fill_opacity": fill_opacity,
                    "stroke_opacity": stroke_opacity,
                    "rotate": float(rotate),
                },
            )
        )

    def draw_rect(
        self,
        rect,
        *,
        color=(0, 0, 0),
        fill=None,
        width: float = 1,
        fill_opacity: float = 1,
        stroke_opacity: float = 1,
    ) -> None:
        """矩形を描く。color が None なら線を引かず、fill が None なら塗らない（PyMuPDF と同じ既定）。"""
        self._operations.append(
            (
                "path",
                [("rect", tuple(Box.of(rect)))],
                {
                    "color": _rgb(color),
                    "fill": _rgb(fill),
                    "width": width,
                    "fill_opacity": fill_opacity,
                    "stroke_opacity": stroke_opacity,
                    "close": False,
                },
            )
        )

    def draw_line(self, start, end, *, color=(0, 0, 0), width: float = 1, stroke_opacity: float = 1) -> None:
        self._operations.append(
            (
                "path",
                [("line", tuple(start), tuple(end))],
                {
                    "color": _rgb(color),
                    "fill": None,
                    "width": width,
                    "fill_opacity": 1,
                    "stroke_opacity": stroke_opacity,
                    "close": False,
                },
            )
        )

    def draw_polyline(self, points, *, color=(0, 0, 0), width: float = 1) -> None:
        self._operations.append(
            (
                "path",
                [("polyline", [tuple(point) for point in points])],
                {
                    "color": _rgb(color),
                    "fill": None,
                    "width": width,
                    "fill_opacity": 1,
                    "stroke_opacity": 1,
                    "close": False,
                },
            )
        )

    def new_shape(self) -> Shape:
        return Shape(self)

    def insert_image(
        self, rect, *, filename: str | Path | None = None, stream: bytes | None = None, keep_proportion: bool = True
    ) -> None:
        """画像を rect に置く。keep_proportion が真なら縦横比を保って中央に置く（PyMuPDF と同じ既定）。"""
        data = Path(filename).read_bytes() if filename is not None else bytes(stream)
        self._operations.append(("image", tuple(Box.of(rect)), data, keep_proportion))

    def insert_link(self, link: dict) -> None:
        """外部へのリンク（{"from": 矩形, "uri": 宛先}）を加える。"""
        self._operations.append(("link", tuple(Box.of(link["from"])), str(link["uri"])))

    # 書き出し

    def _render(self, canvas) -> None:
        for operation in self._operations:
            getattr(self, "_draw_" + operation[0])(canvas, *operation[1:])

    def _y(self, value: float) -> float:
        return self.height - value

    def _draw_text(self, canvas, point, text: str, style: dict) -> None:
        canvas.saveState()
        mode = style["mode"]
        if style["fill"] is not None:
            canvas.setFillColorRGB(*style["fill"])
        if style["color"] is not None:
            canvas.setStrokeColorRGB(*style["color"])
        canvas.setFillAlpha(style["fill_opacity"])
        canvas.setStrokeAlpha(style["stroke_opacity"])
        if mode in (1, 2):
            canvas.setLineWidth(style["border"] * style["size"])
        x, baseline = point[0], self._y(point[1])
        if style["rotate"]:
            # 原点を point に移して左回りに回し、回した座標の原点から書く。
            canvas.translate(x, baseline)
            canvas.rotate(style["rotate"])
            x, baseline = 0.0, 0.0
        for number, line in enumerate(text.split("\n")):
            content = canvas.beginText()
            content.setTextOrigin(x, baseline - number * style["size"] * 1.2)
            content.setFont(style["font"], style["size"])
            content.setTextRenderMode(mode)
            content.textOut(line)
            canvas.drawText(content)
        canvas.restoreState()

    def _draw_path(self, canvas, items: list[tuple], style: dict) -> None:
        canvas.saveState()
        path = canvas.beginPath()
        for item in items:
            if item[0] == "rect":
                x0, y0, x1, y1 = item[1]
                path.rect(x0, self._y(y1), x1 - x0, y1 - y0)
            elif item[0] == "line":
                path.moveTo(item[1][0], self._y(item[1][1]))
                path.lineTo(item[2][0], self._y(item[2][1]))
            elif item[0] == "polyline":
                (x, y), *rest = item[1]
                path.moveTo(x, self._y(y))
                for x, y in rest:
                    path.lineTo(x, self._y(y))
        if style.get("close"):
            path.close()
        stroke, fill = style["color"] is not None, style["fill"] is not None
        if stroke:
            canvas.setStrokeColorRGB(*style["color"])
            canvas.setLineWidth(style["width"])
            canvas.setStrokeAlpha(style["stroke_opacity"])
        if fill:
            canvas.setFillColorRGB(*style["fill"])
            canvas.setFillAlpha(style["fill_opacity"])
        canvas.drawPath(path, stroke=int(stroke), fill=int(fill))
        canvas.restoreState()

    def _draw_image(self, canvas, rect, data: bytes, keep: bool) -> None:
        from reportlab.lib.utils import ImageReader

        x0, y0, x1, y1 = rect
        canvas.drawImage(
            ImageReader(io.BytesIO(data)),
            x0,
            self._y(y1),
            x1 - x0,
            y1 - y0,
            preserveAspectRatio=keep,
            anchor="c",
            mask="auto",
        )

    def _draw_link(self, canvas, rect, uri: str) -> None:
        x0, y0, x1, y1 = rect
        canvas.linkURL(uri, (x0, self._y(y1), x1, self._y(y0)), relative=0, thickness=0)


class Document:
    """作る PDF（`with Document() as document:` で使う）。"""

    def __init__(self):
        self.pages: list[Page] = []
        self.metadata: dict[str, str] = {}

    def __enter__(self) -> Document:
        return self

    def __exit__(self, *_exc) -> None:
        return None

    def __len__(self) -> int:
        return len(self.pages)

    def __getitem__(self, index: int) -> Page:
        return self.pages[index]

    def new_page(self, width: float = 595, height: float = 842) -> Page:
        page = Page(width, height)
        self.pages.append(page)
        return page

    def set_metadata(self, values: dict[str, str]) -> None:
        self.metadata.update(values)

    def tobytes(self) -> bytes:
        from reportlab.pdfgen import canvas as canvases

        buffer = io.BytesIO()
        canvas = canvases.Canvas(buffer, pageCompression=0, invariant=1)
        setters: dict[str, Callable[[str], Any]] = {
            "author": canvas.setAuthor,
            "title": canvas.setTitle,
            "subject": canvas.setSubject,
            "keywords": canvas.setKeywords,
            "creator": canvas.setCreator,
            "producer": canvas.setProducer,
        }
        for key, value in self.metadata.items():
            setters[key](value)
        for page in self.pages or [Page(595, 842)]:
            canvas.setPageSize((page.width, page.height))
            page._render(canvas)
            canvas.showPage()
        canvas.save()
        return buffer.getvalue()

    def save(self, path: str | Path) -> None:
        Path(path).write_bytes(self.tobytes())


# 読む側


def metadata(path: str | Path) -> dict[str, str]:
    """PDF の文書情報（author・title などの小文字の名前から値）。"""
    import pypdf

    info = pypdf.PdfReader(str(path)).metadata or {}
    return {str(key).lstrip("/").lower(): str(value) for key, value in info.items()}


def links(path: str | Path, index: int) -> list[str]:
    """PDF の index ページ目（0 始まり）のリンクの宛先（URI）。"""
    import pypdf

    page = pypdf.PdfReader(str(path)).pages[index]
    found = []
    for annotation in page.get("/Annots") or []:
        annotation = annotation.get_object()
        action = annotation.get("/A")
        action = action.get_object() if action is not None else None
        if action is not None and action.get("/URI") is not None:
            found.append(str(action["/URI"]))
    return found


def page_text(page) -> str:
    """ページの文字（描画順。見えない文字を除く）。"""
    return "".join(span.text for span in page.spans() if span.kind != 3)


def render_region(page, scale: float, box):
    """ページを scale 倍で画像にし、box（pt）の範囲を切り出した RGB の画像。"""
    image = page.render(scale)
    left, top, right, bottom = Box.of(box).pixels(scale)
    return image.crop((left, top, right, bottom))


def _page_objects(path: str | Path, index: int, kind: int, read: Callable[[Any, tuple, float], Any]) -> list:
    """PDF の index ページ目（0 始まり）の種類 kind（PDFium の FPDF_PAGEOBJ_*）の要素を、フォーム XObject の中も含めて
    描画順に読む。read(要素, ページへの行列, ページの高さ) の戻り値の列を返す。行列は PDF の [a b c d e f]（原点は左下）。"""
    import pypdfium2 as pdfium
    import pypdfium2.raw as raw

    def compose(inner, outer):
        a, b, c, d, e, f = inner
        A, B, C, D, E, F = outer
        return (a * A + b * C, a * B + b * D, c * A + d * C, c * B + d * D, e * A + f * C + E, e * B + f * D + F)

    document = pdfium.PdfDocument(str(path))
    try:
        page = document[index]
        height = page.get_height()
        found = []

        def walk(container, count, get, outer):
            for position in range(count(container)):
                handle = get(container, position)
                m = raw.FS_MATRIX()
                raw.FPDFPageObj_GetMatrix(handle, m)
                matrix = compose((m.a, m.b, m.c, m.d, m.e, m.f), outer)
                object_kind = raw.FPDFPageObj_GetType(handle)
                if object_kind == raw.FPDF_PAGEOBJ_FORM:
                    walk(handle, raw.FPDFFormObj_CountObjects, raw.FPDFFormObj_GetObject, matrix)
                elif object_kind == kind:
                    found.append(read(handle, matrix, height))

        walk(page.raw, raw.FPDFPage_CountObjects, raw.FPDFPage_GetObject, (1.0, 0.0, 0.0, 1.0, 0.0, 0.0))
        page.close()
        return found
    finally:
        document.close()


def image_boxes(path: str | Path, index: int) -> list[Box]:
    """PDF の index ページ目（0 始まり）に描かれた画像の範囲（pt。左上が原点）の描画順の列。

    PyMuPDF の get_image_info の bbox に当たる（画像の行列で単位正方形を移した範囲）。
    """
    import pypdfium2.raw as raw

    def read(_handle, matrix, height):
        a, b, c, d, e, f = matrix
        xs, ys = (e, e + a, e + c, e + a + c), (f, f + b, f + d, f + b + d)
        return Box(min(xs), height - max(ys), max(xs), height - min(ys))

    return _page_objects(path, index, raw.FPDF_PAGEOBJ_IMAGE, read)


def line_segments(path: str | Path, index: int) -> list[tuple[tuple[float, float, float], tuple, tuple]]:
    """PDF の index ページ目（0 始まり）の、線を引く経路の直線の区間（線の色, 始点, 終点）の描画順の列。

    座標は pt で左上が原点、色は RGB（0〜1）。PyMuPDF の get_drawings の線の描画の項目 "l" に当たる（矩形の辺も直線の区間
    として返す）。
    """
    import ctypes

    import pypdfium2.raw as raw

    def read(handle, matrix, height):
        fill, stroke = ctypes.c_int(), ctypes.c_int()
        raw.FPDFPath_GetDrawMode(handle, fill, stroke)
        if not stroke.value:
            return []
        red, green, blue, alpha = (ctypes.c_uint() for _ in range(4))
        raw.FPDFPageObj_GetStrokeColor(handle, red, green, blue, alpha)
        color = (red.value / 255, green.value / 255, blue.value / 255)
        a, b, c, d, e, f = matrix
        x, y = ctypes.c_float(), ctypes.c_float()
        segments, current = [], None
        for number in range(raw.FPDFPath_CountSegments(handle)):
            segment = raw.FPDFPath_GetPathSegment(handle, number)
            raw.FPDFPathSegment_GetPoint(segment, x, y)
            point = (a * x.value + c * y.value + e, height - (b * x.value + d * y.value + f))
            if raw.FPDFPathSegment_GetType(segment) == raw.FPDF_SEGMENT_LINETO and current is not None:
                segments.append((color, current, point))
            current = point
        return segments

    return [segment for segments in _page_objects(path, index, raw.FPDF_PAGEOBJ_PATH, read) for segment in segments]


__all__ = [
    "BASE_FONTS",
    "Document",
    "LINK_URI",
    "Page",
    "Shape",
    "image_boxes",
    "line_segments",
    "links",
    "metadata",
    "open_pdf",
    "page_text",
    "render_region",
    "text_length",
]
