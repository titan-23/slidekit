"""PDF の読み取り（slidekit.pdf）の規則を確かめる。PDF は reportlab で作る（pdf_tools）か、描画命令を直接書く。"""

import io

import pytest

from pdf_tools import Document, text_length
from slidekit.pdf import Box, _read_to_unicode, open_pdf, render_to_pixels


def write(path, draw, size=(400, 300)):
    with Document() as document:
        draw(document.new_page(*size))
        document.save(path)
    return path


def test_box_union_intersection_and_pixels():
    box = Box(0, 0, 10, 10)
    # 空の矩形は和に含めない（MuPDF の fz_union_rect と同じ）。
    assert box | Box(5, 5, 5, 20) == box
    assert box | Box(5, 5, 20, 12) == Box(0, 0, 20, 12)
    # 辺が接するだけなら重ならない。
    assert box.intersects(Box(9, 9, 12, 12)) and not box.intersects(Box(10, 0, 12, 10))
    # 画素の範囲は外側へ丸め、0.001 画素の余裕を持たせる（MuPDF の fz_round_rect と同じ）。
    assert Box(0.1, 0.2, 10.2, 10.3).pixels(2) == (0, 0, 21, 21)
    assert Box(0.5, 0.5, 1.0, 1.0).pixels(2) == (1, 1, 2, 2)


@pytest.mark.parametrize("embedded", [False, True], ids=["base14", "embedded"])
def test_overlapping_identical_strings_are_both_read(tmp_path, dejavu, embedded):
    # PDFium は近い位置に同じ文字列を描く文字の要素を重ね描きとみなして、文字の一覧から除く。除いた字も描画命令から補い、
    # 文字は符号化か /ToUnicode から、原点は幅（幅の表がない標準 14 書体は PDFium の幅）から求める。
    fontfile = dejavu.path if embedded else None
    text = "Same wörds"

    def draw(page):
        for baseline in (100, 90):
            page.insert_text((40, baseline), text, fontsize=20, fontfile=fontfile)

    with open_pdf(write(tmp_path / "same.pdf", draw)) as document:
        first, second = document[0].spans()
    assert (first.text, second.text) == (text, text)
    assert [char.origin[1] for char in second.chars] == pytest.approx([90] * len(text))
    assert [char.origin[0] for char in second.chars] == pytest.approx([char.origin[0] for char in first.chars])
    assert second.chars[-1].ink[2] == pytest.approx(40 + text_length(text, fontsize=20, fontfile=fontfile))
    # インクの上下は字形の輪郭で、同じ字なら基線の差だけずれる。
    assert [first.chars[0].ink[1] - 10, first.chars[0].ink[3] - 10] == pytest.approx(
        [second.chars[0].ink[1], second.chars[0].ink[3]], abs=0.5
    )
    assert second.chars[0].ink[1] < 80


def test_to_unicode_reads_single_characters():
    class Stream:
        def __init__(self, text):
            self.text = text

        def get_data(self):
            return self.text.encode("latin-1")

    cmap = Stream(
        "2 beginbfchar\n<0003> <0020>\n<0004> <00660069>\nendbfchar\n"
        "2 beginbfrange\n<0010> <0012> <0041>\n<0020> <0021> [<D835DC00> <00E9>]\nendbfrange\n"
    )
    # 2 字に対応する符号（合字の fi）は含めない。基本多言語面の外の文字は代用符号の組から 1 字にする。
    assert _read_to_unicode(cmap) == {3: 0x20, 0x10: 0x41, 0x11: 0x42, 0x12: 0x43, 0x20: 0x1D400, 0x21: 0xE9}
    assert _read_to_unicode(None) == {} and _read_to_unicode(Stream("<0001> <00>")) == {}


def test_text_spans_drop_characters_outside_the_page(tmp_path):
    # 検査に使う文字片は、ページの完全に外にある字を除く（PyMuPDF の TEXT_MEDIABOX_CLIP と同じ）。一部が入る字は残す。
    path = write(tmp_path / "edge.pdf", lambda page: page.insert_text((370, 40), "Sample", fontsize=20))
    with open_pdf(path) as document:
        page = document[0]
        assert [span.text for span in page.spans()] == ["Sample"]
        [span] = page.text_spans()
    assert span.text == "Sam"
    assert span.bbox.x1 == pytest.approx(370 + text_length("Sam", fontsize=20))


def test_drawing_order_counts_fill_and_stroke_separately(tmp_path):
    def draw(page):
        page.draw_rect((10, 10, 50, 50), color=(1, 0, 0), fill=(0, 0, 1))
        page.insert_text((60, 40), "A", fontsize=20, render_mode=2)
        page.draw_line((100, 10), (150, 60), color=(0, 1, 0), width=2)

    with open_pdf(write(tmp_path / "order.pdf", draw)) as document:
        page = document[0]
        drawings, spans = page.drawings(), page.spans()
    # 1 つの要素で塗りと線を描く図形と、塗りと輪郭で描く文字は、描画順の番号を 2 つ使う（PyMuPDF と同じ）。
    assert [(drawing.seqno, drawing.kind) for drawing in drawings] == [(0, "fs"), (4, "s")]
    assert [(span.seqno, span.kind) for span in spans] == [(2, 0), (3, 1)]
    square, line = drawings
    assert (square.rect, square.fill, square.color) == (Box(10, 10, 50, 50), (0.0, 0.0, 1.0), (1.0, 0.0, 0.0))
    assert (line.rect, line.width, line.color) == (Box(100, 10, 150, 60), 2.0, (0.0, 1.0, 0.0))


def raw_pdf(path, content, *, form=None, size=(400, 300)):
    """描画命令を直接書いた 1 ページの PDF。書体 F1 は埋め込まない Helvetica（幅の表なし）、form はフォーム XObject Fm の
    描画命令。"""

    def stream(data, dictionary=b""):
        return b"<< %s /Length %d >>\nstream\n%s\nendstream" % (dictionary, len(data), data)

    resources = b"<< /Font << /F1 5 0 R >> /XObject << /Fm 6 0 R >> >>"
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 %d %d] /Resources %s /Contents 4 0 R >>" % (*size, resources),
        stream(content),
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
        stream(form or b"", b"/Type /XObject /Subtype /Form /BBox [0 0 %d %d] /Resources %s" % (*size, resources)),
    ]
    data = bytearray(b"%PDF-1.7\n")
    offsets = []
    for number, body in enumerate(objects, 1):
        offsets.append(len(data))
        data += b"%d 0 obj\n%s\nendobj\n" % (number, body)
    start = len(data)
    data += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objects) + 1)
    data += b"".join(b"%010d 00000 n \n" % offset for offset in offsets)
    data += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objects) + 1, start)
    path.write_bytes(bytes(data))
    return path


def origins(path):
    with open_pdf(path) as document:
        return [(chr(char.unicode), char.origin) for span in document[0].spans() for char in span.chars]


@pytest.mark.parametrize(
    "content, form, baseline, first",
    [
        # 文字の上下の移動（Ts）は、PDFium の文字の要素の位置に含まれる。
        (b"BT /F1 20 Tf 20 200 Td 5 Ts (a  b) Tj ET", None, 95, 20),
        # 最初の字の前の TJ の数も、文字の要素の位置に含まれる。
        (b"BT /F1 20 Tf 20 200 Td [-1000 (a  b)] TJ ET", None, 100, 40),
        # フォーム XObject は、呼んだ時点の文字の状態（ここでは Tc）を受け継ぐ。
        (b"3 Tc /Fm Do", b"BT /F1 20 Tf 20 200 Td (a  b) Tj ET", 100, 20),
    ],
    ids=["rise", "leading-kerning", "form-state"],
)
def test_origins_of_characters_pdfium_does_not_return(tmp_path, content, form, baseline, first):
    # PDFium は続く空白の一部を文字として返さないので、その字は描画命令から補う。補った字の原点も PyMuPDF と同じにする。
    spacing = 3 if form else 0
    widths = [text_length(char, fontsize=20) + spacing for char in "a  "]
    expected = [first + sum(widths[:position]) for position in range(4)]
    read = origins(raw_pdf(tmp_path / "text.pdf", content, form=form))
    assert [char for char, _origin in read] == list("a  b")
    assert [origin[0] for _char, origin in read] == pytest.approx(expected)
    assert [origin[1] for _char, origin in read] == pytest.approx([baseline] * 4)


def test_text_spans_keep_outline_text_drawn_as_its_own_object(tmp_path):
    # 輪郭だけで描く文字の要素は、直前の塗りの文字の要素とは別に残す。除くのは、塗りと輪郭の両方で描く要素の輪郭だけである。
    def draw(page):
        page.insert_text((20, 100), "Fill", fontsize=20)
        page.insert_text((120, 100), "Outline", fontsize=20, render_mode=1)
        page.insert_text((20, 150), "Both", fontsize=20, render_mode=2)

    with open_pdf(write(tmp_path / "outline.pdf", draw)) as document:
        page = document[0]
        assert [(span.text, span.kind) for span in page.spans()] == [
            ("Fill", 0),
            ("Outline", 1),
            ("Both", 0),
            ("Both", 1),
        ]
        assert [span.text for span in page.text_spans()] == ["Fill", "Outline", "Both"]


def embedded_font(source, target, change):
    """source の PDF の埋め込んだ書体の書体記述と書体のファイルを change(書体記述, 書体のファイル) で書き換えて target に書く。"""
    import pypdf

    writer = pypdf.PdfWriter(clone_from=source)
    for font in writer.pages[0]["/Resources"]["/Font"].values():
        descriptor = font.get_object().get("/FontDescriptor")
        if descriptor is not None and "/FontFile2" in descriptor.get_object():
            descriptor = descriptor.get_object()
            change(descriptor, descriptor["/FontFile2"].get_object())
    writer.write(target)
    return target


@pytest.mark.parametrize("metrics", [None, (1079, -210)], ids=["within-1em", "over-1em"])
def test_text_spans_use_the_text_extraction_box(tmp_path, dejavu, metrics):
    # 検査に使う文字片の字の範囲は、PyMuPDF の get_text("dict") と同じく、書体のアセンダーとディセンダーの合計が 1 em 以上なら
    # そのまま使い、満たなければ 1 em に広げる。spans() の範囲は常に合計 1 em に縮める（PyMuPDF の texttrace）。
    from pypdf.generic import NameObject, NumberObject

    path = write(
        tmp_path / "box.pdf", lambda page: page.insert_text((20, 100), "Ag", fontsize=20, fontfile=dejavu.path)
    )
    if metrics is not None:

        def change(descriptor, _stream):
            descriptor[NameObject("/Ascent")] = NumberObject(metrics[0])
            descriptor[NameObject("/Descent")] = NumberObject(metrics[1])

        path = embedded_font(path, tmp_path / "tall.pdf", change)
    with open_pdf(path) as document:
        page = document[0]
        [trace], [span] = page.spans(), page.text_spans()
    total = span.ascender - span.descender
    assert (total > 1) == (metrics is not None)
    top, bottom = (span.ascender, span.descender) if total >= 1 else (span.ascender / total, span.descender / total)
    for traced, char in zip(trace.chars, span.chars):
        assert char.bbox[1] == pytest.approx(100 - top * 20, abs=1e-3)
        assert char.bbox[3] == pytest.approx(100 - bottom * 20, abs=1e-3)
        assert (char.bbox[0], char.bbox[2]) == (traced.ink[0], traced.ink[2])
        assert traced.bbox[3] - traced.bbox[1] == pytest.approx(20, abs=1e-3)


def test_text_spans_drop_characters_whose_outline_is_below_the_page(tmp_path, dejavu):
    # ページの下の外では、上端が x の高さの a は輪郭が外にあるので除き、l は輪郭の上の部分がページに入るので残す。
    path = write(
        tmp_path / "below.pdf", lambda page: page.insert_text((20, 312), "al", fontsize=20, fontfile=dejavu.path)
    )
    with open_pdf(path) as document:
        page = document[0]
        assert [span.text for span in page.spans()] == ["al"]
        assert [span.text for span in page.text_spans()] == ["l"]


def test_type0_font_with_other_cmap_uses_pdfium_values(tmp_path):
    # Identity 以外の CMap の CID の書体は、符号の長さを決めることができないので、PDFium の値で読む。
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfgen import canvas

    pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
    path = tmp_path / "cid.pdf"
    page = canvas.Canvas(str(path), pagesize=(400, 300))
    page.setFont("HeiseiKakuGo-W5", 20)
    page.drawString(20, 200, "日本語 ABC")
    page.save()
    with open_pdf(path) as document:
        assert [span.text for span in document[0].text_spans()] == ["日本語 ABC"]


def test_unreadable_embedded_font_falls_back_to_pdfium_values(tmp_path, dejavu):
    # 埋め込まれた書体の字形の表を読むことができなくても、PDF の読み取りは止めない（その書体は PDFium の値で近似する）。
    from fontTools.ttLib import TTFont

    def change(_descriptor, stream):
        data = bytearray(stream.get_data())
        table = TTFont(io.BytesIO(bytes(data))).reader.tables["glyf"]
        data[table.offset : table.offset + table.length] = b"\xff" * table.length
        stream.set_data(bytes(data))

    source = write(
        tmp_path / "font.pdf", lambda page: page.insert_text((20, 100), "abcd", fontsize=20, fontfile=dejavu.path)
    )
    with open_pdf(embedded_font(source, tmp_path / "unreadable.pdf", change)) as document:
        assert "".join(span.text for span in document[0].spans()) == "abcd"


def test_right_to_left_text_keeps_each_glyph_with_its_character(tmp_path, dejavu):
    # PDFium は右から左に書く文字を逆の順に返す。描画命令の符号と食い違うときは PDFium の値を使い、字と位置の組を保つ。
    text = "שלום"
    path = write(
        tmp_path / "rtl.pdf", lambda page: page.insert_text((20, 100), text, fontsize=20, fontfile=dejavu.path)
    )
    widths = [text_length(char, fontsize=20, fontfile=dejavu.path) for char in text]
    expected = {char: 20 + sum(widths[:position]) for position, char in enumerate(text)}
    read = dict(origins(path))
    # PDFium の値は幅の表を整数に丸めて置くので、0.05 pt の差を許す。
    assert {char: origin[0] for char, origin in read.items()} == pytest.approx(expected, abs=0.05)


def test_render_keeps_content_at_the_page_edge(tmp_path):
    # 画像化は倍率を正確に保つ。寸法と倍率の積が整数をわずかに超えても、端に接する線を失わない。
    def draw(page):
        page.draw_rect((458.3, 0, 460.8, 345.6), color=None, fill=(1, 0, 0))
        page.draw_rect((0, 343.1, 460.8, 345.6), color=None, fill=(1, 0, 0))

    path = write(tmp_path / "edge.pdf", draw, size=(460.8, 345.6))
    image = render_to_pixels(path, 1, (240, 180))
    assert image.size == (240, 180)
    for point in ((239, 90), (120, 179)):
        red, green, blue = image.getpixel(point)
        assert red > 200 and green < 100 and blue < 100, point
