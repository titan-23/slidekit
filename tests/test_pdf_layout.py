"""PDF の字形と数式の線を、描画順と位置で要素へ割り当てて読むことを確かめる。PDF は試験用の補助（pdf_tools）で作る。

PDF は描画の関数を並べた順に描く（PowerPoint と同じく、要素の順に描き、数式の線はその構造の最初の字形の前に描く）。
"""

import pdf_tools
import pytest

from slidekit.pdf import open_pdf
from slidekit.pdf_layout import read_pdf, read_slide

COLOR = "17212D"
RGB = tuple(int(COLOR[i : i + 2], 16) / 255 for i in (0, 2, 4))
RED = (0.8, 0.1, 0.1)


def element(id, area, content, *, kind="text", size=20, line_spacing=1.0, **extra):
    style = {
        "size_pt": size,
        "line_spacing": line_spacing,
        "color": COLOR,
        "padding_pt": 0,
        "font_latin": "Helvetica",
        "font_japanese": "Helvetica",
        "font_math": "Times-Roman",
    }
    return {"id": id, "kind": kind, "area": list(area), "role": "body", "style": style, "content": content, **extra}


def text(value):
    return {"kind": "text", "text": value}


def math(value):
    return {"kind": "math", "text": value}


def scene(*pages):
    return {
        "schema_version": 1,
        "size": [720, 540],
        "slides": [
            {"id": f"page{number}", "number": number, "elements": list(elements)}
            for number, elements in enumerate(pages, 1)
        ],
    }


def write_pdf(path, *pages):
    """ページごとに、描画の関数を並べた順に描いた PDF を作る。"""
    with pdf_tools.Document() as document:
        for commands in pages:
            page = document.new_page(width=720, height=540)
            for command in commands:
                command(page)
        document.save(path)
    return path


def glyphs(x, y, value, *, font="helv", size=20, **options):
    return lambda page: page.insert_text((x, y), value, fontname=font, fontsize=size, color=RGB, **options)


def fill(rect, color=RGB):
    return lambda page: page.draw_rect(rect, color=None, fill=color)


def invisible(rect, color=RGB):
    """線の不透明度が 0 の矩形（説明のある要素に PowerPoint が加える見えない矩形と同じ形）。"""
    return lambda page: page.draw_rect(rect, color=color, width=0.14, stroke_opacity=0)


def polygon(points, color=RGB):
    def draw(page):
        shape = page.new_shape()
        shape.draw_polyline(points)
        shape.finish(color=None, fill=color, closePath=True)
        shape.commit()

    return draw


def stroke(start, end):
    return lambda page: page.draw_line(start, end, color=RGB, width=1)


def width(value, font="helv", size=20):
    return pdf_tools.text_length(value, fontname=font, fontsize=size)


def flat(lines):
    return [value for line in lines for value in (line.baseline, line.left, line.right)]


def test_text_is_attributed_in_drawing_order_even_for_identical_strings(tmp_path):
    # first の文字は枠からはみ出して second の枠の中に描かれるが、描画順で first に割り当てる。
    data = scene(
        [
            element("first", (40, 60, 400, 40), [text("Same words")]),
            element("second", (40, 140, 400, 40), [text("Same words")]),
            element("third", (40, 260, 400, 60), [text("Two\nlines")]),
        ]
    )
    path = write_pdf(
        tmp_path / "deck.pdf",
        [
            glyphs(44, 176, "Same words"),
            glyphs(44, 164, "Same words"),
            glyphs(44, 290, "Two"),
            glyphs(44, 314, "lines"),
        ],
    )
    reading = read_pdf(path, data)[0]
    assert reading.unassigned == 0
    first, second, third = (reading.elements[key] for key in ("first", "second", "third"))
    assert (first.glyphs, second.glyphs, third.glyphs) == (9, 9, 8)
    assert flat(first.lines) == pytest.approx([176, 44, 44 + width("Same words")])
    assert flat(second.lines) == pytest.approx([164, 44, 44 + width("Same words")])
    assert flat(third.lines) == pytest.approx([290, 44, 44 + width("Two"), 314, 44, 44 + width("lines")])
    assert first.ink[0] == pytest.approx(44) and first.ink[2] == pytest.approx(44 + width("Same words"))
    assert first.ink[1] < 176 < first.ink[3] < 180
    assert (first.drawings, first.sizes, first.fonts["S"]) == (0, (20.0,), "Helvetica")


def test_math_glyphs_and_drawings_are_attributed_by_position(tmp_path):
    data = scene(
        [
            element("intro", (40, 20, 300, 40), [text("Intro")]),
            element("left", (40, 100, 200, 60), [math("a")], kind="equation"),
            element("right", (300, 100, 200, 60), [math("b+c")], kind="equation"),
            element("marker", (470, 105, 10, 10), [], kind="box"),
            element("pointer", (420, 150, 60, 0), [], kind="arrow", start=[420, 150], end=[480, 150]),
            element("outro", (40, 300, 300, 40), [text("Outro")]),
        ]
    )
    path = write_pdf(
        tmp_path / "deck.pdf",
        [
            glyphs(44, 44, "Intro"),
            # right を left より先に描く。数式の字形は前後の通常文字の要素の範囲で、位置から要素を決める。
            glyphs(310, 140, "b", font="tiro"),
            fill((330, 118, 334, 142)),  # 字形の間の、文字色の線（含める）
            fill((340, 110, 344, 142), RED),  # 字形の間でも、文字色と違う（含めない）
            glyphs(330, 140, "+c", font="tiro"),
            fill((470, 105, 480, 115)),  # 直後でも、要素の枠と同じ矩形（含めない）
            polygon([(476, 148), (480, 150), (476, 152)]),  # 直後でも、矢印の描画範囲の中（含めない）
            stroke((420, 150), (476, 150)),
            glyphs(243, 140, "a", font="tiro"),  # 枠の外だが 1 em 以内
            fill((60, 150, 100, 151)),  # 直後の、文字色の線（含める）
            glyphs(44, 324, "Outro"),
        ],
    )
    reading = read_pdf(path, data)[0]
    assert reading.unassigned == 0
    assert set(reading.elements) == {"intro", "left", "right", "outro"}
    left, right = reading.elements["left"], reading.elements["right"]
    assert (left.glyphs, left.drawings, len(left.lines)) == (1, 1, 1)
    assert (right.glyphs, right.drawings, len(right.lines)) == (3, 1, 1)
    # 赤い描画（上端 110）、目印の枠（上端 105）、矢じり（右端 480）を含めない。
    assert right.ink == pytest.approx((310, 118, 330 + width("+c", "tiro"), 142), abs=0.5)
    assert left.ink[0] == pytest.approx(60) and left.ink[3] == pytest.approx(151)
    assert flat(right.lines) == pytest.approx([140, 310, 330 + width("+c", "tiro")])


def test_math_glyphs_are_searched_between_neighbouring_text(tmp_path):
    # wide の枠は narrow の字形も含むが、narrow の字形は前後の通常文字の要素の範囲（between〜after）で探す。
    data = scene(
        [
            element("wide", (40, 100, 600, 300), [math("u")], kind="equation"),
            element("between", (40, 420, 300, 40), [text("Between")]),
            element("narrow", (300, 300, 100, 40), [math("v")], kind="equation"),
            element("after", (40, 470, 300, 40), [text("After")]),
        ]
    )
    path = write_pdf(
        tmp_path / "deck.pdf",
        [
            glyphs(50, 130, "u", font="tiro"),
            glyphs(44, 444, "Between"),
            glyphs(310, 330, "v", font="tiro"),
            glyphs(44, 494, "After"),
        ],
    )
    reading = read_pdf(path, data)[0]
    assert reading.unassigned == 0
    counts = {key: value.glyphs for key, value in reading.elements.items()}
    assert counts == {"wide": 1, "between": 7, "narrow": 1, "after": 5}
    assert [line.baseline for line in reading.elements["narrow"].lines] == pytest.approx([330])


def test_math_glyph_just_outside_frames_goes_to_the_nearest_frame(tmp_path):
    # 下揃えの要素の分母や下の添字は、枠の下端をわずかに出て描かれる。1 em 広げた left の枠は right の枠と重なり、
    # right の b の中心も含むが、中心から元の枠までの距離が最も小さい right に割り当てる（候補の順で先の left にしない）。
    data = scene(
        [
            element("left", (40, 100, 146, 60), [math("a")], kind="equation"),
            element("right", (194, 100, 146, 60), [math("b")], kind="equation"),
        ]
    )
    path = write_pdf(
        tmp_path / "deck.pdf",
        [
            glyphs(170, 159.9, "a", font="tiro"),  # left の枠の中
            glyphs(196, 160.1, "b", font="tiro"),  # right の枠の下端より 0.1 pt 下（left の枠から 1 em 以内）
            glyphs(184, 161, "c", font="tiro"),  # 両方の枠の外。中心は left の枠に近い
        ],
    )
    assert 194 + width("b", "tiro") / 2 < 186 + 20
    reading = read_pdf(path, data)[0]
    assert reading.unassigned == 0
    left, right = reading.elements["left"], reading.elements["right"]
    assert (left.glyphs, right.glyphs) == (2, 1)
    assert flat(right.lines) == pytest.approx([160.1, 196, 196 + width("b", "tiro")])
    assert [line.left for line in left.lines] == pytest.approx([170])


def test_line_count_ignores_fraction_parts(tmp_path):
    data = scene(
        [
            element("sum", (40, 100, 400, 60), [text("Sum = "), math(r"\frac{a}{b}")]),
            element("ratio", (40, 220, 300, 60), [math(r"x = \frac{a}{b}")], kind="equation"),
            element("two", (40, 320, 300, 80), [text("First\n"), math("y^2")]),
        ]
    )
    # 行送りは 24 pt。ratio の分子・分母は主な基線から 16 pt（行送りの 0.6 倍より遠く、0.9 倍より近い）離れる。
    path = write_pdf(
        tmp_path / "deck.pdf",
        [
            glyphs(44, 140, "Sum ="),
            glyphs(110, 129, "a", font="tiro", size=14),
            fill((108, 133, 122, 134)),
            glyphs(110, 151, "b", font="tiro", size=14),
            glyphs(44, 250, "x =", font="tiro"),
            glyphs(80, 234, "a", font="tiro", size=14),
            fill((78, 244, 92, 245)),
            glyphs(80, 266, "b", font="tiro", size=14),
            glyphs(44, 340, "First"),
            glyphs(44, 364, "y", font="tiro"),
            glyphs(54, 357, "2", font="tiro", size=14),
        ],
    )
    reading = read_pdf(path, data)[0]
    assert reading.unassigned == 0
    lines = {key: [line.baseline for line in value.lines] for key, value in reading.elements.items()}
    # 分子・分母は通常文字の行か、数式のまとまりの主な基線の行に含める。数式だけの 2 行目は別の行とみなす。
    assert lines == {"sum": pytest.approx([140]), "ratio": pytest.approx([250]), "two": pytest.approx([340, 364])}
    assert [reading.elements[key].drawings for key in ("sum", "ratio", "two")] == [1, 1, 0]
    assert reading.elements["sum"].sizes == (14.0, 20.0)


def test_unassigned_glyphs_are_counted(tmp_path):
    data = scene(
        [
            element("title", (40, 20, 600, 40), [text("Overview")]),
            element("formula", (40, 200, 200, 40), [math("z")], kind="equation"),
            element("body", (40, 100, 600, 40), [text("Details here")]),
        ]
    )
    path = write_pdf(
        tmp_path / "deck.pdf",
        [
            glyphs(600, 520, "Draft"),  # 雛形の図形の文字
            glyphs(44, 44, "Overview"),
            glyphs(44, 230, "z", font="tiro"),
            glyphs(44, 124, "Details here"),
            glyphs(650, 300, "z", font="tiro"),  # どの数式の枠にも入らない数式の字形
        ],
    )
    reading = read_pdf(path, data)[0]
    assert reading.unassigned == len("Draft") + 1
    assert {key: value.glyphs for key, value in reading.elements.items()} == {"title": 8, "formula": 1, "body": 11}
    # 数式の書体を指定しなければ、数式の字形は通常文字として照合され、割り当てられない。
    with open_pdf(path) as document:
        plain = read_slide(document[0], data["slides"][0], math_fonts=())
    assert plain.unassigned == len("Draft") + 2 and plain.elements["formula"].glyphs == 0


def test_rotated_and_rasterized_elements_are_not_read(tmp_path):
    data = scene(
        [
            element("turned", (100, 100, 200, 40), [text("Turned")], rotation=90),
            element("formula", (40, 300, 300, 40), [text("Area "), math("r^2")]),
            element("plain", (40, 400, 300, 40), [text("Plain")]),
        ]
    )
    standard = write_pdf(
        tmp_path / "standard.pdf",
        [
            glyphs(200, 160, "Turned", rotate=90),
            glyphs(44, 330, "Area"),
            glyphs(100, 330, "r", font="tiro"),
            glyphs(108, 322, "2", font="tiro", size=14),
            glyphs(44, 430, "Plain"),
        ],
    )
    # 互換表示では、数式を含む要素は画像（ここでは塗り）になり、字形がない。
    fallback = write_pdf(
        tmp_path / "fallback.pdf",
        [
            glyphs(200, 160, "Turned", rotate=90),
            fill((40, 300, 340, 340), (0.9, 0.9, 0.9)),
            glyphs(44, 430, "Plain"),
        ],
    )
    reading = read_pdf(standard, data)[0]
    assert set(reading.elements) == {"formula", "plain"} and reading.unassigned == 0
    assert (reading.elements["formula"].glyphs, len(reading.elements["formula"].lines)) == (6, 1)
    rasterized = read_pdf(fallback, data, rasterized_math=True)[0]
    assert set(rasterized.elements) == {"plain"} and rasterized.unassigned == 0
    # 数式を画像にしない読み方では、字形が見つからない要素も結果に含める。
    missing = read_pdf(fallback, data)[0].elements["formula"]
    assert (missing.glyphs, missing.lines, missing.ink, missing.drawings, missing.sizes) == (0, (), None, 0, ())


def test_text_in_math_is_attributed_to_the_math(tmp_path):
    # 数式の \text{...} は文字用の書体で描かれる。後の要素の同じ文字列と取り違えず、数式の一部として行に含める。
    data = scene(
        [
            element("case", (40, 100, 300, 40), [math(r"x \text{ if } y")], kind="equation"),
            element("word", (40, 200, 300, 40), [text("if")]),
        ]
    )
    path = write_pdf(
        tmp_path / "deck.pdf",
        [
            glyphs(44, 130, "x", font="tiro"),
            glyphs(60, 130, "if"),
            glyphs(80, 130, "y", font="tiro"),
            glyphs(44, 230, "if"),
        ],
    )
    reading = read_pdf(path, data)[0]
    case, word = reading.elements["case"], reading.elements["word"]
    assert reading.unassigned == 0
    assert (case.glyphs, len(case.lines), case.fonts) == (4, 1, {"i": "Helvetica", "f": "Helvetica"})
    assert (case.math_glyphs, word.math_glyphs) == (4, 0)  # 数式の書体の字形と \text の字形を数式の一部に含める
    assert word.glyphs == 2
    assert [line.baseline for line in word.lines] == pytest.approx([230])


def test_invisible_and_doubly_painted_glyphs_are_ignored(tmp_path):
    data = scene([element("bold", (40, 100, 300, 40), [text("Bold")])])
    path = write_pdf(
        tmp_path / "deck.pdf",
        [
            glyphs(
                44, 130, "Bold", render_mode=2, border_width=0.05
            ),  # 塗りと輪郭で同じ字形を 2 回描く（太字の模倣など）
            glyphs(44, 230, "Hidden", render_mode=3),  # 見えない文字
        ],
    )
    reading = read_pdf(path, data)[0]
    assert reading.unassigned == 0
    bold = reading.elements["bold"]
    assert (bold.glyphs, len(bold.lines), bold.sizes) == (4, 1, (20.0,))


def test_read_pdf_reads_each_page_and_checks_the_page_count(tmp_path):
    data = scene(
        [element("one", (40, 100, 300, 40), [text("One")])], [element("two", (40, 100, 300, 40), [text("Two")])]
    )
    path = write_pdf(tmp_path / "deck.pdf", [glyphs(44, 130, "One")], [glyphs(44, 130, "Two")])
    readings = read_pdf(path, data)
    assert [list(reading.elements) for reading in readings] == [["one"], ["two"]]
    assert [reading.elements[key].glyphs for reading, key in zip(readings, ("one", "two"))] == [3, 3]
    with pytest.raises(ValueError, match="ページ数"):
        read_pdf(path, scene([element("one", (40, 100, 300, 40), [text("One")])]))


def test_math_lines_are_drawn_just_before_the_first_glyph(tmp_path):
    """PowerPoint は数式の線をその構造の最初の字形の 1〜2 つ前に描く。最初の字形の 2 つ前から（前の要素の最後の字形より後に
    限る）割り当て、前の数式の要素の直後にも当たる線は、次の要素の前の線として扱う。"""
    data = scene(
        [
            element("upper", (40, 100, 200, 40), [math("a")], kind="equation"),
            element("lower", (40, 145, 200, 40), [math(r"\sqrt{b}")], kind="equation"),
            element("intro", (40, 300, 300, 40), [text("Intro")]),
            element("third", (40, 350, 200, 40), [math(r"\frac{c}{d}")], kind="equation"),
        ]
    )
    path = write_pdf(
        tmp_path / "deck.pdf",
        [
            glyphs(44, 130, "a", font="tiro"),
            fill((44, 146, 70, 147)),  # upper の直後で枠の 0.5 em 以内だが、lower の最初の字形の 2 つ前（lower の線）
            fill((44, 147, 46, 172)),  # lower の最初の字形の 1 つ前
            glyphs(50, 170, "b", font="tiro"),
            fill((44, 352, 60, 353)),  # third の 2 つ前だが、前の要素 intro の最後の字形より前（含めない）
            glyphs(44, 324, "Intro"),
            glyphs(44, 370, "c", font="tiro"),
        ],
    )
    reading = read_pdf(path, data)[0]
    assert reading.unassigned == 0
    upper, lower, third = (reading.elements[key] for key in ("upper", "lower", "third"))
    assert (upper.drawings, lower.drawings, third.drawings) == (0, 2, 0)
    assert lower.ink[:2] == pytest.approx((44, 146))
    assert upper.ink[3] < 140


def test_outer_lines_of_nested_math_are_followed_back(tmp_path):
    """入れ子の構造では外側の線（伸びた括弧）が最初の字形の 3 つ以上前に描かれる。要素の線の条件に合う描画が続く間は
    さかのぼって割り当て、要素の線の条件に合わない描画（色の違う飾りなど）で止める。"""
    data = scene(
        [
            element("intro", (40, 40, 300, 40), [text("Intro")]),
            element("nested", (40, 100, 200, 60), [math(r"\left(\frac{\sqrt{x}}{y}\right)")], kind="equation"),
        ]
    )
    path = write_pdf(
        tmp_path / "deck.pdf",
        [
            glyphs(44, 64, "Intro"),
            fill((42, 96, 44, 158)),  # 色は同じだが、間に色の違う描画があるので含めない
            fill((40, 162, 240, 164), RED),  # 飾り（色が違う）：ここで止める
            fill((44, 98, 50, 158)),  # 括弧：最初の字形の 3 つ前
            fill((52, 128, 80, 129)),  # 分数線：2 つ前
            fill((54, 104, 80, 106)),  # 根号：1 つ前
            glyphs(60, 124, "x", font="tiro"),
            glyphs(60, 150, "y", font="tiro"),
        ],
    )
    reading = read_pdf(path, data)[0]
    nested = reading.elements["nested"]
    assert reading.unassigned == 0 and nested.drawings == 3
    assert nested.ink[:2] == pytest.approx((44, 98))


def test_invisible_drawings_are_not_math_lines(tmp_path):
    """不透明度 0 の描画（説明のある要素の見えない矩形）は、文字色と同じでも数式の線にしない。"""
    data = scene([element("formula", (40, 100, 300, 60), [math(r"\frac{a}{b}")], kind="equation", description="分数")])
    path = write_pdf(
        tmp_path / "deck.pdf",
        [
            invisible((42, 102, 300, 150)),
            fill((50, 128, 64, 129)),
            glyphs(52, 124, "a", font="tiro", size=14),
            glyphs(52, 144, "b", font="tiro", size=14),
        ],
    )
    formula = read_pdf(path, data)[0].elements["formula"]
    assert formula.drawings == 1
    assert formula.ink[0] == pytest.approx(50) and formula.ink[2] < 70


def test_display_fraction_is_one_line_but_wrapped_math_is_two(tmp_path):
    """独立した数式の分数は分子と分母の基線の差が行送りに近いが、間に分数線があるので 1 行とみなす。
    PowerPoint が改行した数式の次の行（分数線がない）は別の行とみなす。"""
    data = scene(
        [
            element("ratio", (40, 100, 200, 80), [math(r"\frac{a}{b}")], kind="equation", line_spacing=1.2),
            element("long", (40, 200, 200, 80), [math("a+b+c+d")], kind="equation", line_spacing=1.2),
            element("apart", (300, 100, 300, 80), [math(r"\frac{a}{b}")], kind="equation", line_spacing=1.2),
        ]
    )
    # 行送りは 1.2 × 1.2 × 20 = 28.8 pt。分子と分母の基線の差 28 pt はその 0.9 倍（25.92 pt）より大きい。
    path = write_pdf(
        tmp_path / "deck.pdf",
        [
            fill((48, 127, 62, 128.5)),
            glyphs(50, 120, "a", font="tiro"),
            glyphs(50, 148, "b", font="tiro"),
            glyphs(44, 230, "a+b+", font="tiro"),
            glyphs(44, 258.8, "c+d", font="tiro"),
            fill((400, 127, 414, 128.5)),
            glyphs(310, 120, "a", font="tiro"),
            glyphs(310, 148, "b", font="tiro"),
        ],
    )
    reading = read_pdf(path, data)[0]
    ratio, long, apart = (reading.elements[key] for key in ("ratio", "long", "apart"))
    assert (len(ratio.lines), ratio.drawings) == (1, 1)
    assert (ratio.lines[0].left, ratio.lines[0].right) == pytest.approx((48, 62))  # 分数線を行の範囲に含める
    assert [line.baseline for line in long.lines] == pytest.approx([230, 258.8])
    # 分数線が分子と分母に横に重ならなければ、分数とみなさない。
    assert len(apart.lines) == 2


def parentheses(*rects):
    """複数の部分からなる 1 つの描画（PowerPoint は伸びた左右の括弧を 1 つの描画にする）。"""

    def draw(page):
        shape = page.new_shape()
        for rect in rects:
            shape.draw_rect(rect)
        shape.finish(color=None, fill=RGB)
        shape.commit()

    return draw


def test_rows_enclosed_by_stretched_parentheses_are_one_line(tmp_path):
    """独立した数式の \\binom は、分数線のない 2 段の基線の差が行送りの 0.9 倍を超えるが、伸びた括弧（1 つの描画）が
    2 段を縦に覆い、横に挟むので 1 行とみなす。挟まない描画（片側だけ、縦に覆わない）では 2 行とみなす。"""
    data = scene(
        [
            element("binom", (40, 100, 200, 80), [math(r"\binom{n}{k}")], kind="equation", line_spacing=1.2),
            element("side", (300, 100, 200, 80), [math(r"\binom{n}{k}")], kind="equation", line_spacing=1.2),
            element("short", (40, 300, 200, 80), [math(r"\binom{n}{k}")], kind="equation", line_spacing=1.2),
            element("bare", (300, 300, 200, 80), [math(r"\binom{n}{k}")], kind="equation", line_spacing=1.2),
        ]
    )
    # 行送りは 1.2 × 1.2 × 20 = 28.8 pt。2 段の基線の差 34 pt はその 0.9 倍（25.92 pt）より大きい。
    path = write_pdf(
        tmp_path / "deck.pdf",
        [
            parentheses((44, 98, 48, 158), (76, 98, 80, 158)),
            glyphs(54, 116, "n", font="tiro"),
            glyphs(54, 150, "k", font="tiro"),
            parentheses((304, 98, 308, 158)),
            glyphs(314, 116, "n", font="tiro"),
            glyphs(314, 150, "k", font="tiro"),
            parentheses((44, 298, 48, 320), (76, 298, 80, 320)),
            glyphs(54, 316, "n", font="tiro"),
            glyphs(54, 350, "k", font="tiro"),
            glyphs(314, 316, "n", font="tiro"),
            glyphs(314, 350, "k", font="tiro"),
        ],
    )
    reading = read_pdf(path, data)[0]
    binom, side, short, bare = (reading.elements[key] for key in ("binom", "side", "short", "bare"))
    assert (len(binom.lines), binom.drawings) == (1, 1)
    assert (binom.lines[0].left, binom.lines[0].right) == pytest.approx((44, 80))  # 括弧を行の範囲に含める
    assert (len(side.lines), side.drawings) == (2, 1)
    assert (len(short.lines), short.drawings) == (2, 1)
    assert [line.baseline for line in bare.lines] == pytest.approx([316, 350])


def test_math_line_extents_include_radical_sign(tmp_path):
    """根号の記号は線で描かれ、字形でない。行の範囲（左端・右端）に、その行の数式の線を含める。"""
    data = scene(
        [
            element("root", (40, 100, 200, 40), [text("Root "), math(r"\sqrt{n}")]),
            element("two", (40, 200, 200, 80), [math(r"\sqrt{x}"), text("\nNext")]),
        ]
    )
    path = write_pdf(
        tmp_path / "deck.pdf",
        [
            glyphs(44, 130, "Root"),
            polygon([(90, 124), (94, 134), (100, 112), (120, 112), (120, 113), (100, 113), (94, 136), (89, 125)]),
            glyphs(102, 130, "n", font="tiro"),
            polygon([(44, 224), (48, 234), (54, 212), (70, 212), (70, 213), (54, 213), (48, 236), (43, 225)]),
            glyphs(56, 230, "x", font="tiro"),
            glyphs(44, 254, "Next"),
        ],
    )
    reading = read_pdf(path, data)[0]
    root, two = reading.elements["root"], reading.elements["two"]
    assert root.drawings == 1 and len(root.lines) == 1
    assert (root.lines[0].left, root.lines[0].right) == pytest.approx((44, 120))
    # 2 行の要素では、線を縦の位置が近い行（1 行目）にだけ含める。
    assert two.drawings == 1 and len(two.lines) == 2
    assert two.lines[0].left == pytest.approx(43) and two.lines[0].right == pytest.approx(70)
    assert two.lines[1].left == pytest.approx(44) and two.lines[1].right == pytest.approx(44 + width("Next"))
