"""寸法計算が指定された実フォントを使うこと、既定の書体の探し方を確かめる。

文字の値（送り幅・kern・輪郭・OS/2 の usWin の上下）を表から読む規則は、合成フォント（tests/synthetic_fonts.py）で確かめる。
"""

from concurrent.futures import ThreadPoolExecutor
import logging
import os
from pathlib import Path
import re
import struct
from types import SimpleNamespace
import unicodedata

from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTFont
import pytest

from slidekit import fonts, renderers
from slidekit.fonts import FontRegistry
from slidekit.math import metrics
from slidekit.math.office_math import text_size
from slidekit.measure import MEASURE_VERSION, FaceMetrics, FontSource, MathBox, MeasureError, layout_text
from synthetic_fonts import (
    JAPANESE_GLYPHS,
    JAPANESE_KERNING,
    LATIN_GLYPHS,
    LATIN_KERNING,
    SyntheticFonts,
    make_font,
    make_standard_fonts,
)


def project(faces, **roles):
    """faces は実行環境の fonts の形（書体名から path・bold_path）。roles はテーマの fonts。"""
    return SimpleNamespace(theme={"fonts": roles}, environment={"fonts": faces}, config={"math": {}})


@pytest.fixture
def registry(dejavu):
    config = {"path": str(dejavu.path), "bold_path": str(dejavu.bold_path)}
    return FontRegistry(
        project({"DejaVu Sans": config}, japanese="DejaVu Sans", latin="DejaVu Sans", math="DejaVu Sans")
    )


@pytest.fixture
def synthetic_faces(tmp_path):
    return make_standard_fonts(tmp_path / "fonts")


@pytest.fixture
def synthetic(synthetic_faces):
    """欧文 Test Latin と和文 Test Japanese の合成フォントを使う寸法計算器。"""
    return FontRegistry(project(synthetic_faces, japanese="Test Japanese", latin="Test Latin"))


def test_metrics_preserve_newlines_and_bold_font(registry):
    one = registry.measure_text("AB", 20)
    two = registry.measure_text("AB\nAB", 20)
    assert two[0] == one[0]
    # 明示改行で増える高さは、行間 1.0 の行送り（1.2 × 文字サイズ）。
    assert two[1] - one[1] == pytest.approx(1.2 * 20)
    assert registry.measure_text("AB", 20, bold=True)[0] > one[0]
    with pytest.raises(ValueError, match="正の有限値"):
        registry.measure_text("AB", 0)


def test_compatible_measure_text_uses_layout_model(synthetic_faces, synthetic):
    # 幅は送り幅の和で、和欧文の境界に空きを入れない（A は 0.5 em、漢は 1.0 em）。
    assert synthetic.measure_text("A漢", 20)[0] == pytest.approx(30)
    # 行末の空白は幅に含めない（PowerPoint の行の幅と同じ）。行頭の空白（0.25 em）は含める。
    assert synthetic.measure_text("A  ", 20)[0] == synthetic.measure_text("A", 20)[0] == pytest.approx(10)
    assert synthetic.measure_text(" A", 20)[0] == pytest.approx(15)
    assert synthetic.measure_text(" ", 20)[0] == 0
    # 1 行の高さは、行の書体の組（欧文と和文）の上側と下側の和で、行間 1.0 では 1.2 × 文字サイズ（配置のモデル。
    # 上下の割合は書体の OS/2 の usWin の値で決まり、和は書体によらない）。
    assert synthetic.measure_text("A漢", 20)[1] == pytest.approx(1.2 * 20)
    assert synthetic.measure_text("A", 20)[1] == pytest.approx(1.2 * 20)
    # 欧文の kern 表は 12 pt 以上で使う（A と V の対は −60/1000 em）。
    assert synthetic.measure_text("AV", 20)[0] == pytest.approx(20 * (1.0 - 0.06))
    assert synthetic.measure_text("AV", 11)[0] == pytest.approx(11)
    # layout_text を行間 1.0・折り返しなしで呼んだ結果と同じ。
    style = {
        "size_pt": 24,
        "bold": True,
        "line_spacing": 1.0,
        "word_wrap": False,
        "font_latin": "Test Latin",
        "font_japanese": "Test Japanese",
    }
    layout = layout_text([{"kind": "text", "text": "TA 漢字\nAVATAR"}], style, synthetic)
    assert synthetic.measure_text("TA 漢字\nAVATAR", 24, bold=True) == (layout.content_width, layout.content_height)
    # 欧文の文字だけの行は欧文の書体だけで置く。和文の文字を含む行には和文の役割の書体が要り、欧文の書体で
    # 代えない（漢 は欧文の書体 Test Latin にもあるが、和文の書体を探す）。
    alone = FontRegistry(project({"Test Latin": synthetic_faces["Test Latin"]}, latin="Test Latin"))
    assert alone.measure_text("AV", 20) == synthetic.measure_text("AV", 20)
    with pytest.raises(FileNotFoundError, match="実フォントが見つからない: Yu Gothic UI"):
        alone.measure_text("A漢", 20)


def test_font_source_reads_os2_hmtx_kern_and_outlines(synthetic_faces, synthetic):
    assert isinstance(synthetic, FontSource)
    reference = SyntheticFonts.from_environment(synthetic_faces)
    for family, glyphs, kerning in (
        ("Test Latin", LATIN_GLYPHS, LATIN_KERNING),
        ("Test Japanese", JAPANESE_GLYPHS, JAPANESE_KERNING),
    ):
        for bold in (False, True):
            assert synthetic.face_metrics(family, bold) == reference.face_metrics(family, bold)
            for char in glyphs:
                assert synthetic.has_glyph(char, family, bold)
                assert synthetic.advance(char, family, bold) == reference.advance(char, family, bold)
                assert synthetic.glyph_bounds(char, family, bold) == reference.glyph_bounds(char, family, bold)
            for left, right in kerning:
                assert synthetic.kerning(left, right, family, bold) == reference.kerning(left, right, family, bold) != 0
    # OS/2 の usWinAscent・usWinDescent（合成フォントでは hhea と同じ）、hmtx の送り幅、輪郭の範囲、kern 表の値を
    # unitsPerEm で割った値。
    assert synthetic.face_metrics("Test Japanese") == FaceMetrics("Test Japanese", False, 1.08, 0.25, True)
    assert synthetic.advance("A", "Test Latin", True) == 0.52
    assert synthetic.glyph_bounds("g", "Test Latin") == (-0.2, 0.5)
    assert synthetic.glyph_bounds(" ", "Test Latin") == (0.0, 0.0)
    assert synthetic.kerning("T", "A", "Test Latin") == -0.08
    assert synthetic.kerning("A", "A", "Test Latin") == 0.0


def test_face_metrics_use_os2_win_and_bounds_use_curve_extrema(tmp_path):
    # 合成フォントは OS/2 の値が hhea と同じで、字形は長方形なので、OS/2 と hhea、曲線と制御点の範囲を見分けることができない。
    # OS/2 の上下と行間を hhea と違う値にし、制御点が曲線の外に出る二次曲線の字形を置いたフォントで確かめる。
    path = make_font(
        tmp_path / "base.ttf", "Test Curve", {"A": (500, 0, 700), "o": (650, 0, 0)}, ascent=800, descent=200
    )
    with TTFont(path) as font:
        font["hhea"].lineGap = 90
        os2 = font["OS/2"]
        os2.usWinAscent, os2.usWinDescent = 1100, 400
        os2.sTypoAscender, os2.sTypoDescender, os2.sTypoLineGap = 950, -300, 120
        pen = TTGlyphPen(None)
        pen.moveTo((0, 0))
        pen.qCurveTo((300, 1000), (600, 0))  # 曲線の上端は 500（制御点は 1000）
        pen.qCurveTo((300, -600), (0, 0))  # 曲線の下端は −300（制御点は −600）
        pen.closePath()
        font["glyf"][font.getBestCmap()[ord("o")]] = pen.glyph()
        font.save(tmp_path / "curve.ttf")
    registry = FontRegistry(
        project({"Test Curve": {"path": str(tmp_path / "curve.ttf")}}, japanese="Test Curve", latin="Test Curve")
    )
    # ascent・descent は OS/2 の usWinAscent・usWinDescent（PowerPoint が行の上下に使う値）。hhea・sTypo と行間は使わない。
    assert registry.face_metrics("Test Curve") == FaceMetrics("Test Curve", False, 1.1, 0.4, False)
    # 輪郭の範囲は曲線の極値（制御点の範囲ではない）。送り幅は輪郭の幅（0.6 em）でなく hmtx の値。
    assert registry.glyph_bounds("o", "Test Curve") == pytest.approx((-0.3, 0.5))
    assert registry._glyph_box("o", "Test Curve") == pytest.approx((0.0, -0.3, 0.6, 0.5))
    assert registry.advance("o", "Test Curve") == 0.65
    assert registry.glyph_bounds("A", "Test Curve") == (0.0, 0.7)
    # OS/2 の usWin が 0 の書体だけ、hhea の値を使う。
    with TTFont(tmp_path / "curve.ttf") as font:
        font["OS/2"].usWinAscent = font["OS/2"].usWinDescent = 0
        font.save(tmp_path / "zero.ttf")
    zero = FontRegistry(
        project({"Test Curve": {"path": str(tmp_path / "zero.ttf")}}, japanese="Test Curve", latin="Test Curve")
    )
    assert zero.face_metrics("Test Curve") == FaceMetrics("Test Curve", False, 0.8, 0.2, False)


def test_glyph_missing_from_one_font(synthetic):
    # ※ は和文の書体にだけある。欧文の書体では送り幅を求めることができない。
    assert not synthetic.has_glyph("※", "Test Latin")
    with pytest.raises(MeasureError, match=r"Test Latin に字形がない文字です: '※' \(U\+203B\)") as caught:
        synthetic.advance("※", "Test Latin")
    assert caught.value.code == "font_glyph_missing"
    assert caught.value.details == {"characters": ["※"], "fonts": ["Test Latin"]}
    assert synthetic.kerning("※", "A", "Test Latin") == 0.0
    assert synthetic.glyph_bounds("※", "Test Latin") == (0.0, 0.0)


def kern_subtable(pairs, *, flags=0x01, form=0, number=None, padding=b"", length=None):
    """kern 表（版 0）の小表の bytes。format 0 の pairs は字形番号の対から値への辞書（対の順に並べる）か、
    （対、値）の列（その順に置く。同じ対を重ねることができる）。それ以外の format では中身の bytes。

    number は対の数の欄（既定は実際の対の数）、padding は中身の後ろに置くバイト列。長さの欄は、指定しなければ
    padding を含めた長さを 16 ビットで切り捨てた値にする（実フォントと同じ）。
    """
    content = pairs
    if form == 0:
        items = sorted(pairs.items()) if isinstance(pairs, dict) else list(pairs)
        content = struct.pack(">HHHH", len(items) if number is None else number, 0, 0, 0) + b"".join(
            struct.pack(">HHh", left, right, value) for (left, right), value in items
        )
    content += padding
    declared = (6 + len(content)) & 0xFFFF if length is None else length
    return struct.pack(">HHBB", 0, declared, form, flags) + content


def kern_table(*subtables):
    """版 0 の kern 表の bytes。subtables は kern_subtable の値。"""
    return struct.pack(">HH", 0, len(subtables)) + b"".join(subtables)


def test_kern_pairs_sum_horizontal_format_0_subtables():
    order = [".notdef", "A", "V", "T"]
    data = kern_table(
        kern_subtable({(1, 2): -50, (3, 1): -40}),  # 横書き
        kern_subtable(bytes(10), form=2),  # format 2 は読まない
        kern_subtable({(1, 2): -10, (9, 1): -5}),  # 加算する。範囲外の字形番号は飛ばす
        kern_subtable({(2, 3): -77}, flags=0x03),  # 最小値
        kern_subtable({(1, 3): -99}, flags=0x05),  # 交差方向
        kern_subtable({(3, 2): -66}, flags=0x00),
    )  # 縦書き
    assert fonts._kern_pairs(data, order) == {("A", "V"): -60, ("T", "A"): -40}
    # 同じ小表の中で重なる対は後の値を使い（fontTools と同じ。実フォントにもある）、小表をまたぐ対は加算する。
    data = kern_table(kern_subtable([((1, 2), -20), ((1, 2), -30), ((3, 1), -40)]), kern_subtable({(1, 2): -5}))
    assert fonts._kern_pairs(data, order) == {("A", "V"): -35, ("T", "A"): -40}
    # 対の後ろに余分なバイトがある小表は、宣言の長さで次の小表へ進む。余分なバイトは対の形でも読まない。
    data = kern_table(
        kern_subtable({(1, 2): -50}, padding=bytes(2)),
        kern_subtable({(3, 1): -40}, padding=struct.pack(">HHh", 1, 3, -99)),
        kern_subtable({(2, 1): -30}),
    )
    assert fonts._kern_pairs(data, order) == {("A", "V"): -50, ("T", "A"): -40, ("V", "A"): -30}
    # 宣言の長さが対の数から求めた長さより短い小表は、あふれ（65536 の倍数の差）でなければ、宣言の長さの中の対だけを読んで
    # 次の小表へ進む。最後の小表では長さの欄を使わず、対の数だけ読む（fontTools が小表 1 つの表を読むときと同じ）。
    data = kern_table(kern_subtable({(1, 2): -50}, number=2), kern_subtable({(3, 1): -40, (2, 1): -30}, length=20))
    assert fonts._kern_pairs(data, order) == {("A", "V"): -50, ("T", "A"): -40, ("V", "A"): -30}
    # 宣言の長さが小表の頭（14 バイト）より短い小表から後ろは読まない（同じ小表を次の小表として読み直さない）。
    assert (
        fonts._kern_pairs(kern_table(kern_subtable({(1, 2): -50}, length=0), kern_subtable({(3, 1): -40})), order) == {}
    )
    # 対の多い小表で長さの欄があふれても、対の数から長さを求めて後ろの小表まで読む。
    order = [".notdef"] + [f"g{number}" for number in range(1, 121)]
    many = {(left, right): -1 for left in range(1, 111) for right in range(1, 101)}
    pairs = fonts._kern_pairs(kern_table(kern_subtable(many), kern_subtable({(1, 2): -7})), order)
    assert len(pairs) == 11000 and pairs[("g110", "g100")] == -1 and pairs[("g1", "g2")] == -8
    # Apple の形式（版 1.0）と、短すぎる表は読まない。
    assert fonts._kern_pairs(struct.pack(">LL", 0x00010000, 1), order) == {}
    assert fonts._kern_pairs(b"\x00", order) == {}


def test_overflowing_kern_table_is_read_without_warnings(tmp_path, caplog):
    # 対が 10,922 組以上の小表は、長さの欄（16 ビット）があふれる。Calibri などの実フォントにもある。
    # 文字は、幅が曖昧でない欧文の文字にする（幅が曖昧な文字は別の区切りになり、kern を使わない）。
    letters = [char for char in map(chr, range(0x0100, 0x0250)) if unicodedata.east_asian_width(char) == "N"][:110]
    kerning = {(left, right): -10 for left in letters for right in letters[:100]}
    path = make_font(tmp_path / "pairs.ttf", "Test Pairs", {char: (500, 0, 700) for char in letters}, kerning=kerning)
    caplog.clear()
    with caplog.at_level(logging.WARNING, logger="fontTools"):
        registry = FontRegistry(project({"Test Pairs": {"path": str(path)}}, japanese="Test Pairs", latin="Test Pairs"))
        assert registry.face_metrics("Test Pairs").has_kerning
        assert registry.kerning(letters[-1], letters[99], "Test Pairs") == -0.01
        assert registry.measure_text(letters[0] + letters[1], 20)[0] == pytest.approx(20 * 0.99)
    assert [record.getMessage() for record in caplog.records if record.name.startswith("fontTools")] == []


def test_each_face_is_opened_once_and_closed(synthetic, monkeypatch):
    opened = []
    original = fonts.TTFont

    def counting(*args, **kwargs):
        font = original(*args, **kwargs)
        opened.append((Path(args[0]).name, font.reader.file))
        return font

    monkeypatch.setattr(fonts, "TTFont", counting)
    synthetic.measure_text("AV 漢字", 20)
    synthetic.measure_text("TA 漢字\nAVATAR", 24)
    assert synthetic.glyph_bounds("g", "Test Latin") == (-0.2, 0.5)
    # 計測に使った書体を 1 回ずつ開き、表を読んだらファイルを閉じる（輪郭はメモリ上の表から描く）。
    assert [name for name, _file in opened] == ["latin.ttf", "japanese.ttf"]
    assert all(file.closed for _name, file in opened)
    # 記録には計測のときに読んだ値を使う。太字の書体は記録のために名前・版・ハッシュ値だけを読む（表は読まない）。
    assert set(synthetic.inventory()["fonts"]) == {"Test Latin", "Test Japanese"}
    assert [name for name, _file in opened] == ["latin.ttf", "japanese.ttf", "japanese-bold.ttf", "latin-bold.ttf"]
    assert all(file.closed for _name, file in opened)


def test_registry_can_be_shared_between_threads(synthetic_faces):
    texts = ["AVATAR Type 漢字日本語の文章", "gjpqy ※①△☆", "TA。、「」（）", "計測幅 Lift 折返"] * 3

    def registry():
        return FontRegistry(project(synthetic_faces, japanese="Test Japanese", latin="Test Latin"))

    serial = registry()
    expected = [(serial.measure_text(text, 20), serial.measure_text(text, 20, bold=True)) for text in texts]
    shared = registry()
    with ThreadPoolExecutor(max_workers=6) as executor:
        results = list(
            executor.map(lambda text: (shared.measure_text(text, 20), shared.measure_text(text, 20, bold=True)), texts)
        )
    assert results == expected


def test_math_box_is_built_from_math_layout(synthetic, monkeypatch):
    calls = []

    class Measurer:
        def box(self, expr, size, bold=False, *, display=False, text_fonts=None):
            calls.append((expr, size, bold, display, text_fonts.key))
            return {
                "x": metrics.Box(30.0, 26.0, 6.0, (0.5, 29.0, 24.0, 5.0)),
                "y": metrics.Box(12.0, -1.0, 3.0, (0.0, 12.0, -2.0, 4.0)),
            }[expr]

    families = []
    monkeypatch.setattr(synthetic, "_math_measurer", lambda family=None: families.append(family) or Measurer())
    # 送り幅、インクの上下（根号・分数線を含む描画範囲）、行の配置に使う上下（組版上の上下）を、math.metrics の箱から作る。
    box = synthetic.math_box(
        "x", 20, True, family="Test Math", display=True, japanese="Test Japanese", latin="Test Latin"
    )
    assert box == MathBox(30.0, 24.0, 5.0, 26.0, 6.0)
    # 独立した数式の組み方と、\text を計測する文字の書体（既定はテーマの書体）を渡す。
    assert calls == [("x", 20.0, True, True, ("Test Japanese", "Test Latin"))] and families == ["Test Math"]
    # 基線より上か下にインクがない値は 0 にする。
    assert synthetic.math_box("y", 20, latin="Other Latin") == MathBox(12.0, 0.0, 4.0, 0.0, 3.0)
    assert calls[-1] == ("y", 20.0, False, False, ("Test Japanese", "Other Latin")) and families[-1] is None


def test_math_box_uses_math_font(registry):
    width, total, depth = registry.measure_math_metrics("x_i^2 + 1", 20)
    box = registry.math_box("x_i^2 + 1", 20, japanese="DejaVu Sans", latin="DejaVu Sans")
    assert box.width == width > 0
    # 全高と深さは行の配置に使う上下。インクはその中にある（この式には根号などの余白がない）。
    assert box.line_ascent + box.line_descent == pytest.approx(total) and box.line_descent == depth
    assert 0 < box.ascent <= box.line_ascent and box.descent <= box.line_descent
    # 独立した数式の組み方は、分数の分子と分母を縮めないので高くなる。
    assert registry.math_box(r"\frac{a}{b}", 20, display=True).ascent > registry.math_box(r"\frac{a}{b}", 20).ascent
    # 文字と数式を並べた行の幅は、文字の送り幅と数式の幅の和。
    style = {"size_pt": 20, "line_spacing": 1.0, "font_latin": "DejaVu Sans", "font_japanese": "DejaVu Sans"}
    layout = layout_text([{"kind": "text", "text": "a"}, {"kind": "math", "text": "x_i^2 + 1"}], style, registry)
    assert layout.content_width == pytest.approx(registry.advance("a", "DejaVu Sans") * 20 + width)


def test_math_layout_returns_ink_sides(registry):
    """math_layout は math.metrics の箱（描画の左右を含む）を返し、math_box はそれから作る。"""
    full = registry.math_layout("x^2", 20, display=True)
    box = registry.math_box("x^2", 20, display=True)
    assert (full.width, max(0.0, full.ink[2]), max(0.0, full.ink[3])) == (box.width, box.ascent, box.descent)
    assert (max(0.0, full.height), max(0.0, full.depth)) == (box.line_ascent, box.line_descent)
    # 添字で終わる式の送り幅は、描画の右端より添字の後の空き（SpaceAfterScript）以上広い
    space = registry._math_measurer().font.constant("SpaceAfterScript") * 20
    assert space > 0 and full.width - full.ink[1] >= space - 1e-9
    assert 0 <= full.ink[0] < full.ink[1]


def test_inventory_has_roles_faces_versions_and_hashes(registry):
    inventory = registry.inventory(include_unused=True)
    face = inventory["fonts"]["DejaVu Sans"]
    assert face["roles"] == ["japanese", "latin", "math"]
    assert len(face["sha256"]) == len(face["bold"]["sha256"]) == 64
    assert face["sha256"] != face["bold"]["sha256"]
    assert face["version"]
    assert face["index"] == 0
    assert "glyphs" not in face
    # 文字と数式の計測の方式の版を記録する（Pillow の計測の倍率と matplotlib の版は記録しない）。
    assert inventory["measure_version"] == MEASURE_VERSION
    assert inventory["math_engine"] == metrics.ENGINE_VERSION
    assert "measurement_scale" not in inventory


def test_text_line_metrics_count_trailing_spaces_and_ink_sides(synthetic):
    # 合成フォントの字形は送り幅いっぱいの長方形（A は 0〜0.5 em、上 0.7 em。g は下 0.2 em）。空白は 0.25 em。
    # 送り幅は行末の空白も含め、描画の左右は最初と最後の空白でない字形の範囲（文字列の左端が原点）。
    assert synthetic.text_line_metrics(" A  ", 20) == pytest.approx((25.0, 14.0, 0.0, 5.0, 15.0))
    assert synthetic.text_line_metrics("Ag", 20) == pytest.approx((20.0, 14.0, 4.0, 0.0, 20.0))
    # 書体の選び方と kern は文字の計測と同じ（和文は和文の書体、A と V の対は −0.06 em）。
    assert synthetic.text_line_metrics("漢", 20)[:3] == pytest.approx((20.0, 16.4, 1.6))
    assert synthetic.text_line_metrics("AV", 20)[0] == pytest.approx(20 * 0.94) == synthetic.measure_text("AV", 20)[0]
    assert synthetic.text_line_metrics("A", 20, True)[0] == pytest.approx(10.4)
    # 描画する文字がない文字列は、送り幅だけを持つ。
    assert synthetic.text_line_metrics(" ", 20) == pytest.approx((5.0, 0.0, 0.0, 0.0, 5.0))
    assert synthetic.text_line_metrics("", 20) == (0.0, 0.0, 0.0, 0.0, 0.0)
    with pytest.raises(MeasureError) as missing:
        synthetic.text_line_metrics("A🙂", 20)
    assert missing.value.code == "font_glyph_missing"


def test_math_metrics_use_element_text_fonts(synthetic_faces, dejavu):
    # 数式の書体は DejaVu Sans、\text の書体はテーマ（和文 Test Japanese・欧文 Test Latin）か、引数で指定した書体。
    faces = {**synthetic_faces, "DejaVu Sans": {"path": str(dejavu.path), "bold_path": str(dejavu.bold_path)}}
    registry = FontRegistry(project(faces, japanese="Test Japanese", latin="Test Latin", math="DejaVu Sans"))
    assert registry.measure_math_metrics(r"\text{A漢}", 20)[0] == pytest.approx(30.0)
    assert registry.measure_math_metrics(r"\text{A}", 20, latin="DejaVu Sans")[0] == pytest.approx(
        registry.advance("A", "DejaVu Sans") * 20
    )
    # 行の配置の上下は字形の描画範囲（A は 0〜0.7 em）。
    assert registry.measure_math_metrics(r"\text{A}", 20)[1:] == pytest.approx((14.0, 0.0))
    # 添字の中の \text は書き出しと同じ縮めた大きさ（office_math.text_size()。数式の書体によらず 73 %）で計測する。
    # DejaVu Sans の ScriptPercentScaleDown（80 %）は使わない。AA の幅は 1 em。
    assert registry.math_constant("ScriptPercentScaleDown") == pytest.approx(0.8)
    plain = registry.measure_math_metrics("x", 20)[0]
    scripted = registry.math_box(r"x_{\text{AA}}", 20)
    assert scripted.width > plain + text_size(20, 1)
    assert scripted.width < plain + text_size(20, 1) + 0.25 * 20
    # math_box も同じ書体を使う。独立した数式でも \text は文字の書体で計測する。
    assert registry.math_box(r"\text{A漢}", 20, display=True).width == pytest.approx(30.0)


def test_wrong_font_family_is_not_silently_replaced(registry, dejavu):
    registry._specifications["Unknown"] = {"path": str(dejavu.path)}
    with pytest.raises(ValueError, match="一致しない"):
        registry.measure_text("hello", 20, latin="Unknown")


def test_unavailable_glyphs_are_reported(synthetic):
    # 片方の書体にだけある文字（欧文だけの A、和文だけの ※）は報告せず、両方の書体にない文字だけを返す。
    assert synthetic.missing_glyphs("A※漢 🙂✈") == ["✈", "🙂"]
    assert synthetic.missing_glyphs("A※", japanese="Test Latin") == ["※"]
    assert synthetic.missing_glyphs("\U0010ffff", bold=True) == ["\U0010ffff"]


def test_inventory_only_requires_used_fonts(registry):
    registry.roles["math"] = "未導入の数式書体"
    assert not registry.inventory()["fonts"]
    registry.measure_text("Lipták", 20)
    assert set(registry.inventory()["fonts"]) == {"DejaVu Sans"}


def test_font_table_has_default_and_calibration_families():
    # 同梱のテーマの 3 書体、文書の 5 書体、校正用の資料の書体を、設定なしで探すことができる。
    assert {
        "Yu Gothic UI",
        "Yu Gothic",
        "Arial",
        "Cambria Math",
        "Calibri",
        "Segoe UI",
        "Meiryo",
        "BIZ UDGothic",
    } <= set(fonts.WINDOWS_FONT_FILES)
    assert fonts.WINDOWS_FONT_FILES["Cambria Math"] == ("cambria.ttc", 1, None, None)
    for regular, index, bold, bold_index in fonts.WINDOWS_FONT_FILES.values():
        assert regular and index >= 0
        assert (bold is None) == (bold_index is None)


def test_user_font_directory_is_searched_after_system(tmp_path, monkeypatch):
    system, user = (tmp_path / "system").resolve(), (tmp_path / "user").resolve()
    system.mkdir()
    user.mkdir()
    for folder, name in ((system, "arial.ttf"), (user, "arial.ttf"), (user, "arialbd.ttf"), (user, "segoeui.ttf")):
        (folder / name).write_bytes(folder.name.encode())
    monkeypatch.setattr(fonts, "_detect_font_directories", lambda environment: (system, user))
    # 同じファイル名は前（システム）のフォルダーを使い、前になければ後ろ（利用者単位）を使う。
    assert fonts.default_font("Arial", {}) == {
        "path": str(system / "arial.ttf"),
        "index": 0,
        "bold_path": str(user / "arialbd.ttf"),
        "bold_index": 0,
    }
    assert fonts.default_font("Segoe UI", {}) == {"path": str(user / "segoeui.ttf"), "index": 0}
    assert fonts.default_font("Yu Gothic UI", {}) is None
    # 候補表にない書体は探さない。
    assert fonts.default_font("DejaVu Sans", {}) is None
    # 設定のフォルダーがあれば、自動探索の結果より優先する。
    assert fonts.default_font("Arial", {"font_directories": [str(user)]})["path"] == str(user / "arial.ttf")


def test_missing_font_names_searched_folders(tmp_path):
    project = SimpleNamespace(
        theme={"fonts": {}}, environment={"font_directories": [str(tmp_path)]}, config={"math": {}}
    )
    message = (
        r'実フォントが見つからない: Arial。environment.toml の \[fonts."Arial"\] に path を指定する（探したフォルダー: '
        + re.escape(str(tmp_path))
        + "）"
    )
    with pytest.raises(FileNotFoundError, match=message):
        FontRegistry(project).measure_text("A", 20)
    with pytest.raises(FileNotFoundError, match="探したフォルダー: なし"):
        FontRegistry(SimpleNamespace(theme={"fonts": {}}, environment={}, config={"math": {}})).measure_text("A", 20)


def test_wsl_is_detected_from_distribution_or_kernel(monkeypatch):
    monkeypatch.setenv("WSL_DISTRO_NAME", "Ubuntu")
    assert fonts._is_wsl()
    monkeypatch.delenv("WSL_DISTRO_NAME")
    monkeypatch.setattr(fonts.platform, "release", lambda: "6.6.87.2-microsoft-standard-WSL2")
    assert fonts._is_wsl()
    monkeypatch.setattr(fonts.platform, "release", lambda: "6.8.0-45-generic")
    assert not fonts._is_wsl()


@pytest.mark.skipif(os.name == "nt", reason="WSL の探索の試験である")
def test_wsl_font_directories_use_powershell_folders(tmp_path, monkeypatch, font_detection):
    system = tmp_path / "Windows" / "Fonts"
    local = tmp_path / "Local"
    user = local / "Microsoft" / "Windows" / "Fonts"
    system.mkdir(parents=True)
    user.mkdir(parents=True)
    calls = []

    def context(environment):
        calls.append(environment)
        return {
            "ok": True,
            "temp": "C:\\Temp",
            "sid": "S-1-5-21",
            "fonts": "C:\\WINDOWS\\Fonts",
            "local_app_data": "C:\\Users\\example\\AppData\\Local",
        }

    windows = {"C:\\WINDOWS\\Fonts": system, "C:\\Users\\example\\AppData\\Local": local}
    monkeypatch.setattr(fonts, "_is_wsl", lambda: True)
    monkeypatch.setattr(renderers, "windows_context", context)
    monkeypatch.setattr(renderers, "_unix_path", lambda value: windows[value])
    environment = {"renderers": {"powerpoint": {"powershell": "powershell.exe"}}}
    # システム、利用者単位の順。PowerShell への問い合わせはプロセス内で 1 回だけ。
    assert fonts.font_directories(environment) == (system, user)
    assert font_detection(environment) == (system, user)
    assert calls == [environment]
    assert fonts._DETECTED == {"powershell.exe": (system, user)}
    # 存在しないフォルダーは返さない。
    fonts._DETECTED.clear()
    user.rmdir()
    assert font_detection(environment) == (system,)


@pytest.mark.skipif(os.name == "nt", reason="WSL の探索の試験である")
def test_wsl_font_directories_fall_back_without_powershell(tmp_path, monkeypatch, font_detection):
    def missing(environment):
        raise renderers.RendererError("実行ファイルが見つかりません: powershell.exe", code="renderer_missing")

    fallback = tmp_path / "Fonts"
    fallback.mkdir()
    monkeypatch.setattr(renderers, "windows_context", missing)
    monkeypatch.setattr(fonts, "WSL_FONT_DIRECTORY", fallback)
    monkeypatch.setattr(fonts, "_is_wsl", lambda: True)
    assert font_detection({}) == (fallback,)
    fonts._DETECTED.clear()
    fallback.rmdir()
    assert font_detection({}) == ()
    # WSL でも Windows でもなければ探さず、PowerShell も使わない。
    fonts._DETECTED.clear()
    monkeypatch.setattr(fonts, "_is_wsl", lambda: False)
    monkeypatch.setattr(renderers, "windows_context", lambda environment: pytest.fail("PowerShell を使った"))
    assert font_detection({}) == ()
