"""文字の配置を PowerPoint の表示と照合する校正用の資料。内容は汎用の例文だけにする。座標の単位は pt。

ページごとに、確かめる配置（1 行の幅、行送り、折り返し、図形の文字領域、数式の行、背の高い数式の行の上側と下側（3 ページ）、
上下に置くものの行、文の中の根号と添字の行（2 ページ）、揃え、書体、表、同じ文字列）を並べる。
各要素の書式（文字サイズ・行間・書体・折り返し・揃え・余白）は、テーマの既定に頼らず明示する（STYLE・FRAME_STYLES）。
枠の寸法も同じ値で求めるので、同梱のテーマの既定が変わっても、枠と書式が食い違わない。
折り返しの位置を決めたい要素は、枠の幅を context.measure_text・measure_math の計測値から求める
（例：「。」が次の行の先頭に来る幅は、直前までの幅に「。」の幅の半分を加える）。
要素の description の "tags:" に続く語は、照合データのタグになる（tools/make_layout_fixture.py）。
"overflow" のタグは枠より長い語や数式の要素に付ける（照合の試験は、計測が text_overflow を報告すれば一致を求めない）。
"""

import math

from slidekit import Deck, Emphasis, Math, Rect, Text

SIZE = 20.0
SPACING = 1.15
LATIN, JAPANESE, MATH = "Arial", "Yu Gothic UI", "Cambria Math"
BOX_PADDING = 6.0
PITCH_EM = 1.2  # 行送り ÷（文字サイズ × 行間）
TOP, BOTTOM, LEFT, RIGHT, GAP = 14.0, 530.0, 20.0, 940.0, 8.0
# 要素ごとに明示する書式。行の高さ（line_spacing_pt）を指定する要素も、line_spacing を明示する（line_spacing_pt が優先される）。
STYLE = {
    "size_pt": SIZE,
    "line_spacing": SPACING,
    "font_latin": LATIN,
    "font_japanese": JAPANESE,
    "font_math": MATH,
    "align": "left",
    "valign": "top",
}
FRAME_STYLES = {
    "text": {"word_wrap": True, "padding_pt": 0.0},
    "box": {"word_wrap": True, "padding_pt": BOX_PADDING},
    "equation": {"word_wrap": False, "padding_pt": 0.0},
}


def tags(*words):
    return "tags: " + " ".join(words)


def frame_height(lines, size=SIZE, spacing=SPACING, *, spacing_pt=None, padding=0.0, extra=0.0):
    """lines 行の文字ブロックが入る枠の高さ（0.5 pt 単位に切り上げる）。extra は下に加える余り。"""
    pitch = spacing_pt if spacing_pt is not None else PITCH_EM * spacing * size
    return math.ceil((lines * pitch + 0.35 * size + padding + extra) * 2) / 2


class Ruler:
    """要素に明示する書体（LATIN・JAPANESE・MATH）で 1 行の幅を計測する（折り返さない）。"""

    def __init__(self, context):
        self.context = context

    def text(self, value, size=SIZE, *, bold=False, latin=LATIN, japanese=JAPANESE):
        return float(self.context.measure_text(value, size, bold, latin=latin, japanese=japanese)[0])

    def math(self, expression, size=SIZE):
        return float(self.context.measure_math(expression, size, family=MATH)[0])

    def space(self, size=SIZE):
        """語の間の空白の幅（行末の空白は幅に入らないので、語の間で計測する）。"""
        return self.text("a a", size) - self.text("aa", size)

    def between(self, value, size=SIZE):
        """和文の文字の間に置いた文字の幅（全角空白も、行末でなければ幅に入る）。"""
        return self.text("あ" + value + "あ", size) - self.text("ああ", size)


class Column:
    """縦に枠を並べる列。"""

    def __init__(self, page, x, width, gap=GAP):
        self.page, self.x, self.width, self.gap, self.y = page, x, width, gap, TOP

    def take(self, height, width=None):
        width = self.width if width is None else width
        if self.y + height > BOTTOM + 0.01:
            raise ValueError(f"ページ {self.page.slide.id} の列 x={self.x} に高さ {height} の枠が入りません")
        area = Rect(self.x, self.y, width, height)
        self.y += height + self.gap
        return area


class Page:
    """1 ページ分の要素を作る。枠は列（columns）に上から並べるか、place で左から順に並べる。"""

    def __init__(self, deck, id, title, notes):
        self.slide = deck.slide(id=id, kind="blank", title=title, notes=notes)
        self.x, self.y, self.row_height = LEFT, TOP, 0.0

    def columns(self, *widths, gap=GAP):
        """左から順に列を作る。gap は列の中の枠の間隔（列の間隔は GAP）。"""
        result, x = [], LEFT
        for width in widths:
            result.append(Column(self, x, width, gap))
            x += width + GAP
        if x - GAP > RIGHT + 0.01:
            raise ValueError(f"ページ {self.slide.id} の列が幅に入りません")
        return result

    def place(self, width, height):
        """左から順に置き、右端を越えるなら次の行へ送る。"""
        if self.x > LEFT and self.x + width > RIGHT + 0.01:
            self.next_row()
        if self.y + height > BOTTOM + 0.01:
            raise ValueError(f"ページ {self.slide.id} に高さ {height} の枠が入りません")
        area = Rect(self.x, self.y, width, height)
        self.x += width + GAP
        self.row_height = max(self.row_height, height)
        return area

    def next_row(self):
        if self.x > LEFT:
            self.x, self.y, self.row_height = LEFT, self.y + self.row_height + GAP, 0.0

    @staticmethod
    def _style(kind, tag, style):
        """明示する書式（STYLE・FRAME_STYLES）に、要素ごとの指定を重ねる。"""
        result = {**STYLE, **FRAME_STYLES[kind], **style}
        if tag:
            result["description"] = tags(*tag)
        return result

    def text(self, id, content, area, *, tag=(), **style):
        return self.slide.text(id=id, text=content, area=area, **self._style("text", tag, style))

    def box(self, id, content, area, *, tag=(), **style):
        return self.slide.box(id=id, text=content, area=area, **self._style("box", tag, style))

    def equation(self, id, expression, area, *, tag=(), **style):
        return self.slide.equation(
            id=id, expression=expression, area=area, role="sample", **self._style("equation", tag, style)
        )


def widths(deck, ruler):
    page = Page(
        deck,
        "widths",
        "1 行の文字の幅",
        "文字サイズと文字の種類ごとに、1 行の幅を確かめる。折り返さない。11 pt は kern を使わない大きさである。",
    )
    left, right = page.columns(464, 448, gap=4)
    samples = {
        "kana": "あさの空気はすこし冷たい",
        "latin": "Quick brown fox jumps",
        "digits": "0123456789",
        "kerning": "AVATAR Type, LTA.",
        "mixed": "東京から Osaka まで 500 km",
        "fullwidth": "全角　空白　の　幅",
        "symbols": "※①△☆→…“”×±",
        "symbols_text": "注※ 丸① 三角△ 星☆ 矢印→ 省略… “引用” 2×3 ±1",
    }
    extra = {
        "kerning": ("kerning",),
        "symbols": ("ambiguous",),
        "symbols_text": ("ambiguous",),
        "fullwidth": ("fullwidth_space",),
    }
    for name, value in samples.items():
        page.text(
            f"s20_{name}", value, left.take(frame_height(1)), word_wrap=False, size_pt=SIZE, tag=extra.get(name, ())
        )
    page.text("s20_bold", "太字の文 Bold text", left.take(frame_height(1)), word_wrap=False, bold=True)
    page.text(
        "s20_emphasis",
        [Text("通常の文と"), Emphasis("強調の文"), Text(" and more")],
        left.take(frame_height(1)),
        word_wrap=False,
    )
    # 断片の境界をまたぐ kern の対（A|T と T|y）。
    boundary = [Text("AVA"), Text("TAR T"), Text("ype, LTA.")]
    page.text(
        "s20_kerning_runs", boundary, left.take(frame_height(1)), word_wrap=False, tag=("kerning", "run_boundary")
    )
    # 大きさか太さが違う断片の境界の対（PowerPoint も計測も詰めない。太さは右の列に置く）。
    page.text(
        "s20_kerning_size_runs",
        [Text("AVA"), Text("TAR T", size_pt=22), Text("ype, LTA.")],
        left.take(frame_height(1, 22)),
        word_wrap=False,
        tag=("kerning", "run_boundary", "run_size"),
    )
    for name in ("kerning", "latin"):
        page.text(
            f"s11_{name}",
            samples[name],
            left.take(frame_height(1, 11)),
            word_wrap=False,
            size_pt=11,
            tag=("kerning",) if name == "kerning" else (),
        )
    for size, names in (
        (12, ("kana", "latin", "kerning")),
        (14, ("kana", "latin", "kerning")),
        (28, ("kana", "kerning")),
        (36, ("kana", "kerning")),
    ):
        for name in names:
            page.text(
                f"s{size}_{name}",
                samples[name],
                right.take(frame_height(1, size)),
                word_wrap=False,
                size_pt=size,
                tag=extra.get(name, ()),
            )
    page.text(
        "s36_kerning_runs",
        boundary,
        right.take(frame_height(1, 36)),
        word_wrap=False,
        size_pt=36,
        tag=("kerning", "run_boundary"),
    )
    page.text(
        "s20_kerning_bold_runs",
        [Text("AVA"), Text("TAR T", bold=True), Text("ype, LTA.")],
        right.take(frame_height(1)),
        word_wrap=False,
        tag=("kerning", "run_boundary", "run_bold"),
    )


PITCH_TEXT = "一行目 Ag\n二行目 Bp\n三行目 Cq"


def pitch(deck):
    page = Page(
        deck,
        "pitch",
        "行送りと縦の位置",
        "3 行の明示改行で、行間、行の高さの指定、文字サイズ、縦の揃え、余白ごとに、最初の基線と行送りと最後の行の下を確かめる。",
    )
    columns = page.columns(176, 176, 176, 176, 176)
    cases = [
        # (列, 識別子, 文字サイズ, 行間, 行の高さ（pt）, 縦の揃え, 余白, 下に加える余り)
        (0, "s20_sp100", 20, 1.0, None, "top", 0, 0),
        (0, "s20_sp115", 20, 1.15, None, "top", 0, 0),
        (0, "s20_sp120", 20, 1.2, None, "top", 0, 0),
        (0, "s20_sp150", 20, 1.5, None, "top", 0, 0),
        (0, "s20_sp080", 20, 0.8, None, "top", 0, 0),
        (1, "s20_pt24", 20, None, 24, "top", 0, 0),
        (1, "s20_pt30", 20, None, 30, "top", 0, 0),
        (1, "s14_sp115", 14, 1.15, None, "top", 0, 0),
        (1, "s28_sp115", 28, 1.15, None, "top", 0, 0),
        (1, "s20_sp090", 20, 0.9, None, "top", 0, 0),
        (2, "s14_sp150", 14, 1.5, None, "top", 0, 0),
        (2, "s28_sp100", 28, 1.0, None, "top", 0, 0),
        (2, "s20_middle", 20, 1.15, None, "middle", 0, 40),
        (2, "s20_bottom", 20, 1.15, None, "bottom", 0, 40),
        (3, "s28_sp150_bottom", 28, 1.5, None, "bottom", 0, 30),
        (3, "s20_pt30_middle", 20, None, 30, "middle", 0, 40),
        (3, "s20_pad6", 20, 1.15, None, "top", 6, 0),
        (4, "s20_pad_sides", 20, 1.15, None, "top", {"left": 4, "top": 10, "right": 2, "bottom": 0}, 0),
        (4, "s20_pad_bottom", 20, 1.15, None, "bottom", {"left": 0, "top": 2, "right": 0, "bottom": 12}, 30),
        (4, "s20_pad_middle", 20, 1.15, None, "middle", {"left": 8, "top": 14, "right": 8, "bottom": 4}, 30),
        (4, "s20_sp090_bottom", 20, 0.9, None, "bottom", 0, 40),
    ]
    for column, id, size, spacing, spacing_pt, valign, padding, extra in cases:
        sides = padding if isinstance(padding, dict) else {"top": padding, "bottom": padding}
        height = frame_height(
            3,
            size,
            spacing or SPACING,
            spacing_pt=spacing_pt,
            padding=sides.get("top", 0) + sides.get("bottom", 0),
            extra=extra,
        )
        style = {"size_pt": size, "valign": valign, "padding_pt": padding}
        if spacing_pt is None:
            style["line_spacing"] = spacing
        else:
            style["line_spacing_pt"] = spacing_pt
        # 行間 1 未満の行の上側は、計測が 書体の値 − 1.2 ×（1 − 行間）と 行間 × min(書体の値, 0.9035) の大きいほうにする
        # （PowerPoint で確かめたのは 0.9 だけ。0.8 は確かめるための例）。0.8 では 1 行目の字形が枠の上へ出るので、
        # 配置の検査は text_overflow を報告する（意図したもの）。
        below = spacing is not None and spacing < 1
        page.text(
            id, PITCH_TEXT, columns[column].take(height), tag=("c4", "spacing_below_one") if below else (), **style
        )


LONG_JA = "駅前の小さな図書館では、毎週土曜日の午前に絵本の読み聞かせが開かれる。参加は自由で、予約はいらない。"

# （識別子, 前の文字列, 位置を確かめる文字, 後の文字列, タグ）。前の文字列と確かめる文字の半分が 1 行目に入る幅にする。
# 確かめる文字の幅は文の途中で計測する（全角空白は行末では幅に入らず、Ruler.text では 0 になる）。
KINSOKU_START = [
    ("period", "今日は朝から細かい雨が降っていた", "。", "午後は晴れた。", "kinsoku:line_start"),
    ("comma", "窓を開けると冷たい風が入ってきた", "、", "すぐに閉めた。", "kinsoku:line_start"),
    ("close_bracket", "店の人は「本日のおすすめです", "」", "と言った。", "kinsoku:line_start"),
    ("repeat_mark", "朝の駅には急いで歩く人", "々", "の姿が多い。", "kinsoku:line_start"),
    ("small_kana", "日曜日の朝は近所の公園でき", "ょ", "うも子どもが遊ぶ。", "kinsoku:small_kana"),
    ("long_vowel", "駅の近くの店で温かいコ", "ー", "ヒーを買った。", "kinsoku:long_vowel"),
    ("ideographic_space", "見出しの後に全角の空白を置く", "　", "例である。", "kinsoku:fullwidth_space"),
]


def wrap_ja(deck, ruler):
    page = Page(
        deck,
        "wrap_ja",
        "和文の折り返し",
        "和文の長い文を 3 つの幅で折り返す。禁則の文字が行頭や行末に来る幅で、追い出しか追い込みかを確かめる。",
    )
    total = ruler.text(LONG_JA)
    for width in (200, 260, 320):
        lines = math.ceil(total / width) + 1
        page.text(f"long_w{width}", LONG_JA, page.place(width, frame_height(lines)), tag=("c1",))
    page.next_row()
    for id, before, target, after, tag in KINSOKU_START:
        width = ruler.text(before) + 0.5 * ruler.between(target)
        page.text(f"start_{id}", before + target + after, page.place(round(width, 2), frame_height(3)), tag=("c1", tag))
    # 行末の禁則：「「」が 1 行目の最後に来る幅。
    before, rest = "店の前の古い看板には「", "準備中」と書いてあった。"
    width = ruler.text(before) + 0.5 * ruler.text(rest[0])
    page.text(
        "end_open_bracket", before + rest, page.place(round(width, 2), frame_height(3)), tag=("c1", "kinsoku:line_end")
    )


LATIN_PARAGRAPH = (
    "Small towns often keep a weekly market where farmers sell fresh vegetables, bread, "
    "and flowers to their neighbors."
)
HYPHENS = "Our well-known long-term plan uses self-service kiosks."


def wrap_latin(deck, ruler):
    page = Page(
        deck,
        "wrap_latin",
        "欧文の折り返し",
        "欧文の文を語の境界で折り返す。内側より長い語、ハイフン、スラッシュ、数字の区切りの扱いを確かめる。",
    )
    total = ruler.text(LATIN_PARAGRAPH)
    for width in (200, 300):
        lines = math.ceil(total / width) + 1
        page.text(f"paragraph_w{width}", LATIN_PARAGRAPH, page.place(width, frame_height(lines)), tag=("c1",))
    # 内側の幅を語の幅の 8 割にする。
    page.text(
        "long_word",
        "Internationalization",
        page.place(round(0.8 * ruler.text("Internationalization"), 2), frame_height(3)),
        tag=("c1", "overflow", "long_word"),
    )
    page.text(
        "long_word_sentence",
        "A word like internationalization does not fit.",
        page.place(round(0.8 * ruler.text("internationalization"), 2), frame_height(5)),
        tag=("c1", "overflow", "long_word"),
    )
    page.next_row()
    cases = [
        ("hyphen_break", HYPHENS, "Our well-known long-", "c1:hyphen"),
        ("hyphen_late", HYPHENS, "Our well-known long-term plan uses self-", "c1:hyphen"),
        ("slash", "Choose read/write mode.", "Choose read/", "c1:slash"),
        ("digits_hyphen", "Call 0123-456-789 today.", "Call 0123-", "c1:digits"),
        ("digits_comma", "The total was 1,250,000 visitors.", "The total was 1,", "c1:digits"),
    ]
    for id, value, before, tag in cases:
        width = ruler.text(before) + 0.5 * ruler.text(value[len(before)])
        page.text(id, value, page.place(round(width, 2), frame_height(4)), tag=("c1", tag))


def wrap_mixed(deck, ruler):
    page = Page(
        deck,
        "wrap_mixed",
        "和文・欧文・数式の混在",
        "和文と欧文と行内の数式が混ざる文を折り返す。数式が行末・行頭に来る幅、行の残りより広い数式、数式の直後の読点を確かめる。",
    )
    mixed = [
        Text("気温は "),
        Math("T"),
        Text(" 度で、湿度は 60% だった。The forecast said rain, but the sky stayed clear."),
    ]
    total = ruler.text(
        "気温は  度で、湿度は 60% だった。The forecast said rain, but the sky stayed clear."
    ) + ruler.math("T")
    for width in (260, 360):
        lines = math.ceil(total / width) + 1
        page.text(f"paragraph_w{width}", mixed, page.place(width, frame_height(lines)), tag=("c1",))
    value, before = "全部で120個の部品があり、15個を交換した。", "全部で120"
    width = ruler.text(before) + 0.5 * ruler.text("個")
    page.text("digits_kanji", value, page.place(round(width, 2), frame_height(6)), tag=("c1", "c1:digits_kanji"))
    page.next_row()
    # 数式が 1 行目の最後に来る幅（数式の後の空白の後で分ける）。
    expression = r"\pi r^2"
    line = ruler.text("円の面積は") + ruler.space() + ruler.math(expression)
    width = line + 0.3 * (ruler.space() + ruler.text("で"))
    page.text(
        "math_line_end",
        [Text("円の面積は "), Math(expression), Text(" で求める。")],
        page.place(round(width, 2), frame_height(4)),
        tag=("c1", "c1:math_end"),
    )
    # 行の残りより広い数式（1 行には入る）。
    for id, before, expression, after in (
        ("math_rest_ja", "面積の式は ", "a^2+b^2+c^2", " になる。"),
        ("math_rest_latin", "The sum is ", "x_1+x_2+x_3", " for all cases."),
    ):
        formula = ruler.math(expression)
        width = max(ruler.text(before) + 0.5 * formula, formula + 10)
        page.text(
            id,
            [Text(before), Math(expression), Text(after)],
            page.place(round(width, 2), frame_height(4)),
            tag=("c1", "c1:math_rest"),
        )
    # 数式の直後の読点が次の行の先頭に来る幅。
    before, expression = "数列の最後の値 ", "a_n"
    width = ruler.text(before) + ruler.math(expression) + 0.5 * ruler.text("、")
    page.text(
        "math_comma",
        [Text(before), Math(expression), Text("、つまり終わりの値を使う。")],
        page.place(round(width, 2), frame_height(4)),
        tag=("c1", "c1:math_kinsoku"),
    )
    page.text(
        "math_line_start",
        [Math("x+y"), Text(" は和を表す。足し算の結果である。")],
        page.place(200, frame_height(4)),
        tag=("c1",),
    )


LONG_EXPRESSION = "a+b+c+d+e+f+g+h+i+j+k+m+n"


def wrap_math(deck, ruler):
    page = Page(
        deck,
        "wrap_math",
        "数式の折り返し",
        "行の残りに入らない文中の数式を、式ごと次の行へ送るか、式の途中で改行するかを確かめる。"
        "枠より長い独立した数式と文中の数式の改行も確かめる。",
    )
    # 枠より長い数式は、右側を空けて置く（PowerPoint が改行しない場合のはみ出しを読むことができるように）。
    long_column, _empty, first, second = page.columns(200, 236, 222, 222)
    height = frame_height(4)
    page.text("long_inline", Math(LONG_EXPRESSION), long_column.take(height), tag=("c1", "c1:math_long", "overflow"))
    page.text(
        "long_inline_text",
        [Text("式 "), Math(LONG_EXPRESSION), Text(" を使う。")],
        long_column.take(height),
        tag=("c1", "c1:math_long", "overflow"),
    )
    page.equation(
        "long_display_wrap",
        LONG_EXPRESSION,
        long_column.take(height),
        word_wrap=True,
        tag=("c1", "c1:math_long", "overflow"),
    )
    page.equation(
        "long_display_nowrap",
        LONG_EXPRESSION,
        long_column.take(height),
        word_wrap=False,
        tag=("c1", "c1:math_long", "overflow"),
    )
    cases = [
        (first, "rest_operators", "合計は ", "a+b+c+d+e+f", " である。"),
        (first, "rest_operators_latin", "The total is ", "a+b+c+d+e+f", " here."),
        (first, "rest_letters", "積は ", "abcdefghij", " である。"),
        (second, "rest_fraction", "二つの和の比は ", r"\frac{a+b}{c+d}", " である。"),
        (second, "rest_sum", "総和は ", r"\sum_{i=1}^{n} a_i b_i", " である。"),
    ]
    for column, id, before, expression, after in cases:
        formula = ruler.math(expression)
        width = min(column.width, max(ruler.text(before) + 0.4 * formula, formula + 10))
        page.text(
            id,
            [Text(before), Math(expression), Text(after)],
            column.take(height, round(width, 2)),
            tag=("c1", "c1:math_rest"),
        )
    page.equation("display_fits_wrap", "a+b+c", second.take(height), word_wrap=True, tag=("c1",))


def explicit_wrap(deck):
    page = Page(
        deck,
        "explicit_wrap",
        "明示改行と折り返し",
        "明示改行と自動の折り返しを併用する。空の行、先頭と末尾の空の行、折り返さない要素の明示改行を確かめる。"
        "末尾の空の行と空白だけの行の書体の組も、Meiryo の下揃えで確かめる。",
    )
    ja = "最初の段落は短い。\n二つ目の段落は長いので、枠の幅で自動的に折り返される。\n\n空の行の後の段落。"
    latin = "First line.\nThe second paragraph is long enough to wrap inside the frame.\n\nAfter an empty line."
    page.text("ja_paragraphs", ja, page.place(260, frame_height(7)))
    page.text("latin_paragraphs", latin, page.place(240, frame_height(7)))
    page.text("ja_paragraphs_bottom", ja, page.place(260, frame_height(7, extra=30)), valign="bottom")
    page.next_row()
    page.text("leading_empty", "\n先頭の行が空の要素", page.place(200, frame_height(3)))
    page.text("trailing_empty", "末尾の行が空の要素\n", page.place(200, frame_height(3)))
    page.text("double_empty", "一行目\n\n\n四行目", page.place(200, frame_height(5)))
    # 末尾の空の行と、末尾の空白（U+0020）だけの行の書体の組。PowerPoint は欧文の書体だけで置き、計測もそうする。
    # 和文の書体を含む組とした場合と比べて、Meiryo では下揃えの 1 行目の基線が約 1.3 pt 違う（20 pt・行間 1.15）。
    for id, value in (("trailing_empty_meiryo", "末尾に空の行\n"), ("trailing_space_meiryo", "末尾に空白行\n ")):
        page.text(
            id,
            value,
            page.place(140, frame_height(2, extra=40)),
            font_japanese="Meiryo",
            valign="bottom",
            tag=("c4", "empty_line_fonts"),
        )
    page.next_row()
    page.text(
        "nowrap_explicit",
        "折り返さない長い行の例です。This line is long.\n二行目",
        page.place(470, frame_height(3)),
        word_wrap=False,
    )


SHAPE_TEXT = "図形の中では、文字を置ける範囲が枠より狭くなる。Text area is smaller."
SHAPE_SHORT = "楕円とひし形の文。Shape text."


def shapes(deck):
    page = Page(
        deck,
        "shapes",
        "図形の文字領域",
        "四角形・角丸・楕円・ひし形の中で長い文を折り返し、縦は中央に揃える。図形ごとの文字領域の幅を確かめる。",
    )
    style = {"size_pt": 16, "valign": "middle", "padding_pt": 6}
    # rounded_default は、照合データを測ったときの PowerPoint の図形の既定（短い辺 140 pt の 1/6）を明示する。
    # 同梱のテーマの既定（corner_radius_pt = 6）は、rounded_r4・rounded_r12 と同じく半径を指定した図形として確かめる。
    for id, shape, radius, align in (
        ("rectangle", "rectangle", None, "left"),
        ("rounded_default", "rounded_rectangle", 140 / 6, "center"),
        ("rounded_r4", "rounded_rectangle", 4, "left"),
        ("rounded_r12", "rounded_rectangle", 12, "center"),
    ):
        extra = {} if radius is None else {"corner_radius_pt": radius}
        page.box(id, SHAPE_TEXT, page.place(220, 140), shape=shape, align=align, tag=("c3",), **style, **extra)
    page.next_row()
    for id, shape, width, height, text in (
        ("ellipse_wide", "ellipse", 280, 170, SHAPE_TEXT),
        ("ellipse_tall", "ellipse", 170, 230, SHAPE_SHORT),
        ("diamond_wide", "diamond", 260, 170, SHAPE_SHORT),
        ("diamond_tall", "diamond", 180, 240, SHAPE_SHORT),
    ):
        page.box(id, text, page.place(width, height), shape=shape, align="center", tag=("c3",), **style)


def callouts(deck):
    page = Page(
        deck,
        "callouts",
        "吹き出しの文字領域",
        "長方形と角丸の本体の 4 辺にしっぽを付け、長い文を折り返す。文字領域がしっぽのない図形と同じであることを確かめる。",
    )
    style = {"size_pt": 16, "valign": "middle", "padding_pt": 6, "line_color": "183153"}
    # しっぽ（長さ 8 pt）が隣の吹き出しやページの端に届かないように、間を 24 pt あける。
    for row, (shape, radius, align) in enumerate((("rectangle", None, "left"), ("rounded_rectangle", 12, "center"))):
        for column, side in enumerate(("left", "right", "top", "bottom")):
            extra = {} if radius is None else {"corner_radius_pt": radius}
            area = Rect(LEFT + 12 + column * 228, TOP + 16 + row * 170, 200, 140)
            page.box(
                f"{shape}_{side}",
                SHAPE_TEXT,
                area,
                shape=shape,
                align=align,
                tail_side=side,
                tail_at=0.3,
                tag=("c3", "callout"),
                **style,
                **extra,
            )


MATH_ONLY = [
    ("square", "x^2"),
    ("fraction", r"\frac{a}{b}"),
    ("root", r"\sqrt{n}"),
    ("sum", r"\sum_{i=1}^{n} a_i"),
    ("paren_fraction", r"\left(\frac{1}{2}\right)"),
    ("order", r"O(n\log n)"),
]


def math_lines(deck):
    page = Page(
        deck,
        "math_lines",
        "数式の行",
        "数式だけの要素を、独立した数式と文中の数式で、行間と文字サイズを変えて置く。根号を含む行の基線の下がり方も確かめる。",
    )
    width = 146
    for name, expression in MATH_ONLY:
        page.equation(
            f"display_{name}", expression, page.place(width, 90), size_pt=24, line_spacing=1.2, tag=("c4", "math_only")
        )
    page.next_row()
    for name, expression in MATH_ONLY:
        page.text(
            f"inline_{name}",
            Math(expression),
            page.place(width, 70),
            size_pt=24,
            line_spacing=1.2,
            tag=("c4", "math_only"),
        )
    page.next_row()
    for spacing in (1.0, 1.5):
        for name, expression in MATH_ONLY[:3]:
            page.equation(
                f"display_{name}_sp{round(spacing * 100)}",
                expression,
                page.place(width, 90),
                size_pt=24,
                line_spacing=spacing,
                valign="middle",
                tag=("c4", "math_only"),
            )
    page.next_row()
    for size in (18, 28):
        for name, expression in (MATH_ONLY[2], MATH_ONLY[3]):
            page.equation(
                f"display_{name}_s{size}",
                expression,
                page.place(width, 100),
                size_pt=size,
                line_spacing=1.2,
                tag=("c4", "math_only"),
            )
    root = [Text("面積は "), Math(r"\sqrt{n}"), Text(" に比例する。")]
    page.text("root_in_text", root, page.place(300, frame_height(2)), tag=("c4", "radical"))
    page.next_row()
    page.text(
        "root_in_text_sp150",
        root,
        page.place(300, frame_height(2, spacing=1.5)),
        line_spacing=1.5,
        tag=("c4", "radical"),
    )
    page.text(
        "root_first_line",
        [Text("一行目 "), Math(r"\sqrt{x}"), Text(" の値\n二行目の文")],
        page.place(300, frame_height(3)),
        tag=("c4", "radical"),
    )
    page.text(
        "root_second_line",
        [Text("一行目の文\n二行目 "), Math(r"\sqrt{x}"), Text(" の値")],
        page.place(300, frame_height(3)),
        tag=("c4", "radical"),
    )


# 行の上側を数式で決める規則の境目。計測は、分数・伸びる括弧・大型演算子・下に置くものを含み（添字の中と上下に置くものの
# 中を除く）、インクの上下の和が文字サイズを超える式の行を、背の高い数式の規則（measure._math_line）で置く。和が 1.3 em 以上の
# 式と下に置くものを含む式はインクの上下と行間から上側と下側を決め、ほかの 1.0〜1.3 em の式は上側を 1.2 × インクの上 × 行間に
# する。文中の背の高い根号は根号の規則で置き、それ以外は通常の数式の行の規則（measure._math_line）で置く。
STACK_DISPLAY = [MATH_ONLY[1], MATH_ONLY[3], MATH_ONLY[4]]  # 分数・総和・伸びる括弧の中の分数
STACK_INLINE = [
    ("script_fraction", r"a^{\frac{a}{c}}"),  # 添字の中の分数（インクの上下の和は 1 em を少し超える）
    ("script_fraction_small", r"a^{\frac{.}{.}}"),  # 同（1 em を少し下回る）
    ("root_fraction", r"\sqrt{\frac{a}{b}}"),  # 根号の中の分数
    ("binom", r"\binom{n}{k}"),
    ("paren_square", r"\left(x^2\right)"),  # 伸びる括弧（中身は分数でない）
    ("integral", r"\int f"),  # 大型演算子（1 em を少し超える）
]
STACK_LIMITS = [("over", r"\overset{n}{x}"), ("over_dot", r"\overset{\cdot}{a}")]  # 1 em を少し超える式と下回る式


def math_stack(deck):
    page = Page(
        deck,
        "math_stack",
        "数式の行の上側",
        "分数・積んだ添字・伸びる括弧・大型演算子を含む数式の行を、行間と縦の揃えを変えて置く。"
        "インクの上下の和が 1 em 前後の式、根号の中の分数、分数の行の次の行送りも確かめる。",
    )
    width = 146
    # 独立した数式だけの行。上揃えで行の上側を、下揃えで行の下側（文字ブロックの下端）を確かめる。
    for spacing, valign, height in ((1.0, "top", 86), (1.0, "bottom", 86), (1.5, "top", 106), (1.2, "bottom", 106)):
        for name, expression in STACK_DISPLAY:
            page.equation(
                f"display_{name}_sp{round(spacing * 100)}_{valign}",
                expression,
                page.place(width, height),
                size_pt=24,
                line_spacing=spacing,
                valign=valign,
                tag=("c4", "math_stack"),
            )
    page.next_row()
    for name, expression in STACK_INLINE:
        page.text(
            f"inline_{name}",
            Math(expression),
            page.place(width, 80),
            size_pt=24,
            line_spacing=1.2,
            tag=("c4", "math_stack"),
        )
    page.next_row()
    for name, expression in [STACK_INLINE[5], *STACK_LIMITS]:
        page.text(
            f"inline_{name}_sp150",
            Math(expression),
            page.place(width, 90),
            size_pt=24,
            line_spacing=1.5,
            tag=("c4", "math_stack"),
        )
    for name, expression in (STACK_INLINE[2], STACK_INLINE[3], STACK_INLINE[5]):
        page.equation(
            f"display_{name}", expression, page.place(width, 90), size_pt=24, line_spacing=1.2, tag=("c4", "math_stack")
        )
    page.next_row()
    # 文字と分数の行。分数の行の次の行の行送りと、最後の行の分数の下側（下揃え）を確かめる。
    fraction = r"\frac{a}{b}"
    page.text(
        "fraction_first_line",
        [Text("比は "), Math(fraction), Text(" である。\n次の行の文")],
        page.place(300, frame_height(3)),
        tag=("c4", "math_stack"),
    )
    page.text(
        "fraction_last_line_bottom",
        [Text("一行目の文\n比は "), Math(fraction), Text(" である。")],
        page.place(300, frame_height(3, extra=20)),
        valign="bottom",
        tag=("c4", "math_stack"),
    )
    page.text(
        "inline_sub_fraction_bottom",
        Math(r"x_{\frac{1}{2}}"),
        page.place(width, 80),
        size_pt=24,
        line_spacing=1.2,
        valign="bottom",
        tag=("c4", "math_stack"),
    )


# math_stack で確かめていない行間・揃え・式（20 pt）。
STACK_MORE_INLINE = [
    ("scripts", r"x_i^2"),  # 添字だけの式（規則に当たらない）
    ("scripts_deep", r"x_{i,j}^{2}"),  # 同（下が深い）
    ("paren_square", r"\left(x^2\right)"),  # 伸びる括弧（インクの上下の和が 1.0〜1.3 em）
    ("binom", r"\binom{n}{k}"),  # 同
]
BASE_FRACTION = r"(\frac{a}{b})^2"  # 添字の土台の中の分数（規則に当たる）
UNDERSET = r"\underset{x \in S}{\arg\min} f(x)"  # 下に置くもので、下が深い式（規則に当たる。MEASURE_VERSION 6 から）


def math_stack_more(deck):
    page = Page(
        deck,
        "math_stack_more",
        "数式の行の上側と下側",
        "背の高い数式の行を、行間 1.5 の下揃えと中央揃え、行間 2.0、行間 1.5 の文中の分数で置く。"
        "添字だけの式と伸びる括弧の式の下揃え、添字の土台の中の分数、下が深い上下に置くものも確かめる。",
    )
    width = 146
    # 独立した数式だけの行。行間 1.5 の下揃えと中央揃えで、行間で増えた分の上下の分け方を確かめる。
    for spacing, valign, height in ((1.5, "bottom", 96), (1.5, "middle", 96), (2.0, "top", 126), (2.0, "bottom", 126)):
        for name, expression in STACK_DISPLAY:
            page.equation(
                f"display_{name}_sp{round(spacing * 100)}_{valign}",
                expression,
                page.place(width, height),
                size_pt=SIZE,
                line_spacing=spacing,
                valign=valign,
                tag=("c4", "math_stack"),
            )
    page.next_row()
    # 行間 1.5 の文中の分数。分数の行の上側、次の行の行送り、最後の行の分数の下側（下揃え）を確かめる。
    fraction = r"\frac{a}{b}"
    page.text(
        "fraction_first_line_sp150",
        [Text("比は "), Math(fraction), Text(" である。\n次の行の文")],
        page.place(300, frame_height(3, spacing=1.5)),
        line_spacing=1.5,
        tag=("c4", "math_stack"),
    )
    page.text(
        "fraction_last_line_sp150_bottom",
        [Text("一行目の文\n比は "), Math(fraction), Text(" である。")],
        page.place(300, frame_height(3, spacing=1.5, extra=10)),
        line_spacing=1.5,
        valign="bottom",
        tag=("c4", "math_stack"),
    )
    for valign in ("top", "bottom"):
        page.text(
            f"inline_fraction_sp150_{valign}",
            Math(fraction),
            page.place(width, 80),
            line_spacing=1.5,
            valign=valign,
            tag=("c4", "math_stack"),
        )
    page.next_row()
    # 規則に当たらない添字の式と、1.0〜1.3 em の伸びる括弧の式の、下揃えの下側（行間 1.2）。
    for name, expression in STACK_MORE_INLINE:
        page.text(
            f"inline_{name}_bottom",
            Math(expression),
            page.place(width, 60),
            line_spacing=1.2,
            valign="bottom",
            tag=("c4", "math_stack"),
        )
    page.equation(
        "display_scripts_bottom",
        STACK_MORE_INLINE[0][1],
        page.place(width, 60),
        line_spacing=1.2,
        valign="bottom",
        tag=("c4", "math_stack"),
    )
    page.next_row()
    # 添字の土台の中の分数（行間 1.0・1.2 の上揃え）と、下が深い上下に置くもの（行間 1.15 の独立した数式）。
    for spacing in (1.0, 1.2):
        page.text(
            f"inline_base_fraction_sp{round(spacing * 100)}",
            Math(BASE_FRACTION),
            page.place(width, 70),
            line_spacing=spacing,
            tag=("c4", "math_stack"),
        )
    page.equation(
        "display_base_fraction", BASE_FRACTION, page.place(width, 70), line_spacing=1.2, tag=("c4", "math_stack")
    )
    for valign in ("top", "middle", "bottom"):
        page.equation(
            f"display_underset_{valign}",
            UNDERSET,
            page.place(width, 70),
            line_spacing=1.15,
            valign=valign,
            tag=("c4", "math_stack"),
        )


# math_stack・math_stack_more で確かめていない行間（20 pt）。行間で増える分の上下の分け方は、行間 1.2 以下ですべて上、
# 1.5・2.0 で 4 分の 3 を上と確かめた。その間（1.3・1.4）と、文中の分数の行間 2.0、1.0〜1.3 em の式の行間 1.5・2.0 の
# 下揃えを確かめる（MEASURE_VERSION 7 の計測は、行間で増える分が 1.17 × インクの下以下ならすべて上に置く）。
STACK_SPACING_INLINE = [
    STACK_MORE_INLINE[2],
    STACK_MORE_INLINE[3],
]  # 伸びる括弧と \binom（インクの上下の和が 1.0〜1.3 em）


def math_stack_spacing(deck):
    page = Page(
        deck,
        "math_stack_spacing",
        "数式の行の行間",
        "背の高い数式の行を、行間 1.3・1.4 の上揃えと下揃え、文中の分数の行間 2.0 で置く。"
        "インクの上下の和が 1.0〜1.3 em の式の、行間 1.5・2.0 の下揃えも確かめる。",
    )
    width = 146
    for spacing in (1.3, 1.4):
        for valign in ("top", "bottom"):
            for name, expression in (MATH_ONLY[1], MATH_ONLY[3]):
                page.equation(
                    f"display_{name}_sp{round(spacing * 100)}_{valign}",
                    expression,
                    page.place(width, 100),
                    size_pt=SIZE,
                    line_spacing=spacing,
                    valign=valign,
                    tag=("c4", "math_stack"),
                )
    page.next_row()
    fraction = r"\frac{a}{b}"
    for valign in ("top", "bottom"):
        page.text(
            f"inline_fraction_sp200_{valign}",
            Math(fraction),
            page.place(width, 110),
            line_spacing=2.0,
            valign=valign,
            tag=("c4", "math_stack"),
        )
    page.text(
        "fraction_first_line_sp200",
        [Text("比は "), Math(fraction), Text(" である。\n次の行の文")],
        page.place(300, frame_height(2, spacing=2.0, extra=20)),
        line_spacing=2.0,
        tag=("c4", "math_stack"),
    )
    page.text(
        "fraction_last_line_sp200_bottom",
        [Text("一行目の文\n比は "), Math(fraction), Text(" である。")],
        page.place(300, frame_height(2, spacing=2.0, extra=20)),
        line_spacing=2.0,
        valign="bottom",
        tag=("c4", "math_stack"),
    )
    page.next_row()
    for spacing in (1.5, 2.0):
        for name, expression in STACK_SPACING_INLINE:
            page.text(
                f"inline_{name}_sp{round(spacing * 100)}_bottom",
                Math(expression),
                page.place(width, 90),
                line_spacing=spacing,
                valign="bottom",
                tag=("c4", "math_stack"),
            )


# 上下に置くもの（20 pt）。下に置くものがある式は、インクの上下の和によらず背の高い数式の規則に当たり、a に行の書体の組の
# 行間 1.0 の上側の下限を設ける（MEASURE_VERSION 7。このページの UNDERSET・SHALLOW_UNDERSET・LIMIT で確かめた）。
# 上だけに置くもの（\overset）は規則に当たらない（行間 1.5 で確かめた。行間がちょうど 1 の上側はインクの上 + 0.147 em 以上）。
SHALLOW_UNDERSET = r"\underset{i}{\max} x_i"  # 下が浅い \underset
LIMIT = r"\lim\limits_{n \to \infty} a_n"  # \limits の下の添字
OVERSET = STACK_LIMITS[0][1]  # \overset{n}{x}


def math_limits(deck):
    page = Page(
        deck,
        "math_limits",
        "上下に置くものの行",
        "下に置くもの（\\underset、\\limits の下の添字）と上に置くもの（\\overset）の式を、行間と縦の揃えを変えて置く。"
        "下が深い式と浅い式、独立した数式と文中の数式を確かめる。",
    )
    width = 146
    for valign in ("top", "bottom"):
        for spacing in (1.0, 1.5, 2.0):
            page.equation(
                f"display_underset_sp{round(spacing * 100)}_{valign}",
                UNDERSET,
                page.place(width, 100),
                size_pt=SIZE,
                line_spacing=spacing,
                valign=valign,
                tag=("c4", "math_limits"),
            )
    page.next_row()
    for valign in ("top", "middle", "bottom"):
        page.equation(
            f"display_shallow_underset_{valign}",
            SHALLOW_UNDERSET,
            page.place(width, 80),
            size_pt=SIZE,
            line_spacing=1.15,
            valign=valign,
            tag=("c4", "math_limits"),
        )
    for valign in ("top", "bottom"):
        page.text(
            f"inline_underset_{valign}",
            Math(UNDERSET),
            page.place(width, 80),
            line_spacing=1.15,
            valign=valign,
            tag=("c4", "math_limits"),
        )
    page.next_row()
    for valign in ("top", "bottom"):
        page.equation(
            f"display_limit_{valign}",
            LIMIT,
            page.place(width, 80),
            size_pt=SIZE,
            line_spacing=1.15,
            valign=valign,
            tag=("c4", "math_limits"),
        )
        page.text(
            f"inline_limit_{valign}",
            Math(LIMIT),
            page.place(width, 80),
            line_spacing=1.15,
            valign=valign,
            tag=("c4", "math_limits"),
        )
    for valign in ("top", "bottom"):
        page.text(
            f"inline_over_sp100_{valign}",
            Math(OVERSET),
            page.place(width, 80),
            line_spacing=1.0,
            valign=valign,
            tag=("c4", "math_limits"),
        )
    page.next_row()
    for valign in ("top", "bottom"):
        page.text(
            f"inline_over_sp200_{valign}",
            Math(OVERSET),
            page.place(width, 100),
            line_spacing=2.0,
            valign=valign,
            tag=("c4", "math_limits"),
        )
        page.text(
            f"text_limit_{valign}",
            [Text("極限 "), Math(LIMIT), Text(" の値")],
            page.place(width * 2 + GAP, 100),
            line_spacing=1.2,
            valign=valign,
            tag=("c4", "math_limits"),
        )


# 文の中の背の高い根号と添字の行。計測は、文中の背の高い根号（インクの上下の和が文字サイズを超え、分数などを含まない式）の
# 行の上側を 1.17 × 行間 × インクの上、最後の行の下側を d にし、規則に当たらない数式の下が書体の組の下より深い行は、
# 行の上側に 0.5 × 行間 × 深い分を加える（どちらも実際の資料の PDF の行間 1.2 で確かめた）。MEASURE_VERSION 7 では、
# このページの結果から、根号の行の行間 1.0 の上側（1.22 × インクの上）、次の行への行送り、下の深い根号の行を直した。
TALL_ROOT = r"O(\sqrt{n\log n})"  # 根号の中に上と下へ出る文字がある式（インクの上下の和が約 1.24 em）
DEEP_SCRIPT = "y_j"  # 下付きの添字の下が深い式
ROOT_TEXT = [Text("計算量 "), Math(TALL_ROOT), Text(" の例")]
SCRIPT_TEXT = [Text("各項 "), Math(DEEP_SCRIPT), Text(" を足す。")]


def math_text_lines(deck):
    page = Page(
        deck,
        "math_text_lines",
        "文の中の根号と添字の行",
        "和文と、中身の背の高い根号の行と、下が深い添字の行を、文字サイズ 18・20・24 pt と縦の揃えを変えて置く（行間 1.2）。"
        "根号と添字の行の次の行の行送りと、文の行の次の添字の行も確かめる。",
    )
    width = 300
    for size in (18, 20, 24):
        for name, value in (("root", ROOT_TEXT), ("script", SCRIPT_TEXT)):
            for valign in ("top", "middle", "bottom"):
                page.text(
                    f"{name}_s{size}_{valign}",
                    value,
                    page.place(width, frame_height(1, size, 1.2, extra=16)),
                    size_pt=size,
                    line_spacing=1.2,
                    valign=valign,
                    tag=("c4", "math_text"),
                )
            page.next_row()
    page.text(
        "root_then_text",
        [*ROOT_TEXT, Text("\n次の行の文")],
        page.place(width, frame_height(2, spacing=1.2)),
        line_spacing=1.2,
        tag=("c4", "math_text"),
    )
    page.text(
        "script_then_text",
        [*SCRIPT_TEXT, Text("\n次の行の文")],
        page.place(width, frame_height(2, spacing=1.2)),
        line_spacing=1.2,
        tag=("c4", "math_text"),
    )
    page.text(
        "text_then_script_bottom",
        [Text("一行目の文\n"), *SCRIPT_TEXT],
        page.place(width, frame_height(2, spacing=1.2, extra=10)),
        line_spacing=1.2,
        valign="bottom",
        tag=("c4", "math_text"),
    )


EA_SYMBOLS = "○□●■▲▼°§♪÷‘’"  # PowerPoint が既定で和文の書体で描く幅の曖昧な記号（計測も和文の書体）
LATIN_QUOTES = "He said “yes” and ‘no’ twice."


def math_text_spacing(deck):
    page = Page(
        deck,
        "math_text_spacing",
        "文の中の根号と添字の行間",
        "中身の背の高い根号と下が深い添字の行を、行間 1.0・1.5 の上揃えと下揃えで置く。下の深い根号、"
        "数式だけの文の添字と根号、和文の書体で描く記号と欧文の中の引用符も確かめる。",
    )
    width = 300
    for spacing in (1.0, 1.5):
        for name, value in (("root", ROOT_TEXT), ("script", SCRIPT_TEXT)):
            for valign in ("top", "bottom"):
                page.text(
                    f"{name}_sp{round(spacing * 100)}_{valign}",
                    value,
                    page.place(width, frame_height(1, spacing=spacing, extra=16)),
                    line_spacing=spacing,
                    valign=valign,
                    tag=("c4", "math_text"),
                )
    page.text(
        "deep_root_in_text",
        [Text("根号 "), Math(r"\sqrt{y_j}"), Text(" の値")],
        page.place(width, frame_height(1, spacing=1.2, extra=16)),
        line_spacing=1.2,
        tag=("c4", "math_text"),
    )
    page.next_row()
    small = 146
    for name, expression in (("script", DEEP_SCRIPT), ("script_deep", STACK_MORE_INLINE[1][1]), ("root", TALL_ROOT)):
        for valign in ("top", "bottom"):
            page.text(
                f"inline_{name}_{valign}",
                Math(expression),
                page.place(small, 60),
                line_spacing=1.2,
                valign=valign,
                tag=("c4", "math_text"),
            )
    page.next_row()
    page.text("s20_symbols_ea", EA_SYMBOLS, page.place(width, frame_height(1)), word_wrap=False, tag=("c5", "ea_font"))
    page.text(
        "s20_quotes_latin", LATIN_QUOTES, page.place(width, frame_height(1)), word_wrap=False, tag=("c5", "ea_font")
    )


SCRIPT_FRACTIONS = (("quarter", "m^{1/4}"), ("ceiling", r"2^{\lceil a/2\rceil}"))


def math_script_lines(deck):
    page = Page(
        deck,
        "math_script_lines",
        "上付きの分数の行",
        "上付きの横分数・斜線・上下の分数を、12・20・32 pt、行間 1.0・1.2 で比べる。"
        "通常文字の角括弧を目印にして、基線は 0.5 pt の許容差で照合する。",
    )
    expressions = (*SCRIPT_FRACTIONS, ("slash", r"m^{1\slash4}"), ("stack", r"x^{\frac{a}{b}}"))
    for size in (12, 20, 32):
        for spacing in (1.0, 1.2):
            for name, expression in expressions:
                page.text(
                    f"{name}_s{size}_sp{round(spacing * 100)}",
                    [Text("["), Math(expression), Text("]")],
                    page.place(146, 112),
                    size_pt=size,
                    line_spacing=spacing,
                    word_wrap=False,
                    tag=("c4", "text_baseline"),
                )


def math_script_spacing(deck):
    page = Page(
        deck,
        "math_script_spacing",
        "上付きの横分数の行間と揃え",
        "20 pt の上付きの横分数を、行間 1.15・1.5、上・中・下揃えで置く。12・20・32 pt の和文の中の行も比べる。",
    )
    for name, expression in SCRIPT_FRACTIONS:
        for spacing in (1.15, 1.5):
            for valign in ("top", "middle", "bottom"):
                page.text(
                    f"{name}_sp{round(spacing * 100)}_{valign}",
                    [Text("["), Math(expression), Text("]")],
                    page.place(300, 60),
                    line_spacing=spacing,
                    valign=valign,
                    word_wrap=False,
                    tag=("c4", "text_baseline"),
                )
    for size in (12, 20, 32):
        for name, expression in SCRIPT_FRACTIONS:
            page.text(
                f"{name}_s{size}_japanese",
                [Text("値は "), Math(expression), Text(" です。")],
                page.place(300, 60),
                size_pt=size,
                line_spacing=1.2,
                word_wrap=False,
                tag=("c4", "text_baseline"),
            )


def math_script_multiline(deck):
    page = Page(
        deck,
        "math_script_multiline",
        "上付きの横分数を含む複数行",
        "12・20・32 pt、行間 1.5 の 3 行の中央に上付きの横分数を置き、前後の行への行送りを比べる。",
    )
    for size in (12, 20, 32):
        for name, expression in SCRIPT_FRACTIONS:
            page.text(
                f"{name}_s{size}",
                [Text("First line\nValue ["), Math(expression), Text("]\nLast line")],
                page.place(300, 200),
                size_pt=size,
                line_spacing=1.5,
                word_wrap=False,
                tag=("c4", "text_baseline"),
            )


def math_nested(deck):
    page = Page(
        deck,
        "math_nested",
        "入れ子の分数と伸びる括弧",
        "高い式と上下に偏る式を囲む括弧、独立した式の入れ子の分数を比べる。枠には描画範囲の上下に余裕を持たせる。",
    )
    expressions = (
        ("numerator", r"\frac{\frac{a}{b}}{c}"),
        ("denominator", r"\frac{a}{\frac{b}{c}}"),
        ("fenced_numerator", r"\left(\frac{\frac{a}{b}}{c}\right)"),
        ("fenced_denominator", r"\left(\frac{a}{\frac{b}{c}}\right)"),
        ("fenced_root", r"\left(\sqrt{\frac{a}{b}}\right)"),
        ("fenced_sum", r"\left(\sum_{i=1}^{n} a_i\right)"),
        ("bracket_numerator", r"\left[\frac{\frac{a}{b}}{c}\right]"),
        ("deep_denominator", r"\left(\frac{x}{y^2}\right)"),
    )
    for name, expression in expressions:
        for mode in ("inline", "display"):
            area = page.place(224, 112)
            if mode == "display":
                page.equation(
                    f"{mode}_{name}", expression, area, size_pt=20, line_spacing=1.2, tag=("c4", "math_nested")
                )
            else:
                page.text(
                    f"{mode}_{name}",
                    [Text("["), Math(expression), Text("]")],
                    area,
                    size_pt=20,
                    line_spacing=1.2,
                    word_wrap=False,
                    tag=("c4", "math_nested"),
                )


ALIGN_TEXT = "揃えの確認に使う文です。Align the lines of this paragraph in different ways."
BAND_TEXTS = (("ja", "誤差の幅を確かめる短い文"), ("latin", "Check the narrow band"))


def align(deck, ruler):
    page = Page(
        deck,
        "align",
        "揃えと誤差の幅",
        "左・中央・右・両端揃えで折り返す。行末が内側の幅の ±0.1 % と ±0.3 % に来る例で、折り返しの判定の誤差の幅を確かめる。",
    )
    for value in ("left", "center", "right"):
        page.text(f"paragraph_{value}", ALIGN_TEXT, page.place(260, frame_height(4)), align=value)
    page.next_row()
    page.text("paragraph_justify", ALIGN_TEXT, page.place(260, frame_height(4)), align="justify")
    page.text(
        "justify_latin",
        "Justified text spreads the spaces between words so that each line ends at the same edge.",
        page.place(260, frame_height(6)),
        align="justify",
    )
    page.text(
        "justify_ja",
        "両端揃えでは、文字の間を少しずつ広げて、各行の終わりを右端に揃える。",
        page.place(260, frame_height(4)),
        align="justify",
    )
    for name, value in BAND_TEXTS:
        page.next_row()
        natural = ruler.text(value)
        for label, ratio in (("m03", -0.003), ("m01", -0.001), ("p01", 0.001), ("p03", 0.003)):
            # 文字の幅が内側の幅の (1 + ratio) 倍になる幅。
            page.text(
                f"band_{name}_{label}",
                value,
                page.place(round(natural / (1 + ratio), 3), frame_height(2)),
                tag=("c7", f"band:{ratio * 100:+.1f}%"),
            )


FONT_PAIRS = [
    ("Segoe UI", "Yu Gothic UI"),
    ("Calibri", "Yu Gothic UI"),
    ("Arial", "Meiryo"),
    ("Arial", "BIZ UDGothic"),
    ("Segoe UI", "Meiryo"),
    ("Calibri", "BIZ UDGothic"),
]
FONTS_TEXT = "日本語 Ag\n欧文 Bp\n記号 ※① Cq"


def fonts(deck):
    page = Page(
        deck,
        "fonts",
        "書体の違い",
        "欧文と和文の書体の組み合わせを変え、上揃えと下揃えで、最初の基線と最後の行の下を確かめる。"
        "記号の書体と、欧文の書体で測る記号の前後の kern も確かめる。",
    )
    columns = page.columns(222, 222, 222, 222)
    for number, (latin, japanese) in enumerate(FONT_PAIRS):
        column = columns[number // 2]
        key = f"{latin}_{japanese}".lower().replace(" ", "")
        for valign, extra in (("top", 0), ("bottom", 40)):
            page.text(
                f"{key}_{valign}",
                FONTS_TEXT,
                column.take(frame_height(3, extra=extra)),
                font_latin=latin,
                font_japanese=japanese,
                valign=valign,
                tag=("c4", "c5"),
            )
    page.text(
        "segoeui_meiryo_wrap",
        "メイリオと Segoe UI の組み合わせで、長い文を折り返す例です。",
        columns[3].take(frame_height(5)),
        font_latin="Segoe UI",
        font_japanese="Meiryo",
        tag=("c4",),
    )
    page.text(
        "calibri_bizudgothic_bold",
        FONTS_TEXT,
        columns[3].take(frame_height(3)),
        font_latin="Calibri",
        font_japanese="BIZ UDGothic",
        bold=True,
        tag=("c4", "c5"),
    )
    # 行間 1 未満。PowerPoint の行間 0.9 の観測では、Meiryo と Arial の和文の行（書体の値が上限 0.9035 より小さい組）の
    # 上側は 行間 × 書体の値で、計測もそうする（ほかの組は書体の値から 1.2 ×（1 − 行間）を除いた値）。
    for valign, extra in (("top", 0), ("bottom", 40)):
        area = columns[3].take(frame_height(3, spacing=0.9, extra=extra))
        page.text(
            f"arial_meiryo_sp090_{valign}",
            FONTS_TEXT,
            area,
            font_latin="Arial",
            font_japanese="Meiryo",
            line_spacing=0.9,
            valign=valign,
            tag=("c4", "spacing_below_one"),
        )
    # 欧文の文字と、kern の対（A’ と L’）を作る幅の曖昧な記号 ’ の境界。’ は PowerPoint が既定で和文の書体で描くので、
    # MEASURE_VERSION 6 から和文の書体で計測し、書き出しは forced の区切り（a:latin と a:ea の両方を和文の書体）にする。
    # 書体が違うので、PowerPoint も計測も詰めない（5 までは欧文の書体の forced で、区切りの境界をまたぐ対を詰めないことを確かめた）。
    page.text(
        "s20_kerning_forced",
        "DATA’S MALL’S",
        columns[0].take(frame_height(1)),
        word_wrap=False,
        tag=("kerning", "forced_boundary"),
    )


TABLE_ROWS = [
    ["本", "図書館で借りた本を、期限までに返す。", "3"],
    ["Map", "A paper map of the old town with walking routes.", "1"],
    ["傘", "雨の日に備えて、玄関に置いておく。", "2"],
]


def tables(deck):
    page = Page(
        deck,
        "tables",
        "表のセルの折り返し",
        "行の高さを内容から決める表で、説明の列を折り返す。列の幅が違う 2 つの表を置く。",
    )
    for id, x, widths in (("wide", 20, [100, 260, 80]), ("narrow", 500, [80, 170, 60])):
        page.slide.table(
            id=f"table_{id}",
            headers=["項目", "説明", "数量"],
            rows=TABLE_ROWS,
            row_keys=["book", "map", "umbrella"],
            column_keys=["item", "note", "count"],
            area=Rect(x, TOP, sum(widths), 400),
            column_widths=widths,
            row_heights="auto",
            role="sample",
            **STYLE,
            **FRAME_STYLES["box"],
        )


def same_text(deck):
    page = Page(
        deck,
        "same_text",
        "同じ文字列の要素",
        "同じ文字列を持つ要素を並べる。作成の順と位置の上下を入れ替えた組も置き、描画の順で割り当てることを確かめる。",
    )
    value = "同じ文字列 Same text"
    page.text("first", value, Rect(40, 40, 320, frame_height(1)), tag=("c8",))
    page.text("second", value, Rect(40, 120, 320, frame_height(1)), tag=("c8",))
    page.text("larger", value, Rect(420, 40, 420, frame_height(1, 28)), size_pt=28, tag=("c8",))
    page.box("in_box", value, Rect(420, 120, 320, 60), padding_pt=6, valign="middle", tag=("c8",))
    order = "順序の確認 Order check"
    page.text("created_first_lower", order, Rect(40, 400, 320, frame_height(1)), tag=("c8", "order"))
    page.text("created_second_upper", order, Rect(40, 300, 320, frame_height(1)), tag=("c8", "order"))


def build(context):
    deck = Deck(context)
    ruler = Ruler(context)
    widths(deck, ruler)
    pitch(deck)
    wrap_ja(deck, ruler)
    wrap_latin(deck, ruler)
    wrap_mixed(deck, ruler)
    wrap_math(deck, ruler)
    explicit_wrap(deck)
    shapes(deck)
    callouts(deck)
    math_lines(deck)
    math_stack(deck)
    math_stack_more(deck)
    math_stack_spacing(deck)
    math_limits(deck)
    math_text_lines(deck)
    math_text_spacing(deck)
    math_script_lines(deck)
    math_script_spacing(deck)
    math_script_multiline(deck)
    math_nested(deck)
    align(deck, ruler)
    fonts(deck)
    tables(deck)
    same_text(deck)
    return deck
