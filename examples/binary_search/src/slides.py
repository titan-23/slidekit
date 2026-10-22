"""二分探索の考え方を説明する 4:3 の 10 枚の資料。座標の単位はポイント。

雛形、表紙、章扉の飾り、分岐図、計算から作る状態図、配列の区間の図、編集可能な数式、準備済みの図、
数式を含む表、文献一覧を示す。ページは章ごとの関数で作り、部分の検査・生成（--slides）では、
選んだページを含む章だけを作る（context.selected）。
"""

from slidekit import Deck, Emphasis, Math, Rect, Text, align, distribute, key_id
from slidekit.markup import parse

# 図を作るスクリプトと同じ座標の計算で、図に重ねるラベルを置く。
from make_figure import FIGURE_SIZE, LARGEST, binary_comparisons, plot_point

# 具体例の配列と探す値。4 ページ目の表と状態図はこの値から計算する。
VALUES = (2, 3, 5, 7, 11, 13, 17, 19)
TARGET = 13
# 文献一覧に載せる文献キー。
REFERENCES = ("knuth1998", "cormen2022")


def trace(values, target):
    """x 以上の最初の位置を探し、各比較の (l, r, m) と、終了時の l を返す。"""
    left, right, steps = 0, len(values), []
    while left < right:
        middle = (left + right) // 2
        steps.append((left, right, middle))
        if values[middle] < target:
            left = middle + 1
        else:
            right = middle
    return steps, left


def section_page(deck, *, id, title, description, notes, seconds):
    """章扉を作る。題名の下の線とひし形は、テーマのページ種別 section の飾りが置く。"""
    slide = deck.slide(id=id, kind="section", title=title, notes=notes, seconds=seconds)
    content = slide.content
    slide.text(
        id="description", area=Rect(content.x, content.y, content.w, None), text=description, role="note", align="right"
    )
    return slide


# ---- 表紙 ----


def opening(deck):
    cover = deck.slide(
        id="cover",
        kind="cover",
        title="二分探索の考え方",
        notes="二分探索の考え方を説明します。整列した配列で、比較のたびに探す範囲を半分にする方法です。"
        "最後に、比較回数を線形探索と比べます。",
        seconds=20,
    )
    cover.line(id="title_rule", start=(120, 256), end=(600, 256), line_color="primary", line_width_pt=2)
    # 副題は 1 行で見せるので、折り返しを無効にする（枠に収まらなければ配置の検査が知らせる）。
    cover.text(
        id="subtitle",
        area=Rect(60, 270, 600, 40),
        text="整列した配列から値を探す",
        role="heading",
        align="center",
        word_wrap=False,
    )
    cover.text(id="author", area=Rect(60, 420, 600, 32), text="slidekit example", align="center", color="muted")


# ---- 前半：考え方 ----


def branching_page(deck):
    slide = deck.slide(
        id="branching",
        title="比較による分岐",
        notes="昇順に並んだ配列から、x 以上の最初の位置を探します。探す範囲を l 以上 r 未満とし、"
        "中央の位置 m の値と x を比べます。a_m が x より小さければ、答えは m より右にあるので l を m＋1 にします。"
        "そうでなければ、答えは m 以下にあるので r を m にします。どちらの場合も範囲は半分以下になり、"
        "l と r が等しくなったところで終わります。",
        seconds=60,
    )
    # 高さを None にした領域は、文章を幅で折り返した高さになる。
    premise = slide.text(
        id="premise",
        area=Rect(36, 84, 648, None),
        sources="knuth1998",
        text=parse(r"昇順の配列 $a_0 < a_1 < \cdots < a_{n-1}$ から、$x$ 以上の最初の位置を探す"),
    )
    middle = slide.text(
        id="middle",
        area=Rect(36, premise.area.bottom + 8, 648, None),
        text=parse(r"探す範囲を $[l, r)$ とし、中央 $m = \lfloor (l + r) / 2 \rfloor$ の値 $a_m$ を $x$ と比べる"),
    )
    # 分岐図は 1 つのまとまりにし、PowerPoint でもグループとして動くようにする。
    # 数式を含むまとまりはグループにならないので、条件と更新の数式はまとまりの外に置く。
    diagram = slide.group(id="decision", area=Rect(36, middle.area.bottom + 20, 648, 220))
    compare = diagram.box(
        id="compare",
        area=Rect(234, 0, 180, 90),
        shape="diamond",
        text="比べる",
        align="center",
        valign="middle",
        fill="background",
        line_color="primary",
        line_width_pt=2,
    )
    to_left = diagram.box(
        id="to_left",
        area=Rect(0, 160, 268, 60),
        shape="rounded_rectangle",
        text="左の半分を探す",
        align="center",
        valign="middle",
    )
    to_right = diagram.box(
        id="to_right",
        area=Rect(380, 160, 268, 60),
        shape="rounded_rectangle",
        text="右の半分を探す",
        align="center",
        valign="middle",
    )
    # 矢印の先は箱の上辺から 6 pt 離す。
    diagram.arrow(
        id="if_not_less",
        start=compare.anchor("left"),
        end=to_left.anchor("top", gap=6),
        line_color="primary",
        line_width_pt=2,
    )
    diagram.arrow(
        id="if_less",
        start=compare.anchor("right"),
        end=to_right.anchor("top", gap=6),
        line_color="primary",
        line_width_pt=2,
    )
    top = diagram.area.y
    slide.text(id="not_less", area=Rect(56, top + 66, 120, 30), text=Math(r"a_m \ge x"), align="right")
    slide.text(id="less", area=Rect(544, top + 66, 120, 30), text=Math(r"a_m < x"))
    for name, box, expression in (
        ("left_update", to_left, r"r \leftarrow m"),
        ("right_update", to_right, r"l \leftarrow m + 1"),
    ):
        area = box.area
        slide.text(id=name, area=Rect(area.x, area.bottom + 8, area.w, None), text=Math(expression), align="center")


def trace_page(deck):
    steps, final = trace(VALUES, TARGET)
    slide = deck.slide(
        id="trace",
        title="探す範囲の推移",
        notes="具体例として、8 個の値から 13 以上の最初の位置を探します。"
        "表の下の段は、各位置の値が x 以上かどうかで、答えは最初の丸の位置です。"
        "最初の範囲は 0 以上 8 未満です。中央の a_4 は 11 で 13 より小さいので、l を 5 にします。"
        "次の中央 a_6 は 17 なので r を 6 にし、その次の a_5 は 13 なので r を 5 にします。"
        "l と r が 5 で等しくなり、答えは位置 5 です。",
        seconds=60,
    )
    example = slide.text(
        id="example",
        area=Rect(36, 84, 648, None),
        text=parse(rf"例：$x = {TARGET}$ を探す　○ は $a_i \ge x$ を満たす位置"),
    )
    # 列キーは位置から作る（key_id("i", 3) は "i_3"）。
    positions = [key_id("i", index) for index in range(len(VALUES))]
    table = slide.table(
        id="values",
        area=Rect(36, example.area.bottom + 10, 448, None),
        headers=[Math("i"), *(Math(str(index)) for index in range(len(VALUES)))],
        rows=[
            [Math("a_i"), *(Math(str(value)) for value in VALUES)],
            [Math(r"a_i \ge x"), *("○" if value >= TARGET else "×" for value in VALUES)],
        ],
        row_keys=["value", "test"],
        column_keys=["label", *positions],
        column_widths=[96] + [44] * len(VALUES),
        row_heights=30,
        align="center",
        valign="middle",
        padding_pt=2,
        header_role="body",
        header_style={"fill": "primary", "color": "FFFFFF", "bold": True},
        band_fills=["background", "surface"],
        line_color="line",
    )
    # 状態図：各段は探す範囲 [l, r) の位置を塗り、比べる位置 m を強調する。表の列とそろえる。
    cell = 44
    left_edge = table.area.x + 96
    chain = slide.group(id="states", area=Rect(left_edge - 4, table.area.bottom + 14, cell * len(VALUES) + 8, 234))
    states = [*steps, (final, final, None)]
    previous = None
    for number, (left, right, middle) in enumerate(states):
        y = number * 66
        frame = chain.box(
            id=key_id("range", left, right), area=Rect(0, y, cell * len(VALUES) + 8, 36), fill="none", line_color="line"
        )
        for index in range(len(VALUES)):
            marked = index == middle if middle is not None else index == final
            fill = "accent" if marked else "surface" if left <= index < right else "background"
            chain.box(
                id=key_id("range", left, right, index),
                area=Rect(4 + cell * index, y + 4, cell, 28),
                fill=fill,
                line_color="line",
                padding_pt=0,
            )
        if previous is not None:
            chain.arrow(
                id=key_id("step", left, right),
                start=previous.anchor("bottom"),
                end=frame.anchor("top"),
                line_color="primary",
                line_width_pt=2,
            )
        previous = frame
    # 範囲の値と比較は数式なので、まとまりの外に置く。
    # 数字の間のカンマは数の一部になり後ろが空かないので、範囲の値の間には \, を挟む。
    right_edge = chain.area.right + 16
    for number, (left, right, middle) in enumerate(states):
        y = chain.area.y + number * 66
        label = Math(rf"[{left},\,{right})") if middle is not None else [Math(f"l = r = {final}"), Text(" で終了")]
        slide.text(
            id=key_id("label", left, right),
            area=Rect(right_edge, y + 3, 684 - right_edge, 30),
            text=label,
            role="heading" if middle is None else None,
        )
        if middle is not None:
            smaller = VALUES[middle] < TARGET
            relation = "<" if smaller else r"\ge"
            update = rf"l \leftarrow {middle + 1}" if smaller else rf"r \leftarrow {middle}"
            slide.text(
                id=key_id("compare", middle),
                area=Rect(chain.area.x + 196, y + 38, 360, 28),
                size_pt=18,
                text=[Math(rf"a_{middle} = {VALUES[middle]} {relation} x"), Text("　なので　"), Math(update)],
            )


def invariant_page(deck):
    slide = deck.slide(
        id="invariant",
        title="繰り返しの間の条件",
        notes="繰り返しの間は、l より左の値はすべて x 未満、r 以降の値はすべて x 以上です。"
        "l 以上 r 未満の区間だけが、まだ比べていない部分です。"
        "終了時は l と r が等しく中央の区間が空になるので、l が x 以上の最初の位置になります。",
        seconds=45,
    )
    content = slide.content
    statement_text = parse("繰り返しの間、探す範囲の外の値について次が常に成り立つ")
    conclusion_text = parse(r"終了時は $l = r$ で中央の区間が空になり、$l$ が $x$ 以上の最初の位置になる")
    conditions = (Math(r"a_i < x"), Text("まだ比べていない"), Math(r"a_i \ge x"))
    # 縦の並び（文、境界の印、区間、条件、結論）は、計測した高さで内容の領域に等間隔で置く。
    heights = [
        slide.measure(statement_text, width=content.w).height,
        64,
        56,
        max(slide.measure(text).height for text in conditions),
        slide.measure(conclusion_text, width=content.w).height,
    ]
    rows = distribute(
        [Rect(content.x, 0, content.w, height) for height in heights], content.inset(top=4, bottom=8), axis="vertical"
    )
    statement_row, marker_row, bar_row, condition_row, conclusion_row = rows
    slide.text(id="statement", area=statement_row, text=statement_text)
    # 配列を 3 つの区間に分ける。幅の比は区間の長さの目安で、区間の間を空けて境界を見せる。
    names = ("below", "range", "above")
    boxes = {}
    for name, interval, column in zip(names, (r"[0, l)", r"[l, r)", r"[r, n)"), content.columns([3, 2, 3], gap=16)):
        outlined = name == "range"
        boxes[name] = slide.box(
            id=name,
            area=Rect(column.x, bar_row.y, column.w, bar_row.h),
            text=Math(interval),
            align="center",
            valign="middle",
            fill="background" if outlined else "surface",
            line_color="primary" if outlined else "line",
            line_width_pt=2 if outlined else 1,
        )
    # 各区間の条件を区間の箱の下の中央にそろえる。
    for name, condition in zip(names, conditions):
        measured = slide.measure(condition)
        (area,) = align(
            [Rect(0, condition_row.y, measured.width, measured.height)], boxes[name].area, horizontal="center"
        )
        slide.text(id=f"{name}_condition", area=area, text=condition, color="muted" if name == "range" else None)
    # 境界の印 l と r から、区間の箱の左上の角（辺の上の位置 at=0）へ矢印を引く。
    for name, symbol in (("range", "l"), ("above", "r")):
        corner = boxes[name].area.x
        marker = slide.text(
            id=f"{symbol}_marker", area=Rect(corner - 20, marker_row.y, 40, 28), text=Math(symbol), align="center"
        )
        slide.arrow(
            id=f"{symbol}_boundary",
            start=marker.anchor("bottom", gap=2),
            end=boxes[name].anchor("top", at=0, gap=4),
            line_color="primary",
            line_width_pt=2,
        )
    slide.text(id="conclusion", area=conclusion_row, text=conclusion_text)


def idea_chapter(deck):
    section_page(
        deck,
        id="idea",
        title="考え方",
        description="比較のたびに探す範囲を半分にする",
        notes="前半では、比較による分岐と、具体例での範囲の変化を説明します。",
        seconds=10,
    )
    branching_page(deck)
    trace_page(deck)
    invariant_page(deck)


# ---- 後半：比較回数 ----


def recurrence_page(deck):
    slide = deck.slide(
        id="recurrence",
        title="比較回数の漸化式",
        notes="範囲の要素数を n とします。1 回の比較で、要素数は n を 2 で割って切り捨てた値以下になります。"
        "要素数 n での最悪の比較回数を T(n) とすると、次の漸化式が成り立ちます。T(0) は 0 です。"
        "これを解くと、T(n) は log2 n の切り捨てに 1 を足した値になります。要素数が 100 万なら 20 回です。",
        seconds=50,
    )
    # 文と式を上から順に積む。高さは内容から決まる。
    column = slide.stack(slide.content.inset(top=4), gap=20)
    column.text(
        id="shrink",
        text=parse(r"範囲の要素数を $n$ とすると、$1$ 回の比較で要素数は " r"$\lfloor n / 2 \rfloor$ 以下になる"),
    )
    # 役割を指定しない equation() は、テーマの役割 equation（この例では 28 pt の主の色）を使う。
    column.equation(
        id="relation",
        indent=48,
        sources="knuth1998",
        expression=r"T(n) = T\left(\left\lfloor \frac{n}{2} \right\rfloor\right) + 1",
    )
    column.text(id="base", text=parse(r"$T(0) = 0$ から順に求めると"))
    column.equation(
        id="solution",
        align="center",
        sources="cormen2022",
        expression=r"T(n) = \left\lfloor \log_2 n \right\rfloor + 1 \quad (n \ge 1)",
    )
    column.equation(id="million", align="center", expression=r"T(10^6) = 20\ \text{回}")
    # 注は左だけ広い余白（辺ごとの余白）で、本文から少し下げて見せる。
    column.box(
        id="meaning",
        role="note",
        fill="surface",
        line_color="none",
        padding_pt={"left": 16, "right": 8, "top": 4, "bottom": 4},
        text=parse(r"$T(n)$：要素数 $n$ での最悪の比較回数"),
    )


def growth_page(deck):
    slide = deck.slide(
        id="growth",
        title="比較回数の増え方",
        notes="要素数ごとの最悪の比較回数を図にしました。灰色の線が線形探索、青の階段が二分探索です。"
        "要素数が 64 のとき、線形探索は 64 回ですが、二分探索は 7 回です。"
        "図の値は式から計算したもので、測定結果ではありません。",
        seconds=45,
        # 軸名の役割 chart_label は、使う場所を限った小さい文字の役割なので、このページで使うと宣言する。
        allow_roles=["chart_label"],
    )
    slide.text(id="y_label", area=Rect(36, 84, 432, None), text="最悪の比較回数", role="chart_label")
    chart = slide.image(
        id="chart",
        path="figures/comparisons.svg",
        area=Rect(36, 108, *FIGURE_SIZE),
        vector=True,
        sources="figure_script",
        description="線形探索と二分探索の、要素数ごとの最悪の比較回数のグラフ",
    )
    area = chart.area
    slide.text(
        id="x_label",
        area=Rect(area.x, area.bottom + 4, area.w, None),
        text=[Text("要素数 "), Math("n")],
        role="chart_label",
        align="center",
    )
    # 図の線の終わりの値を、図に重ねたラベルで示す。重なりは意図したものなので、理由を付けて許容する。
    # 位置は線の終わりの点からのずれ（dx, dy）と幅で決め、線に重ならない所に置く。
    for name, value, color, (dx, dy, width) in (
        ("linear_value", LARGEST, "muted", (-130, -2, 90)),
        ("binary_value", binary_comparisons(LARGEST), "primary", (-70, -32, 64)),
    ):
        x, y = plot_point(LARGEST, value)
        label = slide.text(
            id=name,
            area=Rect(area.x + x + dx, area.y + y + dy, width, 26),
            text=[Math(str(value)), Text(" 回")],
            role="note",
            color=color,
            align="right",
        )
        slide.allow_warning(
            "element_overlap",
            element=label,
            related=chart,
            reason="線の終わりの値を示すラベルを、図の空いた所に重ねている",
        )
    # 凡例は線と文字のまとまり（数式を含まないので PowerPoint のグループになる）。
    legend = slide.group(id="legend", area=Rect(492, 160, 192, 96))
    legend.line(id="linear_sample", start=(0, 18), end=(40, 18), line_color="muted", line_width_pt=2)
    legend.text(id="linear_label", area=Rect(50, 0, 142, 36), text="線形探索", valign="middle")
    legend.line(id="binary_sample", start=(0, 78), end=(40, 78), line_color="primary", line_width_pt=2.5)
    legend.text(id="binary_label", area=Rect(50, 60, 142, 36), text="二分探索", valign="middle")
    slide.text(
        id="computed", area=Rect(36, 448, 648, None), text="※ 値は式から計算した値で、測定結果ではない", role="note"
    )


def comparison_page(deck):
    slide = deck.slide(
        id="comparison",
        title="線形探索との比較",
        notes="線形探索は並び順を問いませんが、最悪の場合は n 回比べます。"
        "二分探索は配列が昇順に整列済みである必要がありますが、比較回数は log2 n の切り捨てに 1 を足した回数です。"
        "要素数が 100 万のとき、二分探索は 20 回で済みます。"
        "Python では、標準ライブラリの bisect モジュールで二分探索を使えます。",
        seconds=45,
    )
    table = slide.table(
        id="methods",
        area=Rect(36, 92, 648, None),
        sources="cormen2022",
        headers=["", "線形探索", "二分探索"],
        rows=[
            ["前提", "並び順を問わない", "昇順に整列済み"],
            ["最悪の比較回数", Math("n"), Math(r"\lfloor \log_2 n \rfloor + 1")],
            [[Math("n = 10^6"), Text(" のとき")], Math("10^6"), Math("20")],
        ],
        row_keys=["premise", "worst", "million"],
        column_keys=["item", "linear", "binary"],
        column_widths=[200, 224, 224],
        row_heights=[48, 60, 60, 60],
        align="center",
        valign="middle",
        padding_pt=8,
        line_color="line",
        header_style={"fill": "primary", "color": "FFFFFF", "bold": True},
        column_styles={"item": {"align": "left", "bold": True}},
        row_styles={"million": {"fill": "surface"}},
        # 数式だけのセルは役割 table_math（主の色）にする。
        math_role="table_math",
    )
    # まとめの箱は文章を計測した寸法にして、内容の領域の中央に置く。
    message = [Emphasis("整列の手間"), Text("と引き換えに、比較回数が対数で済む")]
    measured = slide.measure(message, kind="box", role="heading", padding_pt=12)
    (area,) = align(
        [Rect(0, table.area.bottom + 28, measured.width, measured.height)], slide.content, horizontal="center"
    )
    slide.box(id="takeaway", area=area, text=message, role="heading", padding_pt=12, fill="surface", line_color="none")
    slide.text(
        id="library",
        area=Rect(36, area.bottom + 24, 648, None),
        role="note",
        align="center",
        text=[
            Text("Python では標準ライブラリの "),
            Text("bisect モジュール", link="https://docs.python.org/3/library/bisect.html"),
            Text(" で二分探索を使える"),
        ],
    )


def cost_chapter(deck):
    section_page(
        deck,
        id="cost",
        title="比較回数",
        description="最悪の場合を線形探索と比べる",
        notes="後半では、最悪の場合の比較回数を式で求め、線形探索と比べます。",
        seconds=10,
    )
    recurrence_page(deck)
    growth_page(deck)
    comparison_page(deck)


# ---- 文献 ----


def closing(deck):
    slide = deck.slide(
        id="references",
        kind="references",
        title="参考文献",
        notes="二分探索と、その比較回数の考え方は、次の教科書で扱われています。",
        seconds=15,
    )
    # 項目は文献ファイルから作り、出典に文献キーを付ける。
    slide.references(id="entries", keys=list(REFERENCES), gap_pt=24)


# 章ごとのページの識別子と、章を作る関数（ページの順）。
CHAPTERS = (
    (("cover",), opening),
    (("idea", "branching", "trace", "invariant"), idea_chapter),
    (("cost", "recurrence", "growth", "comparison"), cost_chapter),
    (("references",), closing),
)


def build(context):
    deck = Deck(context)
    for slide_ids, make in CHAPTERS:
        # --slides で選んだページを含む章だけを作る（選択がなければすべて作る）。
        if context.selected(*slide_ids):
            make(deck)
    return deck
