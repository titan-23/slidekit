"""線形探索と二分探索の、最悪の比較回数の図 figures/comparisons.svg を作る。

値は式から計算し、測定はしない。線形探索は n 回、二分探索は ⌊log2 n⌋ + 1 回である。
図の文字は目盛りの数字だけにする。軸名と凡例はスライドの文字と線で置く。目盛りの数字は DejaVu Sans の字形の輪郭で描き、
SVG に文字を残さない（表示する環境の書体に左右されず、PDF の検査でスライドの文字と混ざらない）。
色は theme.toml から読み、スライドの凡例とそろえる。作成日時などの変わる情報は入れないので、同じ入力からは同じファイルになる。

slides.py は図に重ねるラベルの位置を `plot_point()` で求める（図と同じ座標の計算を使う）。

    python make_figure.py            # figures/comparisons.svg を作り直す
    python make_figure.py OUTPUT     # 別の場所へ書き出す
"""

from __future__ import annotations

from pathlib import Path
import sys
import tomllib

from fontTools.pens.svgPathPen import SVGPathPen
from fontTools.pens.transformPen import TransformPen
from fontTools.ttLib import TTFont

HERE = Path(__file__).resolve().parent
# 図の寸法（pt）。スライドでは同じ寸法で置く（拡大しない）。
FIGURE_SIZE = (432.0, 288.0)
# 線を描く範囲（図の左上からの pt。左、上、右、下）。左と下に目盛りの数字を置く。
PLOT = (36.0, 10.0, 420.0, 262.0)
LARGEST = 64
TICKS = range(0, LARGEST + 1, 16)
# 目盛りの数字の書体（Ubuntu などの fonts-dejavu-core）と大きさ（pt）。
TICK_FONT = Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
TICK_SIZE = 11


def binary_comparisons(n: int) -> int:
    """要素数 n の二分探索の最悪の比較回数 ⌊log2 n⌋ + 1 を返す。"""
    return n.bit_length()


def plot_point(n: float, value: float) -> tuple[float, float]:
    """要素数 n と比較回数 value の点の、図の左上からの位置（pt）を返す。"""
    left, top, right, bottom = PLOT
    return (left + (right - left) * n / LARGEST, bottom - (bottom - top) * value / LARGEST)


def _number(value: float) -> str:
    """座標を小数第 3 位までの短い文字列にする。"""
    text = f"{value:.3f}".rstrip("0").rstrip(".")
    return "0" if text == "-0" else text


def _points(points) -> str:
    return " ".join(f"{_number(x)},{_number(y)}" for x, y in points)


def _steps() -> list[tuple[float, float]]:
    """二分探索の階段の頂点。比較回数は要素数が 2 の累乗になるたびに 1 増える。"""
    points = [plot_point(1, 1)]
    for n in range(2, LARGEST + 1):
        if binary_comparisons(n) != binary_comparisons(n - 1):
            points.append(plot_point(n, binary_comparisons(n - 1)))
            points.append(plot_point(n, binary_comparisons(n)))
    points.append(plot_point(LARGEST, binary_comparisons(LARGEST)))
    return points


class _Label:
    """目盛りの数字を、書体の字形の輪郭（SVG の path）にする。"""

    def __init__(self, path: Path, size: float):
        if not path.is_file():
            raise SystemExit(
                f"目盛りの数字の書体がない: {path}\n"
                "DejaVu Sans を入れる（Ubuntu・Debian では sudo apt install fonts-dejavu-core）か、TICK_FONT を書き換える。"
            )
        self.font = TTFont(path)
        self.glyphs = self.font.getGlyphSet()
        self.cmap = self.font.getBestCmap()
        self.scale = size / self.font["head"].unitsPerEm

    def width(self, text: str) -> float:
        """text の送り幅（pt）。"""
        return sum(self.font["hmtx"][self.cmap[ord(char)]][0] for char in text) * self.scale

    def path(self, text: str, x: float, y: float) -> str:
        """基準線の左端を (x, y) に置いた text の輪郭の、path の d 属性。"""
        pen = SVGPathPen(self.glyphs, ntos=_number)
        for char in text:
            name = self.cmap[ord(char)]
            # 書体の座標は上向きなので、y を反転して図の座標（下向き）に移す。
            self.glyphs[name].draw(TransformPen(pen, (self.scale, 0, 0, -self.scale, x, y)))
            x += self.font["hmtx"][name][0] * self.scale
        return pen.getCommands()


def build_figure(output: Path) -> Path:
    """図を output に書き、そのパスを返す。"""
    colors = {
        name: f"#{value}"
        for name, value in tomllib.loads((HERE / "theme.toml").read_text(encoding="utf-8"))["colors"].items()
    }

    def lines(segments) -> str:
        return "".join(f"M{_points([a])}L{_points([b])}" for a, b in segments)

    # 目盛りの線（背景）。
    grid = [(plot_point(tick, 0), plot_point(tick, LARGEST)) for tick in TICKS[1:]]
    grid += [(plot_point(0, tick), plot_point(LARGEST, tick)) for tick in TICKS[1:]]
    # 軸と目盛りの印。
    axes = [(plot_point(0, LARGEST), plot_point(0, 0)), (plot_point(0, 0), plot_point(LARGEST, 0))]
    for tick in TICKS:
        x, y = plot_point(tick, 0)
        axes.append(((x, y), (x, y + 4)))
        x, y = plot_point(0, tick)
        axes.append(((x - 4, y), (x, y)))
    # 目盛りの数字。横軸は目盛りの下の中央、縦軸は目盛りの左に右寄せで置く。
    label = _Label(TICK_FONT, TICK_SIZE)
    numbers = []
    for tick in TICKS:
        text = str(tick)
        width = label.width(text)
        x, y = plot_point(tick, 0)
        numbers.append(label.path(text, x - width / 2, y + 6 + TICK_SIZE))
        x, y = plot_point(0, tick)
        numbers.append(label.path(text, x - 7 - width, y + TICK_SIZE * 0.35))
    width, height = (_number(value) for value in FIGURE_SIZE)
    svg = "\n".join(
        [
            f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}pt" height="{height}pt" '
            f'viewBox="0 0 {width} {height}">',
            "<title>最悪の比較回数</title>",
            f'<path d="{lines(grid)}" fill="none" stroke="{colors["surface"]}" stroke-width="1"/>',
            f'<path d="{lines(axes)}" fill="none" stroke="{colors["muted"]}" stroke-width="1"/>',
            # 線形探索の線と、二分探索の階段。
            f'<path d="{lines([(plot_point(1, 1), plot_point(LARGEST, LARGEST))])}" fill="none" '
            f'stroke="{colors["muted"]}" stroke-width="2" stroke-linecap="round"/>',
            f'<polyline points="{_points(_steps())}" fill="none" stroke="{colors["primary"]}" stroke-width="2.5" '
            'stroke-linecap="round" stroke-linejoin="round"/>',
            f'<path d="{"".join(numbers)}" fill="{colors["text"]}"/>',
            "</svg>",
            "",
        ]
    )
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(svg, encoding="utf-8", newline="\n")
    return output


if __name__ == "__main__":
    print(build_figure(Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "figures" / "comparisons.svg"))
