"""共通のスライド表現を編集可能な PPTX へ出力する。

文字枠は次のように書く（計測の `measure` と PPTX の検査の `pptx_validation` は、この書き方を前提にする）。

- 要素の文章は 1 つの段落に書き、明示改行は a:br にする。折り返しは書式の word_wrap に従い（真なら wrap="square"、
  偽なら wrap="none"）、文字の自動の縮小・調整はしない（a:noAutofit）。
- 雛形の既定に左右されないよう、段落に eaLnBrk="1"・latinLnBrk="0"・hangingPunct="0" を、通常文字に kern="1200" を明示する。
- 通常文字は断片の実効書式（`styles.fragment_style`）で書く。fonts（`measure.FontSource`）があれば `measure.font_runs` で
  書体の区切りに分け、forced の区切り（欧文の書体にない文字、幅の曖昧な文字、PowerPoint が既定で和文の書体で描く記号）は
  a:latin と a:ea の両方を計測した書体にする。○ × ± “ などは和文の書体、→ — などは欧文の書体になる（PowerPoint が既定で
  描く書体と同じ。PowerPoint 16.0、build 20430 で確かめた）。書体の判定は `measure.font_runs` だけで行い、書き出しは
  文字の一覧を持たない。
- 段落の末尾の書式（a:endParaRPr）は、最後の断片の実効書式にする（断片がない段落と独立した数式の段落は要素の書式）。
  書かないと PowerPoint は雛形の既定の書式を使い、段落の最後の断片のリンクを PDF のリンクにしない
  （PowerPoint 16.0、build 20430 で確かめた）。
- 数式は、equation() で空でない断片が数式 1 つだけなら独立した数式（m:oMathPara と m:jc）に、それ以外は文中の数式にする
  （`office_math.math_mode`）。文中の数式で、明示改行で分けた行のうち数式があり通常文字のない行には、その行の後に U+200B だけの
  a:r を加える（加えないと PowerPoint はその行を独立した数式として扱い、段落の揃えを無視して枠の中央に置く。
  PowerPoint 16.0、build 20430 で確かめた）。
- 中央揃えの独立した数式は、a:endParaRPr に要素の書体を書き、PowerPoint が式の後ろに加える幅
  （`office_math.display_trailing_width`）と、末尾が添字の式の箱の末尾の空き（`display_trailing_space`）を a:pPr/@marL に
  して中心をそろえる。幅は fonts の空白の送り幅と数式の箱から求めるので、fonts がなければ補正しない。
- 明示改行の a:br にはその改行がある断片の実効書式を a:rPr に書く。書かないと PowerPoint は空の行を雛形の既定の
  文字サイズ（既定の雛形で 18 pt）で置く（PowerPoint 16.0、build 20430 で確かめた）。

図形は次のように書く。

- 箱と文字枠の枠線は、線の色（line_color）があるときだけ書き、dashed なら破線（a:prstDash val="dash"）にする。
  線・矢印も dashed なら破線にする。
- 場面で pptx_group が真のまとまりは PowerPoint のグループ（p:grpSp）にし、子孫の図形を「最も近い、書き出すまとまり」の
  グループへ入れる。グループの座標変換は恒等（off = chOff、ext = chExt）で、子の座標はページ上の座標のままである。
  数式を含む要素はグループに入れない（互換画像はスライド直下の図形だけに付けるため。グループにするかは作成 API が決める）。
- anchor() で結んだ線の端は、`anchored_site` が接続点の番号を返すとき PowerPoint の接続（a:stCxn・a:endCxn）にする。
  端点の座標は場面の値のまま書く。数式を含む接続先も、互換画像を付ける資料で接続する（PowerPoint は mc:AlternateContent の
  中の図形への接続を認識する。PowerPoint 16.0、build 20430 で確かめた）。ただし LibreOffice が必須で互換画像を付ける資料では、
  数式を含む接続先を接続しない。
- SVG でも埋め込む画像（vector_path がある画像）は、PNG の a:blip に SVG の拡張（asvg:svgBlip）を加える。
- ppt/presentation.xml の p:kinsoku（日本語の行頭・行末の禁止文字）を、計測と同じ一覧（`measure.NO_LINE_START`・
  `NO_LINE_END`）で書く。雛形にあれば置き換える。

書式の値が場面にないときは `styles.LIBRARY_DEFAULTS` を使う。ただし色（既定はテーマの色の名前）は補わず、書かない。
"""

from __future__ import annotations

from collections.abc import Mapping
from datetime import datetime, timezone
import hashlib
from itertools import groupby
from pathlib import Path

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.oxml.xmlchemy import OxmlElement
from pptx.text.text import Font
from pptx.util import Pt

from .fallback import fallback_required
from .files import atomic_output, file_hash
from .layout import Rect
from .math.office_math import append_math, display_trailing_width, math_mode, math_options
from .math.syntax import Function, Nary, Scripts, parse
from .measure import MODEL, NO_LINE_END, NO_LINE_START, FontRun, font_runs, padding_insets

# NS は試験が writer から読み込むので、ここでも持つ。
from .ooxml import EMU_PER_PT, NS, tag
from .styles import COLOR_KEYS, LIBRARY_DEFAULTS, fragment_style

# 文中の数式だけの行に加える文字（幅 0 の空白）。
ZERO_WIDTH_SPACE = "​"
# 欧文の kern を使う最小の文字サイズ（1/100 pt）。計測の kerning_min_size_pt と同じ値にする。
KERN = str(round(MODEL.kerning_min_size_pt * 100))
# 段落の行の分け方。日本語の禁則を使い、欧文の語の途中で分けず、句読点をぶら下げない。
LINE_BREAKING = {"eaLnBrk": "1", "latinLnBrk": "0", "hangingPunct": "0"}
# リンクの文字をテーマのリンクの色でなく文字の色で表示させる拡張（Office 2019 以降）。
HYPERLINK_COLOR_URI = "{A12FA001-AC4F-418D-AE19-62706E023703}"
HYPERLINK_COLOR_NS = "http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor"
# 画像を SVG でも埋め込む拡張（Office 2016 以降）。対応しないソフトは a:blip の PNG を表示する。
SVG_BLIP_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
SVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
SVG_CONTENT_TYPE = "image/svg+xml"
# p:kinsoku の言語。
KINSOKU_LANGUAGE = "ja-JP"
# CT_Presentation で p:kinsoku の後に置く子要素（ECMA-376 の順）。
_KINSOKU_SUCCESSORS = ("p:defaultTextStyle", "p:modifyVerifier", "p:extLst")

_ALIGNMENTS = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}
_ANCHORS = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}
_SHAPES = {
    "rectangle": MSO_SHAPE.RECTANGLE,
    "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
    "ellipse": MSO_SHAPE.OVAL,
    "diamond": MSO_SHAPE.DIAMOND,
}
# 場面にない書式の項目に使う値。色はテーマの色の名前なので使わない。word_wrap がない場面（0.1 で作った場面）は、
# 計測と検査と同じく折り返さない。
_FALLBACK_STYLE = {key: value for key, value in LIBRARY_DEFAULTS.items() if key not in COLOR_KEYS | {"word_wrap"}}
# a:rPr の書体の子要素の後に置く子要素（ECMA-376 の CT_TextCharacterProperties の順）。
_FONT_SUCCESSORS = {
    "latin": ("a:ea", "a:cs", "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst"),
    "ea": ("a:cs", "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst"),
    "cs": ("a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst"),
}
# 既定図形の接続点の番号（DrawingML の presetShapeDefinitions の cxnLst の順）。楕円は 8 点を持つ。
_CONNECTION_SITES = {
    "rectangle": {"top": 0, "left": 1, "bottom": 2, "right": 3},
    "rounded_rectangle": {"top": 0, "left": 1, "bottom": 2, "right": 3},
    "diamond": {"top": 0, "left": 1, "bottom": 2, "right": 3},
    "ellipse": {"top": 0, "left": 2, "bottom": 4, "right": 6},
}


def _color(value):
    return RGBColor.from_string(str(value).removeprefix("#").upper())


def element_style(element: Mapping) -> dict:
    """書き出しに使う要素の書式。場面にない項目（値が None の項目を含む）は styles.LIBRARY_DEFAULTS の値にする。

    色の項目は補わない（塗らない・線なし・色を書かない）。word_wrap がない場面は折り返さない。
    """
    style = dict(_FALLBACK_STYLE)
    style.update({key: value for key, value in (element.get("style") or {}).items() if value is not None})
    style.setdefault("word_wrap", False)
    return style


def _font_runs(text: str, style: Mapping, fonts, bold: bool) -> list[FontRun]:
    """通常文字の 1 行を、書き出す書体の区切りにする。

    fonts がなければ区切らない（0.1 と同じ）。forced の区切りは計測の書体（○ × なら和文の書体、→ — なら欧文の書体）の
    まま残す。forced でない区切りは書き出す書体が同じ（a:latin は欧文、a:ea は和文）なので、隣どうしをまとめる。
    どちらの書体にも字形がない文字と、書体を判定できない行は forced でない区切りにする
    （字形の欠落と書式の誤りは配置の検査が報告する）。
    """
    latin = style.get("font_latin")
    if fonts is None:
        return [FontRun(text, latin, False)]
    try:
        runs = font_runs(text, style, fonts, bold=bold)
    except (OSError, ValueError) as error:
        missing = set(getattr(error, "details", {}).get("characters", ()))
        if getattr(error, "code", None) != "font_glyph_missing" or not missing:
            return [FontRun(text, latin, False)]
        runs = []
        try:
            for absent, group in groupby(text, key=lambda char: char in missing):
                part = "".join(group)
                runs.extend([FontRun(part, latin, False)] if absent else font_runs(part, style, fonts, bold=bold))
        except (OSError, ValueError):
            return [FontRun(text, latin, False)]
    merged: list[FontRun] = []
    for run in runs:
        if run.forced:
            merged.append(run)
        elif merged and not merged[-1].forced:
            merged[-1] = FontRun(merged[-1].text + run.text, latin, False)
        else:
            merged.append(FontRun(run.text, latin, False))
    return merged


def paragraph_content(element: Mapping, fonts=None) -> list[tuple]:
    """要素の段落に書く中身を、書き出す順に（種類、文字列、断片の実効書式、書体の区切り）の組で返す。

    種類は "run"（通常文字の a:r）、"break"（明示改行の a:br）、"math"（数式の a14:m）。書体の区切り（`measure.FontRun`）は
    "run" だけが持つ。文中の数式では、明示改行で分けた行のうち、数式があり通常文字のない行ごとに、その行の後に U+200B だけの
    "run" を加える（要素の書式で書く。数式だけの段落なら段落の最後にも加わる）。
    PPTX の検査は期待する区切りの列をこの関数で求める（fonts は書き出しと同じものを渡す）。
    """
    style = element_style(element)
    items: list[tuple] = []
    for fragment in element.get("content") or []:
        kind = fragment.get("kind")
        text = fragment.get("text", "")
        own = fragment_style(style, fragment)
        if kind == "math":
            items.append(("math", text, own, None))
            continue
        if kind not in {"text", "emphasis"}:
            raise ValueError(f"未対応の文章の断片: {kind}")
        for index, line in enumerate(str(text).split("\n")):
            if index:
                items.append(("break", "\n", own, None))
            if line:
                items.extend(("run", run.text, own, run) for run in _font_runs(line, style, fonts, own["bold"]))
    if not any(item[0] == "math" for item in items) or math_mode(element) != "inline":
        return items
    own = fragment_style(style, {"kind": "text"})
    separator = ("run", ZERO_WIDTH_SPACE, own, FontRun(ZERO_WIDTH_SPACE, style.get("font_latin"), False))
    result, line = [], []
    for item in [*items, None]:
        if item is None or item[0] == "break":
            kinds = {part[0] for part in line}
            result.extend(line + ([separator] if "math" in kinds and "run" not in kinds else []))
            if item is not None:
                result.append(item)
            line = []
        else:
            line.append(item)
    return result


def math_arguments(element: Mapping, own: Mapping, functions) -> dict:
    """数式の断片の append_math()・math_element() の引数。断片の実効書式を、場面にない項目を補った要素の書式に重ねる。"""
    style = {**element_style(element), **{key: value for key, value in own.items() if value is not None}}
    return math_options(element, functions, style=style)


def _ends_with_script(nodes) -> bool:
    """構文木の最後の原子が添字か（大型演算子の添字は除く。引数のある大型演算子と関数名は引数の最後の原子で決める）。"""
    while nodes:
        node = nodes[-1]
        if isinstance(node, Scripts):
            return True
        if isinstance(node, Nary):
            # 大型演算子の添字の後には SpaceAfterScript を入れない（`math.metrics` の nary_space_after_script）。
            nodes = node.body
        elif isinstance(node, Function):
            nodes = node.body or node.name
        else:
            return False
    return False


def display_trailing_space(element: Mapping, fonts=None) -> float:
    """中央揃えの独立した数式で、PowerPoint が式の箱の末尾に含める、最後の添字の後の空き（pt）。末尾が添字でない式は 0。

    末尾が添字の式は、描いた字形が枠の中心より左に寄る（PowerPoint 16.0、build 20430 で生成した資料で 0.4〜0.8 pt）。
    計測した値は、PowerPoint が添字の後の空き（数式の書体の MATH 表の SpaceAfterScript × 文字サイズ）と末尾の添字の斜体補正
    （`math.metrics.Box.italic`）を箱に含めるとすると 0.2 pt 以内で合うので、その和を返す（補正の効果は PowerPoint で
    確かめる）。末尾が括弧などの式は今の補正で 0.2 pt 以内にそろうので加えない。計測器が数式の箱と MATH 表の読み取りに
    対応していない（`math_layout` か `math_constant` のない FontSource）ときと、式を組む処理が失敗したときは 0 にする。
    計測器の中で起きた AttributeError は 0 にせず例外にする（書き出しと PPTX の検査がそろって補正をやめるのを防ぐ）。
    """
    layout = getattr(fonts, "math_layout", None)
    constant = getattr(fonts, "math_constant", None)
    if layout is None or constant is None or math_mode(element) != "display":
        return 0.0
    style = element_style(element)
    fragment = next(item for item in element["content"] if item.get("kind") == "math" and str(item.get("text", "")))
    own = fragment_style(style, fragment)
    expression = str(fragment["text"])
    try:
        if not _ends_with_script(parse(expression, getattr(fonts, "functions", ()))):
            return 0.0
        box = layout(
            expression,
            own["size_pt"],
            own["bold"],
            family=style.get("font_math"),
            display=True,
            japanese=style.get("font_japanese"),
            latin=style.get("font_latin"),
        )
        space = constant("SpaceAfterScript", family=style.get("font_math"))
    except (KeyError, OSError, TypeError, ValueError):
        return 0.0
    return max(0.0, space * own["size_pt"] + max(0.0, box.italic))


def display_indent(element: Mapping, fonts=None) -> float:
    """中央揃えの独立した数式の段落の左の余白（pt）。描いた式の中心を枠の中心にそろえる。

    余白は、PowerPoint が式の後ろに加える幅と、末尾が添字の式の箱の末尾の空きの和である。後ろに加える幅は
    a:endParaRPr に書く欧文の書体（要素の font_latin）の空白の送り幅 × 数式の文字サイズ、末尾の空きは
    `display_trailing_space` である。中央揃えでない段落、独立した数式でない要素、fonts がないとき、空白の送り幅を
    求める処理が失敗したときは 0（補正しない）。
    """
    if fonts is None or math_mode(element) != "display":
        return 0.0
    style = element_style(element)
    latin = style.get("font_latin")
    if style.get("align") != "center" or not latin:
        return 0.0
    fragment = next(item for item in element["content"] if item.get("kind") == "math" and str(item.get("text", "")))
    size = fragment_style(style, fragment).get("size_pt")
    if size is None:
        return 0.0
    try:
        trailing = display_trailing_width(fonts.advance(" ", latin, bool(style.get("bold"))), size)
    except (OSError, ValueError):
        return 0.0
    return trailing + display_trailing_space(element, fonts)


def _typeface(properties, kind: str, family: str) -> None:
    child = properties.find(tag("a", kind))
    if child is None:
        child = properties.insert_element_before(OxmlElement("a:" + kind), *_FONT_SUCCESSORS[kind])
    child.set("typeface", family)


def _character_properties(properties, own: Mapping, style: Mapping, *, forced: str | None = None) -> None:
    """a:rPr か a:endParaRPr に、断片の実効書式（文字サイズ・太字・下線・色）と書体を書く。値のない項目は書かない。

    forced を渡すと、a:latin と a:ea の両方をその書体にする。a:cs は常に欧文の書体にする。
    """
    font = Font(properties)
    properties.set("lang", "ja-JP")
    if own.get("size_pt") is not None:
        font.size = Pt(own["size_pt"])
    font.bold = bool(own.get("bold"))
    if own.get("underline"):
        properties.set("u", "sng")
    if own.get("color"):
        font.color.rgb = _color(own["color"])
    latin = style.get("font_latin")
    for kind, family in (("latin", forced or latin), ("ea", forced or style.get("font_japanese")), ("cs", latin)):
        if family:
            _typeface(properties, kind, family)


def _link_uses_text_color(properties) -> None:
    """a:hlinkClick に、リンクを文字の色で表示させる拡張（ahyp:hlinkClr val="tx"）を加える。"""
    click = properties.find(tag("a", "hlinkClick"))
    extensions = OxmlElement("a:extLst")
    extension = OxmlElement("a:ext")
    extension.set("uri", HYPERLINK_COLOR_URI)
    color = etree.SubElement(extension, f"{{{HYPERLINK_COLOR_NS}}}hlinkClr", nsmap={"ahyp": HYPERLINK_COLOR_NS})
    color.set("val", "tx")
    extensions.append(extension)
    click.append(extensions)


def _style_run(run, own: Mapping, font_run: FontRun, style: Mapping) -> None:
    properties = run._r.get_or_add_rPr()
    _character_properties(properties, own, style, forced=font_run.family if font_run.forced else None)
    properties.set("kern", KERN)
    if own.get("link"):
        run.hyperlink.address = own["link"]
        _link_uses_text_color(properties)


def _set_text(shape, element, functions, fonts):
    style = element_style(element)
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = bool(style["word_wrap"])
    frame.auto_size = MSO_AUTO_SIZE.NONE
    insets = padding_insets(style)
    frame.margin_left, frame.margin_top = Pt(insets.left), Pt(insets.top)
    frame.margin_right, frame.margin_bottom = Pt(insets.right), Pt(insets.bottom)
    frame.vertical_anchor = _ANCHORS[style["valign"]]
    paragraph = frame.paragraphs[0]
    paragraph.alignment = _ALIGNMENTS[style["align"]]
    paragraph.space_before = paragraph.space_after = Pt(0)
    if style.get("line_spacing_pt") is not None:
        paragraph.line_spacing = Pt(style["line_spacing_pt"])
    else:
        paragraph.line_spacing = float(style["line_spacing"])
    properties = paragraph._p.get_or_add_pPr()
    for name, value in LINE_BREAKING.items():
        properties.set(name, value)
    end = fragment_style(style, {"kind": "text"})
    for kind, text, own, font_run in paragraph_content(element, fonts):
        end = own
        if kind == "math":
            append_math(paragraph, text, **math_arguments(element, own, functions))
        elif kind == "break":
            # 空の行の大きさが改行のある断片の書式で決まるように、a:br にも書式を書く（リンクは書かない）。
            _character_properties(paragraph._p.add_br().get_or_add_rPr(), own, style)
        else:
            run = paragraph.add_run()
            run.text = text
            _style_run(run, own, font_run, style)
    if math_mode(element) == "display":
        # PowerPoint は式の後ろの幅を段落の末尾の書式の書体で決める。要素の書体を明示し、中央揃えではその幅だけ左を空ける。
        end = fragment_style(style, {"kind": "text"})
        indent = display_indent(element, fonts)
        if indent > 0:
            properties.set("marL", str(round(indent * EMU_PER_PT)))
    # リンクは段落の末尾の書式に写さない（_character_properties はリンクを書かない）。
    _character_properties(paragraph._p.get_or_add_endParaRPr(), end, style)


def _set_outline(shape, style):
    """箱と文字枠の塗りと枠線。枠線は線の色があるときだけ書き、dashed なら破線にする。"""
    if style.get("fill"):
        shape.fill.solid()
        shape.fill.fore_color.rgb = _color(style["fill"])
    else:
        shape.fill.background()
    if style.get("line_color"):
        shape.line.color.rgb = _color(style["line_color"])
        shape.line.width = Pt(style["line_width_pt"])
        if style.get("dashed"):
            shape.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    else:
        shape.line.fill.background()


def _disable_shadow(shape):
    properties = shape._element.spPr
    effects = properties.find(tag("a", "effectLst"))
    if effects is None:
        properties.append(OxmlElement("a:effectLst"))
    for effect in shape._element.findall(".//" + tag("a", "effectRef")):
        effect.set("idx", "0")


def _has_math(element: Mapping) -> bool:
    return any(fragment.get("kind") == "math" for fragment in element.get("content") or [])


def _draw(shapes, element, functions, fonts, svg_parts=None):
    """要素の図形を shapes（スライドかグループの図形の集まり）へ加えて返す。まとまりは図形にしない（None）。

    svg_parts は SVG の中身のハッシュ値から部品への辞書（同じ SVG を資料の中で 1 つの部品にする）。
    """
    kind = element["kind"]
    if kind == "group":
        return None
    x, y, width, height = element["area"]
    style = element_style(element)
    rotation = element.get("rotation", 0)
    if _has_math(element) and (rotation or element.get("flip_h") or element.get("flip_v")):
        raise ValueError(f"数式を含む図形は回転・反転できない: {element['id']}")
    if kind in {"text", "equation"}:
        shape = shapes.add_textbox(Pt(x), Pt(y), Pt(width), Pt(height))
        _set_outline(shape, style)
        _set_text(shape, element, functions, fonts)
    elif kind == "box":
        shape = shapes.add_shape(_SHAPES[element.get("shape", "rectangle")], Pt(x), Pt(y), Pt(width), Pt(height))
        radius = style.get("corner_radius_pt")
        if radius is not None and element.get("shape") == "rounded_rectangle":
            shape.adjustments[0] = min(0.5, max(0.0, radius / min(width, height)))
        _set_outline(shape, style)
        _set_text(shape, element, functions, fonts)
    elif kind in {"line", "arrow"}:
        start, end = element["start"], element["end"]
        shape = shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Pt(start[0]), Pt(start[1]), Pt(end[0]), Pt(end[1]))
        color = style.get("line_color") or style.get("color")
        if color:
            shape.line.color.rgb = _color(color)
        shape.line.width = Pt(style["line_width_pt"])
        if style.get("dashed"):
            shape.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        if kind == "arrow":
            end_node = OxmlElement("a:tailEnd")
            end_node.set("type", "triangle")
            shape.line._get_or_add_ln().append(end_node)
    elif kind == "image":
        source = Path(element["path"])
        # 変換済みの素材は出力領域にあるため、読み込み元の検査は構築の段階で行う。
        with Image.open(source) as image:
            if image.format not in {"PNG", "JPEG"}:
                raise ValueError(f"画像は PNG または JPEG へ変換してから配置する: {source}")
            image_width, image_height = image.size
        fit = element.get("fit", "contain")
        if fit == "contain":
            shown = Rect(x, y, width, height).contain(image_width, image_height)
            shape = shapes.add_picture(str(source), Pt(shown.x), Pt(shown.y), Pt(shown.w), Pt(shown.h))
        elif fit == "cover":
            shape = shapes.add_picture(str(source), Pt(x), Pt(y), Pt(width), Pt(height))
            source_ratio, target_ratio = image_width / image_height, width / height
            if source_ratio > target_ratio:
                shape.crop_left = shape.crop_right = (1 - target_ratio / source_ratio) / 2
            else:
                shape.crop_top = shape.crop_bottom = (1 - source_ratio / target_ratio) / 2
        else:
            raise ValueError(f"未対応の画像の配置: {fit}")
        if element.get("vector_path"):
            _embed_svg(shape, Path(element["vector_path"]), {} if svg_parts is None else svg_parts)
    else:
        raise ValueError(f"未対応の要素: {kind}")
    shape.name = element["id"]
    if element.get("description"):
        shape._element.find(".//" + tag("p", "cNvPr")).set("descr", element["description"])
    if rotation:
        shape.rotation = rotation
    for key, attribute in (("flip_h", "flipH"), ("flip_v", "flipV")):
        if element.get(key):
            shape._element.spPr.get_or_add_xfrm().set(attribute, "1")
    _disable_shadow(shape)
    return shape


def _fills_area(image: Mapping, tolerance: float = 0.01) -> bool:
    """画像の表示範囲が配置領域と一致するか。cover は常に一致し、contain は縦横比が同じときだけ（actual_area で判定）。"""
    fit = image.get("fit", "contain")
    if fit == "cover":
        return True
    actual, area = image.get("actual_area"), image.get("area")
    if fit != "contain" or not actual or not area or len(actual) != 4 or len(area) != 4:
        return False
    return all(abs(float(a) - float(b)) <= tolerance for a, b in zip(actual, area))


def connection_site(target: Mapping, side: str) -> int | None:
    """anchor() で結んだ端を PowerPoint の接続にするときの、接続先の図形の接続点の番号（a:stCxn・a:endCxn の idx）。

    接続するのは、side が top・left・bottom・right のどれかで（center の接続点はない）、接続先が文字枠（text・equation）、
    図形（box）、表示範囲が配置領域と一致する画像のときだけである。まとまり・線・ほかの画像は接続しない（None）。
    番号は図形自身の向きを基準にするので、回転・反転した接続先にも同じ番号を使う。LibreOffice の条件は呼び出し側が判定する。
    """
    kind = target.get("kind")
    if kind in {"text", "equation"}:
        shape = "rectangle"
    elif kind == "box":
        shape = target.get("shape", "rectangle")
    elif kind == "image" and _fills_area(target):
        shape = "rectangle"
    else:
        return None
    return _CONNECTION_SITES.get(shape, {}).get(side)


def anchored_site(reference: Mapping | None, target: Mapping | None, project) -> int | None:
    """線の端（場面の start_ref・end_ref）を接続にするときの、接続先の接続点の番号。接続しないなら None。

    `connection_site` の条件に加え、資料の render.required に libreoffice があるときは、次の端を接続しない。
    LibreOffice は読み込み時に接続線を接続点へ引き直すおそれがあり、手元で確かめることができないためである。

    - gap が 0 でないか at が 0.5 でない端（描いた端が接続点にない端）。
    - 互換画像を付ける資料で、接続先が数式を含む端。互換画像の図形は数式の図形より最大 `fallback.MAX_MARGIN_PT` 広い長方形で、
      楕円の接続点もないので、LibreOffice が互換画像の側で引き直すと端がずれる。

    PowerPoint だけを使う資料では、数式を含む接続先も接続する（互換画像の mc:AlternateContent の中の図形への接続を
    PowerPoint が認識し、開いた直後に引き直さないことを確かめた。PowerPoint 16.0、build 20430）。
    書き出しと PPTX の検査（`connection_mismatch`）が同じ判定を使う。
    """
    if not isinstance(reference, Mapping) or not isinstance(target, Mapping):
        return None
    site = connection_site(target, reference.get("side"))
    if site is None:
        return None
    config = getattr(project, "config", None) or {}
    if "libreoffice" in ((config.get("render") or {}).get("required") or ()):
        gap, at = reference.get("gap"), reference.get("at")
        # 場面は既定以外のときだけ gap・at を書く（既定は gap 0、at 0.5）。
        if (gap is not None and float(gap) != 0) or (at is not None and float(at) != 0.5):
            return None
        # 接続先だけのページとして、互換画像を付けるかを設定から判定する（接続先に数式がなければ偽）。
        if fallback_required({"slides": [{"elements": [target]}]}, project):
            return None
    return site


def _connect(connector, key: str, shape_id: int, site: int) -> None:
    """接続線の端（key は stCxn か endCxn）を、図形 shape_id の接続点 site に結ぶ。端点の座標は変えない。"""
    properties = connector._element.nvCxnSpPr.cNvCxnSpPr
    node = properties.get_or_add_stCxn() if key == "stCxn" else properties.get_or_add_endCxn()
    node.set("id", str(shape_id))
    node.set("idx", str(site))


def _connect_lines(elements: list, by_id: Mapping, drawn: Mapping, project) -> int:
    """ページの線の端を接続にし、書いた接続の数を返す。接続先はページのどの位置にあってもよい（図形を描いた後に結ぶ）。"""
    count = 0
    for element in elements:
        connector = drawn.get(element["id"])
        if element.get("kind") not in {"line", "arrow"} or connector is None:
            continue
        for end, key in (("start", "stCxn"), ("end", "endCxn")):
            reference = element.get(end + "_ref")
            target = by_id.get(reference.get("element_id")) if isinstance(reference, Mapping) else None
            site = anchored_site(reference, target, project)
            shape = drawn.get(target["id"]) if target is not None else None
            if site is not None and shape is not None:
                _connect(connector, key, shape.shape_id, site)
                count += 1
    return count


def pptx_group_of(element: Mapping, by_id: Mapping) -> str | None:
    """要素（まとまりを含む）を入れる PowerPoint のグループ。最も近い、pptx_group が真の祖先のまとまりの識別子。なければ None。

    書き出しと PPTX の検査（`group_mismatch`）が同じ判定を使う。
    """
    seen = {element.get("id")}
    current = by_id.get(element.get("parent_id"))
    while isinstance(current, Mapping) and current.get("id") not in seen:
        if current.get("kind") == "group" and current.get("pptx_group") is True:
            return current["id"]
        seen.add(current.get("id"))
        current = by_id.get(current.get("parent_id"))
    return None


def _group_shapes(slide, identifier: str | None, by_id: Mapping, groups: dict):
    """グループ identifier の図形の集まり（None はスライド）。グループは最初の子を加えるときに作る。

    子孫が要素の作成順でまとまりの直後に続く場面（作成 API が pptx_group を真にする条件）では、作る位置はまとまりの位置と同じ
    重なり順になる。描く子孫がないまとまりはグループにしない。
    """
    if identifier is None:
        return slide.shapes
    group = groups.get(identifier)
    if group is None:
        element = by_id[identifier]
        group = _group_shapes(slide, pptx_group_of(element, by_id), by_id, groups).add_group_shape()
        group.name = identifier
        if element.get("description"):
            group._element.nvGrpSpPr.cNvPr.set("descr", element["description"])
        groups[identifier] = group
    return group.shapes


def _embed_svg(picture, svg_path: Path, svg_parts: dict) -> None:
    """画像の図形（PNG の a:blip）に、SVG の部品を参照する拡張 asvg:svgBlip を加える。同じ中身の SVG は 1 つの部品にする。"""
    data = svg_path.read_bytes()
    key = hashlib.sha256(data).hexdigest()
    slide_part = picture.part
    part = svg_parts.get(key)
    if part is None:
        package = slide_part.package
        part = Part(package.next_image_partname("svg"), SVG_CONTENT_TYPE, package, data)
        svg_parts[key] = part
    relationship = slide_part.relate_to(part, RT.IMAGE)
    blip = picture._element.find(".//" + tag("a", "blip"))
    extensions = blip.find(tag("a", "extLst"))
    if extensions is None:
        extensions = OxmlElement("a:extLst")
        blip.append(extensions)
    extension = OxmlElement("a:ext")
    extension.set("uri", SVG_BLIP_URI)
    svg = etree.SubElement(extension, f"{{{SVG_NS}}}svgBlip", nsmap={"asvg": SVG_NS})
    svg.set(tag("r", "embed"), relationship)
    extensions.append(extension)


def kinsoku_attributes() -> dict:
    """p:kinsoku に書く属性。行頭・行末の禁止文字は計測と同じ一覧を、文字の符号の順に並べる。"""
    return {
        "lang": KINSOKU_LANGUAGE,
        "invalStChars": "".join(sorted(NO_LINE_START)),
        "invalEndChars": "".join(sorted(NO_LINE_END)),
    }


def _set_kinsoku(presentation) -> None:
    """ppt/presentation.xml の p:kinsoku を、計測の禁則の一覧にする。雛形にある p:kinsoku は置き換える（1 つだけ置くことができる）。

    PowerPoint は p:kinsoku の一覧で日本語の行頭・行末の禁則を決める（PowerPoint 16.0、build 20430 で確かめた）。
    """
    root = presentation.part._element
    for node in root.findall(tag("p", "kinsoku")):
        root.remove(node)
    node = OxmlElement("p:kinsoku")
    for name, value in kinsoku_attributes().items():
        node.set(name, value)
    successor = next(
        (child for name in _KINSOKU_SUCCESSORS for child in [root.find(tag(*name.split(":")))] if child is not None),
        None,
    )
    if successor is None:
        root.append(node)
    else:
        successor.addprevious(node)


def _presentation(project):
    template = project.config.get("template", {})
    template_path = project.resolve_input(template["path"]) if template.get("path") else None
    presentation = Presentation(str(template_path)) if template_path else Presentation()
    layout_index = int(template.get("layout_index", 6))
    if not 0 <= layout_index < len(presentation.slide_layouts):
        raise ValueError(f"雛形にレイアウトがない: {layout_index}")
    layout = presentation.slide_layouts[layout_index]
    inspection = template.get("inspection") or {}
    # 資料の読み込み（config）が空白レイアウトを確かめた雛形は、ここで調べ直さない。
    if not (inspection.get("blank_layout") and inspection.get("layout_index") == layout_index):
        allowed_placeholders = {"dt", "ftr", "sldNum"}
        for shape in layout.shapes:
            placeholder = shape._element.find(".//" + tag("p", "ph"))
            if placeholder is None or placeholder.get("type", "obj") not in allowed_placeholders:
                raise ValueError("雛形の layout_index は空白レイアウトを指定する")
    original_slides = len(presentation.slides)
    for slide in list(presentation.slides._sldIdLst):
        presentation.part.drop_rel(slide.rId)
        presentation.slides._sldIdLst.remove(slide)
    return (
        presentation,
        layout,
        {
            "path": str(template_path) if template_path else None,
            "sha256": file_hash(template_path) if template_path else None,
            "layout_index": layout_index,
            "layout_name": layout.name,
            "removed_slides": original_slides,
            "show_master_shapes": bool(template.get("show_master_shapes", False)),
        },
    )


def write_pptx(scene, project, output: Path, *, fonts=None) -> dict:
    """数式本体を通常の図形として含む PPTX を出力する。

    fonts には計測と同じ `measure.FontSource` を渡す。通常文字の書体の区切りと、中央揃えの独立した数式の位置の補正に使う。
    None なら 0.1 と同じく区切らず、補正もしない。
    戻り値の groups・connections・vector_images は、書いた PowerPoint のグループ・接続・SVG でも埋め込んだ画像の数である。
    """
    presentation, layout, template_record = _presentation(project)
    presentation.slide_width, presentation.slide_height = map(Pt, scene["size"])
    _set_kinsoku(presentation)
    metadata = scene.get("metadata", project.config.get("metadata", {}))
    properties = presentation.core_properties
    for field in (
        "title",
        "subject",
        "author",
        "keywords",
        "comments",
        "last_modified_by",
        "category",
        "content_status",
        "identifier",
        "language",
        "version",
    ):
        setattr(properties, field, str(metadata.get(field, "")))
    generated_at = datetime.now(timezone.utc).replace(microsecond=0)
    properties.created = properties.modified = generated_at
    for field in list(properties._element):
        if etree.QName(field).localname == "lastPrinted":
            properties._element.remove(field)
    properties.revision = 1
    functions = tuple(project.config.get("math", {}).get("functions", ()))
    counts = {"slides": 0, "shapes": 0, "native_math": 0, "groups": 0, "connections": 0, "vector_images": 0}
    svg_parts: dict = {}
    for page in scene["slides"]:
        slide = presentation.slides.add_slide(layout)
        slide._element.find(tag("p", "cSld")).set("name", page["id"])
        slide._element.set("showMasterSp", "1" if template_record["show_master_shapes"] else "0")
        for shape in list(slide.shapes):
            shape._element.getparent().remove(shape._element)
        by_id = {element["id"]: element for element in page["elements"]}
        groups: dict = {}
        drawn = {}
        for element in page["elements"]:
            if element["kind"] == "group":
                continue
            group = pptx_group_of(element, by_id)
            if group is not None and _has_math(element):
                raise ValueError(
                    f"数式を含む要素は PowerPoint のグループに入れられない（まとまり {group} の pptx_group を偽にする）: "
                    f"{element['id']}"
                )
            shape = _draw(_group_shapes(slide, group, by_id, groups), element, functions, fonts, svg_parts)
            drawn[element["id"]] = shape
            counts["shapes"] += 1
            counts["native_math"] += len(shape._element.findall(".//" + tag("m", "oMath")))
            counts["vector_images"] += bool(element["kind"] == "image" and element.get("vector_path"))
        counts["groups"] += len(groups)
        counts["connections"] += _connect_lines(page["elements"], by_id, drawn, project)
        slide.notes_slide.notes_text_frame.text = page.get("notes") or ""
        counts["slides"] += 1
    output = Path(output)
    with atomic_output(output) as stream:
        presentation.save(stream)
    return {
        "output": str(output),
        **counts,
        "template": template_record,
        "generated_at": generated_at.isoformat(),
        "sha256": file_hash(output),
    }
