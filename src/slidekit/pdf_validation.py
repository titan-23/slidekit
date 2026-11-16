"""PDF の検査。ページの数と寸法、抽出できる文字と書体、要素ごとの字形のインクと行数を検査し、表示確認を補助する。

字形は `pdf_layout.read_slide` で描画順と位置から要素へ割り当て、インクを要素の文字領域と比べる。
計測の予測（`measure.layout_scene` の結果）があれば行数も比べ、予測と実際を並べた記録（`compare`）を作る。
PPTX の検査と共有する文字の正規化（`_normalize_text`。NFKC の後に空白と書式文字を除く。字形の読み取りと同じ規則）は
`pptx_validation` に、PDF の書体名の正規化（`_font_name`）、要素の枠（`_area`）、位置で探すときの 1 em（`_em`）は
`pdf_layout` にある。検査は入力を変更しない。
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Callable, Mapping
from functools import partial
from pathlib import Path
import re
from typing import Any

from .layout import Rect, as_rect
from .math.commands import SPACES
from .measure import MEASURE_VERSION, MODEL, ElementLayout, FontSource, _visible, layout_scene, text_area
from .pdf_layout import RenderedText, _area, _em, _font_name, _text_in_math, read_slide
from .pptx_validation import _normalize_text
from .validators import _contains, _content, _has_math, _normal_text, make_diagnostic, merge_diagnostics

RECORD_VERSION = 1  # renders/powerpoint/<表示>/layout.json の形式。形を変えたら上げる
_TEXT_LENGTH = 40  # 診断に添える要素の文字の長さ
_EPSILON = 1e-6
# 字形を描かない数式の命令（空白の命令と \text の名前）。数式からこれと波括弧と空白を除いて何も残らなければ、字形を求めない。
_NO_GLYPH = re.compile("|".join(map(re.escape, (*SPACES, "\\text"))))


def validate_pdf(
    path: Path,
    scene: dict,
    project: Any,
    variant: str = "standard",
    *,
    fonts: FontSource | None = None,
    layouts: Mapping | None = None,
    page_evidence: Mapping[str, str] | None = None,
    layout_record: dict | None = None,
) -> list[dict]:
    """PDF のページの数と寸法、抽出できる文字と書体、要素ごとの字形のインクと行数を検査する。

    `layouts` は `measure.layout_scene()` の結果である。渡さなければ `fonts`（`measure.FontSource`）から計算し、
    `fonts` もなければ予測を使う検査（行数）を省く。インクは予測を使わずに、要素の文字領域と比べる
    （許容幅は、塗りか線のある枠 0.5 pt、ない枠 1.5 pt）。
    `page_evidence` はページの識別子から表示の証拠（そのページ・表示方法）への対応である。渡すと、
    ページの診断の証拠をその値と表示方法にし、要素の診断はそれに要素を加える。ページ番号が変わっても、
    表示が同じなら証拠は変わらない。渡さなければ、要素かページの全体と表示環境を証拠にする。
    `layout_record` に辞書を渡すと、`renders/powerpoint/<表示>/layout.json` の内容（予測と実際）をそこへ書く。
    欠落の候補（`pdf_text_missing`）は 2 つの方法で示す。文字の書体で描く部分（通常の文字と数式の `\\text{...}` の中身）は
    ページの抽出した文字と照らし、数式は、数式の字形も数式の線も割り当てられなかった要素の一覧で示す（文字と数式が混じる要素で
    数式だけが描かれない場合を含む。空白の命令だけの数式は除く）。
    字形を 1 つも割り当てられなかった要素は、行数とインクを比べない。文字の照合で欠けた文字があるか、割り当てられない字形がある
    ページでは、インクが要素の枠を 1 em 広げた範囲と交わらない要素も比べない（記録には読んだとおりに書く）。
    同じ文字列の別の要素や雛形の字形を割り当てたとみなすためである。
    """
    from .pdf import open_pdf

    issues: list[dict] = []
    scope = f"pdf:powerpoint:{variant}"
    environment = project.environment.get("powerpoint", {})
    rasterized_math = variant == "fallback"
    if layouts is None and fonts is not None:
        layouts = layout_scene(scene, fonts)
    pages: list[dict] = []

    def report(code, message, *, slide=None, number=None, element=None, observations=None):
        # 重大度は診断の表の値を使う（make_diagnostic が決める）。
        visual = page_evidence.get(slide.get("id")) if slide is not None and page_evidence is not None else None
        if visual is not None:
            evidence = {"page": visual, "renderer": "powerpoint", "variant": variant}
            if element is not None:
                evidence["element"] = element
        else:
            evidence = {
                "element": element,
                "slide": slide if element is None else None,
                "renderer": "powerpoint",
                "variant": variant,
                "environment": environment,
                "fonts": project.environment.get("fonts", {}),
            }
        issues.append(
            make_diagnostic(
                code,
                message,
                scope=scope,
                slide_id=slide.get("id") if slide is not None else None,
                element_id=element.get("id") if element is not None else None,
                observations=observations,
                evidence=evidence,
                slide_number=number,
            )
        )

    try:
        with open_pdf(path) as document:
            slides = scene["slides"]
            if len(document) != len(slides):
                report(
                    "pdf_page_count",
                    "PDF のページ数が一致しない",
                    observations={"expected": len(slides), "actual": len(document)},
                )
            for position, (page, slide) in enumerate(zip(document, slides), 1):
                number = slide.get("number", position)
                pages.append(
                    _check_page(
                        page,
                        slide,
                        number,
                        partial(report, slide=slide, number=number),
                        size=scene["size"],
                        rasterized_math=rasterized_math,
                        layouts=layouts,
                    )
                )
    except (OSError, ValueError, RuntimeError) as exc:
        report("pdf_unreadable", "PDF を検査できない", observations={"reason": str(exc)})
    if layout_record is not None:
        layout_record.update(
            {
                "schema_version": RECORD_VERSION,
                "renderer": "powerpoint",
                "variant": variant,
                "measure_version": MEASURE_VERSION,
                "slides": pages,
            }
        )
    return merge_diagnostics(issues)


def _check_page(
    page: Any, slide: Mapping, number: Any, report: Callable, *, size, rasterized_math: bool, layouts: Mapping | None
) -> dict:
    """1 ページを検査し、layout.json のそのページの記録を返す。"""
    if abs(page.width - size[0]) > 0.2 or abs(page.height - size[1]) > 0.2:
        report(
            "pdf_page_size",
            "PDF のページ寸法が一致しない",
            observations={"expected": size, "actual": [page.width, page.height]},
        )
    # 見えない文字片（不透明度 0。PowerPoint が説明のある要素の数式に加える読み上げ用の文字など）は表示されないので、
    # 文字の照合・位置・書体の検査に使わない（字形の読み取りと同じ）。
    spans = [span for span in page.text_spans() if span.opacity > 0]
    # 数式を画像にした表示では、数式を含む要素は画像なので、文字と書体を求めない。
    present = [element for element in slide["elements"] if not (rasterized_math and _has_math(element))]
    reading = read_slide(page, slide, rasterized_math=rasterized_math)

    expected_text = "".join(_drawn_text(element) for element in present)
    missing = Counter(_normalize_text(expected_text)) - Counter(_normalize_text("".join(span.text for span in spans)))
    # 数式の書体の字形は文字の照合に入らないので、数式の字形を 1 つも割り当てられなかった要素を別に示す。
    without_glyphs = [
        str(element.get("id"))
        for element in present
        if _math_not_drawn(element, reading.elements.get(str(element.get("id"))))
    ]
    if missing or without_glyphs:
        observations = {"missing_characters": dict(sorted(missing.items())), "missing_count": sum(missing.values())}
        if without_glyphs:
            observations["math_elements_without_glyphs"] = without_glyphs
        report("pdf_text_missing", "PDF の文字抽出で欠落の候補が見つかった", observations=observations)

    expected_fonts = _expected_fonts(present)
    for span in spans:
        # 空白と書式文字（数式だけの段落に加える U+200B など）だけの文字片は表示されないので、位置と書体を検査しない。
        if not _normalize_text(span.text):
            continue
        x0, y0, x1, y1 = span.bbox
        span_area = (x0, y0, x1 - x0, y1 - y0)
        if not _contains((0, 0, *size), span_area, tolerance=0.5):
            report(
                "pdf_text_overflow",
                "表示された文字がページの外へはみ出している",
                observations={"text": span.text, "area": span_area},
            )
        if expected_fonts and _font_name(span.font) not in expected_fonts:
            report(
                "pdf_font_substitution",
                "PDF に指定外のフォントが含まれている",
                observations={
                    "actual_font": span.font,
                    "expected_fonts": sorted(expected_fonts),
                    "text": span.text,
                },
            )
    if reading.unassigned:
        report(
            "pdf_text_unassigned",
            "PDF の字形の一部をどの要素にも割り当てられない",
            observations={"count": reading.unassigned},
        )

    # 描くはずの文字が欠けているか、どの要素のものでもない字形があると、字形の文字列だけの対応付けは、
    # 同じ文字列の別の要素や雛形の字形を割り当てることがある。
    uncertain = bool(missing) or reading.unassigned > 0
    records = []
    for element in slide["elements"]:
        actual = reading.elements.get(str(element.get("id")))
        if actual is None:
            continue
        predicted = layouts.get((slide.get("id"), element.get("id"))) if layouts else None
        predicted = predicted if isinstance(predicted, ElementLayout) else None
        records.append({"id": str(element.get("id")), **compare(predicted, actual)})
        # 字形のない要素と、割り当ての誤りとみなす要素は比べない（欠落はページの pdf_text_missing が示す）。
        if actual.glyphs == 0 or (uncertain and _far_from_frame(element, actual)):
            continue
        overflow = _ink_overflow(element, actual, predicted)
        if overflow is not None:
            report(
                "pdf_element_text_overflow",
                "表示された文字のインクが文字領域から許容幅を超えてはみ出している",
                element=element,
                observations=overflow,
            )
        if predicted is not None:
            lines = _content_lines(predicted)
            if len(lines) != len(actual.lines):
                report(
                    "pdf_line_count_mismatch",
                    "PDF の行数が計測と異なる",
                    element=element,
                    observations={
                        "predicted_lines": len(lines),
                        "actual_lines": len(actual.lines),
                        "predicted_line_widths_pt": [_round(line.width) for line in lines],
                        "actual_line_widths_pt": [_round(line.right - line.left) for line in actual.lines],
                        "available_width_pt": _round(predicted.inner.w),
                        "suggestion": "枠の幅を広げるか、改行を明示する",
                    },
                )
    return {"id": slide.get("id"), "number": number, "unassigned_glyphs": reading.unassigned, "elements": records}


def _drawn_text(element: Mapping) -> str:
    """文字の書体で描かれ、PDF から文字として抽出できるはずの部分（通常の文字と、数式の \\text{...} の中身）。"""
    return "".join(
        (
            _text_in_math(str(fragment.get("text", "")))
            if fragment.get("kind") == "math"
            else str(fragment.get("text", ""))
        )
        for fragment in _content(element)
    )


def _draws_glyphs(expression: str) -> bool:
    """数式が字形を描くか。空白の命令（\\quad など）と \\text の名前、波括弧、空白を除いて文字が残れば描くとみなす。"""
    return bool(_normalize_text(_NO_GLYPH.sub("", expression).replace("{", "").replace("}", "")))


def _math_not_drawn(element: Mapping, actual: RenderedText | None) -> bool:
    """字形を描くはずの数式を含む要素に、数式の字形も数式の線も割り当てられなかったか。読まなかった要素は調べない。"""
    return (
        actual is not None
        and actual.math_glyphs == 0
        and actual.drawings == 0
        and any(
            fragment.get("kind") == "math" and _draws_glyphs(str(fragment.get("text", "")))
            for fragment in _content(element)
        )
    )


def _far_from_frame(element: Mapping, actual: RenderedText) -> bool:
    """インク（割り当てた字形と数式の線）が、要素の枠を 1 em 広げた範囲と交わらないか。"""
    area = _area(element)
    if area is None or actual.ink is None:
        return False
    x, y, w, h = area
    margin = _em(element, _content(element))
    x0, y0, x1, y1 = actual.ink
    return x1 < x - margin or x + w + margin < x0 or y1 < y - margin or y + h + margin < y0


def _expected_fonts(elements: list[dict]) -> set[str]:
    """ページの文字に使われるはずの書体（正規化した名前）。数式の \\text{...} は要素の文字の書体で描かれる。"""
    result = set()
    for element in elements:
        style = element.get("style", {})
        content = _content(element)
        math = [str(fragment.get("text", "")) for fragment in content if fragment.get("kind") == "math"]
        if _normal_text(element) or any(_normalize_text(_text_in_math(expression)) for expression in math):
            result.update(_font_name(style[key]) for key in ("font_japanese", "font_latin") if style.get(key))
        if math and style.get("font_math"):
            result.add(_font_name(style["font_math"]))
    return result


def _text_region(element: Mapping, predicted: ElementLayout | None) -> tuple[Rect, bool] | None:
    """インクを比べる文字領域と、塗りか線のある枠か。予測がなければ要素の領域と図形から求める。"""
    if predicted is not None:
        return predicted.text_area, predicted.visible_frame
    style = element.get("style") or {}
    shape = element.get("shape", "rectangle") if element.get("kind") == "box" else "rectangle"
    try:
        return text_area(as_rect(element.get("area")), shape, style.get("corner_radius_pt")), _visible(style)
    except (TypeError, ValueError):
        return None


def _ink_overflow(element: Mapping, actual: RenderedText, predicted: ElementLayout | None) -> dict | None:
    """実際のインクが文字領域から許容幅を超えてはみ出していれば、その観測値を返す。"""
    region = _text_region(element, predicted)
    if region is None or actual.ink is None:
        return None
    area, visible = region
    tolerance = MODEL.visible_tolerance_pt if visible else MODEL.hidden_tolerance_pt
    x0, y0, x1, y1 = actual.ink
    overflow = {"left": area.x - x0, "top": area.y - y0, "right": x1 - area.right, "bottom": y1 - area.bottom}
    side = max(overflow, key=overflow.get)
    if overflow[side] <= tolerance + _EPSILON:
        return None
    return {
        "text": _element_text(element)[:_TEXT_LENGTH],
        "ink_pt": [_round(value) for value in actual.ink],
        "predicted_ink_pt": _bounds(predicted.ink) if predicted is not None else None,
        "text_area_pt": [_round(value) for value in area],
        "overflow_pt": {"side": side, "amount": _round(overflow[side])},
    }


def compare(predicted: ElementLayout | None, actual: RenderedText) -> dict:
    """予測（`measure.ElementLayout`）と実際（`pdf_layout.RenderedText`）を、layout.json の 1 要素分の形で並べる。

    戻り値は `predicted`・`actual`・`difference_pt` を持つ。座標は絶対の pt で、インクは (x0, y0, x1, y1) である。
    予測の行には、内容のある行（空の行と空白だけの行を除く）だけを含める。PDF には字形のある行だけが現れるためである。
    差は実際 − 予測で、`max_line_width` は行数が同じときの行の幅の差のうち、絶対値が最大のものである。
    予測がなければ `predicted` と `difference_pt` を None にする。
    """
    observed = {
        "lines": len(actual.lines),
        "line_widths_pt": [_round(line.right - line.left) for line in actual.lines],
        "baselines_pt": [_round(line.baseline) for line in actual.lines],
        "ink_pt": None if actual.ink is None else [_round(value) for value in actual.ink],
        "glyphs": actual.glyphs,
        "drawings": actual.drawings,
        "sizes_pt": list(actual.sizes),
    }
    if predicted is None:
        return {"predicted": None, "actual": observed, "difference_pt": None}
    lines = _content_lines(predicted)
    expected = {
        "lines": len(lines),
        "line_widths_pt": [_round(line.width) for line in lines],
        "baselines_pt": [_round(predicted.block_top + line.baseline) for line in lines],
        "ink_pt": _bounds(predicted.ink),
    }
    ink = None
    if predicted.ink is not None and actual.ink is not None:
        box = predicted.ink
        ink = [_round(a - b) for a, b in zip(actual.ink, (box.x, box.y, box.right, box.bottom))]
    widths = (
        [(line.right - line.left) - expected_line.width for line, expected_line in zip(actual.lines, lines)]
        if len(lines) == len(actual.lines)
        else []
    )
    difference = {"ink": ink, "max_line_width": _round(max(widths, key=abs)) if widths else None}
    return {"predicted": expected, "actual": observed, "difference_pt": difference}


def _content_lines(predicted: ElementLayout) -> list:
    """予測の行のうち、表示される文字か数式のある行。"""
    return [line for line in predicted.text.lines if _normalize_text(line.text)]


def _element_text(element: Mapping) -> str:
    """診断に添える要素の文字。数式は $…$ で表す（計測の行の表記と同じ）。"""
    return "".join(
        f"${fragment.get('text', '')}$" if fragment.get("kind") == "math" else str(fragment.get("text", ""))
        for fragment in _content(element)
    )


def _bounds(rect: Rect | None) -> list[float] | None:
    return None if rect is None else [_round(rect.x), _round(rect.y), _round(rect.right), _round(rect.bottom)]


def _round(value: float) -> float:
    return round(float(value), 2)
