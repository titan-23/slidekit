"""入力・配置・記述規則の検査と、診断の共通の補助。検査は入力を変更しない。

構造の誤り（識別子・領域・親子関係・役割など）は作成 API が例外にするので、場面の検査では調べない。
診断の重大度・区分・証拠の範囲は `diagnostics` の表にある（この表は `validators` からも同じ名前で読み込むことができる）。
文字と数式の寸法は `measure`、記述規則の判定は `writing` で求める。
PPTX の検査は `pptx_validation`、PDF の検査は `pdf_validation` にある。
"""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Mapping, Sequence
import math
from pathlib import Path
from typing import Any

from . import measure, writing

# 診断の表は validators からも同じ名前で読み込むことができる（DiagnosticSpec・WARNING_CODES はそのために読み込む）。
from .diagnostics import DIAGNOSTICS, WARNING_CODES, DiagnosticSpec, EvidenceScope
from .evidence import EVIDENCE_VERSION, scoped_evidence
from .files import canonical_json, json_digest, json_safe
from .math.syntax import MathSyntaxError, parse, text_contents
from .styles import LIBRARY_DEFAULTS, fragment_style

# 場面の検査で扱う要素の種類。ほかの種類（手で作った場面）は検査しない。
KINDS = frozenset({"text", "box", "line", "arrow", "image", "equation", "group"})
_LINE_KINDS = frozenset({"line", "arrow"})
_SEVERITIES = ("error", "warning", "info")
_SIDES = ("left", "top", "right", "bottom")
# 表にない識別子の証拠の範囲（0.1 と同じ広い範囲）。
_DEFAULT_SCOPE = EvidenceScope()


def _evidence_source_paths(value: Any) -> Any:
    """実行ごとに変わる変換先（path・vector_path）を、元素材のパスで表す。内容は元素材のハッシュ値で追う。"""
    if isinstance(value, dict):
        result = {key: _evidence_source_paths(item) for key, item in value.items()}
        if result.get("kind") == "image" and result.get("source_path"):
            result["path"] = result["source_path"]
            if "vector_path" in result:
                result["vector_path"] = result["source_path"]
        return result
    if isinstance(value, (list, tuple)):
        return [_evidence_source_paths(item) for item in value]
    return value


def make_diagnostic(
    code: str,
    message: str,
    *,
    severity: str | None = None,
    slide_id: str | None = None,
    element_id: str | None = None,
    related_element_id: str | None = None,
    scope: str = "layout",
    observations: Any = None,
    evidence: Any = None,
    unordered: bool = False,
    slide_number: int | None = None,
) -> dict:
    """診断を 1 件作る。照合用の識別子と、再確認の要否を決める証拠を分ける。

    重大度を省くと診断の表（`diagnostics.DIAGNOSTICS`）の値を使い、表にない識別子はエラーにする。
    表にある識別子に表と違う重大度を渡すと ValueError にする（診断の作り方の誤り）。
    証拠のハッシュ値には証拠の形式の版と、表の判定の方法の版（`evidence.version`）を入れる。
    """
    spec = DIAGNOSTICS.get(code)
    if severity is None:
        severity = "error" if spec is None else spec.severity
    elif severity not in _SEVERITIES:
        raise ValueError(f"診断の重大度が不正である: {severity!r}")
    elif spec is not None and severity != spec.severity:
        raise ValueError(f"診断 {code} の重大度は表の {spec.severity} にする（指定: {severity}）")
    if unordered and element_id is not None and related_element_id is not None:
        element_id, related_element_id = sorted((element_id, related_element_id))
    key = [code, slide_id, element_id, related_element_id, scope]
    observed = [] if observations is None else observations
    if not isinstance(observed, list):
        observed = [observed]
    observed = json_safe(observed)
    version = 1 if spec is None else spec.evidence.version
    result = {
        "code": code,
        "slide_id": slide_id,
        "element_id": element_id,
        "related_element_id": related_element_id,
        "scope": scope,
        "severity": severity,
        "message": message,
        "observations": observed,
        "diagnostic_id": "d" + json_digest(key)[:24],
        "evidence_hash": json_digest(
            {
                "format": EVIDENCE_VERSION,
                "version": version,
                "key": key,
                "severity": severity,
                "observations": observed,
                "evidence": _evidence_source_paths(evidence),
            }
        ),
    }
    if slide_number is not None:
        result["slide_number"] = slide_number
    return result


def merge_diagnostics(issues: list[dict]) -> list[dict]:
    """同じ種類・対象・範囲の診断を、観測値を失わず一件にまとめる。"""
    groups: dict[str, list[dict]] = defaultdict(list)
    for issue in issues:
        groups[issue["diagnostic_id"]].append(issue)
    merged = []
    for key in sorted(groups):
        group = groups[key]
        if len(group) == 1:
            merged.append(group[0])
            continue
        result = dict(group[0])
        observations = {canonical_json(obs): obs for issue in group for obs in issue["observations"]}
        result["observations"] = [observations[k] for k in sorted(observations)]
        result["severity"] = min((issue["severity"] for issue in group), key={"error": 0, "warning": 1, "info": 2}.get)
        result["evidence_hash"] = json_digest(sorted({issue["evidence_hash"] for issue in group}))
        result["message"] = " / ".join(sorted({issue["message"] for issue in group}))
        merged.append(result)
    return merged


def _content(element: dict) -> list[dict]:
    """要素の断片の列。断片の列でない値（手で作った場面）は空とみなす。"""
    value = element.get("content")
    if not isinstance(value, (list, tuple)):
        return []
    return [fragment for fragment in value if isinstance(fragment, Mapping)]


def _has_math(element: dict) -> bool:
    return any(fragment.get("kind") == "math" for fragment in _content(element))


def _has_content(element: dict) -> bool:
    return any(str(fragment.get("text", "")).strip() for fragment in _content(element))


def _normal_text(element: dict) -> str:
    return "".join(str(fragment.get("text", "")) for fragment in _content(element) if fragment.get("kind") != "math")


def _number(value: Any) -> bool:
    return not isinstance(value, bool) and isinstance(value, (int, float)) and math.isfinite(value)


def _rect(value: Any, allow_zero: bool = False) -> tuple[float, float, float, float] | None:
    if not isinstance(value, (list, tuple)) or len(value) != 4:
        return None
    if any(not _number(v) for v in value):
        return None
    x, y, w, h = map(float, value)
    if w < 0 or h < 0 or (not allow_zero and (w == 0 or h == 0)):
        return None
    return x, y, w, h


def _point(value: Any) -> tuple[float, float] | None:
    if not isinstance(value, (list, tuple)) or len(value) != 2 or any(not _number(v) for v in value):
        return None
    return float(value[0]), float(value[1])


def _contains(outer: tuple, inner: tuple, tolerance: float = 0.1) -> bool:
    x, y, w, h = outer
    a, b, c, d = inner
    return a >= x - tolerance and b >= y - tolerance and a + c <= x + w + tolerance and b + d <= y + h + tolerance


def _intersection(a: tuple, b: tuple) -> tuple | None:
    x, y = max(a[0], b[0]), max(a[1], b[1])
    right, bottom = min(a[0] + a[2], b[0] + b[2]), min(a[1] + a[3], b[1] + b[3])
    if right - x <= 0.1 or bottom - y <= 0.1:
        return None
    return x, y, right - x, bottom - y


def _overflow(outer: tuple, inner: tuple) -> dict:
    """inner が outer から最も大きくはみ出す辺と量（pt）。"""
    amounts = {
        "left": outer[0] - inner[0],
        "top": outer[1] - inner[1],
        "right": inner[0] + inner[2] - outer[0] - outer[2],
        "bottom": inner[1] + inner[3] - outer[1] - outer[3],
    }
    side = max(_SIDES, key=amounts.get)
    return {"side": side, "amount": round(amounts[side], 2)}


def _drawn_area(element: dict, area: tuple) -> tuple:
    """重なりの判定に使う、実際に描く範囲。画像は作成 API が決めた actual_area（contain では余白を除く。不正なら領域）である。"""
    if element.get("kind") == "image":
        actual = _rect(element.get("actual_area"))
        if actual is not None:
            return actual
    return area


def _line_width(element: dict) -> float:
    width = (element.get("style") or {}).get("line_width_pt", LIBRARY_DEFAULTS["line_width_pt"])
    return float(width) if _number(width) and width >= 0 else float(LIBRARY_DEFAULTS["line_width_pt"])


def _filled(element: dict) -> bool:
    return (element.get("style") or {}).get("fill") not in (None, False, "", "none", "transparent")


def _line_contact(start: tuple, end: tuple, area: tuple, width: float) -> str | None:
    """線分と図形の矩形の関係。交わらなければ None、端だけが入るなら "endpoint"、それ以外は "cross" を返す。

    矩形を線幅の半分だけ広げて線分を切り取る。切り取った区間 [t0, t1]（線分上の位置を 0〜1 で表す）が、
    どちらかの端から長さ max(2 pt, 線幅) + 線幅 / 2 の範囲にあれば "endpoint" とする。辺で止まる線と、
    図形への max(2 pt, 線幅) までの入り込み（辺に垂直な線で計測する）が当たる。境界の値（ちょうど 2 pt など）の判定が
    浮動小数点の誤差で揺れないよう、1e-6 pt だけ許容する。図形に深く入る線、貫通する線、図形の内部から出る線は "cross" になる。
    """
    pad = width / 2
    x, y, w, h = area[0] - pad, area[1] - pad, area[2] + 2 * pad, area[3] + 2 * pad
    dx, dy = end[0] - start[0], end[1] - start[1]
    lower, upper = 0.0, 1.0
    for p, q in ((-dx, start[0] - x), (dx, x + w - start[0]), (-dy, start[1] - y), (dy, y + h - start[1])):
        if p == 0:
            if q < 0:
                return None
            continue
        ratio = q / p
        if p < 0:
            lower = max(lower, ratio)
        else:
            upper = min(upper, ratio)
        if lower > upper:
            return None
    # 区間が線分の端からどこまで届くかを、長さ（pt）で比べる。長さ 0 の線分は交われば端の接触とみなす。
    length = math.hypot(dx, dy)
    reach = max(2.0, width) + pad + 1e-6
    return "endpoint" if upper * length <= reach or (1 - lower) * length <= reach else "cross"


def _ancestors(element: dict, elements: dict) -> set[str]:
    parents: set[str] = set()
    current = element.get("parent_id")
    while current is not None and current not in parents:
        parents.add(current)
        current = elements.get(current, {}).get("parent_id")
    return parents


def _font_evidence(fonts: Any) -> Any:
    try:
        return fonts.inventory()
    except (AttributeError, OSError, ValueError) as exc:
        return {"inventory_error": str(exc)}


def _syntax_error(error: BaseException) -> MathSyntaxError | None:
    """例外の原因をたどり、数式の記法の誤りを探す。"""
    seen = set()
    current: BaseException | None = error
    while current is not None and id(current) not in seen:
        if isinstance(current, MathSyntaxError):
            return current
        seen.add(id(current))
        current = current.__cause__ or current.__context__
    return None


def _syntax_observations(error: MathSyntaxError, expression: str) -> dict:
    # position は式の 0 始まりの文字位置。message の「n 文字目」は 1 始まり。
    return {
        "expression": error.expression or expression,
        "position": error.position,
        "token": error.token,
        "reason": error.reason,
        "hint": error.hint,
        "message": str(error),
    }


def _measurement_observations(error: measure.MeasureError) -> dict:
    # 元の例外から作った誤りは、details の reason（元の例外の文言）を残す。
    return {"reason": str(error), **error.details}


def _inside_frame(
    frame: dict, inner: dict, frame_area: tuple, inner_area: tuple, frame_first: bool, layout: Any
) -> bool:
    """inner が枠 frame の内側に置かれた要素として重なりの判定から外れるか。

    枠は box だけで、inner の領域（線は両端の点）が枠の領域に収まる（許容 0.1 pt）ことが条件である。
    先に描いた枠は、文章がないか、文章（予測したインクを max(1 pt, 0.2 × 文字サイズ) 広げた矩形）と inner が重ならなければ外す。
    計測できなかった枠の文章は矩形を求めず、外さない。後に描く枠は、塗りも文章もない透明な枠だけを外す。
    """
    if frame.get("kind") != "box":
        return False
    line = inner.get("kind") in _LINE_KINDS
    if line:
        points = [_point(inner.get("start")), _point(inner.get("end"))]
        inside = all(point is not None and _contains(frame_area, (*point, 0.0, 0.0)) for point in points)
    else:
        inside = _contains(frame_area, inner_area)
    if not inside:
        return False
    if not frame_first:
        return not _has_content(frame) and not _filled(frame)
    if not _has_content(frame):
        return True
    if not isinstance(layout, measure.ElementLayout):
        return False
    if layout.ink is None:
        return True
    size = (frame.get("style") or {}).get("size_pt")
    margin = max(1.0, 0.2 * size) if _number(size) else 1.0
    ink = layout.ink
    block = (ink.x - margin, ink.y - margin, ink.w + 2 * margin, ink.h + 2 * margin)
    if line:
        return _line_contact(points[0], points[1], block, _line_width(inner)) is None
    return _intersection(block, inner_area) is None


def _check_measurement(element: dict, layout: Any, functions: Sequence[str], fonts: Any, report) -> None:
    """文字と数式の寸法の検査：はみ出し・余白・折り返しの判定と、計測できない入力（字形・記法・書式）。

    数式は出力と同じ構文解析（`math.syntax.parse`）で記法を確かめ、`\\text` の中身も字形を調べる。
    記法の誤りは計測の誤り（text_measurement_failed）と分けて math_syntax_error で報告する。
    """
    style = element.get("style") or {}
    syntax_reported = False
    for fragment in _content(element):
        if fragment.get("kind") != "math" or not str(fragment.get("text", "")).strip():
            continue
        expression = str(fragment["text"])
        try:
            nodes = parse(expression, functions)
        except MathSyntaxError as error:
            report(
                "math_syntax_error",
                "数式の記法が誤っている",
                element,
                scope="input",
                observations=_syntax_observations(error, expression),
            )
            syntax_reported = True
            continue
        except (TypeError, ValueError):
            continue  # 追加の関数名の誤りなどは計測の誤りとして報告する
        bold = fragment_style(style, fragment)["bold"]
        for text in text_contents(nodes):
            try:
                missing = measure.missing_glyphs(text, style, fonts, bold=bold)
            except measure.MeasureError as error:
                report(
                    error.code,
                    "数式の \\text の文字を検査できない",
                    element,
                    scope="input",
                    observations=_measurement_observations(error),
                )
                continue
            if missing:
                report(
                    "font_glyph_missing",
                    "和文と欧文のどちらのフォントにも、必要な字形がない",
                    element,
                    scope="input",
                    observations={
                        "characters": missing,
                        "fonts": [style.get("font_latin"), style.get("font_japanese")],
                        "expression": expression,
                    },
                )
    if isinstance(layout, measure.MeasureError):
        syntax = _syntax_error(layout)
        if syntax is not None:
            if not syntax_reported:
                report(
                    "math_syntax_error",
                    "数式の記法が誤っている",
                    element,
                    scope="input",
                    observations=_syntax_observations(syntax, str(syntax.expression)),
                )
        elif layout.code == "font_glyph_missing":
            report(
                "font_glyph_missing",
                "和文と欧文のどちらのフォントにも、必要な字形がない",
                element,
                scope="input",
                observations=_measurement_observations(layout),
            )
        else:
            report(
                layout.code,
                "文字または数式の寸法を計算できない",
                element,
                scope="input",
                observations=_measurement_observations(layout),
            )
        return
    if isinstance(layout, measure.ElementLayout):
        for finding in measure.fit_findings(layout):
            report(finding["code"], finding["message"], element, observations=finding["observations"])


def _font_sizes(element: dict, roles: Mapping) -> list[dict]:
    """役割の最小値より小さい文字サイズ。要素の大きさと、大きさを指定した断片ごとに調べる。"""
    style = element.get("style") or {}
    role_name = element.get("role") or "body"
    role = roles.get(role_name)
    minimum = (role if isinstance(role, Mapping) else {}).get("min_size_pt", style.get("min_size_pt", 0))
    if not _number(minimum):
        return []
    found = []
    size = style.get("size_pt")
    if _number(size) and size < minimum:
        found.append({"size_pt": size, "min_size_pt": minimum, "role": role_name})
    for number, fragment in enumerate(_content(element), 1):
        if "size_pt" not in (fragment.get("style") or {}):
            continue
        size = fragment_style(style, fragment)["size_pt"]
        if _number(size) and size < minimum:
            found.append({"fragment": number, "size_pt": size, "min_size_pt": minimum, "role": role_name})
    return found


def _check_writing(element: dict, slide: dict, rules: Mapping, report) -> None:
    """記述規則（禁止文字、数式の候補、明示改行の禁則）の検査。"""
    text = _normal_text(element)
    forbidden = writing.forbidden_characters(text, rules.get("forbidden_characters", []))
    if forbidden:
        report(
            "forbidden_character",
            "禁止された文字が含まれている",
            element,
            scope="writing",
            observations={"characters": forbidden},
        )
    config = rules.get("math_candidates", {})
    if writing.math_check_applies(element, slide, config):
        candidates = writing.math_candidates(writing.candidate_text(_content(element)), config)
        if candidates:
            report(
                "math_not_marked",
                "数式として指定すべき文字か確認する必要がある",
                element,
                scope="writing",
                observations={"candidates": candidates},
            )
    problems = writing.line_break_problems(_content(element))
    if problems:
        report(
            "line_break",
            "明示改行の前後に、行頭または行末に置かない文字がある",
            element,
            scope="writing",
            observations=problems,
        )


def _check_sources(element: dict, documents: set, bibliography: set, report) -> None:
    for source in element.get("sources", []) or []:
        if isinstance(source, Mapping):
            document, key = source.get("document_id"), source.get("bib_key")
            valid = (
                bool(document or key)
                and (document is None or document in documents)
                and (key is None or key in bibliography)
            )
        else:
            valid = False
        if not valid:
            report(
                "source_missing", "出典の識別子を参照できない", element, scope="input", observations={"source": source}
            )


def _check_overlaps(valid: list[tuple[dict, tuple]], by_id: dict, layouts: Mapping, slide_id: Any, report) -> None:
    """要素の重なり。valid は描画順の（要素、実際に描く範囲）の列である（画像は actual_area。`_drawn_area`）。

    まとまり、背景と飾り（semantic）、祖先と子孫の組、線どうしの組、`anchor()` で結んだ線と接続先、
    枠の内側に置いた要素（`_inside_frame`）、端が図形の辺に接する線は判定しない。
    """
    candidates = [
        (element, area)
        for element, area in valid
        if element.get("kind") != "group" and element.get("semantic") not in {"background", "decoration"}
    ]
    for index, (first, first_area) in enumerate(candidates):
        for second, second_area in candidates[index + 1 :]:
            if first.get("id") in _ancestors(second, by_id) or second.get("id") in _ancestors(first, by_id):
                continue
            lines = [item for item in (first, second) if item.get("kind") in _LINE_KINDS]
            if len(lines) == 2:
                # 接続線どうしの交差は図の意味に依存するため表示確認で扱う。
                continue
            if lines:
                line = lines[0]
                other = second if line is first else first
                connections = {
                    ref.get("element_id")
                    for ref in (line.get("start_ref"), line.get("end_ref"))
                    if isinstance(ref, Mapping)
                }
                if other.get("id") in connections:
                    continue
            if _inside_frame(
                first, second, first_area, second_area, True, layouts.get((slide_id, first.get("id")))
            ) or _inside_frame(
                second, first, second_area, first_area, False, layouts.get((slide_id, second.get("id")))
            ):
                continue
            if lines:
                start, end = _point(line.get("start")), _point(line.get("end"))
                if (
                    _line_contact(start, end, second_area if line is first else first_area, _line_width(line))
                    != "cross"
                ):
                    continue
                overlap = {"line": [list(start), list(end)]}
            else:
                intersection = _intersection(first_area, second_area)
                if intersection is None:
                    continue
                overlap = {"intersection": intersection}
            opaque = second.get("kind") == "image" or (second.get("kind") == "box" and _filled(second))
            if _has_content(first) and opaque and _contains(second_area, first_area):
                report("text_occluded", "後ろの図形が文章全体を覆っている", first, observations=overlap, related=second)
            else:
                report(
                    "element_overlap", "要素の重なりを確認する必要がある", first, observations=overlap, related=second
                )


def validate_scene(scene: dict, project: Any, fonts: Any, *, layouts: Mapping | None = None) -> list[dict]:
    """配置と記述規則を検査する。文字を縮小したり改行を挿入したりしない。

    fonts は計測に使う `measure.FontSource`（生成では `FontRegistry`）。layouts は `measure.layout_scene(scene, fonts)` の
    結果で、省略するとここで計算する（生成の流れでは 1 回だけ計算し、`layout.json` と PDF の照合でも使う）。
    構造の誤りは作成 API が例外にする。手で作った場面の構造は検査しない（識別子の重複、未対応の種類、不正な領域・端点、
    親子関係、接続先、画像の配置方法、役割の定義と使い方、最小文字サイズの変更。検査できない要素は飛ばす）。
    診断の証拠は診断の表の範囲（`diagnostics.EvidenceScope`）で作り、範囲がフォントを含む種類だけに、
    計測で使ったフォントの一覧を加える。
    """
    issues: list[dict] = []
    rules = project.rules
    roles = project.theme.get("roles", {})
    size = scene.get("size")
    page_area = _rect([0, 0, *size]) if isinstance(size, (list, tuple)) and len(size) == 2 else None
    documents = {document.get("id") for document in project.config.get("inputs", {}).get("documents", [])}
    bibliography = set(project.config.get("_bibliography_keys", ()))
    functions = tuple(project.config.get("math", {}).get("functions", ()))
    if layouts is None:
        layouts = measure.layout_scene(scene, fonts)
    for position, slide in enumerate(scene.get("slides", []), 1):
        slide_id = slide.get("id")
        number = slide.get("number")
        if isinstance(number, bool) or not isinstance(number, int):
            number = position

        def report(
            code: str,
            message: str,
            element: dict | None = None,
            *,
            severity: str | None = None,
            observations: Any = None,
            scope: str = "layout",
            related: dict | None = None,
        ) -> None:
            spec = DIAGNOSTICS.get(code)
            evidence = scoped_evidence(
                _DEFAULT_SCOPE if spec is None else spec.evidence,
                rules=rules,
                roles=roles,
                elements=[item for item in (element, related) if item is not None],
                slide=slide,
                size=size,
            )
            issues.append(
                make_diagnostic(
                    code,
                    message,
                    slide_id=slide_id,
                    slide_number=number,
                    element_id=element.get("id") if element else None,
                    related_element_id=related.get("id") if related else None,
                    scope=scope,
                    severity=severity,
                    observations=observations,
                    unordered=related is not None,
                    evidence=evidence,
                )
            )

        notes = str(slide.get("notes") or "")
        if rules.get("require_notes", False) and not notes.strip():
            report("notes_missing", "発表者ノートが必要である", scope="input")
        if rules.get("notes_style") == "polite":
            candidates = writing.notes_style_candidates(notes)
            if candidates:
                report(
                    "notes_style",
                    "発表者ノートに、です・ます調でない文末がある",
                    scope="writing",
                    observations={"candidates": candidates},
                )
        elements = [element for element in slide.get("elements", []) or [] if isinstance(element, Mapping)]
        by_id = {element.get("id"): element for element in elements}
        content_area = _rect(slide.get("content_area"))
        valid: list[tuple[dict, tuple]] = []
        for element in elements:
            element_id, kind = element.get("id"), element.get("kind")
            if kind not in KINDS:
                continue
            area = _rect(element.get("area"), allow_zero=kind in _LINE_KINDS | {"group"})
            if area is None:
                continue
            painted = area
            if kind in _LINE_KINDS:
                start, end = _point(element.get("start")), _point(element.get("end"))
                if start is None or end is None or start == end:
                    continue
                width = _line_width(element)
                margin = max(width / 2, width * 2 if kind == "arrow" else 0)
                painted = (area[0] - margin, area[1] - margin, area[2] + margin * 2, area[3] + margin * 2)
            if page_area is not None and not _contains(page_area, painted):
                report(
                    "page_overflow",
                    "要素がページの外へはみ出している",
                    element,
                    observations={"area": painted, "page": page_area},
                )
            if (
                content_area
                and kind != "group"
                and not str(element_id).startswith("_theme/")
                and not _contains(content_area, area)
            ):
                report(
                    "content_area_overflow",
                    "要素が本文領域の外に配置されている",
                    element,
                    observations={"area": area, "content_area": content_area},
                )
            bounds = _rect(element.get("layout_bounds"))
            if bounds is not None and not _contains(bounds, area):
                report(
                    "layout_bounds_overflow",
                    "積み上げた要素が指定した領域を超えている",
                    element,
                    observations={"area": list(area), "bounds": list(bounds), "overflow_pt": _overflow(bounds, area)},
                )
            # ページ・本文領域・積み上げの領域は配置した領域で、重なりは実際に描く範囲で判定する。
            valid.append((element, _drawn_area(element, area)))
            if _has_math(element):
                for item in [element] + [
                    by_id[parent] for parent in sorted(_ancestors(element, by_id), key=str) if parent in by_id
                ]:
                    if item.get("rotation") or item.get("flip_h") or item.get("flip_v"):
                        report(
                            "math_transform",
                            "数式を含む要素と親の回転・反転には対応していない",
                            element,
                            scope="input",
                            observations={"transformed_element": item.get("id")},
                        )
            if kind == "image":
                path = Path(str(element.get("path", "")))
                if not path.is_file():
                    report(
                        "asset_missing",
                        "画像ファイルを参照できない",
                        element,
                        scope="input",
                        observations={"path": str(path)},
                    )
            if _has_content(element):
                for found in _font_sizes(element, roles):
                    report("font_too_small", "文字サイズが役割の最小値を下回っている", element, observations=found)
                if kind in {"text", "box", "equation"}:
                    _check_measurement(element, layouts.get((slide_id, element_id)), functions, fonts, report)
                _check_writing(element, slide, rules, report)
            _check_sources(element, documents, bibliography, report)
        _check_overlaps(valid, by_id, layouts, slide_id, report)
    # 計測で実際に読み込んだフォントを、診断の確定時に記録する（証拠の範囲がフォントを含む種類だけ）。
    inventory = _font_evidence(fonts)
    for issue in issues:
        spec = DIAGNOSTICS.get(issue["code"])
        if (_DEFAULT_SCOPE if spec is None else spec.evidence).fonts:
            issue["evidence_hash"] = json_digest({"diagnostic": issue["evidence_hash"], "fonts": inventory})
    return merge_diagnostics(issues)


def _declaration_matches(declaration: Mapping, issue: Mapping) -> bool:
    if declaration.get("code") != issue.get("code"):
        return False
    element, related = declaration.get("element_id"), declaration.get("related_element_id")
    if element is None:
        return True
    if related is None:
        return element in {issue.get("element_id"), issue.get("related_element_id")}
    return {element, related} == {issue.get("element_id"), issue.get("related_element_id")}


def apply_declarations(issues: list[dict], scene: dict) -> list[dict]:
    """資料コードの許容の宣言（ページの `allowed_warnings`）を診断に当てる。入力は変更しない。

    重大度が警告の診断で、同じページの宣言と一致するものの写しに `declared`（`reason`・`location`）を付ける。
    一致の条件は、識別子が同じで、宣言に要素がない（ページ全体）か、宣言の要素が診断の要素か相手の要素であるか、
    宣言の要素の組が診断の組と順不同で一致することである。複数の宣言が一致したときは、要素を多く指定した宣言を使う。
    エラーと参考には付けない。宣言があっても診断は残す。
    どの警告にも一致しない宣言は、参考の診断 `allow_warning_unused` にして末尾に加える。
    """
    declarations = []
    for position, slide in enumerate(scene.get("slides", []), 1):
        number = slide.get("number")
        if isinstance(number, bool) or not isinstance(number, int):
            number = position
        for declaration in slide.get("allowed_warnings") or []:
            if isinstance(declaration, Mapping):
                declarations.append((slide.get("id"), number, declaration))
    used: set[int] = set()
    result = []
    for issue in issues:
        matches = []
        if issue.get("severity") == "warning" and issue.get("slide_id") is not None:
            matches = [
                index
                for index, (slide_id, _number, declaration) in enumerate(declarations)
                if slide_id == issue["slide_id"] and _declaration_matches(declaration, issue)
            ]
        if not matches:
            result.append(issue)
            continue
        used.update(matches)
        chosen = max(
            matches,
            key=lambda index: sum(
                declarations[index][2].get(key) is not None for key in ("element_id", "related_element_id")
            ),
        )
        declaration = declarations[chosen][2]
        result.append(
            {**issue, "declared": {"reason": declaration.get("reason"), "location": declaration.get("location")}}
        )
    unused = []
    for index, (slide_id, number, declaration) in enumerate(declarations):
        if index in used:
            continue
        element, related = declaration.get("element_id"), declaration.get("related_element_id")
        unused.append(
            make_diagnostic(
                "allow_warning_unused",
                "許容の宣言に対応する警告がない",
                scope="declaration",
                slide_id=slide_id,
                slide_number=number,
                element_id=element,
                related_element_id=related,
                unordered=related is not None,
                observations={
                    "code": declaration.get("code"),
                    "reason": declaration.get("reason"),
                    "location": declaration.get("location"),
                },
                evidence=declaration,
            )
        )
    return result + merge_diagnostics(unused)
