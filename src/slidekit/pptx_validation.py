"""PPTX の検査。ZIP 内の参照と、編集可能な本体・ノート・文書情報を照合する。

文字枠は書き出し（`writer`）の規則で書いたものとして照合する。期待する通常文字の区切りと明示改行は
`writer.paragraph_content`、数式の引数は `writer.math_arguments`、中央揃えの独立した数式の左の余白は `writer.display_indent`
で求める。図形は PowerPoint のグループ（p:grpSp）と互換画像（mc:AlternateContent の mc:Choice）の中まで名前で対応づける。
要素を入れるグループは `writer.pptx_group_of`、線の端の接続は `writer.anchored_site` で、書き出しと同じ判定で求める。

PPTX と PDF の検査が共有する補助（`_normalize_text`）はこのモジュールに置き、`pdf_validation` が読み込む。
場面の検査とも共有する補助（`make_diagnostic`・`merge_diagnostics`・`_content` など）は `validators` にある。
"""

from __future__ import annotations

from pathlib import Path
import posixpath
from typing import Any
import unicodedata
from urllib.parse import unquote
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

from .fallback import fallback_required
from .ooxml import EMU_PER_PT, NS
from .styles import fragment_style
from .validators import _content, _has_content, _has_math, _normal_text, make_diagnostic, merge_diagnostics


def _normalize_text(value: str) -> str:
    """照合用に NFKC で正規化し、空白と書式文字（U+200B など）を除く。`pdf_layout._normalize` と同じ規則である。

    書き出しが文中の数式だけの段落に加える U+200B も、ここで除かれる。
    """
    return "".join(
        char
        for char in unicodedata.normalize("NFKC", value)
        if not char.isspace() and unicodedata.category(char) != "Cf"
    )


def _rels_path(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def _target(part: str, target: str) -> str:
    decoded = unquote(target).split("#", 1)[0]
    return (
        posixpath.normpath(posixpath.join(posixpath.dirname(part), decoded))
        if not decoded.startswith("/")
        else decoded.lstrip("/")
    )


def _relationships(archive: ZipFile, part: str) -> dict[str, dict]:
    path = _rels_path(part)
    if path not in archive.namelist():
        return {}
    return {node.attrib["Id"]: dict(node.attrib) for node in ET.fromstring(archive.read(path))}


def _xml_structure(node: Any) -> Any:
    """名前空間の接頭辞や属性順によらず、数式の構造と書式を比較する。"""
    return [node.tag, sorted(node.attrib.items()), node.text or "", [_xml_structure(child) for child in node]]


def _flag(value: str | None) -> bool:
    return value in {"1", "true"}


_P_TAG = f"{{{NS['p']}}}"
_GROUP = _P_TAG + "grpSp"
_ALTERNATE = f"{{{NS['mc']}}}AlternateContent"
# 名前で要素と対応づける図形の要素。
_SHAPES = {_P_TAG + name for name in ("sp", "cxnSp", "pic", "graphicFrame", "contentPart")}
# グループの座標変換を恒等とみなす許容幅（pt。位置と寸法の照合と同じ）。
_IDENTITY_TOLERANCE_PT = 0.1


class _SlideShapes:
    """スライドの図形を、PowerPoint のグループと互換画像の中まで名前で引く索引。

    - shapes：名前から（図形、属する p:grpSp の名前の列（外側から））。互換画像（mc:AlternateContent）は mc:Choice の図形。
    - fallbacks：名前から mc:Fallback。groups：グループの名前から（p:grpSp、属する p:grpSp の名前の列）。
    - names：図形の番号（p:cNvPr の id）から名前。duplicates：重複した名前。
    """

    def __init__(self, tree):
        self.shapes: dict[str, tuple] = {}
        self.fallbacks: dict[str, ET.Element] = {}
        self.groups: dict[str, tuple] = {}
        self.names: dict[str, str] = {}
        self.duplicates: list[str] = []
        if tree is not None:
            self._visit(tree, ())

    def _name(self, node, path: str) -> str | None:
        name_node = node.find(path, NS)
        if name_node is None:
            return None
        if name_node.get("id") is not None:
            self.names.setdefault(name_node.get("id"), name_node.get("name"))
        return name_node.get("name")

    def _add(self, table: dict, name: str | None, value) -> None:
        if name is None:
            return
        if name in self.shapes or name in self.groups:
            self.duplicates.append(name)
        table[name] = value

    def _visit(self, container, chain: tuple) -> None:
        for child in container:
            if child.tag == _ALTERNATE:
                choice = child.find("mc:Choice", NS)
                node = next(iter(choice), None) if choice is not None else None
                if node is not None:
                    name = self._name(node, ".//p:cNvPr")
                    self._add(self.shapes, name, (node, chain))
                    fallback = child.find("mc:Fallback", NS)
                    if name is not None and fallback is not None:
                        self.fallbacks[name] = fallback
            elif child.tag == _GROUP:
                name = self._name(child, "p:nvGrpSpPr/p:cNvPr")
                self._add(self.groups, name, (child, chain))
                self._visit(child, chain + (name,))
            elif child.tag in _SHAPES:
                self._add(self.shapes, self._name(child, ".//p:cNvPr"), (child, chain))


def _group_chain(element: dict, by_id: dict) -> tuple:
    """要素を入れる PowerPoint のグループの名前の列（外側から）。書き出しと同じ `writer.pptx_group_of` で求める。"""
    from .writer import pptx_group_of

    chain, current, seen = [], element, set()
    while True:
        group = pptx_group_of(current, by_id)
        if group is None or group in seen:
            return tuple(reversed(chain))
        seen.add(group)
        chain.append(group)
        current = by_id[group]


def _identity_problem(group) -> dict | None:
    """p:grpSp の座標変換が恒等（off = chOff、ext = chExt、回転・反転なし）でなければ、その観測値（pt）を返す。"""
    transform = group.find("p:grpSpPr/a:xfrm", NS)
    if transform is None:
        return {"field": "a:xfrm", "expected": "恒等の座標変換", "actual": None}
    observed = {
        "rot": int(transform.get("rot", "0")),
        "flipH": _flag(transform.get("flipH")),
        "flipV": _flag(transform.get("flipV")),
    }
    for name, keys in (("off", ("x", "y")), ("chOff", ("x", "y")), ("ext", ("cx", "cy")), ("chExt", ("cx", "cy"))):
        node = transform.find(f"a:{name}", NS)
        observed[name] = None if node is None else [int(node.get(key, 0)) / EMU_PER_PT for key in keys]
    identity = not (observed["rot"] or observed["flipH"] or observed["flipV"]) and all(
        observed[a] is not None
        and observed[b] is not None
        and all(abs(x - y) <= _IDENTITY_TOLERANCE_PT for x, y in zip(observed[a], observed[b]))
        for a, b in (("off", "chOff"), ("ext", "chExt"))
    )
    return None if identity else {"field": "a:xfrm", "expected": "恒等の座標変換", "actual": observed}


def _check_groups(report, slide, elements: list, by_id: dict, tree: _SlideShapes) -> None:
    """場面の pptx_group が真のまとまりが、恒等の座標変換の p:grpSp になり、子孫の図形が最も近い書き出すまとまりの中にあるか。

    描く子孫のないまとまりはグループにしない（書き出しと同じ）。書き出さないまとまりの名前の p:grpSp も不一致にする。
    """
    expected: dict[str, tuple] = {}
    for element in elements:
        if element.get("kind") == "group":
            continue
        chain = _group_chain(element, by_id)
        for index, group in enumerate(chain):
            expected.setdefault(group, chain[:index])
        found = tree.shapes.get(element["id"])
        if found is not None and found[1] != chain:
            report(
                "group_mismatch",
                "図形が、編集元のまとまりに対応する PowerPoint のグループの中にない",
                slide=slide,
                element=element,
                observations={"expected_groups": list(chain), "actual_groups": list(found[1])},
            )
    for group, parents in expected.items():
        found = tree.groups.get(group)
        if found is None:
            report(
                "group_mismatch",
                "書き出すまとまりに対応する PowerPoint のグループがない",
                slide=slide,
                element=by_id[group],
                observations={"group": group},
            )
            continue
        problems = []
        if found[1] != parents:
            problems.append({"field": "parent_groups", "expected": list(parents), "actual": list(found[1])})
        identity = _identity_problem(found[0])
        if identity is not None:
            problems.append(identity)
        for problem in problems:
            report(
                "group_mismatch",
                "PowerPoint のグループの位置か座標変換が編集元と一致しない",
                slide=slide,
                element=by_id[group],
                observations=problem,
            )
    for group in tree.groups:
        if group not in expected:
            element = by_id.get(group)
            report(
                "group_mismatch",
                "書き出さないまとまりか、場面にない名前の PowerPoint のグループがある",
                slide=slide,
                element=element if isinstance(element, dict) and element.get("kind") == "group" else None,
                observations={
                    "group": group,
                    "pptx_group": element.get("pptx_group") if isinstance(element, dict) else None,
                },
            )


def _connection(properties, key: str, names: dict) -> dict | None:
    node = None if properties is None else properties.find(f"a:{key}", NS)
    if node is None:
        return None
    try:
        site = int(node.get("idx", ""))
    except ValueError:
        site = node.get("idx")
    return {"element_id": names.get(node.get("id")), "shape_id": node.get("id"), "site": site}


def _check_connections(report, element: dict, node, by_id: dict, tree: _SlideShapes, project) -> None:
    """anchor() で結んだ線の端ごとに、書き出しと同じ判定（`writer.anchored_site`）で期待する接続と、PPTX の接続を比べる。

    接続しない端（番号がない端、anchor() で結んでいない端）には接続がないことを求める。接続先の図形がない端は
    `element_missing` が報告するので比べない。
    """
    from .writer import anchored_site

    properties = node.find("p:nvCxnSpPr/p:cNvCxnSpPr", NS)
    for end, key in (("start", "stCxn"), ("end", "endCxn")):
        reference = element.get(end + "_ref")
        target = by_id.get(reference.get("element_id")) if isinstance(reference, dict) else None
        site = anchored_site(reference, target, project)
        expected = None
        if site is not None:
            found = tree.shapes.get(target["id"])
            if found is None:
                continue
            name_node = found[0].find(".//p:cNvPr", NS)
            expected = {
                "element_id": target["id"],
                "shape_id": None if name_node is None else name_node.get("id"),
                "site": site,
            }
        actual = _connection(properties, key, tree.names)
        if (expected is None) != (actual is None) or (
            expected is not None and (expected["shape_id"], expected["site"]) != (actual["shape_id"], actual["site"])
        ):
            report(
                "connection_mismatch",
                "anchor() で結んだ線の端の接続が編集元と一致しない",
                observations={"end": end, "expected": expected, "actual": actual},
            )


def _check_line_style(report, node, element) -> None:
    """線・矢印と、枠線のある箱・文字枠の破線（a:prstDash）が、書式の dashed と一致するか。"""
    from .writer import element_style

    style = element_style(element)
    if element["kind"] not in {"line", "arrow"} and not style.get("line_color"):
        return
    dash = node.find("p:spPr/a:ln/a:prstDash", NS)
    custom = node.find("p:spPr/a:ln/a:custDash", NS)
    actual = "custom" if custom is not None else (dash.get("val", "solid") if dash is not None else "solid")
    expected = "dash" if style.get("dashed") else "solid"
    if actual != expected:
        report(
            "line_style_mismatch",
            "線か枠線の破線・実線が編集元の dashed と一致しない",
            observations={"expected": expected, "actual": actual},
        )


def _check_vector(report, node, element, relationships: dict, part: str, names: set) -> None:
    """SVG でも埋め込む画像（vector_path がある画像）の a:blip に asvg:svgBlip があり、その参照先の部品があるか。"""
    from .writer import SVG_BLIP_URI, SVG_NS

    blip = node.find(".//a:blip", NS)
    svg = None
    for extension in [] if blip is None else blip.findall("a:extLst/a:ext", NS):
        if extension.get("uri") == SVG_BLIP_URI:
            svg = extension.find(f"{{{SVG_NS}}}svgBlip")
    relationship = None if svg is None else relationships.get(svg.get(f"{{{NS['r']}}}embed"), {})
    target = _target(part, relationship["Target"]) if relationship and relationship.get("Target") else None
    if svg is None or target not in names:
        report(
            "image_vector_missing",
            "SVG でも埋め込む画像に、SVG の参照か部品がない",
            observations={"svg_blip": svg is not None, "target": target},
        )


def _check_runs(report, node, element, fonts) -> None:
    """通常文字の a:r を、書き出しの区切り（各 a:r の文字）と書式（文字サイズ・太字・下線・色・書体）と照合する。

    区切りの数が同じでも文字の境界がずれれば、ずれた文字はほかの区切りの書式で表示されるので、区切りごとに文字も比べる。
    forced の区切りは、a:latin と a:ea の両方が計測した書体（`measure.font_runs`。○ × なら和文の書体、→ — なら欧文の書体）
    であることを求める。明示改行（a:br）は数と文字サイズを比べる（空の行の大きさは a:br の文字サイズで決まる）。
    """
    from .writer import element_style, paragraph_content

    style = element_style(element)
    items = paragraph_content(element, fonts)
    breaks = [own for kind, _, own, _ in items if kind == "break"]
    actual_breaks = node.findall(".//a:br", NS)
    if len(actual_breaks) != len(breaks):
        report(
            "text_mismatch",
            "PPTX の明示改行の数が編集元と一致しない",
            observations={"expected_breaks": len(breaks), "actual_breaks": len(actual_breaks)},
        )
    else:
        for number, (line_break, own) in enumerate(zip(actual_breaks, breaks), 1):
            properties = line_break.find("a:rPr", NS)
            size = None if properties is None or properties.get("sz") is None else int(properties.get("sz")) / 100
            if own.get("size_pt") is not None and (size is None or abs(size - own["size_pt"]) > 0.01):
                report(
                    "font_size_mismatch",
                    "PPTX の明示改行の文字サイズが指定値と一致しない",
                    observations={"break": number, "expected_size_pt": own["size_pt"], "actual_size_pt": size},
                )
    expected = [(text, own, run) for kind, text, own, run in items if kind == "run"]
    actual = node.findall(".//a:r", NS)
    actual_texts = [run.findtext("a:t", default="", namespaces=NS) for run in actual]
    if len(actual) != len(expected):
        report(
            "font_mismatch",
            "通常文字の区切りが編集元と一致しない",
            observations={"expected_runs": [text for text, _, _ in expected], "actual_runs": actual_texts},
        )
        return
    for number, (run, actual_text, (text, own, font_run)) in enumerate(zip(actual, actual_texts, expected), 1):
        if actual_text != text:
            report(
                "font_mismatch",
                "通常文字の区切りの文字が編集元と一致しない",
                observations={"run": number, "field": "text", "expected": text, "actual": actual_text},
            )
        properties = run.find("a:rPr", NS)
        if properties is None:
            report("font_mismatch", "通常文字の書式が明示されていない", observations={"run": number, "text": text})
            continue
        expected_size = own.get("size_pt")
        actual_size = int(properties.get("sz", 0)) / 100
        if expected_size is not None and abs(actual_size - expected_size) > 0.01:
            report(
                "font_size_mismatch",
                "PPTX の文字サイズが指定値と一致しない",
                observations={
                    "run": number,
                    "text": text,
                    "expected_size_pt": expected_size,
                    "actual_size_pt": actual_size,
                },
            )
        forced = font_run.family if font_run.forced else None
        fields = [
            ("font_latin", forced or style.get("font_latin"), properties.find("a:latin", NS)),
            ("font_japanese", forced or style.get("font_japanese"), properties.find("a:ea", NS)),
        ]
        for field, family, font in fields:
            if family and (font is None or font.get("typeface") != family):
                report(
                    "font_mismatch",
                    "PPTX のフォントが指定値と一致しない",
                    observations={
                        "run": number,
                        "text": text,
                        "field": field,
                        "expected": family,
                        "actual": None if font is None else font.get("typeface"),
                    },
                )
        color = properties.find("a:solidFill/a:srgbClr", NS)
        values = [
            ("bold", bool(own.get("bold")), _flag(properties.get("b"))),
            ("underline", bool(own.get("underline")), properties.get("u", "none") != "none"),
            ("superscript", bool(own.get("superscript")), int(properties.get("baseline", "0") or 0) > 0),
        ]
        if own.get("color"):
            values.append(
                (
                    "color",
                    str(own["color"]).removeprefix("#").upper(),
                    None if color is None else str(color.get("val", "")).upper(),
                )
            )
        for field, value, found in values:
            if value != found:
                report(
                    "font_mismatch",
                    "PPTX の文字の書式が指定値と一致しない",
                    observations={"run": number, "text": text, "field": field, "expected": value, "actual": found},
                )


def _check_math(report, node, element, project, fonts) -> None:
    """数式の数・構造と書式・置き方（独立した数式か文中の数式か、揃え、中央揃えの補正）を照合する。"""
    from .math.office_math import math_element, math_mode
    from .writer import display_indent, element_style, math_arguments

    functions = project.config.get("math", {}).get("functions", ())
    style = element_style(element)
    fragments = [fragment for fragment in _content(element) if fragment.get("kind") == "math"]
    actual_math = node.findall(".//m:oMath", NS)
    if len(actual_math) != len(fragments):
        report(
            "editable_math_missing",
            "編集可能な数式本体の数が一致しない",
            observations={"expected": len(fragments), "actual": len(actual_math)},
        )
        return
    expected = []
    for fragment in fragments:
        try:
            expected.append(
                math_element(fragment["text"], **math_arguments(element, fragment_style(style, fragment), functions))
            )
        except (TypeError, ValueError) as exc:
            report(
                "math_content_mismatch",
                "編集可能な数式の期待値を作れない",
                observations={"expression": fragment["text"], "reason": str(exc)},
            )
            return
    for wrapper, actual in zip(expected, actual_math):
        if _xml_structure(next(wrapper.iter(f"{{{NS['m']}}}oMath"))) != _xml_structure(actual):
            report("math_content_mismatch", "編集可能な数式の構造または書式が編集元と一致しない")
    display = math_mode(element) == "display"
    problems = []
    wrappers = node.findall(".//a14:m", NS)
    for number, (wrapper, actual) in enumerate(zip(expected, wrappers), 1):
        expected_para = wrapper.find(f"{{{NS['m']}}}oMathPara")
        actual_para = actual.find("m:oMathPara", NS)
        if (expected_para is None) != (actual_para is None):
            problems.append(
                {
                    "math": number,
                    "expected": "display" if display else "inline",
                    "actual": "inline" if actual_para is None else "display",
                }
            )
        elif expected_para is not None:
            expected_jc = expected_para.find(f"{{{NS['m']}}}oMathParaPr/{{{NS['m']}}}jc")
            actual_jc = actual_para.find("m:oMathParaPr/m:jc", NS)
            values = [None if item is None else item.get(f"{{{NS['m']}}}val") for item in (expected_jc, actual_jc)]
            if values[0] != values[1]:
                problems.append({"math": number, "field": "m:jc", "expected": values[0], "actual": values[1]})
    for paragraph in node.findall(".//a:p", NS):
        if paragraph.find("a14:m", NS) is None:
            continue
        has_run = paragraph.find("a:r", NS) is not None
        if not display and not has_run:
            problems.append({"paragraph": "文中の数式だけの段落に通常文字の区切り（U+200B）がない"})
        if display:
            if has_run:
                problems.append({"paragraph": "独立した数式の段落に通常文字がある"})
            end_font = paragraph.find("a:endParaRPr/a:latin", NS)
            latin = style.get("font_latin")
            if latin and (end_font is None or end_font.get("typeface") != latin):
                problems.append(
                    {
                        "field": "a:endParaRPr/a:latin",
                        "expected": latin,
                        "actual": None if end_font is None else end_font.get("typeface"),
                    }
                )
            if fonts is not None:
                indent = display_indent(element, fonts)
                properties = paragraph.find("a:pPr", NS)
                actual_indent = int(properties.get("marL", "0")) / EMU_PER_PT if properties is not None else 0.0
                if abs(actual_indent - indent) > 0.01:
                    problems.append(
                        {"field": "a:pPr/@marL", "expected_pt": round(indent, 3), "actual_pt": round(actual_indent, 3)}
                    )
    for problem in problems:
        report("math_layout_mismatch", "数式の置き方（独立・文中）または揃えが編集元と一致しない", observations=problem)


def validate_pptx(
    path: Path, scene: dict, project: Any, *, fonts=None, fallback_expected: bool | None = None
) -> list[dict]:
    """ZIP 内の参照と、編集可能な本体・ノート・文書情報を照合する。

    fonts は書き出し（`writer.write_pptx`）に渡したものと同じ `measure.FontSource` を渡す。通常文字の書体の区切りと、
    中央揃えの独立した数式の左の余白を、それで求めて照合する。None なら書体で分けない区切り（断片と行ごと）を期待し、
    すべての区切りが要素の書体であることを求める。左の余白は照合しない。
    fallback_expected が None なら互換画像の要否を設定から決め、False なら互換画像の欠落（`math_fallback_missing`）を
    報告しない（下書き）。
    """
    issues: list[dict] = []
    numbers = {slide.get("id"): slide.get("number", index) for index, slide in enumerate(scene.get("slides", []), 1)}

    def report(code, message, *, slide=None, element=None, observations=None, severity=None):
        issues.append(
            make_diagnostic(
                code,
                message,
                scope="pptx",
                severity=severity,
                slide_id=slide.get("id") if slide else None,
                element_id=element.get("id") if element else None,
                observations=observations,
                slide_number=numbers.get(slide.get("id")) if slide else None,
                evidence={
                    "element": element,
                    "slide": slide.get("id") if slide else None,
                    "metadata": scene.get("metadata"),
                    "math": project.config.get("math", {}),
                },
            )
        )

    if fallback_expected is None:
        fallback_expected = fallback_required(scene, project)
    try:
        archive = ZipFile(path)
        with archive:
            names = set(archive.namelist())
            if len(names) != len(archive.namelist()):
                report("package_duplicate", "PPTX 内で同じファイル名が重複している")
            xml = {}
            for name in names:
                if name.endswith((".xml", ".rels")):
                    try:
                        xml[name] = ET.fromstring(archive.read(name))
                    except ET.ParseError as exc:
                        report(
                            "invalid_xml", "PPTX 内の XML が不正である", observations={"part": name, "reason": str(exc)}
                        )
            for name, root in xml.items():
                if name.endswith(".xml"):
                    part_relationships = _relationships(archive, name)
                    for node in root.iter():
                        for attribute, value in node.attrib.items():
                            if (
                                attribute in {f"{{{NS['r']}}}id", f"{{{NS['r']}}}embed", f"{{{NS['r']}}}link"}
                                and value not in part_relationships
                            ):
                                report(
                                    "relationship_missing",
                                    "PPTX 内の図形などが存在しない関係を参照している",
                                    observations={"part": name, "relationship_id": value},
                                )
                if not name.endswith(".rels"):
                    continue
                part = (
                    ""
                    if name == "_rels/.rels"
                    else posixpath.join(posixpath.dirname(posixpath.dirname(name)), posixpath.basename(name)[:-5])
                )
                for rel in root:
                    if rel.get("TargetMode") == "External":
                        continue
                    target = _target(part, rel.get("Target", ""))
                    if target not in names:
                        report(
                            "package_reference_missing",
                            "PPTX 内の参照先が存在しない",
                            observations={"part": part, "target": target},
                        )
            core = xml.get("docProps/core.xml")
            expected_metadata = project.config.get("metadata", scene.get("metadata", {}))
            for key, tag in (
                ("title", "dc:title"),
                ("author", "dc:creator"),
                ("last_modified_by", "cp:lastModifiedBy"),
                ("subject", "dc:subject"),
                ("keywords", "cp:keywords"),
                ("comments", "dc:description"),
                ("category", "cp:category"),
                ("content_status", "cp:contentStatus"),
                ("identifier", "dc:identifier"),
                ("language", "dc:language"),
                ("version", "cp:version"),
            ):
                actual = "" if core is None else core.findtext(tag, default="", namespaces=NS)
                expected = str(expected_metadata.get(key, ""))
                if actual != expected:
                    report(
                        "metadata_mismatch",
                        "文書情報が指定値と一致しない",
                        observations={"field": key, "expected": expected, "actual": actual},
                    )
            presentation = xml.get("ppt/presentation.xml")
            if presentation is None:
                report("presentation_missing", "PPTX に資料本体がない")
                return merge_diagnostics(issues)
            dimensions = presentation.find("p:sldSz", NS)
            if dimensions is None or any(
                abs(int(dimensions.get(key, 0)) / EMU_PER_PT - size) > 0.1
                for key, size in zip(("cx", "cy"), scene["size"])
            ):
                report(
                    "page_size_mismatch",
                    "PPTX のページ寸法が指定値と一致しない",
                    observations={
                        "expected": scene["size"],
                        "actual": None if dimensions is None else dimensions.attrib,
                    },
                )
            relationships = _relationships(archive, "ppt/presentation.xml")
            slide_parts = []
            for entry in presentation.findall("p:sldIdLst/p:sldId", NS):
                rel = relationships.get(entry.get(f"{{{NS['r']}}}id"), {})
                slide_parts.append(_target("ppt/presentation.xml", rel.get("Target", "")))
            if len(slide_parts) != len(scene["slides"]):
                report(
                    "slide_count_mismatch",
                    "PPTX のページ数が一致しない",
                    observations={"expected": len(scene["slides"]), "actual": len(slide_parts)},
                )
            for part, slide in zip(slide_parts, scene["slides"]):
                root = xml.get(part)
                if root is None:
                    report("slide_missing", "PPTX のページを読み込めない", slide=slide, observations={"part": part})
                    continue
                csld = root.find("p:cSld", NS)
                if csld is None or csld.get("name") != slide["id"]:
                    report(
                        "slide_order_mismatch",
                        "PPTX のページの識別子または順序が一致しない",
                        slide=slide,
                        observations={
                            "expected": slide["id"],
                            "actual": csld.get("name") if csld is not None else None,
                        },
                    )
                tree = _SlideShapes(root.find("p:cSld/p:spTree", NS))
                for name in tree.duplicates:
                    report(
                        "element_id_duplicate",
                        "PPTX 内の要素の識別子が重複している",
                        slide=slide,
                        observations={"name": name},
                    )
                elements = slide.get("elements", [])
                by_id = {element["id"]: element for element in elements}
                slide_relationships = _relationships(archive, part)
                _check_groups(report, slide, elements, by_id, tree)
                for element in elements:
                    if element["kind"] == "group":
                        continue
                    found = tree.shapes.get(element["id"])
                    if found is None:
                        report("element_missing", "PPTX 内で要素を参照できない", slide=slide, element=element)
                        continue
                    node, groups = found

                    def element_report(code, message, *, observations=None, slide=slide, element=element):
                        report(code, message, slide=slide, element=element, observations=observations)

                    transform = node.find(".//a:xfrm", NS)
                    expected_area = element.get("actual_area", element.get("area"))
                    if (
                        element["kind"] == "image"
                        and element.get("fit", "contain") == "contain"
                        and "actual_area" not in element
                    ):
                        expected_area = None
                    if expected_area and transform is not None:
                        offset, extent = transform.find("a:off", NS), transform.find("a:ext", NS)
                        if offset is not None and extent is not None:
                            actual_area = [
                                int(offset.get("x", 0)) / EMU_PER_PT,
                                int(offset.get("y", 0)) / EMU_PER_PT,
                                int(extent.get("cx", 0)) / EMU_PER_PT,
                                int(extent.get("cy", 0)) / EMU_PER_PT,
                            ]
                            if any(
                                abs(expected - actual) > 0.1 for expected, actual in zip(expected_area, actual_area)
                            ):
                                element_report(
                                    "geometry_mismatch",
                                    "PPTX の位置または寸法が指定値と一致しない",
                                    observations={"expected_area": expected_area, "actual_area": actual_area},
                                )
                    expected_text = _normalize_text(_normal_text(element))
                    actual_text = _normalize_text("".join(n.text or "" for n in node.findall(".//a:t", NS)))
                    if expected_text != actual_text:
                        element_report(
                            "text_mismatch",
                            "PPTX の通常文字が編集元と一致しない",
                            observations={"expected": expected_text, "actual": actual_text},
                        )
                    if _has_content(element):
                        from .writer import element_style

                        wrap = "square" if element_style(element)["word_wrap"] else "none"
                        for body in node.findall(".//a:bodyPr", NS):
                            if body.find("a:noAutofit", NS) is None:
                                element_report(
                                    "automatic_text_adjustment",
                                    "文字枠に a:noAutofit がなく、自動の縮小・調整が無効になっていない",
                                )
                            if body.get("wrap", "square") != wrap:
                                element_report(
                                    "text_wrap_mismatch",
                                    "折り返しの設定が編集元と一致しない",
                                    observations={"expected": wrap, "actual": body.get("wrap", "square")},
                                )
                        _check_runs(element_report, node, element, fonts)
                    if element["kind"] in {"text", "equation", "box", "line", "arrow"}:
                        _check_line_style(element_report, node, element)
                    if element["kind"] in {"line", "arrow"}:
                        _check_connections(element_report, element, node, by_id, tree, project)
                    if element["kind"] == "image" and element.get("vector_path"):
                        _check_vector(element_report, node, element, slide_relationships, part, names)
                    if _has_math(element):
                        _check_math(element_report, node, element, project, fonts)
                        if groups or node.find(".//p:grpSp", NS) is not None:
                            element_report(
                                "math_grouped",
                                "数式が PowerPoint のグループに入っている",
                                observations={"groups": list(groups)},
                            )
                        for xfrm in node.findall(".//a:xfrm", NS):
                            if int(xfrm.get("rot", "0")) != 0 or _flag(xfrm.get("flipH")) or _flag(xfrm.get("flipV")):
                                element_report("math_transform", "数式の文字枠が回転・反転している")
                        fallback = tree.fallbacks.get(element["id"])
                        if fallback_expected and (fallback is None or fallback.find(".//a:blip", NS) is None):
                            element_report("math_fallback_missing", "数式の互換画像がない")
                notes_parts = [
                    _target(part, rel["Target"])
                    for rel in slide_relationships.values()
                    if rel.get("Type", "").endswith("/notesSlide")
                ]
                actual_notes = ""
                for notes_part in notes_parts:
                    notes_root = xml.get(notes_part)
                    if notes_root is None:
                        continue
                    for shape in notes_root.findall(".//p:sp", NS):
                        placeholder = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
                        if placeholder is not None and placeholder.get("type") != "body":
                            continue
                        actual_notes += "".join(n.text or "" for n in shape.findall(".//a:t", NS))
                if _normalize_text(actual_notes) != _normalize_text(str(slide.get("notes", ""))):
                    report(
                        "notes_mismatch",
                        "発表者ノートが編集元と一致しない",
                        slide=slide,
                        observations={"expected": slide.get("notes", ""), "actual": actual_notes},
                    )
    except (BadZipFile, OSError, KeyError, ET.ParseError, ValueError) as exc:
        report("pptx_unreadable", "PPTX を検査できない", observations={"reason": str(exc)})
    return merge_diagnostics(issues)
