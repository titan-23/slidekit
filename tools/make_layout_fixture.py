"""PowerPoint で変換した校正用の資料の実行から、文字の配置の照合データを作る（開発用。パッケージに含めない）。

使い方:
  python tools/make_layout_fixture.py --run RUN_DIR --output PATH [--pdf PDF]

- 実行ディレクトリの deck.json と、標準表示の PDF（既定は renders/powerpoint/standard/deck.pdf）を pdf_layout.read_pdf で読む。
- 生成記録（manifest.json）から、PowerPoint の版、計測に使った書体の版・書体番号・sha256、パッケージの版を写す。パスは写さない。
- 文字を持つ要素（text・box・equation。テーマの要素を除く）ごとに、場面の要素と、実際の行数・基線・行の範囲・インク・
  文字サイズ・字形の書体を書く。PowerPoint は使わない。
- 照合データは公開してよい内容の試験の対象になる。ローカルのパスに見える文字列があれば、書かずに止める。

照合データは tests/test_layout_fixture.py が使う（環境変数 SLIDEKIT_LAYOUT_FIXTURE で場所を変更できる）。
校正用の資料は tests/data/layout/deck/ にある。実行ディレクトリ（ローカルのパスを含む）を slidekit の中に作らないように、
作業用の場所に校正用の資料を継承する deck.toml（schema_version = 1 と extends = "<slidekit>/tests/data/layout/deck/deck.toml"
の 2 行）を置いて生成する。作る手順（PowerPoint を使う生成は一度に一つだけ行う）:

  slidekit build --project WORK/deck.toml
  python tools/make_layout_fixture.py --run WORK/out/runs/NNNN --output tests/data/layout/powerpoint_layout.json

作り直す時期：PowerPoint・フォント・数式の計測の版を変えたとき、measure.MEASURE_VERSION を上げたとき。差分を確かめてから入れ替える。
"""

from __future__ import annotations

import argparse
from collections.abc import Mapping
from datetime import datetime, timezone
import json
from pathlib import Path
import re
import sys

from slidekit import measure, pdf_layout
from slidekit.files import write_json

SCHEMA_VERSION = 1
TAG_PREFIX = "tags:"
TEXT_KINDS = ("text", "box", "equation")
# 照合データへ写すパッケージの版（計測と PDF の読み取りに関わるもの。数式の計測は fontTools で行う）。
PACKAGES = ("pypdfium2", "pypdf", "fonttools", "python-pptx")
# ローカルのパスに見える文字列（照合データに入れない）。
LOCAL_PATH = re.compile(r"(?:^|[\s\"'(=])(?:/(?:home|mnt|tmp|Users|root|var|private|media)/|~/|[A-Za-z]:[\\/])")


class FixtureError(ValueError):
    """照合データを作成できない入力。"""


def _round(value: float) -> float:
    return round(float(value), 3)


def _label(value) -> str:
    """タグに使う値（空白を _ にする。数は余分な 0 を除く）。"""
    if isinstance(value, float):
        return f"{value:g}"
    return str(value).replace(" ", "_")


def _content(element: Mapping) -> list[Mapping]:
    value = element.get("content") or []
    return [fragment for fragment in value if isinstance(fragment, Mapping)] if isinstance(value, list) else []


def text_elements(slide: Mapping) -> list[Mapping]:
    """照合する要素：テーマの要素を除く text・box・equation で、空白以外の文字か数式を持つもの（measure.layout_scene と同じ）。"""
    result = []
    for element in slide.get("elements", []) or []:
        if element.get("kind") not in TEXT_KINDS or str(element.get("id", "")).startswith("_theme/"):
            continue
        if any(str(fragment.get("text", "")).strip() for fragment in _content(element)):
            result.append(element)
    return result


def case_tags(slide: Mapping, element: Mapping) -> list[str]:
    """照合の失敗を分けて見るためのタグ。ページの識別子、description の "tags:" に続く語、要素の書式から求めた語の順。"""
    tags = [str(slide.get("id"))]
    description = element.get("description")
    if isinstance(description, str) and description.startswith(TAG_PREFIX):
        tags += description[len(TAG_PREFIX) :].split()
    style = element.get("style") or {}
    kind = element.get("kind")
    tags.append(f"kind:{kind}")
    if kind == "box":
        tags.append(f"shape:{element.get('shape', 'rectangle')}")
    tags.append("wrap" if style.get("word_wrap") else "nowrap")
    tags += [f"align:{style.get('align') or 'left'}", f"valign:{style.get('valign') or 'top'}"]
    if style.get("size_pt") is not None:
        tags.append(f"size:{_label(float(style['size_pt']))}")
    if style.get("line_spacing_pt") is not None:
        tags.append(f"spacing_pt:{_label(float(style['line_spacing_pt']))}")
    elif style.get("line_spacing") is not None:
        tags.append(f"spacing:{_label(float(style['line_spacing']))}")
    for key, name in (("font_latin", "latin"), ("font_japanese", "japanese")):
        if style.get(key):
            tags.append(f"{name}:{_label(style[key])}")
    content = _content(element)
    if any(fragment.get("kind") == "math" for fragment in content):
        from slidekit.math.office_math import math_mode

        tags += ["math", math_mode(element)]
    if style.get("bold"):
        tags.append("bold")
    if any(fragment.get("kind") == "emphasis" for fragment in content):
        tags.append("emphasis")
    if any("\n" in str(fragment.get("text", "")) for fragment in content if fragment.get("kind") != "math"):
        tags.append("newline")
    return list(dict.fromkeys(tags))


def actual_record(rendered: pdf_layout.RenderedText) -> dict:
    """PDF から読んだ要素の値（照合データの actual）。"""
    return {
        "lines": len(rendered.lines),
        "baselines_pt": [_round(line.baseline) for line in rendered.lines],
        "line_extents_pt": [[_round(line.left), _round(line.right)] for line in rendered.lines],
        "ink_pt": None if rendered.ink is None else [_round(value) for value in rendered.ink],
        "sizes_pt": [_round(value) for value in rendered.sizes],
        "fonts": dict(sorted(rendered.fonts.items())),
        "glyphs": rendered.glyphs,
        "drawings": rendered.drawings,
    }


def _read_json(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise FixtureError(f"ファイルがありません: {path}") from None
    except json.JSONDecodeError as error:
        raise FixtureError(f"JSON として読めません: {path}: {error}") from None


def _inventory(manifest: Mapping) -> dict:
    """生成記録の環境の、書体ごとの記録（FontRegistry.inventory() の fonts）。"""
    environment = manifest.get("environment") or {}
    fonts = environment.get("fonts") or {}
    if isinstance(fonts.get("fonts"), Mapping):
        fonts = fonts["fonts"]
    return {family: info for family, info in fonts.items() if isinstance(info, Mapping) and info.get("sha256")}


def _families(cases: list[dict]) -> list[str]:
    """照合する要素が使う書体（欧文・和文と、数式を含む要素の数式の書体）。"""
    found = []
    for case in cases:
        element = case["element"]
        style = element.get("style") or {}
        keys = ["font_latin", "font_japanese"]
        if any(fragment.get("kind") == "math" for fragment in _content(element)):
            keys.append("font_math")
        found += [style[key] for key in keys if style.get(key)]
    return sorted(set(found))


def _font_records(manifest: Mapping, families: list[str]) -> dict:
    inventory = _inventory(manifest)
    missing = [family for family in families if family not in inventory]
    if missing:
        raise FixtureError(
            f"生成記録に書体の記録がありません: {'、'.join(missing)}（計測に使った書体の sha256 を照合データに入れられない）"
        )
    records = {}
    for family in families:
        info = inventory[family]
        record = {"version": info.get("version", ""), "index": int(info.get("index", 0)), "sha256": info["sha256"]}
        bold = info.get("bold")
        if isinstance(bold, Mapping) and bold.get("sha256"):
            record.update(
                bold_version=bold.get("version", ""), bold_index=int(bold.get("index", 0)), bold_sha256=bold["sha256"]
            )
        records[family] = record
    return records


def _powerpoint_version(manifest: Mapping) -> str | None:
    environment = manifest.get("environment") or {}
    detected = (environment.get("renderers") or {}).get("powerpoint") or {}
    if detected.get("version"):
        return str(detected["version"])
    for name, operation in (manifest.get("operations") or {}).items():
        if str(name).startswith("render:powerpoint") and isinstance(operation, Mapping) and operation.get("version"):
            build = operation.get("build")
            return f"{operation['version']}.{build}" if build else str(operation["version"])
    return None


def _packages(manifest: Mapping) -> dict:
    versions = (manifest.get("environment") or {}).get("packages") or {}
    return {name: str(versions[name]) for name in PACKAGES if versions.get(name)}


def make_fixture(run: Path, *, pdf: Path | None = None, generated_at: str | None = None) -> dict:
    """実行ディレクトリから照合データの辞書を作る。"""
    run = Path(run)
    scene = _read_json(run / "deck.json")
    manifest = _read_json(run / "manifest.json")
    pdf = Path(pdf) if pdf is not None else run / "renders" / "powerpoint" / "standard" / "deck.pdf"
    if not pdf.is_file():
        raise FixtureError(f"標準表示の PDF がありません: {pdf}")
    try:
        readings = pdf_layout.read_pdf(pdf, scene)
    except ValueError as error:
        raise FixtureError(str(error)) from None
    cases, slides, unread = [], [], []
    for slide, reading in zip(scene.get("slides", []), readings):
        slides.append({"id": slide.get("id"), "unassigned_glyphs": reading.unassigned})
        for element in text_elements(slide):
            rendered = reading.elements.get(str(element.get("id")))
            if rendered is None:
                unread.append(f"{slide.get('id')}/{element.get('id')}")
                continue
            cases.append(
                {
                    "slide_id": slide.get("id"),
                    "element_id": element.get("id"),
                    "tags": case_tags(slide, element),
                    "element": element,
                    "actual": actual_record(rendered),
                }
            )
    if not cases:
        raise FixtureError("照合する要素がありません（文字を持つ要素を PDF から読めない）")
    data = {
        "schema_version": SCHEMA_VERSION,
        "measure_version": measure.MEASURE_VERSION,
        "generated_at": generated_at or datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "environment": {
            "powerpoint": _powerpoint_version(manifest),
            "fonts": _font_records(manifest, _families(cases)),
            "packages": _packages(manifest),
        },
        "slides": slides,
        "unread": unread,
        "cases": cases,
    }
    check_public(data, extra=(run.resolve(), Path.home()))
    return data


def _strings(value):
    if isinstance(value, str):
        yield value
    elif isinstance(value, Mapping):
        for key, item in value.items():
            yield str(key)
            yield from _strings(item)
    elif isinstance(value, (list, tuple)):
        for item in value:
            yield from _strings(item)


def check_public(data, *, extra=()) -> None:
    """ローカルのパスに見える文字列があれば FixtureError にする（照合データは公開してよい内容の試験の対象になる）。"""
    fixed = [str(path) for path in extra if str(path) not in ("", "/", ".")]
    for value in _strings(data):
        if LOCAL_PATH.search(value) or any(path in value for path in fixed):
            raise FixtureError(f"照合データにローカルのパスに見える文字列があります: {value[:80]!r}")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="PowerPoint で変換した校正用の資料の実行から、文字の配置の照合データを作る（PowerPoint は使わない）"
    )
    parser.add_argument(
        "--run",
        type=Path,
        required=True,
        help="実行ディレクトリ（deck.json・manifest.json・標準表示の PDF がある場所）",
    )
    parser.add_argument("--output", type=Path, required=True, help="書き出す照合データの JSON")
    parser.add_argument(
        "--pdf",
        type=Path,
        default=None,
        help="読む PDF（既定: 実行ディレクトリの renders/powerpoint/standard/deck.pdf）",
    )
    arguments = parser.parse_args(argv)
    try:
        data = make_fixture(arguments.run, pdf=arguments.pdf)
    except FixtureError as error:
        print(f"照合データを作れません: {error}", file=sys.stderr)
        return 2
    write_json(arguments.output, data)
    unassigned = sum(slide["unassigned_glyphs"] for slide in data["slides"])
    empty = sum(1 for case in data["cases"] if case["actual"]["glyphs"] == 0)
    print(
        f"書いた照合データ: {arguments.output}（要素 {len(data['cases'])} 個、ページ {len(data['slides'])} 枚、"
        f"読めなかった要素 {len(data['unread'])} 個、字形のない要素 {empty} 個、割り当てなしの字形 {unassigned} 個）"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
