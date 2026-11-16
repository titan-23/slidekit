"""文書の試験：相対リンク、設定と書式の一覧、コマンドの一覧、手順の記述を実装と照らす。

表の行は、先頭の縦線の直後のセルが逆引用符で囲んだ名前である行（例：`| `size_pt` | … |`）として読む。
診断の一覧と数式の命令の一覧は、それぞれ `test_diagnostics.py`・`test_math_commands.py` が確かめる。
"""

from __future__ import annotations

import argparse
import ast
import inspect
import itertools
from pathlib import Path
import re
import tomllib
from types import SimpleNamespace
import unicodedata

from lxml import etree
import pytest

import slidekit
from slidekit import (
    Deck,
    Rect,
    cli,
    config,
    evidence,
    fonts,
    measure,
    parallel,
    pipeline,
    progress,
    records,
    report,
    schema,
    styles,
    tables,
    validators,
)
from slidekit.errors import InputError
from slidekit.ooxml import NS
from slidekit.renderers import powerpoint

ROOT = Path(__file__).resolve().parents[1]
# 相対リンクを確かめる文書（使用例の README は test_examples.py が扱う）。
DOCUMENTS = (
    "README.md",
    "AGENTS.md",
    "docs/usage.md",
    "docs/reference.md",
    "docs/math.md",
    "docs/diagnostics.md",
    "docs/issues.md",
    "workflows/authoring.md",
)
REFERENCE = ROOT / "docs" / "reference.md"
USAGE = ROOT / "docs" / "usage.md"
# 手順を書く文書（コードの例のコマンドの並びを確かめる）。
PROCEDURES = ("README.md", "docs/usage.md", "workflows/authoring.md")

_FENCE = re.compile(r"^(```|~~~)")
_INLINE_CODE = re.compile(r"`[^`]*`")
_LINK = re.compile(r"!?\[[^\]]*\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)")
_HEADING = re.compile(r"^(#{1,6})\s+(.*?)\s*#*\s*$")
_ROW_NAME = re.compile(r"^\|\s*`([^`]+)`\s*\|")
_CELL_SPLIT = re.compile(r"(?<!\\)\|")
_CODE_NAME = re.compile(r"`([^`]+)`")


# ---------------------------------------------------------------- 読み取りの補助


def _text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _prose_lines(text: str):
    """コードブロックの外の行（行番号と行）。"""
    fenced = False
    for number, line in enumerate(text.splitlines(), 1):
        if _FENCE.match(line.lstrip()):
            fenced = not fenced
            continue
        if not fenced:
            yield number, line


def _code_blocks(text: str) -> list[list[str]]:
    blocks, current = [], None
    for line in text.splitlines():
        if _FENCE.match(line.lstrip()):
            if current is None:
                current = []
            else:
                blocks.append(current)
                current = None
            continue
        if current is not None:
            current.append(line)
    return blocks


def _slug(heading: str) -> str:
    """GitHub と同じ規則の見出しのアンカー：小文字にし、文字・数字・結合文字・_・-・空白のほかを除き、空白を - にする。"""
    text = heading.replace("`", "").lower()
    kept = "".join(character for character in text if character in " -_" or unicodedata.category(character)[0] in "LMN")
    return kept.replace(" ", "-")


def _anchors(text: str) -> set[str]:
    anchors, seen = set(), {}
    for _, line in _prose_lines(text):
        match = _HEADING.match(line)
        if not match:
            continue
        slug = _slug(match.group(2))
        count = seen.get(slug, 0)
        seen[slug] = count + 1
        anchors.add(slug if count == 0 else f"{slug}-{count}")
    return anchors


def _section(text: str, title: str) -> str:
    """見出し title の節の本文（同じか上の段の次の見出しの前まで）。コードブロックの中の行（`# ` で始まる注釈など）は見出しにしない。"""
    lines = text.splitlines()
    headings = {number - 1: match for number, line in _prose_lines(text) if (match := _HEADING.match(line))}
    for index, match in headings.items():
        if match.group(2) == title:
            level = len(match.group(1))
            end = next(
                (later for later, found in headings.items() if later > index and len(found.group(1)) <= level),
                len(lines),
            )
            return "\n".join(lines[index + 1 : end])
    raise AssertionError(f"見出し「{title}」がない")


def _rows(section: str) -> dict[str, list[str]]:
    """表の行を、先頭のセルの名前からセルの列へ。同じ名前の行が 2 つあれば誤りにする。"""
    rows: dict[str, list[str]] = {}
    for line in section.splitlines():
        match = _ROW_NAME.match(line)
        if not match:
            continue
        name = match.group(1)
        assert name not in rows, f"表に同じ名前の行が 2 つある: {name}"
        rows[name] = [cell.strip() for cell in _CELL_SPLIT.split(line.strip().strip("|"))]
    return rows


def _names(cell: str) -> set[str]:
    return set(_CODE_NAME.findall(cell))


def _toml_literal(value, *, quote: bool = True) -> str:
    """文書の既定値の書き方（TOML の値）。書式の表の値（色の名前や揃え）は quote を偽にして引用符を付けない。"""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, str):
        return f'"{value}"' if quote else value
    if isinstance(value, list):
        return "[" + ", ".join(_toml_literal(item) for item in value) + "]"
    return str(value)


def _compare(documented, expected, what: str) -> None:
    documented, expected = set(documented), set(expected)
    missing, extra = sorted(expected - documented), sorted(documented - expected)
    assert not missing and not extra, f"{what}が実装と一致しない。文書にない: {missing}、実装にない: {extra}"


# ---------------------------------------------------------------- リンク


@pytest.mark.parametrize("name", DOCUMENTS)
def test_relative_links_resolve(name):
    path = ROOT / name
    text = _text(path)
    problems = []
    for number, line in _prose_lines(text):
        for target in _LINK.findall(_INLINE_CODE.sub("", line)):
            if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*:", target):
                continue  # 外部のリンク（https:、mailto:）
            file, _, anchor = target.partition("#")
            destination = (path.parent / file).resolve() if file else path
            if not destination.exists():
                problems.append(
                    f"{number} 行目: {target}（{destination.relative_to(ROOT) if destination.is_relative_to(ROOT) else destination} がない）"
                )
                continue
            if anchor:
                if destination.is_dir() or destination.suffix != ".md":
                    problems.append(f"{number} 行目: {target}（Markdown でない場所に見出しを指定している）")
                elif anchor not in _anchors(_text(destination)):
                    problems.append(f"{number} 行目: {target}（見出し #{anchor} がない）")
    assert not problems, f"{name} の相対リンクをたどれない:\n" + "\n".join(problems)


def test_slug_follows_github_rules():
    assert _slug("0.2 への移行") == "02-への移行"
    assert _slug("資料の設定（deck.toml）") == "資料の設定decktoml"
    assert _slug("`slidekit layout` の出力") == "slidekit-layout-の出力"
    assert _anchors("# 題\n## 例\n## 例\n```\n## 例\n```\n") == {"題", "例", "例-1"}


def test_section_ignores_headings_in_code_blocks():
    text = "# 題\n## 節\n本文\n```toml\n# 注釈\n## 注釈\n```\n後\n## 次\n"
    assert _section(text, "節") == "本文\n```toml\n# 注釈\n## 注釈\n```\n後"
    assert _section(text, "次") == ""
    with pytest.raises(AssertionError):
        _section(text, "注釈")
    # 利用ガイドの節は例の中の注釈（# theme.toml：…）で切れず、次の節の見出しの前まで続く。
    section = _section(_text(USAGE), "書式と優先順位")
    assert "# theme.toml" in section and "[書式の引数](reference.md#書式の引数)" in section
    assert "## 文字の配置と折り返し" not in section


# ---------------------------------------------------------------- 書式と属性


def test_style_keys_are_documented():
    section = _section(_text(REFERENCE), "書式の引数")
    rows = _rows(section)
    expected = {
        name: {kind for kind, keys in styles.STYLE_KEYS.items() if name in keys} for name in styles.ALL_STYLE_KEYS
    }
    _compare(rows, expected, "書式の引数の表の名前")
    for name, kinds in expected.items():
        assert (
            _names(rows[name][1]) == kinds
        ), f"書式 {name} の使える要素が実装と違う: {sorted(_names(rows[name][1]))} / {sorted(kinds)}"
    # ライブラリの既定値（None でないもの）は「指定がないとき」の列にそのまま書く。
    for name, value in styles.LIBRARY_DEFAULTS.items():
        if value is not None:
            assert (
                f"`{_toml_literal(value, quote=False)}`" in rows[name][4]
            ), f"書式 {name} の既定値 {value!r} が文書にない"
    # しっぽの寸法と位置の既定値は LIBRARY_DEFAULTS でなく outline の定数で、同じ列に書く。
    from slidekit import outline

    for name, value in (
        ("tail_at", outline.TAIL_AT),
        ("tail_width_pt", outline.TAIL_WIDTH_PT),
        ("tail_length_pt", outline.TAIL_LENGTH_PT),
    ):
        assert f"`{_toml_literal(value)}`" in rows[name][4], f"書式 {name} の既定値 {value!r} が文書にない"
    # 種類ごとの既定の表、受け付けない名前、書式を持たない種類を説明する。
    for name in styles.DEFAULT_KINDS | set(styles.UNSUPPORTED):
        assert f"`{name}`" in section, f"書式の引数の節に {name} の説明がない"
    assert not any(styles.STYLE_KEYS[kind] for kind in ("image", "group")), "画像とまとまりの書式の説明を直す"


def test_creators_document_style_arguments(tmp_path, write_deck):
    # 要素を作る関数の表で、書式の引数を受け付ける関数だけに「書式」を書く。表の書式の引数はすべてのセルに適用される。
    reference = _text(REFERENCE)
    rows = {}  # 先頭のセルの関数名（`slide.line`・`slide.arrow` のように複数ある行も）からセルの列へ
    for line in _section(reference, "要素を作る関数").splitlines():
        if line.startswith("| `"):
            cells = [cell.strip() for cell in _CELL_SPLIT.split(line.strip().strip("|"))]
            rows.update(dict.fromkeys(_names(cells[0]), cells))
    kinds = {
        "slide.text": "text",
        "slide.box": "box",
        "slide.equation": "equation",
        "slide.line": "line",
        "slide.arrow": "arrow",
        "slide.image": "image",
        "slide.group": "group",
        "slide.table": "table_cell",
    }
    for name, kind in kinds.items():
        assert ("書式" in rows[name][2]) == bool(styles.STYLE_KEYS[kind]), f"{name} の書式の引数の有無が実装と違う"
    assert "表の書式の引数" in _section(
        reference, "書式の引数"
    ), "書式の引数の節に、表の書式の引数がセルに効くことを書く"
    deck = Deck(config.load_project(write_deck(tmp_path / "deck")))
    table = deck.slide(id="page").table(
        id="table",
        headers=["名前", "値"],
        rows=[["件数", "3"]],
        row_keys=["count"],
        column_keys=["name", "value"],
        area=Rect(40, 100, 400, 120),
        bold=True,
        size_pt=24,
    )
    for row, column in itertools.product(("_header", "count"), ("name", "value")):
        style = table.cell(row, column).data["style"]
        assert (style["bold"], style["size_pt"]) == (True, 24), f"表の書式の引数がセル {row}/{column} に効いていない"


def test_table_parameters_are_documented():
    section = _section(_text(USAGE), "表")
    parameters = set(inspect.signature(tables.add_table).parameters) - {"slide", "kwargs"}
    missing = sorted(name for name in parameters if not re.search(rf"(`{name}\b|\b{name}=)", section))
    assert not missing, f"表の引数が docs/usage.md の「表」にない: {missing}"


def test_geometry_section_covers_sizes_from_content():
    section = _section(_text(USAGE), "寸法と配置を決める")
    for form in ("`Rect(x, y, w, None)`", "`Rect(x, y, None, h)`", "`Rect(x, y, None, None)`", "`invalid_geometry`"):
        assert form in section, f"寸法と配置を決めるの節に {form} の説明がない"


def test_line_placement_rules_match_measure():
    section = _section(_text(USAGE), "行の配置")
    model = measure.MODEL
    for phrase in (
        f"今は {measure.MEASURE_VERSION}",
        f"1 より大きければ sp × min(A, {model.line_top_limit_em}) × s",
        f"1 未満なら max(A − {model.line_pitch_em} × (1 − sp), sp × min(A, {model.line_top_limit_em})) × s",
        f"インクの上 + {model.math_stack_gap_em} em",
        f"{model.math_stack_full_em} em 以上",
        f"{model.math_stack_above_full} × d 以下",
        f"{model.math_stack_top_ratio} × sp × インクの上",
        f"{model.math_radical_single_ratio} × インクの上",
        f"{model.math_radical_top_ratio} × sp × インクの上",
        f"{model.math_descent_above} × sp × (d − D × s)",
    ):
        assert phrase in section, f"行の配置の節に「{phrase}」がない"
    assert model.math_stack_above_share == 0.75 and "4 分の 3" in section
    # 文書に書いた行の上側の式（行間がちょうど 1、1 より大きい、1 未満）を実装と照らす。
    limit, pitch = model.line_top_limit_em, model.line_pitch_em
    for ascent, spacing in itertools.product((0.85, limit, 1.0, 1.1), (0.8, 0.9, 1.0, 1.2, 2.0)):
        if spacing == 1.0:
            expected = ascent
        elif spacing > 1.0:
            expected = spacing * min(ascent, limit)
        else:
            expected = max(ascent - pitch * (1 - spacing), spacing * min(ascent, limit))
        assert measure._line_top(ascent, spacing, model) == pytest.approx(expected), (ascent, spacing)


def test_radical_rule_matches_measure():
    # 文中の背の高い根号の規則：当てはまる式（根号の外に上へ出る部分がない）と、行の上側が行間 1.0 の上側 L との大きいほうであること。
    section = _section(_text(USAGE), "行の配置")
    model = measure.MODEL
    single, top = model.math_radical_single_ratio, model.math_radical_top_ratio
    for phrase in (
        "根号の外に上へ出る部分（上付きの添字、上に置くもの、アクセント、上線）がなく",
        f"T1 = max(L, {single} × インクの上)",
        f"max(sp × L, {top} × sp × インクの上)",
        "sp × (T1 + d) − 上側",
        "行間 1 未満の行には当てない",
        "どちらも上側は sp × L より小さくしない",
    ):
        assert phrase in section, f"行の配置の節に「{phrase}」がない"
    assert measure._structure("\\sqrt{\\sqrt{x}}") == "radical"
    assert measure._structure("\\sqrt{x^{2}}") == "radical", "根号の中の添字は数えない"
    for raised in ("\\sqrt{x}^{2}", "\\overset{n}{\\sqrt{x}}", "\\hat{\\sqrt{x}}", "\\overline{\\sqrt{x}}"):
        assert measure._structure(raised) is None, raised
    assert measure._structure("\\frac{1}{\\sqrt{x}}") == "stacked"
    size = 20.0
    for line_ascent, ascent, descent in ((26.0, 18.0, 6.0), (18.0, 20.0, 6.0)):
        box = measure.MathBox(30.0, ascent, descent, line_ascent, 7.0)
        unit = measure._Unit(
            "\\sqrt{x}", True, box.width, 0.0, size, False, False, False, (ascent, descent), box, radical=True
        )
        first = max(line_ascent, single * ascent)
        for spacing in (1.0, 1.2, 1.5):
            above = first if spacing == 1.0 else max(spacing * line_ascent, top * spacing * ascent)
            expected = (above, spacing * (first + descent) - above, True)
            assert measure._math_line(unit, spacing, model) == pytest.approx(expected), (line_ascent, ascent, spacing)
        # 行間 1 未満では根号の規則を当てず、行間 1.0 の上下に行間を掛ける。
        assert measure._math_line(unit, 0.9, model) == pytest.approx((0.9 * line_ascent, 0.9 * 7.0, False))


def test_attribute_keys_are_documented():
    section = _section(_text(REFERENCE), "要素の属性")
    _compare(_rows(section), styles.ATTRIBUTE_KEYS, "要素の属性の表の名前")
    for name in styles.GROUP_ATTRIBUTE_KEYS | styles.SEMANTIC_VALUES:
        assert f"`{name}`" in section, f"要素の属性の節に {name} がない"


def test_public_api_is_documented():
    text = _text(REFERENCE)
    for name in set(slidekit.__all__) - {"__version__"}:
        assert f"`{name}`" in text, f"公開の名前 {name} が docs/reference.md にない"
    assert "`parse`" in text


# ---------------------------------------------------------------- 設定


def test_theme_and_rule_keys_are_documented():
    text = _text(REFERENCE)
    bundled_theme = tomllib.loads((config.PACKAGE / "themes" / "default.toml").read_text(encoding="utf-8"))

    # テーマの表
    rows = _rows(_section(text, "テーマの表"))
    _compare({name.split(".")[0] for name in rows}, schema.THEME_KEYS, "テーマの表の項目")
    for table, keys in (("fonts", schema.FONT_KEYS), ("margins", schema.MARGIN_KEYS)):
        _compare(
            {name.split(".", 1)[1] for name in rows if name.startswith(table + ".")}, keys, f"テーマの {table} の項目"
        )
        for key in keys:
            literal = _toml_literal(bundled_theme[table][key])
            assert (
                f"`{literal}`" in rows[f"{table}.{key}"][3]
            ), f"テーマの {table}.{key} の既定値 {literal} が文書にない"
    defaults = next(cells for name, cells in rows.items() if name.startswith("defaults.") and "種類" in name)
    assert _names(defaults[1]) >= styles.DEFAULT_KINDS, "defaults.<種類> の行に種類の名前がそろっていない"
    page_kinds = next(cells for name, cells in rows.items() if name.startswith("page_kinds."))
    assert _names(page_kinds[1]) >= set(schema.PAGE_KINDS), "page_kinds.<種別> の行に種別の名前がそろっていない"
    roles = next(cells for name, cells in rows.items() if name.startswith("roles."))
    assert _names(roles[3]) == set(bundled_theme["roles"]), "同梱の役割の一覧が同梱のテーマと違う"

    # 役割・ページ種別・飾りの項目
    _compare(_rows(_section(text, "役割の項目")), styles.ROLE_ONLY_KEYS | {"size_pt", "min_size_pt"}, "役割の項目")
    _compare(_rows(_section(text, "ページ種別の項目")), schema.PAGE_KIND_KEYS, "ページ種別の項目")
    decorations = _section(text, "飾りの項目")
    every = set().union(*schema._DECORATION_KEYS.values())
    _compare(_rows(decorations), every - styles.ALL_STYLE_KEYS, "飾りの項目")
    for name in (every & styles.ALL_STYLE_KEYS) | set(schema.DECORATION_KINDS):
        assert f"`{name}`" in decorations, f"飾りの項目の節に {name} がない"

    # 記述規則（既定値は同梱の general.toml の値）
    bundled_rules = tomllib.loads((config.PACKAGE / "rules" / "general.toml").read_text(encoding="utf-8"))
    rows = _rows(_section(text, "記述規則"))
    expected = {name for name in schema.RULES_KEYS if name != "math_candidates"}
    expected |= {f"math_candidates.{name}" for name in schema.MATH_CANDIDATE_KEYS}
    _compare(rows, expected, "記述規則の項目")
    for name in expected:
        table, _, key = name.rpartition(".")
        value = bundled_rules[table][key] if table else bundled_rules[key]
        assert rows[name][1] == f"`{_toml_literal(value)}`", f"記述規則 {name} の既定値が同梱の値 {value!r} と違う"

    # deck.toml
    rows = _rows(_section(text, "資料の設定（deck.toml）"))
    expected = set()
    for name in schema.DECK_KEYS[""]:
        expected |= {f"{name}.{key}" for key in schema.DECK_KEYS[name]} if name in schema.DECK_KEYS else {name}
    _compare(rows, expected, "deck.toml の設定")
    assert _names(rows["inputs.documents"][1]) >= set(schema.DECK_KEYS["inputs.documents"])
    for name in schema.REQUIRED_METADATA:
        assert rows[f"metadata.{name}"][2] == "必須", f"metadata.{name} を必須と書く"
    for name in schema.DECK_LOCAL_KEYS:
        assert f"`{name}`" in rows["extends"][3], f"extends の行に、引き継がない {name} を書く"
    assert schema.PAGE_SIZES[0] in rows["size"][1] and schema.PAGE_SIZES[1] in rows["size"][1]
    assert f"継承は {config.MAX_EXTENDS} 段まで" in _section(_text(USAGE), "設定の継承"), "継承の段数の上限が実装と違う"

    # 実行環境（既定値は読み込みの既定と PowerPoint の待ち時間）
    rows = _rows(_section(text, "実行環境（environment.toml）"))
    expected = set()
    for name in schema.ENVIRONMENT_KEYS[""]:
        if name == "fonts":
            expected |= {f"fonts.<書体>.{key}" for key in schema.ENVIRONMENT_KEYS["fonts.*"]}
        elif name == "powerpoint":
            expected |= {f"powerpoint.{key}" for key in schema.ENVIRONMENT_KEYS["powerpoint"]}
        else:
            expected.add(name)
    _compare(rows, expected, "実行環境の設定")
    environment, problems = config._environment_settings([])
    assert not problems
    for key, value in environment["powerpoint"].items():
        assert f"`{_toml_literal(value)}`" in rows[f"powerpoint.{key}"][2], f"powerpoint.{key} の既定値が違う"
    assert f"`{powerpoint.DEFAULT_WAIT_SECONDS}`" in rows["powerpoint.wait_seconds"][2]
    from slidekit.fonts import WINDOWS_FONT_FILES

    environment_section = _section(text, "実行環境（environment.toml）")
    for family in WINDOWS_FONT_FILES:
        assert family in environment_section, f"既定の書体の候補 {family} が文書にない"


def test_deck_defaults_are_documented(tmp_path):
    # deck.toml の表の「既定値」の列（`…` で書いた値）を、最小の deck.toml を読み込んだ値と照らす。
    from pptx import Presentation

    Presentation().save(tmp_path / "template.pptx")
    (tmp_path / "slides.py").write_text("def build(context):\n    pass\n", encoding="utf-8")
    (tmp_path / "deck.toml").write_text(
        '[metadata]\ntitle = "試験"\nauthor = ""\nlast_modified_by = ""\n\n' '[template]\npath = "template.pptx"\n',
        encoding="utf-8",
    )
    loaded = config.load_project(tmp_path / "deck.toml").config
    root = tmp_path.resolve()

    def relative(value):
        return Path(value).relative_to(root).as_posix()

    checked = set()
    for name, cells in _rows(_section(_text(REFERENCE), "資料の設定（deck.toml）")).items():
        literal = re.fullmatch(r"`([^`]+)`", cells[2])
        if literal is None:
            assert cells[2] in ("なし", "必須"), f"deck.toml の {name} の既定値の書き方を読めない: {cells[2]}"
            continue
        documented = tomllib.loads(f"value = {literal.group(1)}")["value"]
        table, _, key = name.rpartition(".")
        actual = loaded[table][key] if table else loaded[key]
        if name == "size":
            documented = config._size(documented)
        elif name == "output_dir" or table == "deliverables":
            actual = relative(actual)
        elif name == "inputs.read_roots":
            actual = [relative(value) for value in actual]
        elif name == "inputs.python":
            actual = [value for value in map(relative, actual) if value != "slides.py"]  # 入口は自動で追跡する
        assert actual == documented, f"deck.toml の {name} の既定値 {documented!r} が読み込みの既定 {actual!r} と違う"
        checked.add(name)
    assert {
        "schema_version",
        "entrypoint",
        "size",
        "theme",
        "rules",
        "output_dir",
        "inputs.read_roots",
        "math.fallback",
        "math.fallback_scale",
        "images.raster_scale",
        "deliverables.pptx",
        "template.layout_index",
    } <= checked


# ---------------------------------------------------------------- コマンド


def _leaf_commands(parser: argparse.ArgumentParser, prefix: tuple[str, ...] = ()):
    """（コマンドの名前の列、引数の解析器）を、操作を選ぶ最後のコマンドごとに返す。"""
    found = [action for action in parser._actions if isinstance(action, argparse._SubParsersAction)]
    if not found:
        yield prefix, parser
        return
    for name, child in found[0].choices.items():
        yield from _leaf_commands(child, prefix + (name,))


def test_commands_are_documented():
    text = _text(USAGE)
    rows = _rows(_section(text, "コマンドの一覧"))
    leaves = {"slidekit " + " ".join(names): child for names, child in _leaf_commands(cli.parser())}
    _compare(rows, leaves, "コマンドの一覧の行")
    common = {"--json", "--quiet", "--verbose", "-h", "--help"}
    for command, child in leaves.items():
        options = {option for action in child._actions for option in action.option_strings} - common
        row = "|".join(rows[command])
        missing = sorted(option for option in options if not re.search(re.escape(option) + r"(?![\w-])", row))
        assert not missing, f"{command} の引数が docs/usage.md のコマンドの一覧にない: {missing}"
        for action in child._actions:
            if action.option_strings and action.choices and not set(action.option_strings) & common:
                assert all(
                    str(choice) in row for choice in action.choices
                ), f"{command} の {action.option_strings[0]} の値の一覧がない"
    for option in ("--json", "--quiet", "--verbose", "--version"):
        assert re.search(re.escape(option) + r"(?![\w-])", text), f"{option} の説明が docs/usage.md にない"


def test_procedures_do_not_run_doctor_before_every_build():
    # doctor は初回と実行環境を変えたときに使う。手順の例で、毎回 doctor の後に build を続けない。
    for name in PROCEDURES:
        for block in _code_blocks(_text(ROOT / name)):
            commands = [line.split() for line in block if line.startswith("slidekit ")]
            for first, second in zip(commands, commands[1:]):
                assert not (
                    first[1] == "doctor" and second[1] == "build"
                ), f"{name} の手順で doctor の直後に build を実行している"
    workflow = _section(_text(ROOT / "workflows" / "authoring.md"), "生成し、機械検査する")
    first = next(line for block in _code_blocks(workflow) for line in block if line.startswith("slidekit "))
    assert first.startswith("slidekit layout"), "作成の手順は配置だけの検査から始める"
    assert "`slidekit doctor`" in workflow


def test_parallel_authoring_section_names_selection_api():
    section = _section(_text(ROOT / "workflows" / "authoring.md"), "複数人で並行して書く")
    for phrase in ("context.selected(", "--slides", "slidekit layout", "変換待ち", "--wait"):
        assert phrase in section, f"「複数人で並行して書く」に {phrase} の説明がない"
    usage = _text(USAGE)
    assert "`context.selection`" in usage and "`context.selected(*識別子)`" in usage
    # ページ画像の名前で %XX にする文字は、Windows のファイル名で禁止されている文字と %（pipeline._image_name）。
    escaped = {character for character in map(chr, range(32, 127)) if pipeline._UNSAFE_NAME.fullmatch(character)}
    assert escaped == set('<>:"\\|?*%')
    documented = all(f"`{character}`" in section for character in escaped) and "`%XX`" in section
    assert documented and f"`{pipeline._image_name('a:b')}`" in section, "ページ画像の名前の %XX の規則がない"


def test_review_placeholders_are_documented():
    section = _section(_text(USAGE), "確認と完成物の保存")
    for placeholder in (records.REVIEWER_PLACEHOLDER, records.NOTE_PLACEHOLDER, records.REASON_PLACEHOLDER):
        assert f"`{placeholder}`" in section, f"確認の置き場所 {placeholder} の説明がない"


def test_script_section_names_start_methods():
    # 作業プロセスの起動方法は slidekit が選ぶ（Windows 以外は forkserver、Windows は spawn）。Python の版によらない。
    section = _section(_text(USAGE), "Python のスクリプトから呼ぶとき")
    assert 'if __name__ == "__main__":' in section
    named = parallel._start_method() in section and "forkserver" in section and "spawn" in section
    assert (
        named and "3.14" not in section
    ), "作業プロセスの起動方法を、slidekit が選ぶもの（Python の版によらない）として書く"


def test_migration_section_lists_changes():
    section = _section(_text(USAGE), "0.2 への移行")
    header = next((line for line in section.splitlines() if line.startswith("| 変更点 |")), None)
    assert header is not None and [cell.strip() for cell in header.strip("|").split("|")] == ["変更点", "影響", "対処"]
    rows = [
        line for line in section.splitlines() if line.startswith("| ") and not line.startswith(("| 変更点", "|---"))
    ]
    assert len(rows) >= 15
    # 既存の資料と編集元に影響する主な変更を挙げる。
    for phrase in (
        "word_wrap",
        "context.measure_text()",
        "context.measure_math()",
        "\\left(",
        "--json",
        "--carry-from none",
        "Python 3.12",
        "add_legacy_elements",
        ".slidekit-export.lock",
    ):
        assert phrase in section, f"0.2 への移行の表に {phrase} の変化がない"


def test_machine_checks_explain_how_to_avoid_overlap_warnings():
    # 重なりの警告が出ない場合と、意図した重なりの扱いを利用ガイドに書く。
    section = _section(_text(USAGE), "機械検査")
    for phrase in (
        "`element_overlap`",
        "`anchor()`",
        'semantic="background"',
        'semantic="decoration"',
        "同じまとまりの中",
        "警告の許容",
    ):
        assert phrase in section, f"機械検査の節に {phrase} の説明がない"
    allow = _section(_text(USAGE), "警告の許容")
    for phrase in ("slide.allow_warning(", "review accept", "allow_roles", "evidence_hash"):
        assert phrase in allow, f"警告の許容の節に {phrase} の説明がない"


# ---------------------------------------------------------------- 実装との照合


def test_input_errors_explain_collect_errors(tmp_path, write_deck):
    # 続行できる誤りを集めるのは context.collect_errors が真のとき（layout・doctor・build の構築）だけで、
    # Python から直接 Deck を使うと最初の誤りで InputError になる。
    section = _section(_text(USAGE), "基本 API")
    for phrase in ("`context.collect_errors`（既定は偽）", "最初の誤りで `InputError` になる"):
        assert phrase in section, f"基本 API の節に「{phrase}」がない"
    project = config.load_project(write_deck(tmp_path / "deck"))
    assert project.collect_errors is False
    with pytest.raises(InputError):
        Deck(project).slide(id="page").text(id="first", area=Rect(40, 40, 200, 40), text="a", role="unknown")
    project.collect_errors = True
    deck = Deck(project)
    page = deck.slide(id="page")
    page.text(id="first", area=Rect(40, 40, 200, 40), text="a", role="unknown")
    page.text(id="second", area=Rect(40, 100, 200, 40), text="b", role="unknown")
    assert len(deck.issues) == 2


def test_output_example_matches_summary():
    # コマンドの出力の例は layout の要約と同じ行の並び（所要時間の行を含む）にする。数値は「…」で書く。
    example = _code_blocks(_section(_text(USAGE), "コマンドの出力"))[0]
    timings = [
        {"stage": "layout", "label": "配置の検査", "seconds": 1.5},
        {"stage": "construct", "label": "構築", "seconds": 0.1},
    ]
    result = {
        "command": "layout",
        "status": "passed",
        "project": "examples/proposal/src/deck.toml",
        "slides": 4,
        "fallback_required": True,
        "counts": {"error": 0, "warning": 0, "info": 0},
        "issues": [],
        "timings": timings,
        "elapsed_seconds": 2.0,
        "exit_code": 0,
    }
    actual = report.summarize("layout", result).splitlines()
    assert len(example) == len(actual), f"出力の例の行数が要約と違う: {actual}"
    for documented, line in zip(example, actual):
        pattern = re.escape(documented).replace("…", r"[0-9.]+")
        assert re.fullmatch(pattern, line), f"出力の例「{documented}」が要約「{line}」と合わない"


_STAGE_CALL = re.compile(r'(?:stage|step|checked)\(\s*f?"([^"]+)"')
_STAGE_PREFIX = re.compile(r'f"(\w+):[\w:]*\{')


def _stage_names(source_object) -> set[str]:
    """関数かクラスが使う段階の名前。`:` を含む名前（render・pdf・close）は先頭の部分にする。"""
    source = inspect.getsource(source_object)
    names = {name.split(":")[0] for name in _STAGE_CALL.findall(source)}
    return names | set(_STAGE_PREFIX.findall(source))


def test_stage_table_matches_pipeline():
    section = _section(_text(USAGE), "生成の流れと所要時間")
    rows = _rows(section)
    inspected = _stage_names(pipeline._inspect)
    used = {
        "layout": inspected,
        "doctor": inspected | _stage_names(pipeline.doctor),
        "build": _stage_names(pipeline._Build),
    }
    stages = {name.split(":")[0]: name for name in rows}
    _compare(stages, set().union(*used.values()), "段階の表の名前")
    _compare(stages, set(progress.STAGE_LABELS) | {"render", "pdf", "close"}, "段階の表と進捗の名前")
    for stage, name in stages.items():
        cells = rows[name]
        shown = name.replace("<表示>", "standard")
        assert progress.stage_label(shown) in cells[1], f"段階 {name} の進捗の名前が違う"
        assert _names(cells[2]) == {
            command for command, names in used.items() if stage in names
        }, f"段階 {name} を使うコマンドが違う"
    for stage, group in report._TIMING_GROUPS.items():
        assert re.search(
            rf"`{re.escape(stage)}[`:][^。]*「{re.escape(group)}」", section
        ), f"要約の内訳で {stage} をまとめる名前がない"


def _table_lines(section: str) -> list[list[str]]:
    return [
        [cell.strip() for cell in _CELL_SPLIT.split(line.strip().strip("|"))]
        for line in section.splitlines()
        if line.startswith("| ") and not line.startswith("| 変更 |")
    ]


def test_carry_table_matches_evidence():
    rows = _table_lines(_section(_text(USAGE), "確認と完成物の保存"))

    def row(keyword: str) -> list[str]:
        found = [cells for cells in rows if keyword in cells[0]]
        assert len(found) == 1, f"引き継ぎの表に「{keyword}」の行が 1 つでない"
        return found[0]

    # ノートと出典は内容の証拠だけに入り、表示の確認（ページの部品の証拠）は引き継ぐ。
    assert (
        row("ノート")[1:3] == ["そのページだけやり直す", "引き継ぐ"] and "出典" in row("ノート")[0]
    ), "ノート・出典だけの変更の行が実装と違う"
    assert not any("ノート" in cells[0] and "表示する" in cells[0] for cells in rows)
    project = SimpleNamespace(config={}, input_files={})
    page = {
        "id": "a",
        "title": "題名",
        "notes": "説明",
        "elements": [{"id": "body", "kind": "text", "content": "本文"}],
    }
    cited = {**page, "elements": [{**page["elements"][0], "sources": [{"bib_key": "key"}]}]}
    assert evidence.content_digest(page, project) != evidence.content_digest({**page, "notes": "別の説明"}, project)
    assert evidence.content_digest(page, project) != evidence.content_digest(cited, project)
    assert "notesSlide" in evidence._UNFOLLOWED["slide"]
    # ページ番号は桁数が同じなら証拠が変わらない（数字を 0 にして部品を比べる）。桁数・書式・枠が変わっても、
    # 番号の図形を除いた証拠と、番号の領域を除いた画像が一致すれば引き継ぐ。
    assert row("桁数が同じ")[2].startswith("引き継ぐ") and row("桁数・書式・枠")[2].startswith(
        "引き継ぐ（番号の図形を除いた"
    ), "ページ番号の行の表示の確認が実装と違う"
    assert "body_hash" in _section(_text(USAGE), "確認と完成物の保存"), "番号の図形を除いた証拠の説明がない"

    def masked(number: int) -> bytes:
        shape = etree.fromstring(
            f'<p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><p:nvSpPr><p:cNvPr id="2" name="{evidence.FOOTER_ID}"/></p:nvSpPr>'
            f"<p:txBody><a:p><a:r><a:t>{number}</a:t></a:r></a:p></p:txBody></p:sp>"
        )
        evidence._mask_page_number(shape)
        return etree.tostring(shape)

    assert masked(12) == masked(13) and masked(9) != masked(10)
    # 表示されるフィールドがあるページだけ、資料の中の位置を表示の証拠に入れる。
    assert "`a:fld`" in _section(_text(USAGE), "確認と完成物の保存"), "フィールドのあるページの扱いの説明がない"
    for fields, moved in (((), False), (("slidenum",), True)):
        parts = evidence.PageParts("a", 1, "choice", "full", None, fields)
        digests = {
            evidence.visual_digest(parts, "standard", "env", render_date="2026-01-01", number=number)
            for number in (1, 2)
        }
        assert (len(digests) == 2) == moved, fields


def test_declared_warnings_are_listed_in_summaries():
    section = _section(_text(USAGE), "警告の許容")
    assert (
        "`review status` と `check --final` の要約" in section and "`declared_warnings`" in section
    ), "宣言で許容した警告が要約に並ぶことの説明がない"
    entry = {
        "code": "element_overlap",
        "slide_id": "a",
        "slide_number": 2,
        "element_id": "label",
        "related_element_id": "figure",
        "reason": "図の上にラベルを重ねる",
        "location": "slides.py:12",
    }
    assert "`--verbose`" in section, "--verbose で一覧を出すことの説明がない"
    for command, result in (
        ("review status", {"status": "complete", "declared_warnings": [entry]}),
        ("check", {"status": "passed", "final": True, "declared_warnings": [entry]}),
    ):
        shown = report.summarize(command, result, verbose=True)
        assert "宣言で許容した警告:" in shown and "図の上にラベルを重ねる" in shown, f"{command} の要約に並ばない"
        assert "宣言で許容した警告: 1 件" in report.summarize(command, result), f"{command} の要約に件数がない"
    # layout・doctor・build の要約では、件数の内数と行の末尾の印で示す。
    assert (
        "「うち宣言で許容 N 件」" in section and "「（宣言で許容）」" in section
    ), "layout・build の要約での示し方の説明がない"
    issue = {key: entry[key] for key in ("code", "slide_id", "slide_number", "element_id", "related_element_id")}
    issue.update(severity="warning", message="要素の重なりを確認する必要がある", declared={"reason": entry["reason"]})
    shown = report.summarize(
        "layout",
        {"status": "passed", "issues": [issue], "counts": {"error": 0, "warning": 1}, "exit_code": 0},
        verbose=True,
    )
    assert "うち宣言で許容 1 件" in shown and shown.splitlines()[2].endswith("（宣言で許容）"), shown


def test_font_directories_setting_replaces_detection(monkeypatch):
    # font_directories を書いたときはそのフォルダーだけを探し、自動では探さない（二者択一）。
    section = _section(_text(USAGE), "既定の書体の探し方")
    assert (
        "`font_directories` を書いたとき：そのフォルダーだけを探し、自動では探さない" in section
    ), "設定したフォルダーだけを探すことの説明がない"
    assert "書かないとき：自動で探したフォルダーを使う" in section, "設定がないときに自動で探すことの説明がない"

    def detect(environment):
        return (Path("/detected"),)

    monkeypatch.setattr(fonts, "_detect_font_directories", detect)
    assert fonts.font_directories({"font_directories": ["/configured"]}) == (Path("/configured"),)
    assert fonts.font_directories({"font_directories": []}) == ()
    assert fonts.font_directories({}) == (Path("/detected"),)


def test_diagnostics_head_matches_records():
    text = _text(ROOT / "docs" / "diagnostics.md")
    head = text.split("<!-- slidekit:diagnostics:start -->")[0]
    # 確認記録の操作を拒むときの識別子（records が RecordError に付ける code）をすべて挙げる。
    codes = {records.RecordError.code, records.LegacyRecordError.code, "artifact_changed", "artifact_unrecorded"}
    # 複数行に分かれた呼び出しも漏らさないように、構文木から RecordError(…, code="…") の code を集める。
    for node in ast.walk(ast.parse(inspect.getsource(records))):
        name = getattr(node.func, "id", getattr(node.func, "attr", "")) if isinstance(node, ast.Call) else ""
        if name.endswith("RecordError"):
            codes |= {
                keyword.value.value
                for keyword in node.keywords
                if keyword.arg == "code" and isinstance(keyword.value, ast.Constant)
            }
    assert {"build_incomplete", "build_not_exportable"} <= codes
    for code in codes:
        assert f"`{code}`" in head, f"docs/diagnostics.md の冒頭に、操作を拒む識別子 {code} がない"
    # 選んだ章だけを作った場面の診断は slide_number を持たない。
    row = next(line for line in head.splitlines() if line.startswith("| `slide_id`・`slide_number`"))
    assert "`slide_number` を持たず" in row, "選んだ章だけを作った場面の診断が slide_number を持たないことの説明がない"
    issue = validators.make_diagnostic(
        "text_overflow", "はみ出し", scope="layout", slide_id="a", element_id="e", slide_number=1, evidence={}
    )
    scene = {"slides": [{"id": "a", "number": 1, "elements": []}]}
    assert pipeline._finish_issues([issue], scene)[0]["slide_number"] == 1
    selected = pipeline._finish_issues([issue], {**scene, "selection": {"construct": "selected"}})
    assert "slide_number" not in selected[0]
