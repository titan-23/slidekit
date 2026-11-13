"""照合用の資料を設定ファイルから作り、PowerPoint の PDF を計測して、今の計測と照合データと比べる（開発用。パッケージに含めない）。

使い方（PowerPoint を使う処理は一度に 1 つだけ実行する）:
  python tools/calibrate.py run CONFIG --work DIR                 # 新しい作業ディレクトリに資料を作り、slidekit build で変換して計測し、比べる
  python tools/calibrate.py run --work DIR --measure-only [--run N]  # 変換せず、保存した PDF を今のコードで計測し直して比べる
  python tools/calibrate.py run --work DIR                        # 作業ディレクトリの資料を新しい実行に変換し直して比べる
  python tools/calibrate.py adopt --work DIR [--run N] [--only KEY ...] [--output PATH]  # 計測値を照合データに採用する

設定ファイル（TOML）の deck で資料の種類を選ぶ。
- "cases"：[[cases]] の式・文字サイズ・行間・文中か独立か・揃え・前後の文字から資料を作る（値の配列はすべての組み合わせにする）。
  文中の数式は、前後の文字が [ と ] を含めば、数式の照合データと同じ方法で送り幅と描画範囲も計測する。
- "layout"：校正用の資料 tests/data/layout/deck を継承する資料。文字の照合データ（powerpoint_layout.json）を作り直す。
- "math"：数式の照合データの道具（tests/tools/powerpoint_math_data.py）の式の一覧の資料（expressions で一部に絞ることができる）。
  数式の照合データ（powerpoint_math.json）を作り直す。

作業ディレクトリは slidekit の外の新しいディレクトリにする。設定の写し calibration.toml、生成条件 conditions.json、資料、
slidekit build の実行（out/runs/NNNN/。PDF と生成記録）、実行ごとの計測の結果 results/NNNN/ を置く。results/NNNN/ の内容:
- layout_fixture.json：文字の照合データの形の計測値（tools/make_layout_fixture.py の make_fixture）。
- math_fixture.json：数式の照合データの形の計測値（"cases" では文字サイズと書体を項目ごとに持つ）。
- comparison.json：今の計測との差（tests/test_layout_fixture.py の compare_case と許容差の規則、
  slidekit.math.metrics.compare_reference）と、既存の照合データとの差。
照合データの更新は adopt で明示して行う。"cases" の資料は採用しない（校正用の資料か数式の道具の式の一覧に加えてから作り直す）。
--only は既存の照合データの一部（文字は「ページ/要素」、数式は「inline:式」「display:式」か式）だけを置き換える。
"""

from __future__ import annotations

import argparse
from collections.abc import Mapping
from datetime import date, datetime, timezone
import importlib
import importlib.util
import itertools
import json
import math
from pathlib import Path
import platform
import shutil
import subprocess
import sys
import tomllib

ROOT = Path(__file__).resolve().parents[1]
TESTS = ROOT / "tests"
CALIBRATION_DECK = TESTS / "data" / "layout" / "deck" / "deck.toml"
LAYOUT_FIXTURE = TESTS / "data" / "layout" / "powerpoint_layout.json"
MATH_FIXTURE = TESTS / "data" / "powerpoint_math.json"
SCHEMA_VERSION = 1
DECKS = ("cases", "layout", "math")
PDF = Path("renders") / "powerpoint" / "standard" / "deck.pdf"

# "cases" の各ケースの既定値。EXPANDED の値は配列にすると、すべての組み合わせを作る。
CASE_DEFAULTS = {
    "mode": "inline",
    "size_pt": 20.0,
    "line_spacing": 1.0,
    "align": "left",
    "valign": "top",
    "before": "[",
    "after": "]",
    "latin": "Arial",
    "japanese": "Yu Gothic UI",
    "math": "Cambria Math",
    "width_em": 14.4,  # 枠の幅（文字サイズの em）。20 pt で数式の照合データの枠と同じ 288 pt
    "height_em": 5.6,  # 枠の高さ（文字サイズ × max(1, 行間) の em）。20 pt・行間 1 で 112 pt
    "tags": [],
}
EXPANDED = ("mode", "size_pt", "line_spacing", "align", "valign")
CHOICES = {"mode": ("inline", "display"), "align": ("left", "center", "right"), "valign": ("top", "middle", "bottom")}
# 16:9 の資料の枠を置く範囲（pt）と、枠の間隔（数式の照合データの計測は上下に 4 pt 広げた範囲を見る）。
LEFT, TOP, RIGHT, BOTTOM, GAP = 20.0, 14.0, 940.0, 530.0, 8.0
DIFFERENCE_PT = 0.01  # 既存の照合データと同じとみなす差

DECK_TOML = """# 照合用の資料（tools/calibrate.py が作る。内容は汎用の式だけ）。
schema_version = 1
entrypoint = "slides.py:build"
size = "16:9"
theme = "theme.toml"
output_dir = "out"

[metadata]
title = "照合用の資料"
author = "slidekit example"
last_modified_by = "slidekit example"

[render]
primary = "powerpoint"
required = ["powerpoint"]

# PowerPoint が描いた編集可能な数式を照合するので、互換画像は作らない。
[math]
fallback = "off"
"""

THEME_TOML = """# 照合用の資料のテーマ。小さい文字も許容し、ページのほぼ全体を本文の領域にする。
schema_version = 1

[roles.sample]
inherits = "body"
min_size_pt = 6

[page_kinds.blank]
default_role = "sample"
content_area_pt = [12, 10, 936, 524]
"""

SLIDES_PY = '''"""照合用の資料（tools/calibrate.py が作る）。枠と書式は conditions.json の cases と同じ。"""

import json

from slidekit import Deck, Math, Rect, Text

PAGES = json.loads({pages!r})


def build(context):
    deck = Deck(context)
    for page in PAGES:
        slide = deck.slide(id=page["id"], kind="blank", title=page["id"])
        for item in page["items"]:
            style = {{
                "size_pt": item["size_pt"],
                "line_spacing": item["line_spacing"],
                "align": item["align"],
                "valign": item["valign"],
                "word_wrap": False,
                "padding_pt": 0.0,
                "font_latin": item["latin"],
                "font_japanese": item["japanese"],
                "font_math": item["math"],
            }}
            if item["tags"]:
                style["description"] = "tags: " + " ".join(item["tags"])
            area = Rect(*item["area"])
            if item["mode"] == "display":
                slide.equation(id=item["id"], expression=item["expression"], area=area, **style)
            else:
                content = [Math(item["expression"])]
                if item["before"]:
                    content.insert(0, Text(item["before"]))
                if item["after"]:
                    content.append(Text(item["after"]))
                slide.text(id=item["id"], text=content, area=area, **style)
    return deck
'''


class CalibrationError(ValueError):
    """照合用の資料を作れないか、計測・採用できない入力。"""


# ---- 既存の道具と比べる処理 ----


def _load(path: Path):
    """tools/ と tests/tools/ の道具をモジュールとして読む（試験の load_tool と同じ名前で、同じプロセスでは 1 回だけ）。"""
    name = f"_slidekit_tool_{path.stem}"
    if name not in sys.modules:
        spec = importlib.util.spec_from_file_location(name, path)
        module = importlib.util.module_from_spec(spec)
        sys.modules[name] = module
        spec.loader.exec_module(module)
    return sys.modules[name]


def layout_tool():
    """文字の照合データの道具（tools/make_layout_fixture.py）。"""
    return _load(ROOT / "tools" / "make_layout_fixture.py")


def math_tool():
    """数式の照合データの道具（tests/tools/powerpoint_math_data.py）。"""
    return _load(TESTS / "tools" / "powerpoint_math_data.py")


def comparison():
    """文字の配置の比べ方（tests/test_layout_fixture.py の compare_case・predicted_lines と許容差の規則）。"""
    if str(TESTS) not in sys.path:
        sys.path.insert(0, str(TESTS))
    return importlib.import_module("test_layout_fixture")


# ---- 設定と資料 ----


def _number(value, label: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value <= 0:
        raise CalibrationError(f"{label} は正の数にする: {value!r}")
    return float(value)


def _values(case: Mapping, key: str, label: str) -> list:
    """EXPANDED の値の一覧（配列でなければ 1 つ）。"""
    value = case[key]
    values = value if isinstance(value, list) else [value]
    if not values:
        raise CalibrationError(f"{label} の {key} が空")
    if key in CHOICES:
        wrong = [item for item in values if item not in CHOICES[key]]
        if wrong:
            raise CalibrationError(f"{label} の {key} は {'・'.join(CHOICES[key])} のどれかにする: {wrong}")
        return list(dict.fromkeys(values))
    return list(dict.fromkeys(_number(item, f"{label} の {key}") for item in values))


def _label(value: float) -> str:
    return f"{value:g}"


def expand_cases(config: Mapping) -> list[dict]:
    """[[cases]] を、式と EXPANDED の値の組み合わせごとの要素にする（枠の位置は place で決める）。"""
    defaults = config.get("defaults", {})
    if not isinstance(defaults, Mapping) or set(defaults) - set(CASE_DEFAULTS):
        raise CalibrationError(f"defaults の項目は {', '.join(CASE_DEFAULTS)} にする")
    cases = config.get("cases")
    if not isinstance(cases, list) or not cases:
        raise CalibrationError('deck = "cases" では [[cases]] を 1 つ以上書く')
    items, names = [], set()
    for number, entry in enumerate(cases, 1):
        label = f"cases[{number}]"
        if not isinstance(entry, Mapping):
            raise CalibrationError(f"{label} は表にする")
        unknown = set(entry) - {"name", "expressions", *CASE_DEFAULTS}
        if unknown:
            raise CalibrationError(f"{label}: 未対応の項目 {sorted(unknown)}")
        case = {**CASE_DEFAULTS, **defaults, **entry}
        name = case.get("name", f"case{number}")
        if not isinstance(name, str) or not name or any(char in name for char in "/ \t\n"):
            raise CalibrationError(f"{label} の name は / と空白を含まない文字列にする: {name!r}")
        if name in names:
            raise CalibrationError(f"{label} の name が重なる: {name}")
        names.add(name)
        expressions = case.get("expressions")
        if not isinstance(expressions, list) or not expressions or not all(isinstance(e, str) for e in expressions):
            raise CalibrationError(f"{label} の expressions は式の文字列の配列にする")
        for key in ("before", "after", "latin", "japanese", "math"):
            if not isinstance(case[key], str) or (key not in ("before", "after") and not case[key]):
                raise CalibrationError(f"{label} の {key} は文字列にする: {case[key]!r}")
        tags = case["tags"]
        if not isinstance(tags, list) or not all(isinstance(tag, str) and tag and " " not in tag for tag in tags):
            raise CalibrationError(f"{label} の tags は空白を含まない文字列の配列にする: {tags!r}")
        width_em, height_em = (_number(case[key], f"{label} の {key}") for key in ("width_em", "height_em"))
        values = [_values(case, key, label) for key in EXPANDED]
        for index, expression in enumerate(expressions, 1):
            for mode, size, spacing, align, valign in itertools.product(*values):
                items.append(
                    {
                        "id": f"{name}_{index:02d}_{mode}_s{_label(size)}_sp{_label(spacing)}_{align}_{valign}",
                        "expression": expression,
                        "mode": mode,
                        "size_pt": size,
                        "line_spacing": spacing,
                        "align": align,
                        "valign": valign,
                        "before": case["before"],
                        "after": case["after"],
                        "latin": case["latin"],
                        "japanese": case["japanese"],
                        "math": case["math"],
                        "tags": list(tags),
                        "width": round(width_em * size, 2),
                        "height": round(height_em * size * max(1.0, spacing), 2),
                    }
                )
    return items


def place(items: list[dict]) -> list[dict]:
    """枠を左から順に置き、右端を越えるなら次の行へ、下端を越えるなら次のページへ送る。ページの一覧を返す。"""
    pages: list[dict] = []
    x = y = row = 0.0
    for item in items:
        width, height = item["width"], item["height"]
        if width > RIGHT - LEFT or height > BOTTOM - TOP:
            raise CalibrationError(f"枠がページに入らない: {item['id']}（{width} × {height} pt）")
        if pages and x > LEFT and x + width > RIGHT:
            x, y, row = LEFT, y + row + GAP, 0.0
        if not pages or y + height > BOTTOM:
            pages.append({"id": f"p{len(pages) + 1:02d}", "items": []})
            x, y, row = LEFT, TOP, 0.0
        pages[-1]["items"].append({**item, "area": [x, y, width, height]})
        x, row = x + width + GAP, max(row, height)
    return pages


def read_config(path: Path) -> dict:
    try:
        with Path(path).open("rb") as stream:
            config = tomllib.load(stream)
    except FileNotFoundError:
        raise CalibrationError(f"設定ファイルがありません: {path}") from None
    except tomllib.TOMLDecodeError as error:
        raise CalibrationError(f"設定ファイルを TOML として読めません: {path}: {error}") from None
    if config.get("schema_version") != SCHEMA_VERSION:
        raise CalibrationError(f"設定ファイルの schema_version は {SCHEMA_VERSION} にする")
    unknown = set(config) - {"schema_version", "deck", "expressions", "defaults", "cases"}
    if unknown:
        raise CalibrationError(f"設定ファイルの未対応の項目: {sorted(unknown)}")
    if config.get("deck", "cases") not in DECKS:
        raise CalibrationError(f"deck は {'・'.join(DECKS)} のどれかにする: {config.get('deck')!r}")
    return config


def _check_work(work: Path) -> Path:
    work = Path(work).resolve()
    if work == ROOT or ROOT in work.parents:
        raise CalibrationError(
            f"作業ディレクトリは slidekit の外に作る（実行の記録にローカルのパスが入るため）: {work}"
        )
    return work


def plan(config: Mapping) -> dict:
    """設定から生成条件（conditions.json の内容。資料の種類と、"cases"・"math" ではページごとの枠）を求める。"""
    kind = config.get("deck", "cases")
    conditions = {"schema_version": SCHEMA_VERSION, "deck": kind, "created_at": _now(), "versions": _versions()}
    if kind == "layout":
        if "cases" in config or "expressions" in config or "defaults" in config:
            raise CalibrationError('deck = "layout" では cases・defaults・expressions を書かない（校正用の資料を使う）')
        conditions["extends"] = CALIBRATION_DECK.relative_to(ROOT).as_posix()
    elif kind == "math":
        if "cases" in config or "defaults" in config:
            raise CalibrationError('deck = "math" では cases と defaults を書かない（式は expressions で選ぶ）')
        tool = math_tool()
        expressions = config.get("expressions", list(tool.EXPRESSIONS))
        if not isinstance(expressions, list) or not expressions or not all(isinstance(e, str) for e in expressions):
            raise CalibrationError("expressions は式の文字列の配列にする")
        accepted, rejected = tool.supported(expressions)
        if not accepted:
            raise CalibrationError("書き出すことができる式がない")
        conditions.update(
            size_pt=tool.SIZE_PT, pages=tool.layout_items(accepted), rejected=rejected, subset="expressions" in config
        )
    else:
        if "expressions" in config:
            raise CalibrationError('deck = "cases" では式を [[cases]] の expressions に書く')
        conditions["pages"] = place(expand_cases(config))
    return conditions


def write_deck(work: Path, conditions: Mapping) -> None:
    """作業ディレクトリに資料（deck.toml と、"cases" では theme.toml・slides.py）を書く。"""
    kind = conditions["deck"]
    if kind == "layout":
        text = f"schema_version = 1\nextends = {json.dumps((ROOT / conditions['extends']).as_posix())}\n"
        (work / "deck.toml").write_text(text, encoding="utf-8")
    elif kind == "math":
        math_tool().write_deck(work, conditions["pages"])
    else:
        (work / "deck.toml").write_text(DECK_TOML, encoding="utf-8")
        (work / "theme.toml").write_text(THEME_TOML, encoding="utf-8")
        source = SLIDES_PY.format(pages=json.dumps(conditions["pages"], ensure_ascii=False))
        (work / "slides.py").write_text(source, encoding="utf-8")


def _versions() -> dict:
    from slidekit import __version__, measure
    from slidekit.math import metrics

    return {
        "slidekit": __version__,
        "measure": measure.MEASURE_VERSION,
        "math_engine": metrics.ENGINE_VERSION,
        "python": platform.python_version(),
    }


def prepare(config_path: Path, work: Path) -> dict:
    """新しい作業ディレクトリに設定の写し・資料・生成条件を書く。"""
    work = _check_work(work)
    if work.exists() and (not work.is_dir() or any(work.iterdir())):
        raise CalibrationError(f"作業ディレクトリは新しいディレクトリにする（空でない）: {work}")
    conditions = plan(read_config(config_path))
    work.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(config_path, work / "calibration.toml")
    write_deck(work, conditions)
    _write_json(work / "conditions.json", conditions)
    return conditions


def _write_json(path: Path, value) -> None:
    from slidekit.files import write_json

    write_json(path, value)


def _read_json(path: Path) -> dict:
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise CalibrationError(f"ファイルがありません: {path}") from None
    except json.JSONDecodeError as error:
        raise CalibrationError(f"JSON として読めません: {path}: {error}") from None


# ---- 変換 ----


def convert(work: Path, *, runner=None) -> Path:
    """slidekit build で資料を変換し（PowerPoint は slidekit の排他で直列に使う）、実行ディレクトリを返す。

    runner は subprocess.run と同じ呼び方の関数（試験で差し替える）。
    """
    command = [sys.executable, "-m", "slidekit", "build", "--project", str(work / "deck.toml")]
    completed = (runner or subprocess.run)(
        [*command, "--json", "--quiet", "--carry-from", "none"], capture_output=True, text=True, check=False
    )
    try:
        result = json.loads(completed.stdout)
    except (TypeError, json.JSONDecodeError):
        tail = "\n".join(str(completed.stderr or "").splitlines()[-20:])
        raise CalibrationError(
            f"slidekit build の結果を読めません（終了コード {completed.returncode}）:\n{tail}"
        ) from None
    run = Path(result.get("build") or "")
    if result.get("status") != "built" or not (run / PDF).is_file():
        counts = result.get("counts")
        raise CalibrationError(
            f"slidekit build が変換を終えていない: 状態 {result.get('status')}、診断 {counts}、実行 {run}"
        )
    return run


def _runs(work: Path) -> list[Path]:
    directory = work / "out" / "runs"
    if not directory.is_dir():
        return []
    return sorted((path for path in directory.iterdir() if path.name.isdigit()), key=lambda path: int(path.name))


def find_run(work: Path, value: str | None = None) -> Path:
    """計測する実行。value は実行ディレクトリか番号で、省略すると標準表示の PDF がある最新の実行。"""
    if value:
        path = Path(value)
        run = path if path.is_dir() else work / "out" / "runs" / (f"{int(value):04d}" if value.isdigit() else value)
        if not (run / PDF).is_file():
            raise CalibrationError(f"実行に標準表示の PDF がありません: {run}")
        return run.resolve()
    runs = [run for run in _runs(work) if (run / PDF).is_file() and (run / "deck.json").is_file()]
    if not runs:
        raise CalibrationError(f"PDF のある実行がありません（先に run で変換する）: {work / 'out' / 'runs'}")
    return runs[-1].resolve()


# ---- 計測と比べる処理 ----


def _key(case: Mapping) -> str:
    return f"{case['slide_id']}/{case['element_id']}"


# 後から加えた書式の項目と、その既定値。既定値なら対応付けの内容から除く（加える前に作った照合データと対応付けるため）。
_ADDED_STYLE_DEFAULTS = {"phrase_wrap": False}


def _signature(element: Mapping) -> str:
    """位置を除いた要素の内容（既存の照合データの要素と対応付ける）。"""
    area = element.get("area") or [0, 0, 0, 0]
    value = {key: element.get(key) for key in ("kind", "shape", "content", "style")}
    if isinstance(value["style"], Mapping):
        value["style"] = {
            key: item
            for key, item in value["style"].items()
            if not (key in _ADDED_STYLE_DEFAULTS and item == _ADDED_STYLE_DEFAULTS[key])
        }
    return json.dumps({**value, "size": area[2:]}, ensure_ascii=False, sort_keys=True)


def _relative(case: Mapping) -> dict:
    """枠の左上を原点にした実際の値。"""
    x, y = (case["element"].get("area") or [0, 0])[:2]
    actual = case["actual"]
    ink = actual.get("ink_pt")
    return {
        "lines": actual["lines"],
        "baselines": [value - y for value in actual["baselines_pt"]],
        "extents": [[left - x, right - x] for left, right in actual["line_extents_pt"]],
        "ink": None if ink is None else [ink[0] - x, ink[1] - y, ink[2] - x, ink[3] - y],
    }


def _largest(pairs) -> float | None:
    values = [abs(a - b) for a, b in pairs]
    return round(max(values), 3) if values else None


def compare_layout_fixture(candidate: Mapping, existing: Mapping | None, *, complete: bool) -> dict:
    """計測値と既存の文字の照合データの差。要素の内容（位置を除く）で対応付け、枠の左上からの位置で比べる。

    complete が真（校正用の資料）なら、既存の照合データにだけある要素も示す。
    """
    if existing is None:
        return {"available": False}
    pool: dict[str, list[dict]] = {}
    for case in existing.get("cases", []):
        pool.setdefault(_signature(case["element"]), []).append(case)
    changed, new, same = [], [], 0
    for case in candidate.get("cases", []):
        matches = pool.get(_signature(case["element"]))
        if not matches:
            new.append(_key(case))
            continue
        old = matches.pop(0)
        a, b = _relative(case), _relative(old)
        row = {"key": _key(case), "existing": _key(old), "lines": [a["lines"], b["lines"]]}
        if a["lines"] == b["lines"]:
            row["baseline_pt"] = _largest(zip(a["baselines"], b["baselines"]))
            row["extent_pt"] = _largest(
                pair for lines in zip(a["extents"], b["extents"]) for pair in zip(lines[0], lines[1])
            )
        if a["ink"] is not None and b["ink"] is not None:
            row["ink_pt"] = _largest(zip(a["ink"], b["ink"]))
        values = [row.get(key) for key in ("baseline_pt", "extent_pt", "ink_pt")]
        if a["lines"] != b["lines"] or any(value is not None and value > DIFFERENCE_PT for value in values):
            changed.append(row)
        else:
            same += 1
    missing = [_key(case) for cases in pool.values() for case in cases] if complete else []
    return {
        "available": True,
        "environment_changed": candidate.get("environment") != existing.get("environment"),
        "same": same,
        "changed": changed,
        "new": new,
        "missing": missing,
    }


def _math_values(item: Mapping) -> list[float]:
    if item["mode"] == "display":
        return [item["ink_width_pt"], item["ink_height_pt"]]
    return [item["advance_pt"], *(item.get("ink_pt") or [])]


def compare_math_fixture(items: list[Mapping], existing: Mapping | None, *, complete: bool) -> dict:
    """計測値と既存の数式の照合データの差（同じ文字サイズ・書体の、同じ式と mode の項目）。値の差は pt と em。"""
    if existing is None:
        return {"available": False}
    size = existing.get("size_pt")
    family = (existing.get("font") or {}).get("family")
    text = existing.get("text_fonts") or {}
    old = {(item["expression"], item["mode"]): item for item in existing.get("items", [])}
    changed, new, same, seen = [], [], 0, set()
    for item in items:
        if item.get("size_pt", size) != size or item.get("math", family) != family:
            continue
        if (item.get("latin", text.get("latin")), item.get("japanese", text.get("japanese"))) != (
            text.get("latin"),
            text.get("japanese"),
        ):
            continue
        key = (item["expression"], item["mode"])
        if key not in old:
            new.append(f"{key[1]}:{key[0]}")
            continue
        seen.add(key)
        a, b = _math_values(item), _math_values(old[key])
        if len(a) != len(b):
            changed.append({"key": f"{key[1]}:{key[0]}", "values": [a, b]})
            continue
        largest = max((abs(x - y) for x, y in zip(a, b)), default=0.0)
        if largest > DIFFERENCE_PT:
            row = {"key": f"{key[1]}:{key[0]}", "largest_pt": round(largest, 3), "largest_em": round(largest / size, 4)}
            changed.append({**row, "new": a, "existing": b})
        else:
            same += 1
    missing = [f"{mode}:{expression}" for expression, mode in old if (expression, mode) not in seen] if complete else []
    return {"available": True, "same": same, "changed": changed, "new": new, "missing": missing}


def layout_differences(case: Mapping, fonts, compare) -> dict:
    """1 要素の今の計測と実際の差（実際 − 予測。pt）。行数が同じなら行ごとの基線・左端・右端、インクの 4 辺。"""
    from slidekit import measure

    element, actual = case["element"], case["actual"]
    row = {"key": _key(case), "tags": case.get("tags", []), "size_pt": compare._size(element)}
    try:
        layout = measure.layout_element(element, fonts)
        lines = compare.predicted_lines(layout, element, fonts)
    except (measure.MeasureError, ValueError) as error:
        return {**row, "error": str(error)}
    row["lines"] = [len(lines), actual["lines"]]
    largest = []
    if len(lines) == actual["lines"]:
        row["baseline_pt"] = [round(b - line["baseline"], 3) for line, b in zip(lines, actual["baselines_pt"])]
        row["left_pt"] = [round(e[0] - line["left"], 3) for line, e in zip(lines, actual["line_extents_pt"])]
        row["right_pt"] = [round(e[1] - line["right"], 3) for line, e in zip(lines, actual["line_extents_pt"])]
        largest += row["baseline_pt"] + row["left_pt"] + row["right_pt"]
    ink = layout.ink
    if ink is not None and actual.get("ink_pt") is not None:
        expected = (ink.x, ink.y, ink.x + ink.w, ink.y + ink.h)
        row["ink_pt"] = [round(b - a, 3) for a, b in zip(expected, actual["ink_pt"])]
        largest += row["ink_pt"]
    row["largest_pt"] = max((abs(value) for value in largest), default=0.0)
    return row


def _math_items_of_cases(pdf: Path, pages: list[dict]) -> tuple[list[dict], list[str]]:
    """資料の種類が cases のとき、目印の [ と ] で挟んだ文中の数式と独立した数式を、数式の照合データと同じ方法で計測する。"""
    from slidekit.pdf import open_pdf

    tool = math_tool()
    items, failed = [], []
    try:
        with open_pdf(pdf) as document:
            if len(document) != len(pages):
                raise CalibrationError(f"PDF と資料のページ数が一致しない: PDF={len(document)}, 資料={len(pages)}")
            for page, layout in zip(document, pages):
                for item in layout["items"]:
                    if item["mode"] == "inline":
                        if "[" not in item["before"] or "]" not in item["after"]:
                            continue
                        value = tool.measure_inline(page, item["area"], math_font=item["math"])
                    else:
                        value = tool.measure_display(page, item["area"])
                    if value is None:
                        failed.append(item["id"])
                        continue
                    keys = ("id", "expression", "mode", "size_pt", "latin", "japanese", "math")
                    items.append({**{key: item[key] for key in keys}, **value})
    finally:
        tool._RENDERED.clear()
    return items, failed


def compare_math(items: list[dict], fonts, *, size: float | None = None) -> dict:
    """数式の計測値を今の計測と比べる（slidekit.math.metrics.compare_reference。文字サイズと書体ごと）。"""
    from slidekit.math.metrics import compare_reference

    if not hasattr(fonts, "_math_measurer"):
        return {"available": False, "reason": "数式の計測器がない書体"}
    groups: dict[tuple, list[dict]] = {}
    for item in items:
        key = (item.get("size_pt", size), item.get("math"), item.get("japanese"), item.get("latin"))
        groups.setdefault(key, []).append(item)
    results = []
    for (group_size, family, japanese, latin), members in groups.items():
        data = {"size_pt": group_size, "items": members}
        try:
            result = compare_reference(
                data, fonts._math_measurer(family), text_fonts=fonts._text_fonts(japanese, latin)
            )
        except ValueError as error:
            results.append({"size_pt": group_size, "math": family, "error": str(error)})
            continue
        if len(result["items"]) == len(members):  # 文字の書体を渡すので、比べない式はない
            for row, member in zip(result["items"], members):
                if member.get("id"):
                    row["id"] = member["id"]
        results.append({"size_pt": group_size, "math": family, **result})
    passed = all(result.get("passed") for result in results)
    return {"available": True, "passed": passed, "groups": results}


def measure_run(
    work: Path,
    run: Path,
    *,
    fonts=None,
    layout_fixture: Path | None = LAYOUT_FIXTURE,
    math_fixture: Path | None = MATH_FIXTURE,
) -> dict:
    """保存した PDF を計測して results/<実行>/ に書き、比べた結果（comparison.json の内容）を返す。

    fonts は今の計測に使う FontSource で、省略すると作業ディレクトリの資料の書体（FontRegistry）を使う。
    """
    from slidekit.config import load_project
    from slidekit.fonts import FontRegistry

    work, run = Path(work), Path(run)
    conditions = _read_json(work / "conditions.json")
    project = load_project(work / "deck.toml")
    if fonts is None:
        with FontRegistry(project) as registry:
            return measure_run(work, run, fonts=registry, layout_fixture=layout_fixture, math_fixture=math_fixture)
    kind = conditions["deck"]
    output = work / "results" / run.name
    output.mkdir(parents=True, exist_ok=True)
    compare = comparison()
    result: dict = {"run": run.name, "deck": kind, "measured_at": _now(), "versions": _versions()}

    # 文字の配置（"cases" と "layout"）。"math" の資料は式ごとの枠が隣り合い、目印の角括弧で式を測る資料なので、
    # 要素ごとの行の照合はしない（隣の枠の字形を拾い、照合の対象でない違いが出る）。
    if kind == "math":
        result["layout"] = {"skipped": "数式の資料では文字の配置を比べない（式の送り幅と描画範囲だけを比べる）"}
    else:
        candidate = None
    try:
        candidate = None if kind == "math" else layout_tool().make_fixture(run)
    except ValueError as error:
        result["layout"] = {"error": str(error)}
    if candidate is not None:
        _write_json(output / "layout_fixture.json", candidate)
        rows = [layout_differences(case, fonts, compare) for case in candidate["cases"]]
        problems = [message for case in candidate["cases"] for message in compare.compare_case(case, fonts)]
        existing = _optional_json(layout_fixture)
        result["layout"] = {
            "cases": len(rows),
            "unread": candidate["unread"],
            "problems": problems,
            "differences": rows,
            "font_mismatches": compare.font_mismatches(candidate, project.environment),
            "fixture": compare_layout_fixture(candidate, existing, complete=kind == "layout"),
        }

    # 数式の送り幅と描画範囲（"math" と "cases"）
    pdf = run / PDF
    if kind == "math":
        tool = math_tool()
        family = project.theme["fonts"]["math"]
        try:
            items, failed = tool.measure_pdf(pdf, conditions["pages"], math_font=family)
            font = tool.font_record(project)
        except (OSError, ValueError) as error:
            raise CalibrationError(f"数式を計測できない: {error}") from None
        manifest = _read_json(run / "manifest.json")
        data = {
            "schema_version": tool.SCHEMA_VERSION,
            "created": date.today().isoformat(),
            "powerpoint": layout_tool()._powerpoint_version(manifest),
            "font": font,
            "text_fonts": {"japanese": project.theme["fonts"]["japanese"], "latin": project.theme["fonts"]["latin"]},
            "size_pt": conditions["size_pt"],
            "items": items,
        }
        _write_json(output / "math_fixture.json", data)
        theme = project.theme["fonts"]
        keyed = [{**item, "math": family, "japanese": theme["japanese"], "latin": theme["latin"]} for item in items]
        result["math"] = {
            "items": len(items),
            "failed": failed,
            "rejected": conditions.get("rejected", {}),
            "measurement": compare_math(keyed, fonts, size=data["size_pt"]),
            "fixture": compare_math_fixture(
                items, _optional_json(math_fixture), complete=not conditions.get("subset", False)
            ),
        }
    elif kind == "cases":
        items, failed = _math_items_of_cases(pdf, conditions["pages"])
        _write_json(output / "math_fixture.json", {"schema_version": SCHEMA_VERSION, "items": items})
        result["math"] = {
            "items": len(items),
            "failed": failed,
            "measurement": compare_math(items, fonts),
            "fixture": compare_math_fixture(items, _optional_json(math_fixture), complete=False),
        }
    _write_json(output / "comparison.json", result)
    return result


def _now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _optional_json(path: Path | None) -> dict | None:
    return _read_json(path) if path is not None and Path(path).is_file() else None


def passed(result: Mapping) -> bool:
    """今の計測が許容差に入り、計測できない項目がないか。"""
    layout = result.get("layout") or {}
    math_result = result.get("math") or {}
    measurement = math_result.get("measurement") or {}
    return (
        "error" not in layout
        and not layout.get("problems")
        and not math_result.get("failed")
        and measurement.get("passed", True) is not False
    )


def summary(result: Mapping, *, limit: int = 10) -> list[str]:
    """比べた結果の要約の行。"""
    lines = [f"実行 {result['run']}（資料 {result['deck']}）"]
    layout = result.get("layout") or {}
    if "error" in layout:
        lines.append(f"文字の配置: 計測できない: {layout['error']}")
    elif "skipped" in layout:
        lines.append(f"文字の配置: {layout['skipped']}")
    elif layout:
        lines.append(
            f"文字の配置: 要素 {layout['cases']} 個、許容差を超える違い {len(layout['problems'])} 件、"
            f"読めなかった要素 {len(layout['unread'])} 個"
        )
        if layout["font_mismatches"]:
            lines.append(f"  手元の書体が照合データと違う: {'、'.join(layout['font_mismatches'])}")
        worst = sorted(
            (row for row in layout["differences"] if "largest_pt" in row), key=lambda row: -row["largest_pt"]
        )
        for row in worst[:limit]:
            size = row["size_pt"]
            lines.append(
                f"  {row['key']}: 最大の差 {row['largest_pt']:.2f} pt（{row['largest_pt'] / size:.3f} em）、"
                f"基線 {row.get('baseline_pt')}、インク {row.get('ink_pt')}"
            )
        lines += [f"  {message}" for message in layout["problems"][:limit]]
        lines += _fixture_lines("文字の照合データ", layout["fixture"], limit)
    math_result = result.get("math")
    if math_result:
        lines.append(f"数式: 項目 {math_result['items']} 件、計測できない項目 {len(math_result['failed'])} 件")
        lines += [f"  計測できない: {name}" for name in math_result["failed"][:limit]]
        measurement = math_result["measurement"]
        for group in measurement.get("groups", []):
            if "error" in group:
                lines.append(f"  {group['size_pt']} pt・{group['math']}: 計測できない: {group['error']}")
                continue
            lines.append(
                f"  {group['size_pt']} pt・{group['math']}: 幅の差の最大 {group['width_max_em']:.4f} em、"
                f"上下の差の最大 {group['vertical_max_em']:.4f} em、{'合格' if group['passed'] else '不合格'}"
            )
        lines += _fixture_lines("数式の照合データ", math_result["fixture"], limit)
    lines.append("今の計測は許容差に入る" if passed(result) else "今の計測に許容差を超える違いがある")
    return lines


def _fixture_lines(name: str, fixture: Mapping, limit: int) -> list[str]:
    if not fixture.get("available"):
        return [f"{name}: なし"]
    lines = [
        f"{name}との差: 同じ {fixture['same']} 件、変わった {len(fixture['changed'])} 件、"
        f"新しい {len(fixture['new'])} 件、照合データだけにある {len(fixture['missing'])} 件"
    ]
    if fixture.get("environment_changed"):
        lines.append("  PowerPoint・書体・パッケージの版が照合データと違う")
    for row in fixture["changed"][:limit]:
        values = {key: row[key] for key in row if key not in ("key", "new", "existing")}
        lines.append(f"  {row['key']}: {values}")
    return lines


# ---- 採用 ----


def _results(work: Path, run: str | None) -> Path:
    directory = work / "results"
    if run:
        path = Path(run)
        name = path.name if path.is_dir() else (f"{int(run):04d}" if run.isdigit() else run)
        target = directory / name
    else:
        found = sorted(
            (p for p in directory.iterdir() if p.name.isdigit()) if directory.is_dir() else [],
            key=lambda p: int(p.name),
        )
        if not found:
            raise CalibrationError(f"計測の結果がありません（先に run で計測する）: {directory}")
        target = found[-1]
    if not (target / "comparison.json").is_file():
        raise CalibrationError(f"計測の結果がありません: {target}")
    return target


def _math_key(key: str) -> list[tuple[str, str]]:
    mode, _, expression = key.partition(":")
    if mode in ("inline", "display") and expression:
        return [(expression, mode)]
    return [(key, "inline"), (key, "display")]


def adopt(work: Path, *, run: str | None = None, only=(), output: Path | None = None) -> tuple[Path, int]:
    """計測値を照合データに採用する。戻り値は（書いたファイル、採用した項目の数）。

    only を省略すると照合データを計測値で置き換える。指定すると既存の照合データのその項目だけを置き換える
    （計測の環境が照合データと同じときだけ。新しい項目は加えない）。
    """
    from slidekit import measure

    work = _check_work(work)
    conditions = _read_json(work / "conditions.json")
    kind = conditions["deck"]
    if kind == "cases":
        raise CalibrationError(
            '"cases" の資料は照合データに採用しない。校正用の資料（tests/data/layout/deck）か数式の道具の式の一覧に'
            '加え、deck = "layout" か "math" で作り直す'
        )
    results = _results(work, run)
    compared = _read_json(results / "comparison.json")
    if kind == "layout":
        candidate = _read_json(results / "layout_fixture.json")
        if candidate.get("measure_version") != measure.MEASURE_VERSION:
            raise CalibrationError(
                f"計測値の MEASURE_VERSION {candidate.get('measure_version')} が今の {measure.MEASURE_VERSION} と違う。"
                "--measure-only で計測し直す"
            )
        layout_tool().check_public(candidate, extra=(work, Path.home()))
        output = Path(output or LAYOUT_FIXTURE)
        if only:
            existing = _read_json(output)
            for key in ("environment", "measure_version", "schema_version"):
                if existing.get(key) != candidate.get(key):
                    raise CalibrationError(f"--only は計測の {key} が照合データと同じときだけ使う（全体を採用する）")
            replacements = {_key(case): case for case in candidate["cases"] if _key(case) in set(only)}
            present = {_key(case) for case in existing["cases"]}
            unknown = sorted(set(only) - (set(replacements) & present))
            if unknown:
                raise CalibrationError(f"計測値と照合データの両方にない要素: {unknown}")
            existing["cases"] = [replacements.get(_key(case), case) for case in existing["cases"]]
            candidate, count = existing, len(replacements)
        else:
            count = len(candidate["cases"])
    else:
        if (compared.get("math") or {}).get("failed"):
            raise CalibrationError("計測できない式がある。枠を直して計測し直す（採用しない）")
        candidate = _read_json(results / "math_fixture.json")
        output = Path(output or MATH_FIXTURE)
        if conditions.get("subset") and not only:
            raise CalibrationError("式の一覧の一部（expressions）の計測値は、--only で既存の項目だけを置き換える")
        if only:
            existing = _read_json(output)
            for key in ("font", "powerpoint", "size_pt", "text_fonts", "schema_version"):
                if existing.get(key) != candidate.get(key):
                    raise CalibrationError(f"--only は計測の {key} が照合データと同じときだけ使う（全体を採用する）")
            wanted = {pair for key in only for pair in _math_key(key)}
            replacements = {
                (item["expression"], item["mode"]): item
                for item in candidate["items"]
                if (item["expression"], item["mode"]) in wanted
            }
            present = {(item["expression"], item["mode"]) for item in existing["items"]}
            unknown = sorted(wanted - (set(replacements) & present))
            if unknown:
                raise CalibrationError(f"計測値と照合データの両方にない式: {unknown}")
            existing["items"] = [
                replacements.get((item["expression"], item["mode"]), item) for item in existing["items"]
            ]
            candidate, count = existing, len(replacements)
        else:
            count = len(candidate["items"])
    _write_json(output, candidate)
    return output, count


# ---- コマンド ----


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="照合用の資料を作って PowerPoint で変換し、計測と照合データと比べる（開発用）"
    )
    commands = parser.add_subparsers(dest="command", required=True)
    child = commands.add_parser("run", help="資料を作って変換し、計測して比べる（--measure-only なら変換しない）")
    child.add_argument("config", nargs="?", type=Path, help="設定ファイル（TOML。--measure-only では書かない）")
    child.add_argument("--work", type=Path, required=True, help="作業ディレクトリ（slidekit の外の新しいディレクトリ）")
    child.add_argument(
        "--measure-only", action="store_true", help="変換せず、作業ディレクトリの実行の PDF を計測し直す"
    )
    child.add_argument("--run", help="計測する実行（番号かパス。既定は PDF のある最新の実行）")
    child.add_argument("--layout-fixture", type=Path, default=LAYOUT_FIXTURE, help="比べる文字の照合データ")
    child.add_argument("--math-fixture", type=Path, default=MATH_FIXTURE, help="比べる数式の照合データ")
    child = commands.add_parser("adopt", help="計測値を照合データに採用する")
    child.add_argument("--work", type=Path, required=True, help="作業ディレクトリ")
    child.add_argument("--run", help="採用する計測の実行（番号。既定は最新の計測）")
    child.add_argument("--only", nargs="+", default=(), metavar="KEY", help="置き換える項目（ページ/要素、mode:式）")
    child.add_argument("--output", type=Path, help="書き出す照合データ（既定は tests/data の照合データ）")
    arguments = parser.parse_args(argv)
    try:
        if arguments.command == "adopt":
            path, count = adopt(arguments.work, run=arguments.run, only=arguments.only, output=arguments.output)
            print(f"採用した照合データ: {path}（{count} 件）")
            return 0
        work = _check_work(arguments.work)
        if arguments.measure_only:
            if arguments.config is not None:
                parser.error("--measure-only では設定ファイルを書かない（作業ディレクトリの写しを使う）")
            run = find_run(work, arguments.run)
        else:
            if arguments.run:
                parser.error("--run は --measure-only と使う")
            if arguments.config is not None:
                prepare(arguments.config, work)
            elif not (work / "conditions.json").is_file():
                raise CalibrationError(f"作業ディレクトリに生成条件がありません（設定ファイルを指定する）: {work}")
            run = convert(work)
        result = measure_run(work, run, layout_fixture=arguments.layout_fixture, math_fixture=arguments.math_fixture)
    except CalibrationError as error:
        print(f"照合できません: {error}", file=sys.stderr)
        return 2
    print("\n".join(summary(result)))
    print(f"結果: {work / 'results' / result['run']}")
    return 0 if passed(result) else 1


if __name__ == "__main__":
    sys.exit(main())
