"""欠落、変形、確認記録の間違った引き継ぎにつながる不良を検査する。"""

import ast
from copy import deepcopy
from datetime import date
import json
from types import SimpleNamespace

import pytest

from fake_fonts import FakeFonts
from slidekit.diagnostics import DIAGNOSTICS
from slidekit.math.syntax import parse
from slidekit.validators import (
    apply_declarations,
    make_diagnostic,
    merge_diagnostics,
    skip_role_checks,
    validate_scene,
)
from test_diagnostics import REPORTERS, SOURCE, _strings


@pytest.fixture
def project(tmp_path):
    return SimpleNamespace(
        root=tmp_path,
        config={
            "metadata": {"title": "Example", "author": "", "last_modified_by": ""},
            "inputs": {"documents": [{"id": "paper"}]},
            "_bibliography_keys": [],
            "math": {"fallback": "off"},
        },
        theme={"roles": {"body": {"size_pt": 20, "min_size_pt": 18}, "caption": {"size_pt": 20, "min_size_pt": 12}}},
        rules={},
        environment={},
    )


def element(id="body", *, area=(20, 20, 180, 40), text="Sample", **extra):
    return {
        "id": id,
        "kind": "text",
        "area": list(area),
        "role": "body",
        "style": {
            "size_pt": 20,
            "min_size_pt": 18,
            "font_latin": "Arial",
            "font_japanese": "Arial",
            "font_math": "Cambria Math",
            "padding_pt": 0,
            "line_spacing": 1.0,
        },
        "content": [{"kind": "text", "text": text}],
        **extra,
    }


def frame(id="frame", *, area=(20, 20, 300, 200), text="", fill="FFFFFF"):
    box = element(id, kind="box", area=area, text=text, shape="rectangle")
    box["style"].update(fill=fill, line_color="000000")
    return box


def arrow(id, start, end, *, kind="arrow", **extra):
    area = [min(start[0], end[0]), min(start[1], end[1]), abs(end[0] - start[0]), abs(end[1] - start[1])]
    line = element(id, kind=kind, area=area, text="", start=list(start), end=list(end), **extra)
    line["style"]["line_width_pt"] = 1.0
    return line


def scene(*elements, notes="説明します。"):
    return {
        "schema_version": 1,
        "size": [720, 540],
        "metadata": {"title": "Example", "author": "", "last_modified_by": ""},
        "slides": [{"id": "first", "kind": "content", "notes": notes, "elements": list(elements)}],
    }


def validate(data, project, fonts=None, **options):
    return validate_scene(data, project, FakeFonts() if fonts is None else fonts, **options)


def codes(issues):
    return {issue["code"] for issue in issues}


def issue_of(issues, code):
    return next(issue for issue in issues if issue["code"] == code)


def test_diagnostic_identity_is_stable_but_changed_evidence_invalidates_acceptance():
    first = make_diagnostic("text_overflow", "幅不足", slide_id="first", element_id="body", evidence={"text": "before"})
    second = make_diagnostic("text_overflow", "幅不足", slide_id="first", element_id="body", evidence={"text": "after"})
    assert first["diagnostic_id"] == second["diagnostic_id"]
    assert first["evidence_hash"] != second["evidence_hash"]
    other_scope = make_diagnostic(
        "text_overflow", "幅不足", slide_id="first", element_id="body", scope="pdf:powerpoint:standard"
    )
    assert first["diagnostic_id"] != other_scope["diagnostic_id"]


def test_identifiers_match_values_before_split_and_evidence_hashes_are_pinned(project):
    # 識別子の期待値はハッシュ値の計算を files.py にまとめる前のコードで同じ入力から求めた（識別子は変えない）。
    # 証拠のハッシュ値は証拠の形式を診断の種類ごとの範囲にした後の値で固定する。非有限の数と JSON にない値も含める。
    # 場面の検査の text_overflow は、観測値に計測の結果が入る（計測の校正で変わる）ので、識別子だけを固定する。
    single = make_diagnostic(
        "text_overflow",
        "幅不足",
        slide_id="first",
        element_id="body",
        slide_number=1,
        observations={"required_width_pt": 30.5, "名前": "日本語", "values": (1, 2)},
        evidence={"text": "前", "date": date(2026, 1, 2), "limit": float("inf")},
    )
    pair = make_diagnostic(
        "element_overlap",
        "重なり",
        severity="warning",
        element_id="b",
        related_element_id="a",
        unordered=True,
        observations=[{"area": [1, 2.5, 3, 4]}],
        evidence={"value": float("nan")},
    )
    other = make_diagnostic(
        "element_overlap",
        "重なりの別の観測",
        severity="warning",
        element_id="a",
        related_element_id="b",
        unordered=True,
        observations=[{"area": [0, 0, 1, 1]}],
    )
    issues = [single, *merge_diagnostics([pair, other])]
    issues += validate(
        scene(element(area=(20, 20, 25, 10), text="long text"), element("cover", area=(30, 20, 100, 40), text="x")),
        project,
    )
    assert [[issue["code"], issue["severity"], issue["diagnostic_id"]] for issue in issues] == [
        ["text_overflow", "error", "da5630cbca7b7f0c46237341a"],
        ["element_overlap", "warning", "db23748a8af5e61d3db0a24f4"],
        ["element_overlap", "warning", "d9b09559a4913bc3f8622c234"],
        ["text_overflow", "error", "da5630cbca7b7f0c46237341a"],
    ]
    assert [issue["evidence_hash"] for issue in issues[:3]] == [
        "8b7036e402cb97302b976d4bf6986c5f605e6b4b71a8c54010a1cc9300fe1eb6",
        "817ad37fe79e84b2901eb9b3604af9c588bd3d07ec416816cc7fc7bd3830585d",
        "7fe62d0501257b0e2b398b50f9052a7d3719d18ffae4f7ab45d479ae82439ab0",
    ]
    assert len(issues[3]["evidence_hash"]) == 64 and issues[3]["evidence_hash"] != issues[0]["evidence_hash"]


def test_unordered_pairs_and_observation_merging_are_order_independent():
    first = make_diagnostic(
        "overlap", "重なり", element_id="a", related_element_id="b", unordered=True, observations={"a": 1}
    )
    second = make_diagnostic(
        "overlap", "重なり", element_id="b", related_element_id="a", unordered=True, observations={"a": 2}
    )
    assert first["diagnostic_id"] == second["diagnostic_id"]
    assert merge_diagnostics([first, second]) == merge_diagnostics([second, first])
    assert len(merge_diagnostics([first, second])[0]["observations"]) == 2


def test_make_diagnostic_uses_registry_severity():
    assert make_diagnostic("text_overflow", "はみ出し")["severity"] == "error"
    assert make_diagnostic("element_overlap", "重なり")["severity"] == "warning"
    assert make_diagnostic("text_wrap_near_limit", "限界に近い")["severity"] == "info"
    assert make_diagnostic("element_overlap", "重なり", severity="warning")["severity"] == "warning"
    # 表にない識別子は、重大度を省くとエラーにし、渡した重大度はそのまま使う。
    assert make_diagnostic("unregistered_code", "文言")["severity"] == "error"
    assert make_diagnostic("unregistered_code", "文言", severity="info")["severity"] == "info"
    with pytest.raises(ValueError, match="element_overlap"):
        make_diagnostic("element_overlap", "重なり", severity="error")
    with pytest.raises(ValueError, match="text_overflow"):
        make_diagnostic("text_overflow", "はみ出し", severity="warning")
    with pytest.raises(ValueError, match="重大度"):
        make_diagnostic("unregistered_code", "文言", severity="fatal")


# ---- 報告の呼び出しの重大度（AST で調べる） ----

_REQUIRED = object()  # 必須の severity 引数


def _parameters(function):
    arguments = function.args
    return arguments.posonlyargs + arguments.args + arguments.kwonlyargs


def _implied_severity(function):
    """severity を省いて呼んだときの重大度。定数なら値、表の値なら None を返す。"""
    arguments = function.args
    positional = arguments.posonlyargs + arguments.args
    defaults = dict(
        zip([argument.arg for argument in positional[len(positional) - len(arguments.defaults) :]], arguments.defaults)
    )
    defaults.update({argument.arg: default for argument, default in zip(arguments.kwonlyargs, arguments.kw_defaults)})
    if any(argument.arg == "severity" for argument in _parameters(function)):
        default = defaults.get("severity")
        if default is None:
            return _REQUIRED
        return default.value if isinstance(default, ast.Constant) else _REQUIRED
    # severity を受けず、作る診断に決まった重大度を書く関数（records._issue の辞書、定数の severity を渡す包む関数）。
    for node in ast.walk(function):
        if isinstance(node, ast.Dict):
            for key, value in zip(node.keys, node.values):
                if key is not None and _strings(key) == ["severity"] and isinstance(value, ast.Constant):
                    return value.value
        if isinstance(node, ast.Call):
            for keyword in node.keywords:
                if (
                    keyword.arg == "severity"
                    and isinstance(keyword.value, ast.Constant)
                    and isinstance(keyword.value.value, str)
                ):
                    return keyword.value.value
    return None


def _module_reporters(tree):
    """モジュールの最上位で定義した、診断を作る関数の重大度（ほかのモジュールから読み込んで呼ぶもの）。"""
    return {
        node.name: _implied_severity(node)
        for node in tree.body
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name in REPORTERS
    }


_FUNCTIONS = (ast.FunctionDef, ast.AsyncFunctionDef)


def _reporter_name(node):
    """診断を作る関数の呼び出しなら、その関数の名前。属性の report（進捗の表示など）は除く。"""
    function = node.func
    if isinstance(function, ast.Name):
        name = function.id
    elif isinstance(function, ast.Attribute) and function.attr != "report":
        name = function.attr
    else:
        return None
    return name if name in REPORTERS else None


def _own_reporters(scope):
    """関数（またはモジュール）の中で直接定義した、診断を作る関数。入れ子の関数の中の定義は含めない。"""
    found, pending = {}, list(ast.iter_child_nodes(scope))
    while pending:
        node = pending.pop()
        if isinstance(node, _FUNCTIONS):
            if node.name in REPORTERS:
                found[node.name] = _implied_severity(node)
            continue
        pending.extend(ast.iter_child_nodes(node))
    return found


def severity_problems(tree, registry, *, reporters=None, where="<sample>", inspected=None):
    """報告の呼び出しの重大度が表と一致しない箇所を返す。

    severity には定数を渡すか、省いて表の値を使う。ただし、診断を作る関数を包む関数の中で、識別子を変数で受けて
    severity=severity をそのまま渡す呼び出しは除く。severity を省いた呼び出しは、呼ぶ関数の既定値（定数なら）と表を比べる。
    呼ぶ関数は、内側の関数から外側へ、その中で定義した関数を探し、なければ reporters（ほかのモジュールの定義）を使う。
    {"code": 定数, "severity": 定数} の辞書（計測の判定結果など）も表と比べる。
    inspected に一覧を渡すと、調べた呼び出しの位置（where、行、列）を加える（走査の取りこぼしの確認に使う）。
    """
    problems = []

    def check(node, found_codes, severity, how):
        for code in found_codes:
            expected = registry.get(code)
            if expected is not None and expected != severity:
                problems.append(
                    f"{where}:{node.lineno}: {code} の重大度は表では {expected} だが、{how}は {severity} になる"
                )

    def call(node, owner, visible):
        name = _reporter_name(node)
        if name is None:
            return
        if inspected is not None:
            inspected.append((where, node.lineno, node.col_offset))
        keywords = {keyword.arg: keyword.value for keyword in node.keywords if keyword.arg is not None}
        found_codes = _strings(node.args[0]) if node.args else _strings(keywords["code"]) if "code" in keywords else []
        severity = keywords.get("severity")
        if severity is None or (isinstance(severity, ast.Constant) and severity.value is None):
            default = visible.get(name)
            if isinstance(default, str):
                check(node, found_codes, default, f"severity を省いた {name}() の既定値")
        elif isinstance(severity, ast.Constant) and isinstance(severity.value, str):
            # 識別子が変数なら表と比べることができない（make_diagnostic が実行時に比べる）。包む関数の既定値は呼び出し側で比べる。
            check(node, found_codes, severity.value, "渡した重大度")
        elif not (
            isinstance(severity, ast.Name)
            and severity.id == "severity"
            and not found_codes
            and owner is not None
            and any(argument.arg == "severity" for argument in _parameters(owner))
        ):
            problems.append(
                f"{where}:{node.lineno}: severity に定数でない値を渡している（定数を渡すか、省いて表の値を使う）"
            )

    def visit(scope, owner, visible):
        visible = {**visible, **_own_reporters(scope)}
        pending = list(ast.iter_child_nodes(scope))
        while pending:
            node = pending.pop()
            if isinstance(node, _FUNCTIONS):
                visit(node, node, visible)
                continue
            if isinstance(node, ast.Call):
                call(node, owner, visible)
            elif isinstance(node, ast.Dict):
                entries = {
                    key.value: value
                    for key, value in zip(node.keys, node.values)
                    if isinstance(key, ast.Constant) and isinstance(key.value, str)
                }
                if "code" in entries and isinstance(entries.get("severity"), ast.Constant):
                    check(node, _strings(entries["code"]), entries["severity"].value, "辞書の重大度")
            pending.extend(ast.iter_child_nodes(node))

    visit(tree, None, dict(reporters or {}))
    return problems


def test_severity_scan_detects_mismatches_and_non_constant_values():
    sample = """
def make_diagnostic(code, message, *, severity=None):
    return {}

def _issue(code, message):
    return {"code": code, "severity": "error", "message": message}

def check(hidden, code, progress):
    def report(code, message, *, severity="error"):
        make_diagnostic(code, message, severity=severity)
    report("text_overflow", "既定の error と表が一致する")
    report("element_overlap", "既定の error と表の warning が違う")
    report("element_overlap", "表と一致する", severity="warning")
    report("text_overflow", "表と違う", severity="warning")
    report("text_occluded" if hidden else "element_overlap", "定数でない", severity="error" if hidden else "warning")
    make_diagnostic("text_wrap_near_limit", "表の値になる")
    make_diagnostic(code, "識別子が変数なら実行時に比べる", severity="error")
    make_diagnostic(code, "変数を渡す", severity=hidden)
    progress.report("進捗の文言", severity="info")
    _issue("warning_unaccepted", "error と表が一致する")
    _issue("element_overlap", "error と表の warning が違う")
    return [{"code": "text_padding_overlap", "severity": "error"}, {"code": "text_overflow", "severity": "error"}]

def check_package():
    def report(code, message):
        make_diagnostic(code, message, severity="error")
    report("package_duplicate", "包む関数の error と表が一致する")
    report("fallback_crop_clipped", "包む関数の error と表の warning が違う")
"""
    registry = {code: spec.severity for code, spec in DIAGNOSTICS.items()}
    inspected = []
    problems = severity_problems(ast.parse(sample), registry, inspected=inspected)
    assert sorted(int(problem.split(":")[1]) for problem in problems) == [12, 14, 15, 18, 21, 22, 28], problems
    # 進捗の progress.report（19 行）を除く 14 件の呼び出しを調べる。
    assert len(inspected) == 14 and 19 not in {line for _where, line, _column in inspected}


def test_call_site_severity_matches_registry():
    registry = {code: spec.severity for code, spec in DIAGNOSTICS.items()}
    trees = {
        path: ast.parse(path.read_text(encoding="utf-8"), filename=str(path)) for path in sorted(SOURCE.rglob("*.py"))
    }
    reporters = {}
    for tree in trees.values():
        reporters.update(_module_reporters(tree))
    inspected = []
    problems = [
        problem
        for path, tree in trees.items()
        for problem in severity_problems(
            tree, registry, reporters=reporters, where=str(path.relative_to(SOURCE.parent)), inspected=inspected
        )
    ]
    # 集め方がおかしくなって何も見つからないときに、試験が確かめないまま通らないようにする。ほかのモジュールの書き方には依らない次の条件で確かめる。
    # 表の重大度を使う make_diagnostic が集まり、走査が診断を作る関数の呼び出しを、別の方法（ast.walk）で集めた呼び出しと同じだけ調べ、
    # 場面の検査の報告の呼び出しが見つかる。
    assert reporters.get("make_diagnostic", _REQUIRED) is None
    calls = [
        (str(path.relative_to(SOURCE.parent)), node.lineno, node.col_offset)
        for path, tree in trees.items()
        for node in ast.walk(tree)
        if isinstance(node, ast.Call) and _reporter_name(node) is not None
    ]
    assert sorted(inspected) == sorted(calls)
    assert sum(where == "slidekit/validators.py" for where, _line, _column in calls) >= 15
    assert not problems, "報告の呼び出しの重大度が診断の表と一致しない:\n" + "\n".join(problems)


# ---- 文字の計測とはみ出し ----


def test_overflow_reports_both_dimensions_and_does_not_shrink(project):
    original = scene(element(area=(20, 20, 25, 10), text="long text"))
    snapshot = deepcopy(original)
    issues = validate(original, project)
    assert original == snapshot
    overflow = issue_of(issues, "text_overflow")
    assert overflow["observations"][0]["required_width_pt"] > 25
    assert overflow["observations"][0]["required_height_pt"] > 10


def test_explicit_lines_and_mixed_fragments_share_line_width(project):
    text = element(area=(20, 20, 60, 50), text="ABC\nDEF")
    assert "text_overflow" not in codes(validate(scene(text), project))
    text["content"] = [{"kind": "text", "text": "ABC"}, {"kind": "math", "text": "1234"}]
    assert "text_overflow" in codes(validate(scene(text), project))


def test_wrapped_lines_raise_height_overflow(project):
    # 1 文字 10 pt。幅 100 pt の枠で「abc def」と「ghi jkl」の 2 行に折り返し、高さ 30 pt を超える。
    wrapped = element(area=(20, 20, 100, 30), text="abc def ghi jkl")
    wrapped["style"]["word_wrap"] = True
    overflow = issue_of(validate(scene(wrapped), project), "text_overflow")["observations"][0]
    assert overflow["lines"] == 2 and overflow["line_widths_pt"] == [70.0, 70.0] and overflow["word_wrap"] is True
    assert overflow["overflow_pt"]["side"] == "bottom"
    wrapped["area"] = [20, 20, 100, 50]
    assert "text_overflow" not in codes(validate(scene(wrapped), project))
    # 折り返さない要素は、同じ枠で幅を超える。
    wrapped["style"]["word_wrap"] = False
    overflow = issue_of(validate(scene(wrapped), project), "text_overflow")["observations"][0]
    assert overflow["lines"] == 1 and overflow["overflow_pt"]["side"] == "right"


def test_ink_based_overflow_ignores_tight_cells(project):
    # 数式だけのセル：行の箱（24 pt）は枠（18 pt）より高いが、インク（16 pt）は枠に収まる。
    cell = frame("cell", area=(20, 20, 40, 18))
    cell["content"] = [{"kind": "math", "text": "x"}]
    cell["style"].update(align="center", valign="middle")
    issues = validate(scene(cell), project)
    assert "text_overflow" not in codes(issues)
    # インクが枠を超えれば、はみ出しになる（線のある枠の許容は 0.5 pt）。
    cell["area"] = [20, 20, 40, 14]
    overflow = issue_of(validate(scene(cell), project), "text_overflow")["observations"][0]
    assert overflow["required_height_pt"] > overflow["available_height_pt"]


def test_padding_overlap_is_a_warning(project):
    # 余白 10 pt の枠で、インクの下端が下の余白に 10 pt 入り込む（枠からははみ出さない）。
    box = element(area=(20, 20, 200, 30))
    box["style"]["padding_pt"] = 10
    issues = validate(scene(box), project)
    assert "text_overflow" not in codes(issues)
    overlap = issue_of(issues, "text_padding_overlap")
    assert overlap["severity"] == "warning" and overlap["scope"] == "layout"
    assert overlap["observations"][0]["overflow_pt"] == {"side": "bottom", "amount": 10.0}
    box["style"]["padding_pt"] = 2
    assert "text_padding_overlap" not in codes(validate(scene(box), project))


def test_side_padding_changes_available_width(project):
    text = element(area=(20, 20, 100, 40), text="Sample")
    text["style"]["padding_pt"] = {"left": 30, "top": 0, "right": 0, "bottom": 0}
    issues = validate(scene(text), project)
    assert "text_overflow" not in codes(issues)
    text["style"]["padding_pt"] = {"left": 50, "top": 0, "right": 0, "bottom": 0}
    overflow = issue_of(validate(scene(text), project), "text_overflow")["observations"][0]
    assert overflow["available_width_pt"] == 50.0 and overflow["overflow_pt"] == {"side": "right", "amount": 10.0}


def test_near_limit_wrap_is_information(project):
    # 「abc」の 3 文字（30 pt）が内側の幅 30.1 pt に収まるが、誤差の帯の中にある。
    text = element(area=(20, 20, 30.1, 40), text="abc")
    text["style"]["word_wrap"] = True
    issues = validate(scene(text), project)
    near = issue_of(issues, "text_wrap_near_limit")
    assert near["severity"] == "info" and near["observations"][0]["near_limit_lines"] == [1]
    assert "text_overflow" not in codes(issues)


def test_layout_bounds_overflow(project):
    stacked = element(area=(20, 40, 180, 40), layout_bounds=[20, 20, 200, 50])
    issue = issue_of(validate(scene(stacked), project), "layout_bounds_overflow")
    assert issue["severity"] == "error"
    assert issue["observations"] == [
        {
            "area": [20.0, 40.0, 180.0, 40.0],
            "bounds": [20.0, 20.0, 200.0, 50.0],
            "overflow_pt": {"side": "bottom", "amount": 10.0},
            "stack_overflow_pt": 10.0,
            "stack_elements": [{"id": stacked["id"], "height_pt": 40.0}],
        }
    ]
    # 同じ積み上げの領域の要素をまとめ、積み上げの全体で足りない量と各要素の高さを示す。
    first = element("lead", area=(20, 20, 180, 30), layout_bounds=[20, 20, 200, 50])
    second = element("body", area=(20, 56, 180, 30), layout_bounds=[20, 20, 200, 50])
    issue = issue_of(validate(scene(first, second), project), "layout_bounds_overflow")
    (observed,) = issue["observations"]
    assert observed["stack_overflow_pt"] == 16.0
    assert observed["stack_elements"] == [{"id": "lead", "height_pt": 30.0}, {"id": "body", "height_pt": 30.0}]
    stacked["layout_bounds"] = [20, 20, 200, 100]
    assert "layout_bounds_overflow" not in codes(validate(scene(stacked), project))


def test_glyph_missing_only_when_both_fonts_lack_it(project):
    text = element(text="注※")
    text["style"]["font_japanese"] = "Yu Gothic UI"
    # 欧文の書体にない記号は、和文の書体にあれば計測できる。
    assert "font_glyph_missing" not in codes(validate(scene(text), project, FakeFonts(missing={"Arial": {"※"}})))
    issues = validate(scene(text), project, FakeFonts(missing={"Arial": {"※"}, "Yu Gothic UI": {"※"}}))
    issue = issue_of(issues, "font_glyph_missing")
    assert issue["scope"] == "input"
    assert issue["observations"][0]["characters"] == ["※"]
    assert issue["observations"][0]["fonts"] == ["Arial", "Yu Gothic UI"]


def test_text_command_contents_are_checked_for_glyphs(project):
    formula = element("formula", kind="equation", text="")
    formula["style"]["font_japanese"] = "Yu Gothic UI"
    formula["content"] = [{"kind": "math", "text": r"x + \text{注※}"}]
    fonts = FakeFonts(missing={"Arial": {"※"}, "Yu Gothic UI": {"※"}})
    issue = issue_of(validate(scene(formula), project, fonts), "font_glyph_missing")
    assert issue["observations"] == [
        {"characters": ["※"], "fonts": ["Arial", "Yu Gothic UI"], "expression": r"x + \text{注※}"}
    ]
    formula["content"] = [{"kind": "math", "text": r"x + \text{注}"}]
    assert "font_glyph_missing" not in codes(validate(scene(formula), project, fonts))


def test_math_syntax_error_is_reported_with_position(project):
    formula = element("formula", kind="equation", text="")
    formula["content"] = [{"kind": "math", "text": r"\alpah + 1"}]
    issues = validate(scene(formula), project)
    issue = issue_of(issues, "math_syntax_error")
    assert issue["severity"] == "error" and issue["scope"] == "input"
    observed = issue["observations"][0]
    assert observed["expression"] == r"\alpah + 1" and observed["position"] == 0 and observed["token"] == r"\alpah"
    assert observed["reason"] == "未対応の数式命令" and r"\alpha" in observed["hint"]
    assert "text_measurement_failed" not in codes(issues)

    class ParsingFonts(FakeFonts):
        # 実フォントの計測と同じく、計測する前に式を解析する（記法の誤りは計測の誤りの原因になる）。
        def math_box(self, expr, size, bold=False, **kwargs):
            parse(expr)
            return super().math_box(expr, size, bold, **kwargs)

    issues = validate(scene(formula), project, ParsingFonts())
    assert [issue["code"] for issue in issues] == ["math_syntax_error"]
    assert len(issues[0]["observations"]) == 1


def test_text_dimensions_pass_math_mode(project):
    fonts = FakeFonts()
    formula = element("formula", kind="equation", text="")
    formula["content"] = [{"kind": "math", "text": "a+b"}]
    inline = element("inline", area=(20, 100, 300, 40))
    inline["style"].update(font_japanese="Yu Gothic UI", font_latin="Segoe UI")
    inline["content"] = [{"kind": "text", "text": "式 "}, {"kind": "math", "text": "c"}]
    validate(scene(formula, inline), project, fonts)
    calls = {call["expression"]: call for call in fonts.calls}
    assert calls["a+b"]["display"] is True and calls["c"]["display"] is False
    assert (calls["a+b"]["japanese"], calls["a+b"]["latin"], calls["a+b"]["family"]) == (
        "Arial",
        "Arial",
        "Cambria Math",
    )
    assert (calls["c"]["japanese"], calls["c"]["latin"]) == ("Yu Gothic UI", "Segoe UI")


def test_layouts_computed_by_the_caller_are_used(project, monkeypatch):
    from slidekit import measure

    data = scene(element(area=(20, 20, 25, 10), text="long text"))
    fonts = FakeFonts()
    layouts = measure.layout_scene(data, fonts)

    def fail(*args, **kwargs):
        raise AssertionError("渡した計測の結果を使わずに計算し直した")

    monkeypatch.setattr(measure, "layout_scene", fail)
    assert "text_overflow" in codes(validate_scene(data, project, fonts, layouts=layouts))


def test_math_transform_in_parent_and_unknown_math_are_errors(project):
    parent = element("diagram", kind="group", text="", rotation=90)
    formula = element(parent_id="diagram")
    formula["content"] = [{"kind": "math", "text": r"\unknown"}]
    assert {"math_transform", "text_measurement_failed"} <= codes(validate(scene(parent, formula), project))


def test_font_too_small_uses_the_role_minimum(project):
    # 役割の使い方（数式専用、使うページ）と最小文字サイズの変更は、作成 API が例外にする
    # （tests/test_model_config.py の test_custom_role_is_restricted_without_changing_body）。
    project.theme["roles"]["result_formula"] = {
        "min_size_pt": 16,
        "allowed_slide_ids": ["first"],
        "allowed_content": "math",
    }
    formula = element(role="result_formula")
    formula["style"].update(size_pt=16, min_size_pt=16)
    formula["content"] = [{"kind": "math", "text": "x"}]
    assert "font_too_small" not in codes(validate(scene(formula), project))
    formula["role"] = "body"
    issue = issue_of(validate(scene(formula), project), "font_too_small")
    assert issue["observations"] == [{"size_pt": 16, "min_size_pt": 18, "role": "body"}]


def test_font_too_small_checks_fragment_sizes(project):
    text = element(area=(20, 20, 400, 40))
    text["content"] = [
        {"kind": "text", "text": "通常"},
        {"kind": "text", "text": "注", "style": {"size_pt": 12}},
        {"kind": "math", "text": "x", "style": {"size_pt": 24}},
    ]
    issues = [issue for issue in validate(scene(text), project) if issue["code"] == "font_too_small"]
    assert len(issues) == 1
    assert issues[0]["observations"] == [{"fragment": 2, "size_pt": 12, "min_size_pt": 18, "role": "body"}]


def test_slide_number_uses_scene_number(project):
    data = scene(element(area=(20, 20, 25, 10), text="long text"))
    assert issue_of(validate(data, project), "text_overflow")["slide_number"] == 1
    # 一部のページだけの場面でも、元のページ番号を使う。
    data["slides"][0]["number"] = 7
    assert issue_of(validate(data, project), "text_overflow")["slide_number"] == 7


def test_hand_made_structure_is_not_checked(project):
    # 構造の誤りは作成 API が例外にする。手で作った場面の構造は検査せず、検査できない要素は飛ばす。
    project.rules = {"forbidden_characters": ["文"]}
    data = scene(
        element("same", text="本文"),
        element("same", area=(20, 300, 180, 40)),
        element("nan", area=(float("nan"), 0, 10, 10), text="文"),
        {"id": "unknown", "kind": "chart", "area": [0, 0, 10, 10]},
        element("orphan", area=(20, 400, 180, 40), parent_id="missing"),
        element("loop_a", kind="group", text="", parent_id="loop_b"),
        element("loop_b", kind="group", text="", parent_id="loop_a"),
        arrow("broken", (0, 0), (0, 0)),
        element("bad_points", kind="line", text="", area=(0, 0, 10, 10), start=["x", 0], end=[10, 10]),
        element("legacy", area=(400, 20, 200, 40), content="文字列の内容"),
        element("stranger", area=(400, 100, 200, 40), role="undefined_role"),
        arrow("dangling", (500, 300), (600, 300), end_ref={"element_id": "missing", "side": "left"}),
        element("wide", area=(400, 480, 20, 20), text="long text"),
        {
            **element("picture", kind="image", area=(400, 200, 100, 100), text=""),
            "path": "missing.png",
            "fit": "stretch",
            "image_size": [10, 20],
            "actual_area": [400, 200, 100, 30],
        },
    )
    data["size"] = [0, 540]
    issues = validate(data, project)
    removed = {
        "invalid_page_size",
        "invalid_slide_id",
        "invalid_element_id",
        "unknown_element",
        "parent_missing",
        "parent_cycle",
        "invalid_geometry",
        "connection_missing",
        "image_fit",
        "image_distortion",
        "image_overflow",
        "unknown_role",
        "role_usage",
        "minimum_override",
    }
    assert not codes(issues) & removed
    json.dumps(issues, allow_nan=False)
    # 検査できる要素は、同じ場面でも検査する。文字列の content は断片の列として読まない（旧い形式）。
    found = {(issue["code"], issue["element_id"]) for issue in issues}
    assert {("text_overflow", "wide"), ("asset_missing", "picture"), ("forbidden_character", "same")} <= found
    assert not {element_id for code, element_id in found if code == "forbidden_character"} & {"nan", "legacy"}


# ---- 記述規則 ----


def test_writing_rules_and_references(project):
    project.rules = {
        "require_notes": True,
        "forbidden_characters": ["。"],
        "notes_style": "polite",
        "math_candidates": {"enabled": True, "exclude_patterns": [r"2026年"], "operations": ["Access"]},
    }
    text = element(text="2026年に n 個を Access する。", sources=[{"document_id": "missing"}])
    text["area"] = [20, 20, 650, 100]
    issues = validate(scene(text, notes="これは説明である。"), project)
    assert {"forbidden_character", "math_not_marked", "source_missing", "notes_style"} <= codes(issues)
    candidates = next(issue for issue in issues if issue["code"] == "math_not_marked")["observations"][0]["candidates"]
    assert candidates == ["Access", "n"]
    assert "notes_missing" in codes(validate(scene(text, notes=""), project))


def test_sources_use_documents_and_bibliography_keys(project):
    project.config["_bibliography_keys"] = ["smith2020"]
    text = element(sources=[{"document_id": "paper"}, {"bib_key": "smith2020"}])
    assert "source_missing" not in codes(validate(scene(text), project))
    text["sources"] = [{"bib_key": "unknown2020"}]
    issue = issue_of(validate(scene(text), project), "source_missing")
    assert issue["observations"] == [{"source": {"bib_key": "unknown2020"}}]


def test_math_not_marked_skips_excluded_roles_and_page_kinds(project):
    project.rules = {
        "math_candidates": {"enabled": True, "exclude_roles": ["caption"], "exclude_page_kinds": ["references"]}
    }
    body = element("body", area=(20, 20, 300, 40), text="n 個の要素")
    caption = element("caption", area=(20, 100, 300, 40), text="n 個の要素", role="caption")
    listed = element("listed", area=(20, 20, 300, 40), text="n 個の要素")
    data = scene(body, caption)
    data["slides"].append({"id": "refs", "kind": "references", "notes": "", "elements": [listed]})
    flagged = {
        (issue["slide_id"], issue["element_id"])
        for issue in validate(data, project)
        if issue["code"] == "math_not_marked"
    }
    assert flagged == {("first", "body")}


def test_line_break_observations_name_line_and_character(project):
    start = element("start", area=(20, 20, 300, 60), text="前半\n、後半")
    end = element("end", area=(20, 100, 300, 60), text="本文「\n引用」です")
    issues = [issue for issue in validate(scene(start, end), project) if issue["code"] == "line_break"]
    observed = {issue["element_id"]: issue["observations"] for issue in issues}
    assert observed == {
        "start": [{"line": 2, "side": "start", "character": "、"}],
        "end": [{"line": 1, "side": "end", "character": "「"}],
    }
    assert all(issue["severity"] == "warning" and issue["scope"] == "writing" for issue in issues)


# ---- 重なり ----


def test_background_and_ancestor_are_excluded_but_siblings_are_checked(project):
    background = element("bg", kind="box", area=(0, 0, 720, 540), semantic="background", text="")
    group = element("group", kind="group", area=(0, 0, 500, 400), text="")
    first = element("a", parent_id="group")
    second = element("b", parent_id="group", area=(30, 20, 100, 40))
    issues = validate(scene(background, group, first, second), project)
    overlaps = [issue for issue in issues if issue["code"] == "element_overlap"]
    assert len(overlaps) == 1
    assert {overlaps[0]["element_id"], overlaps[0]["related_element_id"]} == {"a", "b"}


def test_line_segment_does_not_use_bounding_box_and_connection_is_exempt(project):
    line = element("arrow", kind="arrow", area=(0, 0, 100, 100), text="", start=[0, 0], end=[100, 100])
    away = element("text", area=(10, 70, 20, 20), text="x")
    assert "element_overlap" not in codes(validate(scene(line, away), project))
    away["area"] = [70, 70, 20, 20]
    assert "element_overlap" in codes(validate(scene(line, away), project))
    line["end_ref"] = {"element_id": "text", "side": "left"}
    assert "element_overlap" not in codes(validate(scene(line, away), project))


def test_opaque_later_shape_hiding_text_is_an_error(project):
    first = element()
    cover = element("cover", kind="box", area=(10, 10, 220, 80), text="")
    cover["style"]["fill"] = "FFFFFF"
    issue = next(issue for issue in validate(scene(first, cover), project) if issue["code"] == "text_occluded")
    assert issue["severity"] == "error"


def test_overlap_uses_drawn_area_of_contained_image(project):
    # 正方形の画像を横長の領域に contain で置くと、実際に描く範囲（actual_area）は中央の 300×300 pt で、左右に 50 pt の余白がある。
    picture = element(
        "picture",
        kind="image",
        area=(100, 100, 400, 300),
        text="",
        path="missing.png",
        fit="contain",
        image_size=[100, 100],
        actual_area=[150, 100, 300, 300],
    )
    label = element("label", area=(105, 150, 40, 30), text="ab")
    # 余白に置いた文字は、画像の前に作っても後に作っても重なりにならない。
    for data in (scene(label, picture), scene(picture, label)):
        assert not codes(validate(data, project)) & {"element_overlap", "text_occluded"}
    # 実際の画像に重なる文字は、従来どおり報告する。重なりの範囲も実際に描く範囲で求める。
    covered = element("covered", area=(200, 150, 40, 30), text="ab")
    assert "text_occluded" in codes(validate(scene(covered, picture), project))
    assert "element_overlap" in codes(validate(scene(picture, covered), project))
    edge = element("edge", area=(130, 150, 40, 30), text="ab")
    overlap = issue_of(validate(scene(edge, picture), project), "element_overlap")
    assert overlap["observations"] == [{"intersection": [150.0, 150.0, 20.0, 30.0]}]
    # ページの外へのはみ出しは、従来どおり配置した領域で判定する（実際に描く範囲 x 400〜700 はページに収まる）。
    picture.update(area=[300, 100, 500, 300], actual_area=[400, 100, 300, 300])
    assert issue_of(validate(scene(picture), project), "page_overflow")["observations"][0]["area"] == [
        300.0,
        100.0,
        500.0,
        300.0,
    ]


def test_empty_box_has_no_text_height_or_occlusion(project):
    empty = element(kind="box", area=[20, 20, 100, 2], text="")
    assert validate(scene(empty), project) == []


def test_label_inside_frame_is_not_overlap(project):
    label = element("label", area=(40, 40, 100, 40))
    assert "element_overlap" not in codes(validate(scene(frame(), label), project))
    same = element("same", area=(20, 20, 300, 200))
    assert "element_overlap" not in codes(validate(scene(frame(), same), project))
    # 枠からはみ出す要素は判定する。
    outside = element("outside", area=(280, 40, 100, 40))
    assert "element_overlap" in codes(validate(scene(frame(), outside), project))


def test_label_on_frame_text_is_still_overlap(project):
    # 枠の文章「Sample」のインクは x 20〜80、y 24〜40 にある。
    titled = frame(text="Sample", area=(20, 20, 300, 100))
    on_text = element("on_text", area=(30, 25, 40, 20), text="x")
    issues = validate(scene(titled, on_text), project)
    assert {
        (issue["element_id"], issue["related_element_id"]) for issue in issues if issue["code"] == "element_overlap"
    } == {("frame", "on_text")}
    away = element("away", area=(150, 60, 100, 30), text="x")
    assert "element_overlap" not in codes(validate(scene(titled, away), project))
    # 文章を計測できない枠は、文章の位置が分からないので除外しない。
    titled["style"]["size_pt"] = -1
    assert "element_overlap" in codes(validate(scene(titled, away), project))


def test_frame_drawn_later_still_hides_text(project):
    text = element("text", area=(40, 40, 100, 30))
    cover = frame("cover")
    hidden = issue_of(validate(scene(text, cover), project), "text_occluded")
    assert {hidden["element_id"], hidden["related_element_id"]} == {"text", "cover"}
    partial = frame("partial", area=(100, 30, 200, 20))
    assert "element_overlap" in codes(validate(scene(text, partial), project))
    # 後に描く透明な空の枠は、内側の文字を隠さない。
    outline = frame("outline", fill=None)
    assert not codes(validate(scene(text, outline), project)) & {"element_overlap", "text_occluded"}


def test_line_touching_shape_edge_is_connection(project):
    target = frame("target", area=(100, 100, 100, 50))
    for end in ((100, 125), (101, 125)):
        assert "element_overlap" not in codes(validate(scene(target, arrow("arrow", (20, 125), end)), project)), end
    for start, end in (((20, 125), (110, 125)), ((20, 125), (250, 125)), ((150, 125), (250, 125))):
        assert "element_overlap" in codes(validate(scene(target, arrow("arrow", start, end)), project)), (start, end)
    # 図形への入り込みは max(2 pt, 線幅) まで接触とみなす（ちょうどその長さも含む）。線幅 1 pt は 2 pt、線幅 4 pt は 4 pt まで。
    for width, intrusion, overlap in ((1.0, 2.0, False), (1.0, 3.0, True), (4.0, 4.0, False), (4.0, 5.0, True)):
        line = arrow("arrow", (20, 125), (100 + intrusion, 125))
        line["style"]["line_width_pt"] = width
        assert ("element_overlap" in codes(validate(scene(target, line), project))) is overlap, (width, intrusion)


def test_line_inside_frame_is_not_overlap(project):
    divider = arrow("divider", (40, 120), (300, 120), kind="line")
    assert "element_overlap" not in codes(validate(scene(frame(), divider), project))
    # 枠の文章（インクは y 24〜40）から離れた区切り線は除外し、文章を横切る線は判定する。
    titled = frame(text="Sample")
    assert "element_overlap" not in codes(validate(scene(titled, divider), project))
    through = arrow("through", (25, 32), (300, 32), kind="line")
    assert "element_overlap" in codes(validate(scene(titled, through), project))


def test_decoration_is_not_overlap(project):
    text = element("text")
    ornament = element("ornament", kind="box", area=(10, 10, 60, 60), text="", semantic="decoration")
    assert "element_overlap" not in codes(validate(scene(text, ornament), project))
    del ornament["semantic"]
    assert "element_overlap" in codes(validate(scene(text, ornament), project))


# ---- 証拠の範囲 ----


def _hashes(issues):
    return {issue["code"]: issue["evidence_hash"] for issue in issues}


def test_rule_change_only_invalidates_warnings_that_use_it(project):
    project.rules = {"math_candidates": {"enabled": True, "exclude_patterns": []}, "forbidden_characters": []}
    data = scene(element("a", text="n 個"), element("b", area=(30, 20, 100, 40), text="説明"))
    before = _hashes(validate(data, project))
    assert {"math_not_marked", "element_overlap"} <= set(before)
    project.rules["math_candidates"]["exclude_patterns"] = [r"\d{4}年"]
    changed = _hashes(validate(data, project))
    assert changed["math_not_marked"] != before["math_not_marked"]
    assert changed["element_overlap"] == before["element_overlap"]
    project.rules["math_candidates"]["exclude_patterns"] = []
    project.rules["forbidden_characters"] = ["※"]
    other = _hashes(validate(data, project))
    assert (
        other["math_not_marked"] == before["math_not_marked"] and other["element_overlap"] == before["element_overlap"]
    )


def test_font_inventory_only_in_measurement_evidence(project):
    data = scene(
        element("a", area=(20, 20, 25, 10), text="long text"), element("b", area=(30, 20, 100, 40), text="説明")
    )
    before = _hashes(validate(data, project, FakeFonts()))
    after = _hashes(validate(data, project, FakeFonts(roles={"latin": "Segoe UI"})))
    assert after["element_overlap"] == before["element_overlap"]
    assert after["text_overflow"] != before["text_overflow"]


def test_writing_warning_ignores_position_and_style(project):
    project.rules = {"math_candidates": {"enabled": True}}
    text = element("a", text="n 個")
    before = _hashes(validate(scene(text), project))["math_not_marked"]
    text["area"] = [100, 200, 300, 60]
    text["style"]["size_pt"] = 24
    assert _hashes(validate(scene(text), project))["math_not_marked"] == before
    text["content"] = [{"kind": "text", "text": "m 個"}]
    assert _hashes(validate(scene(text), project))["math_not_marked"] != before


def test_default_scope_keeps_full_evidence(project):
    padded = element("padded", area=(20, 100, 200, 30))
    padded["style"]["padding_pt"] = 10
    data = scene(element("a", area=(20, 20, 25, 10), text="long text"), padded)
    before = _hashes(validate(data, project))
    project.rules = {"forbidden_characters": ["※"]}
    after = _hashes(validate(data, project))
    # 範囲を決めていない種類（text_overflow）は記述規則のどの変更でも変わり、規則を使わない警告は変わらない。
    assert after["text_overflow"] != before["text_overflow"]
    assert after["text_padding_overlap"] == before["text_padding_overlap"]


def test_converted_image_warning_evidence_ignores_run_path_but_tracks_source_content(project, tmp_path):
    image = tmp_path / "source.svg"
    image.write_text("<svg/>")
    first, second = tmp_path / "0001.png", tmp_path / "0002.png"
    first.write_bytes(b"same pixels")
    second.write_bytes(b"same pixels")
    drawing = element(
        kind="image",
        text="",
        area=(600, 20, 80, 40),
        path=str(first),
        source_path=str(image),
        source_sha256="original",
        image_size=[100, 50],
        actual_area=[600, 20, 80, 40],
        vector_path=str(tmp_path / "runs" / "0001" / "assets" / "source.svg"),
    )
    data = scene(drawing)
    data["slides"][0]["content_area"] = [20, 20, 500, 400]

    def warning():
        return issue_of(validate(data, project), "content_area_overflow")["evidence_hash"]

    before = warning()
    drawing["path"] = str(second)
    assert warning() == before
    # 埋め込むベクター画像（素材の準備で書く vector_path）も、実行ごとに変わるパスである。
    drawing["vector_path"] = str(tmp_path / "runs" / "0002" / "assets" / "source.svg")
    assert warning() == before
    drawing["source_sha256"] = "changed"
    assert warning() != before


# ---- 資料コードでの許容の宣言 ----


def test_apply_declarations_matches_pairs_and_reports_unused():
    overlap = make_diagnostic(
        "element_overlap",
        "重なり",
        slide_id="first",
        slide_number=1,
        element_id="a",
        related_element_id="b",
        unordered=True,
    )
    other_overlap = make_diagnostic(
        "element_overlap",
        "重なり",
        slide_id="first",
        slide_number=1,
        element_id="a",
        related_element_id="c",
        unordered=True,
    )
    candidate = make_diagnostic(
        "math_not_marked", "候補", slide_id="first", slide_number=1, element_id="a", scope="writing"
    )
    notes = make_diagnostic("notes_style", "文末", slide_id="first", slide_number=1, scope="writing")
    elsewhere = make_diagnostic("notes_style", "文末", slide_id="second", slide_number=2, scope="writing")
    error = make_diagnostic("text_overflow", "はみ出し", slide_id="first", slide_number=1, element_id="a")
    issues = [overlap, other_overlap, candidate, notes, elsewhere, error]
    scene_data = {
        "slides": [
            {
                "id": "first",
                "number": 1,
                "allowed_warnings": [
                    {
                        "code": "element_overlap",
                        "element_id": "b",
                        "related_element_id": "a",
                        "reason": "組の宣言",
                        "location": "slides.py:10",
                    },
                    {"code": "math_not_marked", "reason": "ページ全体", "location": "slides.py:11"},
                    {
                        "code": "notes_style",
                        "element_id": None,
                        "related_element_id": None,
                        "reason": "常体の引用",
                        "location": "slides.py:12",
                    },
                    {
                        "code": "text_overflow",
                        "element_id": "a",
                        "reason": "エラーには使えない",
                        "location": "slides.py:13",
                    },
                    {"code": "line_break", "element_id": "missing", "reason": "古い宣言", "location": "slides.py:14"},
                ],
            },
            {"id": "second", "number": 2},
        ]
    }
    snapshot = deepcopy(issues)
    result = apply_declarations(issues, scene_data)
    assert issues == snapshot
    by_key = {
        (issue["code"], issue["slide_id"], issue["element_id"], issue["related_element_id"]): issue for issue in result
    }
    assert by_key[("element_overlap", "first", "a", "b")]["declared"] == {
        "reason": "組の宣言",
        "location": "slides.py:10",
    }
    assert "declared" not in by_key[("element_overlap", "first", "a", "c")]
    assert by_key[("math_not_marked", "first", "a", None)]["declared"]["reason"] == "ページ全体"
    assert by_key[("notes_style", "first", None, None)]["declared"]["reason"] == "常体の引用"
    assert "declared" not in by_key[("notes_style", "second", None, None)]
    assert "declared" not in by_key[("text_overflow", "first", "a", None)]
    # 宣言があっても診断は残し、元の識別子と証拠を変えない。
    assert [issue["diagnostic_id"] for issue in result[:6]] == [issue["diagnostic_id"] for issue in issues]
    assert [issue["evidence_hash"] for issue in result[:6]] == [issue["evidence_hash"] for issue in issues]
    unused = sorted(result[6:], key=lambda issue: issue["element_id"])
    assert [
        (issue["code"], issue["severity"], issue["scope"], issue["element_id"], issue["slide_number"])
        for issue in unused
    ] == [
        ("allow_warning_unused", "info", "declaration", "a", 1),
        ("allow_warning_unused", "info", "declaration", "missing", 1),
    ]
    assert [issue["observations"][0]["code"] for issue in unused] == ["text_overflow", "line_break"]
    assert apply_declarations(issues, {"slides": [{"id": "first"}]}) == issues


def test_role_skip_checks_remove_warnings_and_info_of_that_role():
    scene = {
        "theme": {"roles": {"references": {"skip_checks": ["math_not_marked", "text_wrap_mid_phrase"]}, "body": {}}},
        "slides": [
            {
                "id": "refs",
                "elements": [{"id": "list", "role": "references"}, {"id": "lead"}],
            }
        ],
    }
    issues = [
        {"code": "math_not_marked", "severity": "warning", "slide_id": "refs", "element_id": "list"},
        {"code": "text_wrap_mid_phrase", "severity": "info", "slide_id": "refs", "element_id": "list"},
        {"code": "math_not_marked", "severity": "warning", "slide_id": "refs", "element_id": "lead"},
        {"code": "text_overflow", "severity": "error", "slide_id": "refs", "element_id": "list"},
    ]
    kept = skip_role_checks(issues, scene)
    assert [(issue["code"], issue["element_id"]) for issue in kept] == [
        ("math_not_marked", "lead"),
        ("text_overflow", "list"),
    ]
    assert skip_role_checks(issues, {**scene, "theme": {"roles": {}}}) == issues


# ---- 吹き出し ----


def callout(id="note", *, area=(300, 100, 160, 60), side="left", at=0.5, text="補足", target=None):
    """しっぽのある box（作成 API と同じく outline.tail_outline から作る）。"""
    from slidekit.outline import tail_outline

    box = element(id, kind="box", area=area, text=text, shape="rectangle")
    box["style"].update(fill=None, line_color="000000")
    outline = tail_outline(area, "rectangle", side=side, at=at)
    box["actual_area"] = list(outline.actual_area)
    box["outside_parts"] = [{"part": name, "area": list(rect)} for name, rect in outline.outside]
    box["tail"] = {"side": side, "at": outline.at, "width_pt": 10.0, "length_pt": 8.0}
    if target is not None:
        box["tail"]["target"] = {"element_id": target, "side": "right"}
    return box


def overlaps(issues):
    """重なりの診断の (要素, 相手, [(部分を持つ要素, 部分)])。組の要素は識別子の順に並ぶ。"""
    return [
        (
            issue["element_id"],
            issue["related_element_id"],
            [(observation.get("part_element_id"), observation.get("part")) for observation in issue["observations"]],
        )
        for issue in issues
        if issue["code"] == "element_overlap"
    ]


def test_tail_outside_page_and_content_area_is_reported(project):
    data = scene(callout(area=(4, 100, 160, 60), side="left"))
    data["slides"][0]["content_area"] = [20, 20, 680, 500]
    issues = validate(data, project)
    page = issue_of(issues, "page_overflow")
    assert page["severity"] == "error" and page["observations"][0]["area"] == [-4, 100, 168, 60]
    assert issue_of(issues, "content_area_overflow")["observations"][0]["area"] == [-4, 100, 168, 60]
    # しっぽがページの内側なら、本体の位置だけでははみ出さない。
    assert "page_overflow" not in codes(validate(scene(callout(area=(8, 100, 160, 60))), project))
    # 積み上げの領域は本体（area）で判定する。
    inside = callout(area=(100, 100, 160, 60), side="bottom")
    inside["layout_bounds"] = [100, 100, 160, 60]
    assert "layout_bounds_overflow" not in codes(validate(scene(inside), project))


def test_tail_overlap_is_reported_with_part(project):
    # しっぽ（x 292〜300、y 125〜135）に左の文字の右端（x 296）が 4 pt 入る。本体どうしは重ならない。
    issues = validate(scene(element("body", area=(100, 110, 196, 40)), callout()), project)
    assert overlaps(issues) == [("body", "note", [("note", "tail")])]
    assert issue_of(issues, "element_overlap")["observations"][0]["intersection"] == [292, 125, 4, 10]


def test_tail_does_not_overlap_elements_beside_or_touching_it(project):
    # 本体の外で、しっぽにも触れない要素（しっぽの上下の空き）。全体の外接長方形では判定しない。
    beside = element("beside", area=(100, 100, 197, 20))
    # 先端が相手の辺にちょうど接する配置。
    touching = element("touching", area=(100, 125, 192, 20))
    assert overlaps(validate(scene(beside, touching, callout()), project)) == []


def test_tail_inside_a_large_frame_is_not_overlap(project):
    assert overlaps(validate(scene(frame(area=(200, 50, 400, 300)), callout()), project)) == []


def test_tail_crossing_lines_and_target(project):
    crossing = arrow("crossing", (296, 90), (296, 160), kind="line")
    passing = arrow("passing", (250, 90), (250, 160), kind="line")
    issues = validate(scene(crossing, passing, callout()), project)
    assert overlaps(issues) == [("crossing", "note", [("note", "tail")])]
    # しっぽで指した要素（tail_toward）は、しっぽの重なりの判定から外す。本体が重なれば警告のまま。
    target = element("target", area=(100, 110, 196, 40))
    assert overlaps(validate(scene(target, callout(target="target")), project)) == []
    body = element("target", area=(100, 110, 210, 40))
    assert overlaps(validate(scene(body, callout(target="target")), project)) == [("note", "target", [(None, None)])]


def test_two_tails_overlap(project):
    first = callout("first", area=(100, 100, 100, 40), side="right")
    second = callout("second", area=(214, 100, 100, 40), side="left")
    assert overlaps(validate(scene(first, second), project)) == [("first", "second", [("first", "tail")])]
