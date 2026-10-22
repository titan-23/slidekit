"""診断の表と、表から作る文書を確かめる。"""

import ast
from collections import defaultdict
from pathlib import Path
import sys

import pytest

from slidekit import diagnostics
from slidekit.diagnostics import (
    CATEGORIES,
    DIAGNOSTICS,
    END_MARKER,
    START_MARKER,
    WARNING_CODES,
    DiagnosticSpec,
    EvidenceScope,
    main,
    markdown,
)

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "src" / "slidekit"
DOCUMENT = ROOT / "docs" / "diagnostics.md"
# 診断を作る関数と、それを包む関数。第 1 引数が識別子である（引数が code・message で始まる関数。
# test_reporter_functions_are_scanned が、ソースのそのような関数がすべてここにあることを確かめる）。
REPORTERS = {"report", "make_diagnostic", "_issue", "element_report"}


def _strings(node):
    """定数の文字列と、条件式なら両方の枝の文字列。"""
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return [node.value]
    if isinstance(node, ast.IfExp):
        return _strings(node.body) + _strings(node.orelse)
    return []


def _is_code_target(node):
    return (isinstance(node, ast.Name) and node.id == "code") or (
        isinstance(node, ast.Attribute) and node.attr == "code"
    )


def codes_in_tree(tree):
    """ソースの構文木から、診断の識別子として使う文字列を集める。"""
    found = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Call):
            function = node.func
            # report は検査の中の局所関数だけを対象にする（属性の呼び出しは別の用途がありうる）。
            name = (
                function.id
                if isinstance(function, ast.Name)
                else (function.attr if isinstance(function, ast.Attribute) and function.attr != "report" else None)
            )
            if name in REPORTERS and node.args:
                found += [(code, node.lineno) for code in _strings(node.args[0])]
            found += [
                (code, node.lineno)
                for keyword in node.keywords
                if keyword.arg == "code"
                for code in _strings(keyword.value)
            ]
            if name == "getattr" and len(node.args) == 3 and _strings(node.args[1]) == ["code"]:
                found += [(code, node.lineno) for code in _strings(node.args[2])]
        elif isinstance(node, ast.Assign) and any(_is_code_target(target) for target in node.targets):
            found += [(code, node.lineno) for code in _strings(node.value)]
        elif isinstance(node, ast.AnnAssign) and _is_code_target(node.target) and node.value is not None:
            found += [(code, node.lineno) for code in _strings(node.value)]
        elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.Lambda)):
            # 例外の __init__ に限らず、識別子を既定値で出す補助関数（入れ子の関数を含む）の引数 code の既定値も集める。
            arguments = node.args
            positional = arguments.posonlyargs + arguments.args
            defaults = list(zip(positional[len(positional) - len(arguments.defaults) :], arguments.defaults))
            defaults += [
                (argument, default)
                for argument, default in zip(arguments.kwonlyargs, arguments.kw_defaults)
                if default is not None
            ]
            found += [
                (code, node.lineno)
                for argument, default in defaults
                if argument.arg == "code"
                for code in _strings(default)
            ]
        elif isinstance(node, ast.Dict):
            found += [
                (code, node.lineno)
                for key, value in zip(node.keys, node.values)
                if key is not None and _strings(key) == ["code"]
                for code in _strings(value)
            ]
    return found


def codes_in_source(root=SOURCE):
    """`src/slidekit` の全 `.py` から、識別子とその場所を集める。"""
    places = defaultdict(list)
    for path in sorted(root.rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for code, line in codes_in_tree(tree):
            places[code].append(f"{path.relative_to(root.parent)}:{line}")
    return dict(places)


def test_source_scan_collects_every_form_of_code():
    sample = """
class SampleError(ValueError):
    code = "class_default"
    note: str = "not_a_code"

    def __init__(self, message, *, code="keyword_default", **details):
        self.code = code

class OtherError(RuntimeError):
    def __init__(self, message, code="positional_default"):
        super().__init__(message)

class Builder:
    def _error(self, message, element_id=None, *, code="method_default", **details):
        return SampleError(message, code=code)

def _identifier(value, kind="識別子", *, code="keyword_helper_default"):
    return value

def accept(run, *, code=None):
    def failure(message, code="helper_default", **details):
        return SampleError(message, code=code)
    lambda message, code="lambda_default": failure(message, code)
    return failure

def check(hidden, error, progress):
    report("reported", "文言")
    report("hidden_branch" if hidden else "visible_branch", "文言", severity="error" if hidden else "warning")
    validators.make_diagnostic("attribute_call", "文言")
    _issue("record_issue", "文言")
    element_report("element_issue", "文言")
    progress.report("進捗の文言")
    raise SampleError("文言", code="keyword_argument")
    code = "assigned"
    code: str = "annotated"
    getattr(error, "code", "fallback")
    return {"code": "dictionary", "message": "文言"}
"""
    assert sorted(code for code, _ in codes_in_tree(ast.parse(sample))) == sorted(
        [
            "class_default",
            "keyword_default",
            "positional_default",
            "method_default",
            "keyword_helper_default",
            "helper_default",
            "lambda_default",
            "reported",
            "hidden_branch",
            "visible_branch",
            "attribute_call",
            "record_issue",
            "element_issue",
            "keyword_argument",
            "assigned",
            "annotated",
            "fallback",
            "dictionary",
        ]
    )


def reporter_functions(root=SOURCE):
    """`src/slidekit` で定義した、引数が code・message で始まる関数（入れ子の関数を含む）の名前と場所。"""
    places = defaultdict(list)
    for path in sorted(root.rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                names = [argument.arg for argument in node.args.posonlyargs + node.args.args]
                if names[:1] in (["self"], ["cls"]):
                    names = names[1:]
                if names[:2] == ["code", "message"]:
                    places[node.name].append(f"{path.relative_to(root.parent)}:{node.lineno}")
    return dict(places)


def test_reporter_functions_are_scanned():
    # 診断を包む関数が増えたのに REPORTERS にないと、その呼び出しの識別子を集めることができず、未登録も未使用も検出できない。
    found = reporter_functions()
    assert {"make_diagnostic", "report", "_issue"} <= set(found)
    missing = {name: places for name, places in found.items() if name not in REPORTERS}
    assert not missing, "診断を作る関数が REPORTERS（tests/test_diagnostics.py）にない: " + "、".join(
        f"{name}（{', '.join(places)}）" for name, places in sorted(missing.items())
    )


def test_every_reported_code_is_registered():
    found = codes_in_source()
    # 集め方がおかしくなって何も見つからないときに、試験が確かめないまま通らないようにする。
    assert len(found) >= 50 and {"text_overflow", "element_overlap", "input_error", "renderer_failed"} <= set(found)
    missing = {code: places for code, places in found.items() if code not in DIAGNOSTICS}
    assert (
        not missing
    ), "診断の表（src/slidekit/diagnostics.py）にない識別子がある。登録して文書を作り直す: " + "、".join(
        f"{code}（{', '.join(places)}）" for code, places in sorted(missing.items())
    )


def test_registry_has_no_unused_codes():
    # 作成 API が先に拒む構造の誤りは、場面の検査から除いた後も、作成 API の例外の code として使う。
    # どこからも出なくなった識別子は表から除く（出ない識別子を文書に残さない）。
    found = codes_in_source()
    assert {"invalid_geometry", "geometry_mismatch", "text_wrap_mismatch", "math_fallback_missing"} <= set(found)
    unused = sorted(code for code in DIAGNOSTICS if code not in found)
    assert (
        not unused
    ), (
        "診断の表（src/slidekit/diagnostics.py）に、ソースのどこからも出ない識別子がある。"
        "表から除いて文書を作り直すか、出す処理を加える: " + "、".join(unused)
    )


def test_diagnostics_document_matches_registry():
    command = "python -m slidekit.diagnostics --write docs/diagnostics.md"
    text = DOCUMENT.read_text(encoding="utf-8")
    assert (
        text.count(START_MARKER) == 1 and text.count(END_MARKER) == 1
    ), f"docs/diagnostics.md に生成部分の印がない。印を置いて {command} を実行する"
    generated = text.split(START_MARKER, 1)[1].split(END_MARKER, 1)[0]
    assert (
        generated.strip() == markdown().strip()
    ), f"docs/diagnostics.md の診断の一覧が表と一致しない。{command} で作り直す"
    assert main(["--check", str(DOCUMENT)]) == 0


def test_registry_has_valid_entries_and_document_rows():
    assert len(set(CATEGORIES)) == len(CATEGORIES)
    assert WARNING_CODES == {code for code, spec in DIAGNOSTICS.items() if spec.severity == "warning"}
    rows = [line for line in markdown().splitlines() if line.startswith("| `")]
    assert [row.split("`")[1] for row in rows] == [
        code for category in CATEGORIES for code, spec in DIAGNOSTICS.items() if spec.category == category
    ]
    with pytest.raises(TypeError):
        DIAGNOSTICS["text_overflow"] = DIAGNOSTICS["text_overflow"]
    for arguments in (
        ("fatal", "input", "判定", "直し方"),
        ("error", "unknown", "判定", "直し方"),
        ("error", "input", "", "直し方"),
        ("error", "input", "判定", "2 行の\n直し方"),
    ):
        with pytest.raises(ValueError):
            DiagnosticSpec(*arguments)
    for options in (
        {"element": "style"},
        {"page": ("size", "title")},
        {"page": ["size"]},
        {"rules": ["notes_style"]},
        {"version": 0},
        {"fonts": 1},
    ):
        with pytest.raises(ValueError):
            EvidenceScope(**options)
    with pytest.raises(ValueError, match="重複"):
        diagnostics._registry([("same", DIAGNOSTICS["text_overflow"]), ("same", DIAGNOSTICS["text_overflow"])])


def test_warning_evidence_scopes_follow_the_records_table():
    # 範囲を決めていない種類は、 0.1 と同じ広い範囲になる。
    assert EvidenceScope() == EvidenceScope(
        rules=None, roles=False, fonts=True, element="all", page=("size", "notes"), pair=False, version=1
    )
    scope = {code: DIAGNOSTICS[code].evidence for code in WARNING_CODES}
    assert scope["notes_style"].rules == ("notes_style",) and scope["notes_style"].page == ("notes",)
    assert scope["math_not_marked"].rules == ("math_candidates",) and scope["math_not_marked"].element == "content"
    assert scope["line_break"].rules == () and scope["line_break"].element == "content"
    assert scope["content_area_overflow"].page == ("size", "content_area")
    assert scope["text_padding_overlap"] == EvidenceScope(rules=(), fonts=True, element="all", page=("size",))
    # PDF の警告はページの値の代わりに validate_pdf に渡すページの表示の証拠を使う（行数の警告も同じ扱い）。
    pdf = EvidenceScope(rules=(), fonts=False, element="all", page=())
    assert scope["pdf_text_missing"] == scope["pdf_font_substitution"] == scope["pdf_line_count_mismatch"] == pdf
    assert scope["fallback_crop_clipped"] == EvidenceScope(rules=(), fonts=False, element="all", page=())
    assert {code for code, value in scope.items() if value.pair} == {"element_overlap"}
    # フォントの一覧を入れる警告は、計測の結果で判定する余白の警告だけである。
    assert {code for code, value in scope.items() if value.fonts} == {"text_padding_overlap"}
    assert all(value.rules is not None and not value.roles and value.version == 1 for value in scope.values())


def test_main_writes_and_checks_document(tmp_path, capsys):
    path = tmp_path / "diagnostics.md"
    path.write_text(f"# 前置き\n\n{START_MARKER}\n古い表\n{END_MARKER}\n\n後書き\n", encoding="utf-8")
    assert main(["--check", str(path)]) == 1
    assert "--write" in capsys.readouterr().err
    assert main(["--write", str(path)]) == 0
    text = path.read_text(encoding="utf-8")
    assert text.startswith(f"# 前置き\n\n{START_MARKER}\n\n## 入力\n") and text.endswith(
        f"\n\n{END_MARKER}\n\n後書き\n"
    )
    assert main(["--check", str(path)]) == 0
    assert main([]) == 0 and capsys.readouterr().out == markdown()
    path.write_text("印のない文書\n", encoding="utf-8")
    assert main(["--write", str(path)]) == 2 and main(["--check", str(path)]) == 2
    assert path.read_text(encoding="utf-8") == "印のない文書\n"
    assert main(["--check", str(tmp_path / "missing.md")]) == 2


def test_diagnostics_imports_only_standard_library():
    tree = ast.parse(Path(diagnostics.__file__).read_text(encoding="utf-8"))
    imports = [node for node in ast.walk(tree) if isinstance(node, (ast.Import, ast.ImportFrom))]
    assert not any(isinstance(node, ast.ImportFrom) and node.level for node in imports)
    modules = {alias.name.split(".")[0] for node in imports if isinstance(node, ast.Import) for alias in node.names}
    modules |= {node.module.split(".")[0] for node in imports if isinstance(node, ast.ImportFrom)}
    assert modules - {"__future__"} <= sys.stdlib_module_names
