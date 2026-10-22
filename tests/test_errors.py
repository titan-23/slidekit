"""入力の誤りの例外と、例外の発生場所・失敗の診断を確かめる。"""

import functools
import importlib.util
import json
from pathlib import Path
import pickle
import sys
from types import SimpleNamespace
import uuid

import pytest

from slidekit.errors import LIBRARY_ROOT, ArgumentNameError, InputError, exception_location, failure_diagnostic
from slidekit.parallel import Workers

SLIDES = """from slidekit.selection import parse

from parts.helper import broken_input


def build(context):
    labels = {"a": "3-"}
    return parse(labels["a"])


def through_helper(context):
    return broken_input()


def through_string(context):
    return exec(compile("parse('3-')", "<string>", "exec"), {"parse": parse})


def missing_key(context):
    labels = {"a": 1}
    return labels["b"]
"""

HELPER = """from pathlib import Path

from slidekit.errors import InputError


def broken_input():
    raise InputError("ページ intro の要素 body: 色が未定義です", code="undefined_color", slide_id="intro",
                     element_id="body", color="red", source=Path("theme.toml"))
"""


def load_project(tmp_path):
    """資料のディレクトリに slides.py と補助モジュールを書いて読み込み、(資料のディレクトリ, モジュール) を返す。"""
    project = tmp_path / "project"
    (project / "parts").mkdir(parents=True)
    (project / "slides.py").write_text(SLIDES, encoding="utf-8")
    (project / "parts" / "helper.py").write_text(HELPER, encoding="utf-8")
    # 資料の生成と同じく、資料のディレクトリから補助モジュールを読み込む。
    sys.path.insert(0, str(project))
    try:
        name = "_slidekit_test_" + uuid.uuid4().hex
        spec = importlib.util.spec_from_file_location(name, project / "slides.py")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
    finally:
        sys.path.remove(str(project))
        sys.modules.pop("parts.helper", None)
        sys.modules.pop("parts", None)
    return project.resolve(), module


def raised(function, *args):
    with pytest.raises(BaseException) as caught:
        function(*args)
    return caught.value


def line_of(path, text):
    return next(
        number for number, line in enumerate(Path(path).read_text(encoding="utf-8").splitlines(), 1) if text in line
    )


def test_input_error_attributes_and_defaults():
    error = InputError("ページ intro: 不正な入力です")
    assert (str(error), error.code, error.slide_id, error.element_id, error.details) == (
        "ページ intro: 不正な入力です",
        "input_error",
        None,
        None,
        {},
    )
    error = InputError("色が未定義です", code="undefined_color", slide_id="intro", element_id="body", color="red")
    assert (error.code, error.slide_id, error.element_id, error.details) == (
        "undefined_color",
        "intro",
        "body",
        {"color": "red"},
    )
    assert InputError.code == "input_error" and ArgumentNameError.code == "unknown_argument"
    # 既存の ValueError を捕まえる処理でも扱うことができる。
    with pytest.raises(ValueError, match="未定義"):
        raise error


def test_argument_name_error_is_type_and_value_error():
    for kind in (TypeError, ValueError, InputError):
        with pytest.raises(kind):
            raise ArgumentNameError("box() は z_index を受け付けません")
    error = ArgumentNameError(
        "line() は fill を受け付けません", code="unsupported_argument", slide_id="intro", element_id="arrow"
    )
    assert (error.code, error.slide_id, error.element_id) == ("unsupported_argument", "intro", "arrow")
    assert ArgumentNameError("名前").code == "unknown_argument"


def test_codes_are_checked():
    with pytest.raises(ValueError, match="unknown_argument か unsupported_argument"):
        ArgumentNameError("名前", code="input_error")
    for value in ("", 3):
        with pytest.raises(ValueError, match="空でない文字列"):
            InputError("入力", code=value)


class SettingsLikeError(InputError):
    """__init__ の引数の形が InputError と違う派生（設定の誤りのように、問題の一覧から文言を作る）。"""

    code = "settings_invalid"

    def __init__(self, problems):
        self.problems = tuple(problems)
        super().__init__("\n".join(self.problems), problems=list(self.problems))


def test_errors_keep_attributes_through_pickle():
    errors = [
        InputError("色が未定義です", code="undefined_color", slide_id="intro", element_id="body", color="red"),
        ArgumentNameError("fill は使えません", code="unsupported_argument", element_id="arrow"),
        SettingsLikeError(["theme.toml: colours は不明な設定です", "rules.toml: numbers は真偽値にしてください"]),
    ]
    errors[0].add_note("補足")
    for error in errors:
        restored = pickle.loads(pickle.dumps(error))
        assert type(restored) is type(error)
        assert (
            str(restored),
            restored.args,
            restored.code,
            restored.slide_id,
            restored.element_id,
            restored.details,
        ) == (str(error), error.args, error.code, error.slide_id, error.element_id, error.details)
    assert pickle.loads(pickle.dumps(errors[0])).__notes__ == ["補足"]
    assert pickle.loads(pickle.dumps(errors[2])).problems == errors[2].problems
    # 子プロセスから戻しても同じ値になる（並列の作業関数の例外を親へ渡す処理）。
    make = functools.partial(ArgumentNameError, code="unsupported_argument", slide_id="intro", element_id="arrow")
    with Workers(2) as workers:
        returned = workers.map(make, [("fill は使えません",), ("rotation は使えません",)], minimum=1)
    assert [(type(item), str(item), item.code, item.slide_id, item.element_id) for item in returned] == [
        (ArgumentNameError, "fill は使えません", "unsupported_argument", "intro", "arrow"),
        (ArgumentNameError, "rotation は使えません", "unsupported_argument", "intro", "arrow"),
    ]


def test_location_points_to_last_project_frame(tmp_path, monkeypatch):
    project, module = load_project(tmp_path)
    shared = tmp_path / "shared"
    shared.mkdir()
    slides = project / "slides.py"
    error = raised(module.build, None)
    location = exception_location(error, roots=[shared, project], library_root=LIBRARY_ROOT, base=project)
    assert location == {
        "path": str(slides),
        "relative_path": "slides.py",
        "line": line_of(slides, 'return parse(labels["a"])'),
        "function": "build",
        "code": 'return parse(labels["a"])',
        "in_project": True,
        "raised_at": {
            "path": str(LIBRARY_ROOT / "selection.py"),
            "line": location["raised_at"]["line"],
            "function": "parse",
        },
    }
    # 補助モジュールから例外が出たときは、その場所を示す。
    location = exception_location(raised(module.through_helper, None), roots=[project], base=project)
    assert (location["relative_path"], location["function"], location["in_project"]) == (
        "parts/helper.py",
        "broken_input",
        True,
    )
    # base の外の参照範囲にあるファイルは、base からの相対パスで示す。
    location = exception_location(error, roots=[project], base=shared)
    assert location["relative_path"] == "../project/slides.py"
    # slidekit を含む範囲を参照範囲にしても、slidekit の中のフレームは資料コードとみなさない。
    location = exception_location(error, roots=[project, LIBRARY_ROOT.parent], base=project)
    assert (location["relative_path"], location["function"]) == ("slides.py", "build")
    # "<string>" はファイルではないので、資料のディレクトリがカレントでも資料コードとみなさない。
    monkeypatch.chdir(project)
    location = exception_location(raised(module.through_string, None), roots=[project], base=project)
    assert (location["relative_path"], location["function"], location["line"]) == (
        "slides.py",
        "through_string",
        line_of(slides, "exec(compile("),
    )


def test_syntax_error_location_points_to_faulty_module(tmp_path):
    # 章のモジュールの構文の誤りは、それを読み込んだ行でなく、誤りのファイルと行を示す（並行して書く作業者が自分の誤りを探すことができる）。
    project = tmp_path / "project"
    project.mkdir()
    (project / "chapter_broken.py").write_text("def build(deck):\n    return 1 +\n", encoding="utf-8")
    (project / "slides.py").write_text(
        "import importlib\n\n\ndef build(deck):\n    importlib.import_module('chapter_broken').build(deck)\n",
        encoding="utf-8",
    )
    sys.path.insert(0, str(project))
    try:
        spec = importlib.util.spec_from_file_location("_slidekit_test_" + uuid.uuid4().hex, project / "slides.py")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        error = raised(module.build, None)
    finally:
        sys.path.remove(str(project))
        sys.modules.pop("chapter_broken", None)
    assert isinstance(error, SyntaxError)
    root = project.resolve()
    location = exception_location(error, roots=[root], base=root)
    assert {key: location[key] for key in ("path", "relative_path", "line", "function", "code", "in_project")} == {
        "path": str(root / "chapter_broken.py"),
        "relative_path": "chapter_broken.py",
        "line": 2,
        "function": None,
        "code": "return 1 +",
        "in_project": True,
    }
    # 誤りのファイルが参照範囲の外なら、従来どおりフレームから探す。
    elsewhere = exception_location(error, roots=[tmp_path / "other"], base=root)
    assert elsewhere["in_project"] is False and elsewhere["relative_path"] != "chapter_broken.py"


def test_location_without_project_frame_uses_raising_frame(tmp_path):
    from slidekit.selection import parse

    error = raised(parse, "3-")
    location = exception_location(error, roots=[tmp_path], library_root=LIBRARY_ROOT)
    selection = LIBRARY_ROOT / "selection.py"
    assert location["path"] == str(selection) and location["relative_path"] == "slidekit/selection.py"
    assert (location["function"], location["in_project"]) == ("parse", False)
    assert location["code"] == Path(selection).read_text(encoding="utf-8").splitlines()[location["line"] - 1].strip()
    assert location["raised_at"] == {"path": str(selection), "line": location["line"], "function": "parse"}
    # 資料の参照範囲がなくても、例外が起きた所を示す。
    assert exception_location(error, roots=[])["function"] == "parse"
    # 送出されていない例外には場所がない。
    assert exception_location(ValueError("未送出"), roots=[tmp_path]) == {
        "path": None,
        "relative_path": None,
        "line": None,
        "function": None,
        "code": None,
        "in_project": False,
        "raised_at": None,
    }


def test_failure_diagnostic_has_location_ids_and_traceback(tmp_path):
    project, module = load_project(tmp_path)
    context = SimpleNamespace(read_roots=(project,), root=project)
    run = tmp_path / "out" / "runs" / "0014"
    error = raised(module.through_helper, None)
    diagnostic = failure_diagnostic(error, stage="construct", project=context, traceback_path=run / "traceback.txt")
    assert {key: diagnostic[key] for key in ("code", "slide_id", "element_id", "scope", "severity", "message")} == {
        "code": "undefined_color",
        "slide_id": "intro",
        "element_id": "body",
        "scope": "construct",
        "severity": "error",
        "message": "ページ intro の要素 body: 色が未定義です",
    }
    assert diagnostic["diagnostic_id"].startswith("d") and diagnostic["evidence_hash"]
    [observation] = diagnostic["observations"]
    assert observation["exception"] == "InputError" and observation["message"] == str(error)
    assert observation["traceback"] == "traceback.txt"
    # 例外の付加情報も観測値に入れ、JSON の値にする。
    assert (observation["color"], observation["source"]) == ("red", "theme.toml")
    assert (observation["location"]["relative_path"], observation["location"]["in_project"]) == (
        "parts/helper.py",
        True,
    )
    json.dumps(diagnostic, ensure_ascii=False)
    # code を持たない例外は build_failed にする。実行ディレクトリからの相対パスはそのまま使う。
    diagnostic = failure_diagnostic(
        raised(module.missing_key, None), stage="construct", project=context, traceback_path=Path("traceback.txt")
    )
    assert (diagnostic["code"], diagnostic["message"], diagnostic["slide_id"]) == ("build_failed", "'b'", None)
    assert diagnostic["observations"][0]["location"]["code"] == 'return labels["b"]'
    assert diagnostic["observations"][0]["traceback"] == "traceback.txt"
    # 文字列でない code（SystemExit など）は使わない。資料がなければ例外の起きた所を示す。
    diagnostic = failure_diagnostic(raised(sys.exit, 3), stage="doctor")
    assert (diagnostic["code"], diagnostic["scope"], diagnostic["observations"][0]["traceback"]) == (
        "build_failed",
        "doctor",
        None,
    )
    assert diagnostic["observations"][0]["location"]["in_project"] is False
    assert failure_diagnostic(KeyboardInterrupt(), stage="render:powerpoint:standard")["message"] == "KeyboardInterrupt"
