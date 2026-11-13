"""診断の並べ替えと件数、コマンドの結果の要約を確かめる。"""

import importlib.util
from types import SimpleNamespace
import uuid

import pytest

from slidekit.errors import failure_diagnostic
from slidekit.report import count_issues, sort_issues
from slidekit.report import summarize as full_summarize


def summarize(*args, **kwargs) -> str:
    """要約から最後の結果の行を除いた本体。最後の行が「結果: 」で始まることも確かめる（行の形は個別の試験で見る）。"""
    body, _, last = full_summarize(*args, **kwargs).rpartition("\n")
    assert last.startswith("結果: "), last
    return body


def issue(
    code, severity="warning", *, slide=None, number=None, element=None, related=None, message="", observations=None
):
    value = {
        "code": code,
        "severity": severity,
        "slide_id": slide,
        "element_id": element,
        "related_element_id": related,
        "message": message or f"{code} の文言",
        "observations": observations if observations is not None else [],
    }
    if number is not None:
        value["slide_number"] = number
    return value


def test_sort_issues_by_page_then_severity():
    issues = [
        issue("pdf_page_count", "error"),
        issue("element_overlap", slide="b", number=2, element="x", related="y"),
        issue("text_overflow", "error", slide="b", number=2, element="z"),
        issue("font_mismatch", "error", slide="a"),
        issue("notes_style", "info", slide="b", number=2),
        issue("element_overlap", slide="b", number=2, element="a", related="y"),
        issue("unknown_page", slide="zzz"),
        issue("element_overlap", slide="b", number=2, element="a"),
    ]
    original = list(issues)
    slides = [{"id": "a", "number": 1}, {"id": "b", "number": 2}]
    ordered = sort_issues(issues, slides)
    assert [(item["code"], item["element_id"], item["related_element_id"]) for item in ordered] == [
        ("font_mismatch", None, None),  # 1 ページ（slide_id から番号を引く）
        ("text_overflow", "z", None),  # 2 ページのエラー
        ("element_overlap", "a", None),
        ("element_overlap", "a", "y"),
        ("element_overlap", "x", "y"),
        ("notes_style", None, None),  # 2 ページの参考
        ("pdf_page_count", None, None),  # ページのない診断は最後。その中も重大度の順
        ("unknown_page", None, None),
    ]
    assert issues == original and ordered is not issues
    # slides がなければ、slide_number のない診断はページのない診断として扱う。
    assert [item["code"] for item in sort_issues(issues)][-3:] == ["font_mismatch", "pdf_page_count", "unknown_page"]
    assert count_issues(issues) == {"error": 3, "warning": 4, "info": 1}
    assert count_issues([{"severity": "fatal"}, "壊れた値", None]) == {"error": 0, "warning": 0, "info": 0}


def test_summary_limits_warnings_but_keeps_errors(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    run = tmp_path / "out" / "runs" / "0012"
    issues = [issue("text_overflow", "error", slide=f"s{n}", number=n, element="body") for n in (40, 3, 12)]
    issues += [issue("element_overlap", slide=f"s{n}", number=n, element="note", related="table") for n in range(1, 31)]
    result = {
        "command": "build",
        "status": "built",
        "build": str(run),
        "stage": "complete",
        "mode": "full",
        "selection": None,
        "exportable": True,
        "slides": 45,
        "fallback": True,
        "issues": issues,
        "exit_code": 1,
    }
    lines = summarize("build", result, limit=5).splitlines()
    assert lines[:2] == ["生成: 完了（エラー 3 件、警告 30 件）", "実行: out/runs/0012（45 ページ、互換画像あり）"]
    assert lines[2:6] == [
        "エラー:",
        "  p.3 エラー text_overflow body text_overflow の文言",
        "  p.12 エラー text_overflow body text_overflow の文言",
        "  p.40 エラー text_overflow body text_overflow の文言",
    ]
    assert lines[6] == "警告:"
    assert lines[7:12] == [f"  p.{n} 警告 element_overlap note / table element_overlap の文言" for n in range(1, 6)]
    assert lines[12] == "  ほか 25 件（diagnostics.json）"
    # エラーのあるページを示して、配置の検査で直す手順を案内する。
    assert lines[-1] == "次の操作: 編集元を直して slidekit layout で確かめる"
    result["project"] = str(tmp_path / "deck.toml")
    assert (
        summarize("build", result).splitlines()[-1]
        == "次の操作: 編集元を直して slidekit layout --project deck.toml --slides 3,12,40"
    )
    # エラーは limit によらず全件を示す。ただし 50 件で打ち切る。
    result["issues"] = [issue("page_overflow", "error", slide=f"s{n}", number=n) for n in range(1, 56)]
    lines = summarize("build", result, limit=0).splitlines()
    assert sum(line.startswith("  p.") for line in lines) == 50
    assert "  ほか 5 件（diagnostics.json）" in lines and "警告:" not in lines
    # 配置の検査では、残りは --json で見る。
    layout = {
        "command": "layout",
        "status": "failed",
        "project": "deck.toml",
        "slides": 3,
        "issues": issues[3:],
        "exit_code": 0,
    }
    assert summarize("layout", layout, limit=1).splitlines()[-3:-1] == [
        "  p.1 警告 element_overlap note / table element_overlap の文言",
        "  ほか 29 件（--json で全件を表示）",
    ]


def test_check_summary_groups_missing_reviews(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    run = tmp_path / "out" / "runs" / "0012"
    pages = [f"p{n:02d}" for n in range(1, 46)]

    def missing(kind, slide, renderer=None, variant=None):
        return {
            "code": "review_missing",
            "severity": "error",
            "message": "必要な確認が完了していない",
            "target": {"kind": kind, "slide_id": slide, "renderer": renderer, "variant": variant, "evidence_hash": "e"},
        }

    warnings = [
        dict(issue("element_overlap", slide="p05", number=5, element="label", related="frame"), diagnostic_id="d1"),
        dict(issue("element_overlap", slide="p07", number=7, element="label", related="frame"), diagnostic_id="d2"),
        dict(issue("math_not_marked", slide="p30", number=30, element="summary"), diagnostic_id="d3"),
    ]
    issues = [
        {
            "code": "input_changed",
            "severity": "error",
            "message": "参照入力が生成時と一致しないため、再生成が必要である",
            "path": str(tmp_path / name),
            "expected": "a",
            "actual": "b",
        }
        for name in ("slides.py", "theme.toml")
    ]
    issues.append(
        issue(
            "text_overflow",
            "error",
            slide="p02",
            number=2,
            element="body",
            message="文章の幅または高さが指定領域を超えている",
        )
    )
    issues += [missing("content", page) for page in pages]
    issues += [missing("visual", page, "powerpoint", "standard") for page in pages]
    issues.append(missing("overview", None, "powerpoint", "standard"))
    issues += [missing("visual", page, "powerpoint", "fallback") for page in ("p10", "p11", "p30")]
    issues += [
        {
            "code": "warning_unaccepted",
            "severity": "error",
            "message": "警告が未対応である",
            "diagnostic_id": item["diagnostic_id"],
        }
        for item in warnings
    ]
    result = {
        "schema_version": 1,
        "run_id": "0012",
        "status": "failed",
        "final": True,
        "issues": issues,
        "warnings": warnings,
        "invalid_records": [{"record_id": "r_0123", "reason": "確認対象の証拠が変わった"}],
        "exit_code": 2,
    }
    assert summarize("check", {**result, "build": str(run)}).splitlines() == [
        "最終検査: 失敗（エラー 3 件、警告 3 件、未完了の確認 94 件、未対応の警告 3 件）",
        "実行: out/runs/0012",
        "生成記録との不一致:",
        "  input_changed 参照入力が生成時と一致しないため、再生成が必要である（slides.py ほか 1 件）",
        "エラー:",
        "  p.2 エラー text_overflow body 文章の幅または高さが指定領域を超えている",
        "未完了の確認:",
        "  内容 45 件（先頭: p01）",
        "  表示（powerpoint・標準） 45 件（先頭: p01）",
        "  一覧画像（powerpoint・標準） 1 件",
        "  表示（powerpoint・互換） 3 件（先頭: p10）",
        "未対応の警告:",
        "  element_overlap 2 件（p.5,7）",
        "  math_not_marked 1 件（p.30）",
        "無効な確認記録: 1 件（r_0123: 確認対象の証拠が変わった）",
        "次の操作: slidekit build で生成し直す（入力または成果物が生成記録と一致しない）",
    ]
    # 確認だけが残っているときは、確認の状態を見る操作を示す。実行の場所がなければ識別子で示す。
    remaining = {
        **result,
        "issues": [item for item in issues if item["code"] in {"review_missing", "warning_unaccepted"}],
        "invalid_records": [],
        "exit_code": 1,
    }
    lines = summarize("check", remaining).splitlines()
    assert lines[:2] == [
        "最終検査: 不合格（エラー 0 件、警告 3 件、未完了の確認 94 件、未対応の警告 3 件）",
        "実行: 0012",
    ]
    assert lines[-1] == "次の操作: slidekit review status --build 0012"
    passed = {
        "run_id": "0012",
        "build": str(run),
        "status": "passed",
        "final": False,
        "issues": [],
        "warnings": warnings[:1],
        "invalid_records": [],
        "exit_code": 0,
    }
    assert summarize("check", passed).splitlines() == [
        "検査: 合格（エラー 0 件、警告 1 件）",
        "実行: out/runs/0012",
        "警告:",
        "  p.5 警告 element_overlap label / frame element_overlap の文言",
        "次の操作: slidekit review status --build out/runs/0012",
    ]
    assert (
        summarize("check", {**passed, "final": True}).splitlines()[-1]
        == "次の操作: slidekit export --build out/runs/0012"
    )


class Unusual(dict):
    def get(self, *args):
        raise RuntimeError("読めない")


@pytest.mark.parametrize(
    "command, result",
    [
        ("build", None),
        ("build", {"issues": "壊れた値", "counts": [], "timings": "x", "images": 3, "selection": 5}),
        ("layout", {"status": 3, "issues": [1, None, {"code": None}], "selection": {"numbers": "x"}}),
        ("doctor", {"environment": [], "renderers": "x", "status": "passed"}),
        ("doctor", {"environment": {"fonts": {"fonts": []}, "renderers": {"powerpoint": None}}}),
        ("check", {"issues": [{"code": "review_missing", "target": None}], "warnings": None, "invalid_records": "x"}),
        ("export", {"status": "exported", "deliverables": ["x"]}),
        ("export", {"status": "failed", "check": []}),
        ("clean", {"candidates": 5, "protected": "x", "removed": None}),
        ("clean", {"candidates": [None, 3, {"bytes": "x"}]}),
        ("renderer status", {"holder": "x", "pending": [], "status": None}),
        ("renderer clear", {"status": ["x"]}),
        ("review record", {"status": "recorded", "records": [{}]}),
        ("unknown", {"status": "ok", "message": "知らない結果"}),
        (
            "review status",
            {"reviews": [None, {"pending": "x"}], "warnings": [], "next": "x", "blocking": [{"paths": 3}]},
        ),
        ("review accept", {"records": "x", "skipped": [None]}),
        ("review revoke", {"revoked": 3, "cascaded": [{}]}),
        ("review carry", {"status": "carried", "skipped": {"x": "y"}, "sources": [], "remaining": [1]}),
        ("build", {"status": "built", "carry": {"status": "carried", "remaining": "x", "renumbered": 3}}),
        (
            "layout",
            {
                "status": "failed",
                "exception": "ValueError",
                "message": None,
                "issues": [{"observations": [{"location": 3}]}],
            },
        ),
        (None, {}),
        (["list"], {"status": "failed"}),
        ("build", Unusual(status="built")),
    ],
)
def test_unknown_result_shape_does_not_raise(command, result):
    text = summarize(command, result, limit="x")
    assert isinstance(text, str) and text
    assert ": " in text.splitlines()[0]


def test_unknown_command_shows_status_and_message():
    assert summarize("unknown", {"status": "failed", "message": "読み込めません", "exception": "ValueError"}) == (
        "unknown: 失敗\n  ValueError: 読み込めません"
    )
    assert summarize("review carry", {"status": "carried", "records": []}) == (
        "確認の引き継ぎ: 引き継いだ（0 件）\n次の操作: slidekit review status"
    )


def test_known_codes_show_observations():
    issues = [
        issue(
            "text_overflow",
            "error",
            number=1,
            element="body",
            message="文章が領域を超えている",
            observations=[
                {
                    "required_width_pt": 120.46,
                    "available_width_pt": 100,
                    "required_height_pt": 30,
                    "available_height_pt": 24.04,
                }
            ],
        ),
        issue(
            "page_overflow",
            "error",
            number=2,
            element="figure",
            message="ページの外へはみ出している",
            observations=[{"area": [900, -5, 100, 40], "page": [0, 0, 960, 540]}],
        ),
        issue(
            "content_area_overflow",
            number=3,
            element="note",
            message="本文領域の外にある",
            observations=[{"area": [40, 100, 900, 40], "content_area": [40, 80, 880, 400]}],
        ),
        issue(
            "math_not_marked",
            number=4,
            element="summary",
            message="数式として指定すべきか確認する",
            observations=[{"candidates": ["n", "x"]}, {"candidates": ["x", "2"]}],
        ),
        issue(
            "font_glyph_missing",
            "error",
            number=5,
            element="mark",
            message="字形がない",
            observations=[{"characters": ["※", "\u2028"]}],
        ),
        issue(
            "forbidden_character",
            "error",
            number=6,
            element="title",
            message="禁止された文字がある",
            observations={"characters": ["…"]},
        ),
        issue(
            "pdf_element_text_overflow",
            "error",
            number=7,
            element="label",
            message="文字枠からはみ出している",
            observations=[
                {
                    "text": "とても長い説明の文章で、四十文字を超えるので途中で省略して表示する例です。",
                    "text_area": [0, 0, 1, 1],
                }
            ],
        ),
        issue(
            "element_overlap",
            number=8,
            element="a",
            related="b",
            message="要素の重なりを確認する必要がある",
            observations=[{"intersection": [1, 2, 3, 4]}],
        ),
    ]
    lines = summarize(
        "layout", {"status": "failed", "project": "deck.toml", "slides": 8, "issues": issues, "exit_code": 1}
    ).splitlines()
    assert [line for line in lines if line.startswith("  p.")] == [
        "  p.1 エラー text_overflow body 文章が領域を超えている（必要な寸法 120.5×30.0 pt、使える寸法 100.0×24.0 pt）",
        "  p.2 エラー page_overflow figure ページの外へはみ出している（ページから上へ 5.0 pt、右へ 40.0 pt はみ出し）",
        "  p.5 エラー font_glyph_missing mark 字形がない（文字: ※、U+2028）",
        "  p.6 エラー forbidden_character title 禁止された文字がある（文字: …）",
        "  p.7 エラー pdf_element_text_overflow label 文字枠からはみ出している（文字: 「とても長い説明の文章で、四十文字を超えるので途中で省略して表…」）",
        "  p.3 警告 content_area_overflow note 本文領域の外にある（本文領域から右へ 20.0 pt はみ出し）",
        "  p.4 警告 math_not_marked summary 数式として指定すべきか確認する（候補: n、x、2）",
        "  p.8 警告 element_overlap a / b 要素の重なりを確認する必要がある",
    ]
    assert lines[0] == "配置の検査: 不合格（エラー 5 件、警告 3 件）"
    assert lines[-1] == "次の操作: 編集元を直して slidekit layout --project deck.toml --slides 1-2,5-7"


SLIDES = """def build_method(context):
    labels = {"a": 1}
    label = labels["b"]
    return label
"""


def test_failure_summary_shows_location_code_and_traceback(tmp_path, monkeypatch):
    project = tmp_path / "project"
    project.mkdir()
    (project / "slides_method.py").write_text(SLIDES, encoding="utf-8")
    spec = importlib.util.spec_from_file_location("_slidekit_report_" + uuid.uuid4().hex, project / "slides_method.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    with pytest.raises(KeyError) as raised:
        module.build_method(None)
    run = project / "out" / "runs" / "0014"
    diagnostic = failure_diagnostic(
        raised.value,
        stage="construct",
        project=SimpleNamespace(read_roots=(project,), root=project),
        traceback_path=run / "traceback.txt",
    )
    monkeypatch.chdir(project)
    result = {
        "command": "build",
        "status": "failed",
        "build": str(run),
        "stage": "construct",
        "issues": [diagnostic],
        "exit_code": 2,
    }
    assert summarize("build", result).splitlines() == [
        "生成: 失敗（構築）",
        "実行: out/runs/0014",
        "エラー:",
        "  build_failed slides_method.py:3（build_method） KeyError: 'b'",
        '      label = labels["b"]',
        "  トレースバック: out/runs/0014/traceback.txt",
        "次の操作: 原因を直して slidekit build を再実行する",
    ]
    # slidekit の内部で起きた例外は、その旨を添える。失敗した段階は括弧を含まない短い名前で示す。
    location = {
        "path": "/opt/slidekit/model.py",
        "relative_path": "slidekit/model.py",
        "line": 120,
        "function": "_add",
        "code": "raise InputError(message)",
        "in_project": False,
    }
    internal = {
        "code": "invalid_style",
        "severity": "error",
        "slide_id": "intro",
        "slide_number": 3,
        "element_id": "body",
        "message": "ページ intro の要素 body: 余白が不正です",
        "observations": [
            {
                "exception": "InputError",
                "message": "ページ intro の要素 body: 余白が不正です",
                "location": location,
                "traceback": None,
            }
        ],
    }
    lines = summarize("build", {**result, "stage": "render:powerpoint:standard", "issues": [internal]}).splitlines()
    assert lines[0] == "生成: 失敗（PowerPoint の変換）"
    assert lines[3:5] == [
        "  p.3 invalid_style slidekit/model.py:120（_add、slidekit の内部で発生） InputError: ページ intro の要素 body: 余白が不正です",
        "      raise InputError(message)",
    ]
    assert (
        summarize("build", {**result, "stage": "doctor", "issues": []}).splitlines()[0]
        == "生成: 失敗（表示ソフトの事前確認）"
    )
    # --slides の指定の誤り（invalid_selection）だけは利用者の入力の誤りなので、場所を持たず、指定を直す操作を示す。
    selected = {
        "code": "invalid_selection",
        "severity": "error",
        "slide_id": None,
        "element_id": None,
        "scope": "construct",
        "message": "ページの識別子がありません: tracee（候補: trace）",
        "observations": [
            {"exception": "SelectionError", "message": "ページの識別子がありません: tracee（候補: trace）"}
        ],
    }
    assert summarize("build", {**result, "issues": [selected]}).splitlines()[2:] == [
        "エラー:",
        "  エラー invalid_selection ページの識別子がありません: tracee（候補: trace）",
        "次の操作: --slides の指定を直して slidekit build を再実行する",
    ]
    layout = {
        "command": "layout",
        "status": "failed",
        "project": str(project / "deck.toml"),
        "issues": [selected],
        "exit_code": 2,
    }
    assert (
        summarize("layout", layout).splitlines()[-1] == "次の操作: --slides の指定を直して slidekit layout を再実行する"
    )
    assert (
        summarize("layout", {**layout, "issues": [internal]}).splitlines()[-1]
        == "次の操作: 原因を直して slidekit layout を再実行する"
    )


def test_command_summaries_show_state_target_and_next_step(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    run = str(tmp_path / "out" / "runs" / "0012")
    project = str(tmp_path / "deck.toml")
    timings = [
        {"stage": "construct", "label": "構築", "seconds": 0.9},
        {"stage": "layout", "label": "配置の検査", "seconds": 2.0},
        {"stage": "render:powerpoint:standard", "label": "PowerPoint で変換（標準表示）", "seconds": 16.5},
        {"stage": "render:powerpoint:fallback", "label": "PowerPoint で変換（互換表示）", "seconds": 13.0},
        {"stage": "pdf:powerpoint:standard", "label": "PDF の検査と画像化", "seconds": 1.5},
        {"stage": "math_fallback", "label": "互換画像の作成", "seconds": 0.8},
    ]
    selection = {"spec": "2,5-6", "slide_ids": ["b", "e", "f"], "numbers": [2, 5, 6], "total_slides": 45}
    built = {
        "command": "build",
        "status": "built",
        "build": run,
        "run_id": "0012",
        "stage": "complete",
        "mode": "full",
        "selection": None,
        "exportable": True,
        "slides": 45,
        "fallback": True,
        "counts": {"error": 0, "warning": 0, "info": 0},
        "issues": [],
        "timings": timings,
        "elapsed_seconds": 38.2,
        "exit_code": 0,
        "images": {"powerpoint:standard": [f"renders/powerpoint/standard/contact-00{n}.png" for n in (1, 2, 3, 4)]},
    }
    assert summarize("build", built).splitlines() == [
        "生成: 完了（エラー 0 件、警告 0 件）",
        "実行: out/runs/0012（45 ページ、互換画像あり）",
        "所要時間: 38.2 秒（変換 29.5、配置の検査 2.0、PDF の検査と画像化 1.5、構築 0.9、その他 4.3）",
        "画像: out/runs/0012/renders/powerpoint/standard/contact-001.png ほか 3 枚",
        "次の操作: slidekit review status --build out/runs/0012",
    ]
    draft = {
        **built,
        "mode": "draft",
        "selection": selection,
        "exportable": False,
        "fallback": False,
        "project": project,
        "timings": [],
        "elapsed_seconds": None,
    }
    assert summarize("build", draft).splitlines() == [
        "生成: 完了（エラー 0 件、警告 0 件）",
        "実行: out/runs/0012（45 ページ中 3 ページ（2,5-6）、下書き）",
        "画像: out/runs/0012/renders/powerpoint/standard/contact-001.png ほか 3 枚",
        "次の操作: 表示を確かめ、全ページを slidekit build --project deck.toml で生成する",
    ]
    stopped = {
        **built,
        "status": "stopped",
        "stage": "layout",
        "counts": None,
        "project": project,
        "exit_code": 1,
        "issues": [issue("text_overflow", "error", number=4, element="body")],
        "timings": None,
        "elapsed_seconds": 1.2,
    }
    assert summarize("build", stopped).splitlines() == [
        "生成: 変換の前に停止（エラー 1 件、警告 0 件）",
        "実行: out/runs/0012（45 ページ、互換画像あり）",
        "エラー:",
        "  p.4 エラー text_overflow body text_overflow の文言",
        "所要時間: 1.2 秒",
        "次の操作: 編集元を直して slidekit layout --project deck.toml --slides 4",
    ]
    # 0.1 の結果の形（エラーのある完了を status "failed" で返す）も、完了として示す。
    legacy = {"status": "failed", "build": run, "stage": "complete", "issues": stopped["issues"], "exit_code": 1}
    assert summarize("build", legacy).splitlines()[0] == "生成: 完了（エラー 1 件、警告 0 件）"
    # 表示ソフトの時間切れ・使用中・残存記録で失敗したときは、資料でなく表示ソフトの状態を確かめる操作を示す。
    timeout = {
        **built,
        "status": "failed",
        "stage": "render:powerpoint:standard",
        "exit_code": 2,
        "timings": None,
        "elapsed_seconds": None,
        "issues": [issue("renderer_timeout", "error")],
    }
    assert summarize("build", timeout).splitlines()[-3:] == [
        "次の操作:",
        "  変換の終了を待ち、slidekit renderer status で確かめる",
        "  時間切れが続くなら、実行環境の timeout を延ばしてから slidekit build を再実行する",
    ]
    busy = {**timeout, "stage": "doctor", "issues": [issue("renderer_busy", "error")]}
    assert (
        summarize("build", busy).splitlines()[-1]
        == "次の操作: 別の変換の終了を待ってから slidekit build を再実行する（待つ上限は --wait）"
    )
    other = {**timeout, "issues": [issue("build_failed", "error")]}
    assert summarize("build", other).splitlines()[-1] == "次の操作: 原因を直して slidekit build を再実行する"

    layout = {
        "command": "layout",
        "status": "passed",
        "project": project,
        "slides": 45,
        "selection": selection,
        "fallback_required": True,
        "counts": {"error": 0, "warning": 0, "info": 1},
        "issues": [issue("text_wrap_near_limit", "info", number=5, element="body")],
        "timings": timings[:2],
        "elapsed_seconds": 3.0,
        "exit_code": 0,
    }
    assert summarize("layout", layout).splitlines() == [
        "配置の検査: 合格（エラー 0 件、警告 0 件、参考 1 件）",
        "資料: deck.toml（45 ページ中 3 ページ（2,5-6）、互換画像あり）",
        "参考:",
        "  p.5 参考 text_wrap_near_limit body text_wrap_near_limit の文言",
        "所要時間: 3.0 秒（配置の検査 2.0、構築 0.9、その他 0.1）",
        "次の操作: slidekit build --draft --project deck.toml",
    ]
    failed_layout = {
        "command": "layout",
        "status": "failed",
        "project": project,
        "slides": 45,
        "exit_code": 2,
        "issues": [issue("build_failed", "error", message="生成コードを読み込めません")],
    }
    assert summarize("layout", failed_layout).splitlines()[0] == "配置の検査: 失敗（エラー 1 件、警告 0 件）"

    font = {"path": "/fonts/yu.ttc", "index": 0, "bold": {"path": "/fonts/yub.ttc", "index": 0}}
    doctor = {
        "status": "passed",
        "project": project,
        "slides": 45,
        "fallback_required": True,
        "issues": [],
        "exit_code": 0,
        "environment": {
            "fonts": {"fonts": {"Yu Gothic UI": font, "Arial": font, "Cambria Math": {"path": "/fonts/cambria.ttc"}}},
            "renderers": {
                "powerpoint": {
                    "ok": True,
                    "version": "16.0",
                    "executable": "C:\\Office\\POWERPNT.EXE",
                    "font_matches": [
                        {"family": family, "variant": variant}
                        for family, variant in (
                            ("Yu Gothic UI", "regular"),
                            ("Yu Gothic UI", "bold"),
                            ("Arial", "regular"),
                            ("Arial", "bold"),
                            ("Cambria Math", "regular"),
                        )
                    ],
                }
            },
        },
    }
    assert summarize("doctor", doctor).splitlines() == [
        "事前確認: 合格（エラー 0 件、警告 0 件）",
        "資料: deck.toml（45 ページ、互換画像あり）",
        "フォント: Yu Gothic UI（通常・太字）、Arial（通常・太字）、Cambria Math。Windows の登録書体と一致",
        "表示ソフト: PowerPoint 16.0（C:\\Office\\POWERPNT.EXE）",
        "次の操作: slidekit build --project deck.toml",
    ]
    busy = {
        **doctor,
        "status": "failed",
        "exit_code": 2,
        "renderers": {"powerpoint": {"status": "failed", "message": "使用中です"}},
        "issues": [issue("renderer_busy", "error", message="同じ Windows 利用者による変換が実行中です")],
    }
    busy["environment"] = {**doctor["environment"], "renderers": {}}
    lines = summarize("doctor", busy).splitlines()
    assert lines[0] == "事前確認: 失敗（エラー 1 件、警告 0 件）"
    assert lines[2:4] == [
        "フォント: Yu Gothic UI（通常・太字）、Arial（通常・太字）、Cambria Math",
        "表示ソフト: PowerPoint を確認できない（使用中です）",
    ]

    exported = {
        "status": "exported",
        "exit_code": 0,
        "run": run,
        "record": str(tmp_path / "deliverables.json"),
        "deliverables": {
            "pptx": str(tmp_path / "presentation.pptx"),
            "pdf": str(tmp_path / "presentation.pdf"),
            "notes": str(tmp_path / "notes.md"),
            "record": str(tmp_path / "deliverables.json"),
        },
        "history": str(tmp_path / "out" / "exports" / "20261012T010203-abcd1234"),
    }
    assert summarize("export", exported).splitlines() == [
        "反映: 完了",
        "実行: out/runs/0012",
        "完成物: presentation.pptx、presentation.pdf、notes.md",
        "採用記録: deliverables.json",
        "旧版の退避先: out/exports/20261012T010203-abcd1234",
    ]
    refused = {
        "status": "failed",
        "exit_code": 1,
        "check": {
            "run_id": "0012",
            "status": "failed",
            "final": True,
            "warnings": [],
            "issues": [
                {"code": "review_missing", "severity": "error", "target": {"kind": "content", "slide_id": "intro"}}
            ],
        },
    }
    assert summarize("export", refused).splitlines() == [
        "反映: 中止（最終検査に不合格）",
        "実行: 0012",
        "未完了の確認:",
        "  内容 1 件（先頭: intro）",
        "次の操作: slidekit review status --build 0012",
    ]

    preview = {
        "command": "clean",
        "status": "preview",
        "exit_code": 0,
        "bytes": 12_300_000,
        "removed": [],
        "candidates": [
            {
                "run": str(tmp_path / "out/runs/0001"),
                "run_id": "0001",
                "status": "failed",
                "mode": "full",
                "bytes": 300_000,
                "reason": "失敗した実行",
            },
            {
                "run": str(tmp_path / "out/runs/0002"),
                "run_id": "0002",
                "status": "built",
                "mode": "draft",
                "bytes": 12_000_000,
                "reason": "下書き",
            },
            {
                "run": str(tmp_path / "out/runs/0004"),
                "run_id": "0004",
                "status": None,
                "mode": None,
                "bytes": 0,
                "reason": "中断した生成",
            },
        ],
        "protected": [{"run": str(tmp_path / "out/runs/0003"), "reasons": ["採用中", "確認記録の引き継ぎ元"]}],
    }
    # 生成記録のない実行（中断した生成）は、状態と種類の代わりに「生成記録なし」と示す。
    assert summarize("clean", preview).splitlines() == [
        "整理: 削除の候補 3 件（12.3 MB）。まだ削除していません",
        "  0001  failed  full  300.0 KB  失敗した実行",
        "  0002  built  draft  12.0 MB  下書き",
        "  0004  生成記録なし  0 B  中断した生成",
        "保護: 1 件",
        "  0003  採用中、確認記録の引き継ぎ元",
        "次の操作: 同じ条件に --apply を付けると削除する",
    ]
    # 0.1 の形（候補はパスの文字列）も読むことができる。
    cleaned = {
        "status": "cleaned",
        "before": "0003",
        "candidates": [str(tmp_path / "out/runs/0001")],
        "removed": [str(tmp_path / "out/runs/0001")],
        "protected": [],
    }
    assert summarize("clean", cleaned).splitlines() == ["整理: 1 件を削除", "  0001"]

    holder = {
        "schema_version": 1,
        "pid": 4321,
        "run_id": "abc",
        "started_at": "2026-10-12T01:00:00+00:00",
        "holder": {
            "project": project,
            "run": str(tmp_path / "out/runs/0011"),
            "command": "build",
            "host": "desk",
            "python_pid": 1234,
        },
    }
    status = {
        "command": "renderer status",
        "renderer": "powerpoint",
        "status": "busy",
        "busy": True,
        "holder": holder,
        "pending": None,
        "pending_process_alive": None,
        "staged_input_locked": None,
        "powerpoint_running": True,
        "state_dir": "C:\\State",
        "exit_code": 1,
    }
    assert summarize("renderer status", status).splitlines() == [
        "表示ソフトの状態: 使用中（PowerPoint）",
        "使用中の処理: PID 1234、実行 out/runs/0011、資料 deck.toml、コマンド build、端末 desk、開始 2026-10-12T01:00:00+00:00",
        "残存記録: なし",
        "PowerPoint: 起動中",
        "状態フォルダー: C:\\State",
        "次の操作: 変換の終了を待つ（slidekit renderer status で確かめる）",
    ]
    pending = {
        **status,
        "status": "pending",
        "busy": False,
        "holder": None,
        "powerpoint_running": False,
        "pending": {"run_id": "abc", "status": "timeout", "started_at": "2026-10-12T01:00:00+00:00"},
        "pending_process_alive": False,
        "staged_input_locked": False,
    }
    assert summarize("renderer status", pending).splitlines() == [
        "表示ソフトの状態: 残存記録あり（PowerPoint）",
        "残存記録: 時間切れ、記録 abc、開始 2026-10-12T01:00:00+00:00",
        "  作業プロセス: 終了",
        "  検査用の複製: 開かれていない",
        "PowerPoint: 起動していない",
        "状態フォルダー: C:\\State",
        "次の操作: slidekit renderer clear",
    ]
    locked = summarize("renderer status", {**pending, "staged_input_locked": True}).splitlines()
    assert locked[-1] == "次の操作: PowerPoint で検査用の複製を閉じてから slidekit renderer clear を実行する"
    # PowerPoint は読み取り専用で開いた複製のファイルを開いたままにしないので、起動中は開いているかを確かめることができない。
    # 解除は --force を求める（複製が見つからないときも同じ）。前の処理が排他を放さずに終わっていたことも示す。
    running = summarize("renderer status", {**pending, "powerpoint_running": True, "abandoned": True}).splitlines()
    assert running[2:] == [
        "  作業プロセス: 終了",
        "  検査用の複製: 開いているかは確かめられない（PowerPoint が起動中）",
        "PowerPoint: 起動中",
        "排他: 前の処理が放さずに終了していた（slidekit が起動した PowerPoint が残っていることがある）",
        "状態フォルダー: C:\\State",
        "次の操作: PowerPoint の終了を待つか、PowerPoint が検査用の複製を開いていないことを確かめてから slidekit renderer clear --force を実行する",
    ]
    missing = summarize(
        "renderer status", {**pending, "powerpoint_running": True, "staged_input_locked": None}
    ).splitlines()
    assert "  検査用の複製" not in "\n".join(missing) and missing[-1].endswith(
        "slidekit renderer clear --force を実行する"
    )
    assert summarize(
        "renderer clear",
        {
            "command": "renderer clear",
            "status": "cleared",
            "reason": "",
            "exit_code": 0,
            "archived": "C:\\State\\cleared\\20261012T010000Z-pending.json",
            "removed_work_dir": True,
        },
    ).splitlines() == ["残存記録の解除: 解除", "退避先: C:\\State\\cleared\\20261012T010000Z-pending.json", "作業フォルダー: 削除した"]
    assert summarize(
        "renderer clear", {"status": "refused", "reason": "変換の実行中のため解除できません", "exit_code": 1}
    ).splitlines() == ["残存記録の解除: 解除を拒否（変換の実行中のため解除できません）", "次の操作: slidekit renderer status で状態を確かめる"]
    assert (
        summarize("renderer clear", {"status": "nothing", "reason": "", "exit_code": 0})
        == "残存記録の解除: 解除する記録なし"
    )


def test_review_summaries_show_groups_warnings_and_next_steps(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    run = str(tmp_path / "out" / "runs" / "0012")
    record = 'slidekit review record --build out/runs/0012 --kind visual --slides 3,5-7 --reviewer "<確認者>" --note "<確認した内容>"'
    status = {
        "command": "review status",
        "status": "incomplete",
        "exit_code": 1,
        "build": run,
        "run_id": "0012",
        "mode": "full",
        "exportable": True,
        "slides": 45,
        "blocking": [],
        "errors": [],
        "reviews": [
            {
                "kind": "visual",
                "renderer": "powerpoint",
                "variant": "standard",
                "required": 45,
                "satisfied": 41,
                "carried": 38,
                "renumbered": 2,
                "slides": "3,5-7",
                "pending": [{"slide_id": f"p{n}", "number": n} for n in (3, 5, 6, 7)],
            },
            {
                "kind": "overview",
                "renderer": "powerpoint",
                "variant": "standard",
                "required": 1,
                "satisfied": 0,
                "carried": 0,
                "renumbered": 0,
                "slides": None,
                "pending": [{"slide_id": None, "number": None}],
            },
            {
                "kind": "content",
                "renderer": None,
                "variant": None,
                "required": 45,
                "satisfied": 45,
                "carried": 40,
                "renumbered": 0,
                "slides": None,
                "pending": [],
            },
        ],
        "warnings": {
            "total": 3,
            "accepted": 1,
            "declared": 1,
            "pending": [
                {
                    "diagnostic_id": "d1a2b3",
                    "code": "element_overlap",
                    "slide_id": "p5",
                    "slide_number": 5,
                    "element_id": "label",
                    "related_element_id": "frame",
                    "message": "重なり",
                }
            ],
        },
        "declared_warnings": [
            {
                "diagnostic_id": "d9",
                "code": "math_not_marked",
                "slide_id": "p8",
                "slide_number": 8,
                "element_id": "refs",
                "related_element_id": None,
                "reason": "文献の年は数式ではない",
                "location": "slides.py:42",
            }
        ],
        "invalid_records": [],
        "next": [record, "slidekit review record --build out/runs/0012 --kind overview ..."],
    }
    assert summarize("review status", status, verbose=True).splitlines() == [
        "確認: 未完了（実行 0012、45 ページ、完全な生成）",
        "  エラー 0 件、警告 3 件（許容 1、宣言 1、未対応 1）",
        "  確認 91 件のうち 86 件が有効（引き継ぎ 78 件、ページ番号だけ変わったもの 2 件）",
        "  未確認:",
        "    表示（powerpoint・標準） 3,5-7（4 ページ）",
        "    一覧画像（powerpoint・標準） 未確認",
        "  未対応の警告:",
        "    p.5 element_overlap label / frame（d1a2b3）",
        "  宣言で許容した警告:",
        "    p.8 math_not_marked refs（理由: 文献の年は数式ではない、宣言: slides.py:42）",
        "次の操作:",
        f"  {record}",
        "  slidekit review record --build out/runs/0012 --kind overview ...",
    ]
    # 既定では、宣言で許容した警告は件数だけを示す。最後の行は合否・残りの件数・終了コードである。
    text = full_summarize("review status", status)
    assert "  宣言で許容した警告: 1 件（一覧は --verbose）" in text.splitlines()
    assert text.splitlines()[-1] == "結果: 未解決の事項あり（エラー 0、未完了の確認 5、未対応の警告 1、終了コード 1）"
    blocked = {
        **status,
        "status": "blocked",
        "exit_code": 2,
        "mode": "draft",
        "exportable": False,
        "reviews": [],
        "blocking": [
            {"code": "input_changed", "message": "入力が変わった", "paths": [str(tmp_path / "slides.py")], "count": 3}
        ],
        "errors": [
            {
                "diagnostic_id": "d2",
                "code": "text_overflow",
                "slide_id": "p2",
                "slide_number": 2,
                "element_id": "body",
                "message": "はみ出し",
            }
        ],
        "warnings": {"total": 0, "pending": []},
        "declared_warnings": [],
        "next": ["slidekit build --project deck.toml"],
    }
    assert summarize("review status", blocked).splitlines() == [
        "確認: 止まっている（実行 0012、45 ページ、下書き）",
        "  エラー 1 件、警告 0 件（未対応 0）",
        "  止めている問題:",
        "    input_changed 入力が変わった（slides.py ほか 2 件）",
        "  エラー:",
        "    p.2 text_overflow body はみ出し",
        "次の操作: slidekit build --project deck.toml",
    ]

    notice = {
        "code": "input_changed",
        "message": "入力が生成後に変わっている。この実行は反映できない",
        "paths": [str(tmp_path / "slides.py")],
        "count": 2,
    }
    recorded = {
        "command": "review record",
        "status": "recorded",
        "exit_code": 0,
        "build": run,
        "batch": "b_01",
        "records": [{}, {}, {}],
        "targets": {"visual:powerpoint:fallback": ["p5", "p6"], "overview:powerpoint": [None]},
        "notices": [notice],
    }
    assert summarize("review record", recorded).splitlines() == [
        "確認の記録: 記録した（3 件、b_01）",
        "実行: out/runs/0012",
        "  表示（powerpoint・互換） 2 件（p5、p6）",
        "  一覧画像（powerpoint） 1 件",
        "注意: 入力が生成後に変わっている。この実行は反映できない（slides.py ほか 1 件）",
        "次の操作: slidekit review status --build out/runs/0012",
    ]
    accepted = {
        "command": "review accept",
        "status": "accepted",
        "exit_code": 0,
        "build": run,
        "batch": "b_02",
        "records": [
            {
                "kind": "accept",
                "code": "element_overlap",
                "slide_id": "p5",
                "element_id": "label",
                "related_element_id": "frame",
            }
        ],
        "skipped": [{"diagnostic_id": "d3", "reason": "資料コードで許容している"}],
        "notices": [],
    }
    assert summarize("review accept", accepted).splitlines() == [
        "警告の許容: 許容した（1 件、b_02）",
        "実行: out/runs/0012",
        "  ページ p5 element_overlap label / frame",
        "記録しなかった警告: 1 件（資料コードで許容している 1）",
        "次の操作: slidekit review status --build out/runs/0012",
    ]
    revoked = {
        "command": "review revoke",
        "status": "revoked",
        "exit_code": 0,
        "build": run,
        "batch": "b_03",
        "revoked": ["r_1", "r_2"],
        "cascaded": [{"run": str(tmp_path / "out/runs/0010"), "record_id": "r_0"}],
        "notices": [{"message": "元の記録の実行がないため、写しだけを取り消した", "record_id": "r_2"}],
    }
    assert summarize("review revoke", revoked).splitlines() == [
        "記録の取り消し: 取り消した（2 件、b_03）",
        "実行: out/runs/0012",
        "  取り消した記録: r_1、r_2",
        "  元の記録も取り消した: out/runs/0010 の r_0",
        "注意: 元の記録の実行がないため、写しだけを取り消した（r_2）",
        "次の操作: slidekit review status --build out/runs/0012",
    ]
    carried = {
        "command": "review carry",
        "status": "carried",
        "exit_code": 0,
        "build": run,
        "mode": "auto",
        "records": [{}] * 5,
        "sources": {str(tmp_path / "out/runs/0011"): 5},
        "renumbered": ["p4"],
        "skipped": {"表示が引き継ぎ元と大きく異なる": 1},
        "skipped_runs": {"古い形式（schema_version 1）": 2},
        "remaining": {"visual": 1, "overview": 1, "content": 0, "warnings": 0},
        "notices": [],
    }
    assert summarize("review carry", carried).splitlines() == [
        "確認の引き継ぎ: 引き継いだ（5 件）",
        "実行: out/runs/0012",
        "  引き継ぎ元: out/runs/0011 5 件",
        "  ページ番号だけ変わったページ: p4",
        "  引き継がなかった確認: 表示が引き継ぎ元と大きく異なる 1 件",
        "  候補にしなかった実行: 古い形式（schema_version 1） 2 件",
        "  未確認: 表示 1、一覧画像 1、内容 0、警告 0",
        "次の操作: slidekit review status --build out/runs/0012",
    ]
    manual = {
        **carried,
        "skipped": [{"source_record_id": "r_1", "reason": "すでに引き継いでいる"}] * 2,
        "sources": None,
        "renumbered": [],
        "skipped_runs": None,
        "remaining": None,
    }
    assert summarize("review carry", manual).splitlines()[2] == "  引き継がなかった確認: すでに引き継いでいる 2 件"
    # 数式のページは標準表示と互換表示の 2 件の写しで挙がるが、ページは 1 回だけ示す。
    assert summarize("review carry", {**manual, "renumbered": ["p4", "p4", "p6"]}).splitlines()[2] == (
        "  ページ番号だけ変わったページ: p4、p6"
    )
    assert summarize(
        "review carry", {"status": "skipped", "message": "別の処理が実行中のため引き継がなかった"}
    ).splitlines() == ["確認の引き継ぎ: 省略", "  別の処理が実行中のため引き継がなかった"]


def test_build_summary_shows_carry_and_groups_final_form(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    timings = [
        {"stage": "final_form", "seconds": 0.5},
        {"stage": "math_fallback", "seconds": 0.8},
        {"stage": "render:powerpoint:standard", "seconds": 10.0},
        {"stage": "close:powerpoint", "seconds": 0.4},
    ]
    result = {
        "command": "build",
        "status": "built",
        "build": str(tmp_path / "out/runs/0012"),
        "stage": "complete",
        "mode": "full",
        "selection": None,
        "slides": 3,
        "issues": [],
        "timings": timings,
        "elapsed_seconds": 12.0,
        "exit_code": 0,
        "carry": {
            "status": "carried",
            "count": 6,
            "renumbered": ["a", "b", "a"],
            "remaining": {"visual": 1, "overview": 1, "content": 1, "warnings": 0},
        },
    }
    lines = summarize("build", result).splitlines()
    assert lines[2:4] == [
        "所要時間: 12.0 秒（変換 10.0、互換画像 1.3、表示ソフトの終了 0.4、その他 0.3）",
        "確認の引き継ぎ: 6 件（ページ番号だけ変わったページ 2 件、未確認 表示 1、一覧画像 1、内容 1、警告 0）",
    ]
    # 候補にしなかった実行（形式 1 の実行など）も示す。
    legacy = {
        "status": "carried",
        "count": 0,
        "remaining": {"visual": 82, "overview": 1, "content": 45, "warnings": 9},
        "skipped_runs": {"古い形式（schema_version 1）": 10, "証拠の形式が異なる": 1},
    }
    assert summarize("build", {**result, "carry": legacy}).splitlines()[3] == (
        "確認の引き継ぎ: 0 件（未確認 表示 82、一覧画像 1、内容 45、警告 9、"
        "候補にしなかった実行 古い形式（schema_version 1） 10 件・証拠の形式が異なる 1 件）"
    )
    for carry, line in (
        ({"status": "skipped", "message": "別の処理が実行中"}, "確認の引き継ぎ: 省略（別の処理が実行中）"),
        ({"status": "failed", "message": "指定した実行がない"}, "確認の引き継ぎ: 失敗（指定した実行がない）"),
    ):
        assert line in summarize("build", {**result, "carry": carry}).splitlines()
    assert not any(
        line.startswith("確認の引き継ぎ") for line in summarize("build", {**result, "carry": None}).splitlines()
    )


def test_layout_and_build_summaries_mark_declared_warnings(tmp_path, monkeypatch):
    # slide.allow_warning() で宣言した警告は、件数に内数を示し、行に「宣言で許容」と添える（宣言のない警告と見分ける）。
    monkeypatch.chdir(tmp_path)
    declared = {
        **issue("element_overlap", number=8, element="label", related="chart"),
        "declared": {"reason": "図の空いた所に重ねる", "location": "slides.py:12"},
    }
    plain = issue("text_padding_overlap", number=3, element="body")
    layout = {
        "command": "layout",
        "status": "passed",
        "project": str(tmp_path / "deck.toml"),
        "slides": 10,
        "counts": {"error": 0, "warning": 2, "info": 1},
        "issues": [declared, plain, issue("text_wrap_near_limit", "info", number=5, element="note")],
        "exit_code": 0,
    }
    lines = summarize("layout", layout, verbose=True).splitlines()
    assert lines[0] == "配置の検査: 合格（エラー 0 件、警告 2 件、うち宣言で許容 1 件、参考 1 件）"
    assert lines[2:6] == [
        "警告と参考:",
        "  p.3 警告 text_padding_overlap body text_padding_overlap の文言",
        "  p.5 参考 text_wrap_near_limit note text_wrap_near_limit の文言",
        "  p.8 警告 element_overlap label / chart element_overlap の文言（宣言で許容）",
    ]
    # 既定では宣言で許容した警告を件数だけにし、新しい警告が埋もれないようにする。
    assert summarize("layout", layout).splitlines()[2:6] == [
        "警告と参考:",
        "  p.3 警告 text_padding_overlap body text_padding_overlap の文言",
        "  p.5 参考 text_wrap_near_limit note text_wrap_near_limit の文言",
        "宣言で許容した警告 1 件（一覧は --verbose）",
    ]
    built = {
        "command": "build",
        "status": "built",
        "build": str(tmp_path / "out/runs/0003"),
        "stage": "complete",
        "mode": "full",
        "selection": None,
        "slides": 10,
        "counts": {"error": 0, "warning": 1, "info": 0},
        "issues": [declared],
        "exit_code": 0,
    }
    assert summarize("build", built).splitlines()[0] == "生成: 完了（エラー 0 件、警告 1 件、うち宣言で許容 1 件）"
    # 宣言がなければ、これまでどおり件数だけを示す。
    assert (
        summarize("layout", {**layout, "issues": [plain]}).splitlines()[0]
        == "配置の検査: 合格（エラー 0 件、警告 2 件、参考 1 件）"
    )


def test_partial_construction_shows_pages_by_identifier(tmp_path, monkeypatch):
    # 資料のコードが選んだ章だけを作った検査では、ページ番号がないので識別子で示す。
    monkeypatch.chdir(tmp_path)
    selection = {
        "spec": "a1,2024",
        "slide_ids": ["a1", "2024"],
        "numbers": None,
        "total_slides": None,
        "construct": "selected",
    }
    result = {
        "command": "layout",
        "status": "failed",
        "project": str(tmp_path / "deck.toml"),
        "slides": 2,
        "selection": selection,
        "exit_code": 1,
        "issues": [
            issue("text_overflow", "error", slide="2024", element="body"),
            issue("text_overflow", "error", slide="a1", element="body"),
        ],
    }
    lines = summarize("layout", result).splitlines()
    assert lines[:4] == [
        "配置の検査: 不合格（エラー 2 件、警告 0 件）",
        "資料: deck.toml（2 ページ（a1、2024。選んだ章だけを構築））",
        "エラー:",
        "  ページ 2024 エラー text_overflow body text_overflow の文言",
    ]
    # 数字だけの識別子は id: を付けて、番号と区別する。
    assert lines[-1] == "次の操作: 編集元を直して slidekit layout --project deck.toml --slides id:2024,a1"


def test_check_guidance_follows_blocking_code(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    base = {
        "run_id": "0012",
        "build": str(tmp_path / "out/runs/0012"),
        "status": "failed",
        "final": True,
        "warnings": [],
        "invalid_records": [],
        "exit_code": 2,
    }
    for code, operation in (
        ("build_incomplete", "slidekit build で生成し直す（生成が完了していない）"),
        ("build_not_exportable", "全ページを通常の slidekit build で生成し直す（下書き・部分生成は完成物にできない）"),
        ("artifact_changed", "slidekit build で生成し直す（入力または成果物が生成記録と一致しない）"),
    ):
        lines = summarize(
            "check", {**base, "issues": [{"code": code, "severity": "error", "message": "止まる"}]}
        ).splitlines()
        assert lines[-1] == f"次の操作: {operation}"


def test_command_failure_summary_lists_message_lines_and_location():
    settings = {
        "command": "layout",
        "status": "failed",
        "exception": "SettingsError",
        "code": "settings_invalid",
        "message": "deck.toml: unknown_key は不明な設定です\ndeck.toml: metadata.subtitel は不明な設定です（候補: subtitle）",
        "exit_code": 2,
    }
    assert summarize("layout", settings).splitlines() == [
        "配置の検査: 失敗",
        "  SettingsError: deck.toml: unknown_key は不明な設定です",
        "    deck.toml: metadata.subtitel は不明な設定です（候補: subtitle）",
    ]
    location = {
        "path": "/opt/slidekit/config.py",
        "relative_path": "slidekit/config.py",
        "line": 540,
        "function": "load_project",
        "code": "raise ValueError(message)",
        "in_project": False,
    }
    failed = {
        "command": "review status",
        "status": "failed",
        "exception": "ValueError",
        "message": "参照範囲がない",
        "exit_code": 2,
        "issues": [{"code": "build_failed", "severity": "error", "observations": [{"location": location}]}],
    }
    assert summarize("review status", failed).splitlines() == [
        "確認: 失敗",
        "  ValueError: 参照範囲がない",
        "  場所: slidekit/config.py:540（load_project、slidekit の内部で発生）",
        "      raise ValueError(message)",
        "  トレースバック: 標準エラーに出力",
    ]
    assert summarize(
        "build", {"status": "failed", "exception": "KeyboardInterrupt", "message": "処理を中断しました"}
    ) == ("生成: 中断\n  処理を中断しました")


def test_build_summary_explains_carry_sources_reasons_and_environment(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    carry = {
        "status": "carried",
        "count": 40,
        "sources": {str(tmp_path / "out/runs/0133"): 38, str(tmp_path / "out/runs/0120"): 2},
        "renumbered": ["a", "b"],
        "renumbered_from": {"0133": 2},
        "remaining": {"visual": 2, "overview": 1, "content": 3, "warnings": 0},
        "pending_reasons": [
            {"kind": "visual", "slide_id": "c", "reason": "表示環境の変更"},
            {"kind": "visual", "slide_id": "d", "reason": "ページの表示の変更"},
            {"kind": "content", "slide_id": "c", "reason": "本文の変更"},
            {"kind": "content", "slide_id": "e", "reason": "元資料の変更", "documents": ["proofs"]},
            {"kind": "content", "slide_id": "f", "reason": "元資料の変更", "documents": ["proofs"]},
        ],
        "compared_run": "0133",
        "environment_changes": [{"item": "pypdfium2", "before": "5.13.0", "after": "5.14.0"}],
    }
    lines = summarize("build", {"status": "built", "stage": "complete", "carry": carry, "exit_code": 0}).splitlines()
    assert (
        "確認の引き継ぎ: 40 件（引き継ぎ元 0133 38 件、0120 2 件、ページ番号だけ変わったページ 2 件（実行 0133 と比較）、"
        "未確認 表示 2、一覧画像 1、内容 3、警告 0）"
    ) in lines
    assert (
        "  引き継がなかった理由: 表示: 表示環境の変更 1、ページの表示の変更 1。内容: 本文の変更 1、元資料の変更 2（proofs）"
    ) in lines
    assert "  表示環境の変更（実行 0133 から）: pypdfium2 5.13.0 → 5.14.0" in lines


def test_review_status_groups_pending_pages_by_reason():
    status = {
        "status": "incomplete",
        "exit_code": 1,
        "reviews": [
            {
                "kind": "content",
                "renderer": None,
                "variant": None,
                "required": 7,
                "satisfied": 3,
                "carried": 3,
                "renumbered": 0,
                "slides": "3,5-7",
                "pending": [
                    {"slide_id": "c", "number": 3, "reason": "本文の変更"},
                    {"slide_id": "e", "number": 5, "reason": "本文の変更"},
                    {"slide_id": "f", "number": 6, "reason": "元資料の変更", "documents": ["proofs"]},
                    {"slide_id": "g", "number": 7, "reason": "元資料の変更", "documents": ["proofs"]},
                ],
            }
        ],
        "warnings": {"total": 0, "pending": []},
        "next": [],
    }
    assert (
        "    内容 3,5-7（4 ページ。本文の変更 p.3,5、元資料の変更 p.6-7（proofs））"
        in summarize("review status", status).splitlines()
    )
