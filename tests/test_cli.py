"""コマンドの引数・振り分け・出力・終了コードを確かめる。

確認記録の操作は `test_records.py` の合成した実行で、配置の検査は小さな資料で確かめる。生成と事前確認は
`pipeline` の関数を差し替え、コマンドが引数をそのまま渡すことを確かめる（生成の流れそのものは `test_pipeline.py`）。
要約の文言の細部は `test_report.py` が確かめる。
"""

import argparse
import json
import shlex

import pytest

from slidekit import cli, pipeline, records
from slidekit.cli import main
from test_pipeline import DECK, SLIDES
from test_records import complete, fixture_run, warning, set_diagnostics

REVIEWER = ["--reviewer", "検査担当"]


def review_records(run):
    return json.loads((run / "review.json").read_text(encoding="utf-8"))["records"]


@pytest.fixture
def deck(tmp_path, write_deck, dejavu, dejavu_environment):
    """DejaVu Sans だけを使う 1 ページの資料。"""
    return write_deck(tmp_path / "deck", deck=DECK, slides=SLIDES, theme=dejavu.theme)


def output(capsys):
    captured = capsys.readouterr()
    return captured.out, captured.err


# ---------------------------------------------------------------- 引数と出力


def leaf_parsers(parser, path=()):
    """操作を選ぶ最後のコマンドの（名前、引数の解析器）を並べる。"""
    actions = [action for action in parser._actions if isinstance(action, argparse._SubParsersAction)]
    if not actions:
        yield " ".join(path), parser
        return
    for name, child in actions[0].choices.items():
        yield from leaf_parsers(child, (*path, name))


def test_every_command_accepts_json_and_quiet():
    commands = dict(leaf_parsers(cli.parser()))
    assert set(commands) == {
        "layout",
        "measure",
        "doctor",
        "build",
        "check",
        "export",
        "clean",
        "review status",
        "review sheet",
        "review record",
        "review accept",
        "review revoke",
        "review carry",
        "renderer status",
        "renderer clear",
    }
    for name, parser in commands.items():
        options = {option for action in parser._actions for option in action.option_strings}
        assert {"--json", "--quiet"} <= options, name
    build = {option for action in commands["build"]._actions for option in action.option_strings}
    assert {
        "--project",
        "--slides",
        "--draft",
        "--stop-on-error",
        "--no-stop-on-error",
        "--jobs",
        "--wait",
        "--keep-intermediate",
        "--carry-from",
    } <= build
    args = cli.parser().parse_args(["review", "status", "--json", "--quiet"])
    assert (args.command, args.action, args.json, args.quiet, args.build) == ("review", "status", True, True, "latest")


def test_layout_command_summary_and_json(deck, capsys):
    assert main(["layout", "--project", str(deck)]) == 0
    out, err = output(capsys)
    lines = out.splitlines()
    assert lines[0] == "配置の検査: 合格（エラー 0 件、警告 0 件）" and lines[1].endswith("（1 ページ）")
    assert lines[-2].startswith("次の操作: slidekit build --draft --project ")
    # 最後の行は合否と終了コードである。
    assert lines[-1] == "結果: 合格（エラー 0、終了コード 0）"
    assert main(["layout", "--project", str(deck), "--slides", "intro", "--json"]) == 0
    out, err = output(capsys)
    result = json.loads(out)
    assert (result["command"], result["status"], result["selection"]["slide_ids"]) == ("layout", "passed", ["intro"])
    assert [element["id"] for element in result["layout"]["slides"][0]["elements"]] == ["body"]
    # 表示ソフトを使わず、ファイルを書かない。
    assert not (deck / "out").exists()


def test_default_output_is_summary_and_json_flag_restores_dict(deck, monkeypatch, capsys):
    result = {
        "command": "doctor",
        "status": "passed",
        "project": str(deck / "deck.toml"),
        "slides": 1,
        "issues": [],
        "fallback_required": False,
        "environment": {},
        "renderers": {},
        "exit_code": 0,
    }
    monkeypatch.setattr(pipeline, "doctor", lambda path, **options: result)
    assert main(["doctor", "--project", str(deck)]) == 0
    out, _ = output(capsys)
    with pytest.raises(json.JSONDecodeError):
        json.loads(out)
    assert out.startswith("事前確認: 合格（エラー 0 件、警告 0 件）\n")
    assert main(["doctor", "--project", str(deck), "--json"]) == 0
    out, _ = output(capsys)
    assert json.loads(out) == result


def test_progress_goes_to_stderr(deck, capsys):
    assert main(["layout", "--project", str(deck)]) == 0
    out, err = output(capsys)
    assert "秒] 構築" in err and "秒] 配置の検査" in err
    assert "秒]" not in out
    assert main(["layout", "--project", str(deck), "--quiet"]) == 0
    out, err = output(capsys)
    assert err == "" and out.startswith("配置の検査: 合格")


def test_build_and_doctor_receive_all_options_and_progress(deck, monkeypatch, capsys):
    calls = {}

    def build(path, **options):
        calls["build"] = (path, options)
        options["progress"].event("変換待ち: 別の処理が PowerPoint を使用中（PID 1234、経過 5 秒）")
        return {"command": "build", "status": "built", "build": str(deck / "out/runs/0001"), "exit_code": 0}

    def doctor(path, **options):
        calls["doctor"] = (path, options)
        return {"command": "doctor", "status": "passed", "exit_code": 0}

    monkeypatch.setattr(pipeline, "build", build)
    monkeypatch.setattr(pipeline, "doctor", doctor)
    assert (
        main(
            [
                "build",
                "--project",
                str(deck),
                "--slides",
                "2,intro",
                "--draft",
                "--no-stop-on-error",
                "--jobs",
                "3",
                "--wait",
                "2.5",
                "--keep-intermediate",
                "--carry-from",
                "none",
            ]
        )
        == 0
    )
    path, options = calls["build"]
    assert path == deck and {key: value for key, value in options.items() if key != "progress"} == {
        "slides": "2,intro",
        "draft": True,
        "stop_on_error": False,
        "jobs": 3,
        "wait_seconds": 2.5,
        "keep_intermediate": True,
        "carry_from": "none",
    }
    # 進捗（変換待ちを含む）は標準エラーへ出る。
    _, err = output(capsys)
    assert "変換待ち: 別の処理が PowerPoint を使用中（PID 1234、経過 5 秒）" in err
    assert main(["build", "--project", str(deck)]) == 0
    options = calls["build"][1]
    assert (
        options["slides"],
        options["draft"],
        options["stop_on_error"],
        options["jobs"],
        options["wait_seconds"],
        options["keep_intermediate"],
        options["carry_from"],
    ) == (None, False, None, None, None, False, "auto")
    # 引き継ぎ元の番号・latest・既存の実行ディレクトリは、書式だけを確かめて渡す（latest は生成の前に pipeline が解決する）。
    (deck / "out/runs/0001").mkdir(parents=True)
    for value in ("12", "latest", str(deck / "out/runs/0001")):
        assert main(["build", "--project", str(deck), "--carry-from", value, "--quiet"]) == 0
        assert calls["build"][1]["carry_from"] == value
    assert main(["doctor", "--project", str(deck), "--wait", "0", "--quiet"]) == 0
    path, options = calls["doctor"]
    assert path == deck and options["wait_seconds"] == 0 and options["progress"] is not None


@pytest.mark.parametrize(
    "arguments",
    [
        ["build", "--jobs", "0"],
        ["build", "--jobs", "x"],
        ["build", "--wait", "-1"],
        ["doctor", "--wait", "inf"],
        ["build", "--slides", "3-"],
        ["layout", "--slides", " , "],
        ["clean", "--keep", "-1"],
        ["build", "--carry-from", "atuo"],
        ["build", "--carry-from", "/nonexistent/out/runs/0099"],
    ],
)
def test_invalid_arguments_are_rejected_before_running(deck, monkeypatch, capsys, arguments):
    def refuse(*args, **kwargs):
        raise AssertionError("引数の誤りでは処理を始めない")

    for name in ("build", "doctor", "layout"):
        monkeypatch.setattr(pipeline, name, refuse)
    with pytest.raises(SystemExit) as raised:
        main([arguments[0], "--project", str(deck), *arguments[1:]])
    assert raised.value.code == 2
    assert "slidekit" in capsys.readouterr().err


def test_traceback_goes_to_stderr_even_when_quiet(deck, monkeypatch, capsys):
    # 資料コードの例外：配置の検査は失敗の診断を返し、トレースバックを標準エラーへ出す。
    (deck / "slides.py").write_text("def build(context):\n    raise KeyError('b')\n", encoding="utf-8")
    assert main(["layout", "--project", str(deck), "--quiet"]) == 2
    out, err = output(capsys)
    assert "Traceback" in err and "KeyError" in err
    assert out.startswith("配置の検査: 失敗") and "slides.py:2（build） KeyError: 'b'" in out

    # コマンドの最上位で捕まえた想定外の例外も、トレースバックを標準エラーへ出し、要約に場所を示す。
    def broken(path, **options):
        raise RuntimeError("内部の誤り")

    monkeypatch.setattr(pipeline, "layout", broken)
    assert main(["layout", "--project", str(deck), "--quiet", "--json"]) == 2
    out, err = output(capsys)
    result = json.loads(out)
    assert (result["status"], result["exception"], result["exit_code"]) == ("failed", "RuntimeError", 2)
    assert result["issues"][0]["observations"][0]["location"]["function"] == "broken"
    assert "Traceback" in err and "RuntimeError: 内部の誤り" in err
    assert main(["layout", "--project", str(deck), "--quiet"]) == 2
    out, _ = output(capsys)
    lines = out.splitlines()
    assert lines[:2] == ["配置の検査: 失敗", "  RuntimeError: 内部の誤り"]
    assert lines[2].startswith("  場所: ") and "（broken" in lines[2]


def test_settings_problems_are_listed_without_traceback(deck, capsys):
    settings = deck / "deck.toml"
    settings.write_text(
        "unknown_key = 1\n"
        + settings.read_text(encoding="utf-8").replace("[metadata]\n", '[metadata]\nsubtitel = "x"\n'),
        encoding="utf-8",
    )
    assert main(["layout", "--project", str(deck)]) == 2
    out, err = output(capsys)
    lines = out.splitlines()
    assert lines[0] == "配置の検査: 失敗" and lines[1].startswith("  SettingsError: ")
    # 設定の問題は 1 件 1 行で示す。説明のための例外なので、トレースバックと場所は出さない。
    problems = [line for line in lines[1:] if "deck.toml" in line]
    assert len(problems) >= 2 and any("unknown_key" in line for line in problems)
    assert "Traceback" not in err and not any(line.startswith("  場所: ") for line in lines)
    assert main(["layout", "--project", str(deck), "--json"]) == 2
    out, _ = output(capsys)
    result = json.loads(out)
    assert result["code"] == "settings_invalid" and len(result["problems"]) >= 2 and "issues" not in result


def test_interrupt_ends_with_code_2_and_shows_run(deck, monkeypatch, capsys):
    run = str(deck / "out/runs/0003")

    def interrupted(path, **options):
        error = KeyboardInterrupt()
        error.result = {
            "command": "build",
            "status": "failed",
            "build": run,
            "stage": "render:powerpoint:standard",
            "issues": [{"code": "build_interrupted", "severity": "error", "message": "利用者が生成を中断しました"}],
            "exit_code": 2,
        }
        raise error

    monkeypatch.setattr(pipeline, "build", interrupted)
    assert main(["build", "--project", str(deck)]) == 2
    lines = output(capsys)[0].splitlines()
    assert lines[0] == "生成: 失敗（PowerPoint の変換）" and lines[1].startswith("実行: ") and lines[1].endswith("0003")
    assert "  エラー build_interrupted 利用者が生成を中断しました" in lines

    def stopped(path, **options):
        raise KeyboardInterrupt

    monkeypatch.setattr(pipeline, "layout", stopped)
    assert main(["layout", "--project", str(deck)]) == 2
    assert output(capsys)[0].splitlines() == ["配置の検査: 中断", "  処理を中断しました", "結果: 中断（終了コード 2）"]


def test_renderer_status_without_project_uses_common_environment(
    deck, dejavu, isolated_environment, monkeypatch, capsys
):
    from slidekit import renderers

    isolated_environment.write_text(dejavu.environment + "[renderers.powerpoint]\ntimeout = 33\n", encoding="utf-8")
    seen = []

    def status(name, environment, **options):
        seen.append((name, environment))
        return {
            "command": "renderer status",
            "renderer": name,
            "status": "busy",
            "holder": None,
            "pending": None,
            "exit_code": 1,
        }

    def clear(name, environment, *, force=False, **options):
        seen.append((name, environment, force))
        return {"command": "renderer clear", "renderer": name, "status": "nothing", "reason": "", "exit_code": 0}

    monkeypatch.setattr(renderers, "renderer_status", status)
    monkeypatch.setattr(renderers, "clear_pending", clear)
    assert main(["renderer", "status"]) == 1
    assert output(capsys)[0].startswith("表示ソフトの状態: 使用中（PowerPoint）")
    name, environment = seen[-1]
    assert name == "powerpoint" and environment["renderers"]["powerpoint"]["timeout"] == 33
    # 資料を指定すると、その資料の実行環境（共通の設定を含む）を読む。
    assert main(["renderer", "status", "--project", str(deck), "--renderer", "libreoffice", "--json"]) == 1
    name, environment = seen[-1]
    assert name == "libreoffice" and "DejaVu Sans" in environment["fonts"]
    assert json.loads(output(capsys)[0])["status"] == "busy"
    assert main(["renderer", "clear", "--force"]) == 0
    assert seen[-1][2] is True and output(capsys)[0] == "残存記録の解除: 解除する記録なし\n結果: 成功（終了コード 0）\n"


# ---------------------------------------------------------------- 検査・反映・整理


def test_check_exit_code_follows_result(tmp_path):
    run = fixture_run(tmp_path, warning=True)
    assert main(["check", "--build", str(run)]) == 0
    assert json.loads((run / "check.json").read_text(encoding="utf-8"))["status"] == "passed"
    # 最終検査では、確認と警告への対応が残っている。
    assert main(["check", "--build", str(run), "--final"]) == 1
    # 番号と latest は、--project の資料の実行保存先から探す。
    project = ["--project", str(run.parent.parent.parent)]
    assert main(["check", "--build", "1", *project]) == 0
    assert main(["check", "--build", "latest", *project, "--final"]) == 1
    assert main(["check", "--build", "2", *project]) == 2
    (run / "deck.pptx").write_text("変更", encoding="utf-8")
    assert main(["check", "--build", str(run)]) == 2
    assert main(["check", "--build", str(tmp_path / "missing")]) == 2


def test_check_and_status_list_declared_warnings(tmp_path, capsys):
    run = fixture_run(tmp_path)
    declared = warning(
        "d_declared",
        "intro",
        "label",
        "frame",
        slide_number=1,
        declared={"reason": "ラベルを枠に重ねて示す", "location": "slides.py:42"},
    )
    set_diagnostics(run, [declared])
    complete(run)
    assert main(["check", "--build", str(run), "--final", "--verbose"]) == 0
    out, _ = output(capsys)
    assert (
        "宣言で許容した警告:\n  p.1 element_overlap label / frame（理由: ラベルを枠に重ねて示す、宣言: slides.py:42）"
        in out
    )
    # --verbose を付けなければ件数だけを示す。
    assert main(["check", "--build", str(run), "--final"]) == 0
    assert "宣言で許容した警告: 1 件（一覧は --verbose）" in output(capsys)[0]
    assert main(["review", "status", "--build", str(run), "--verbose"]) == 0
    out, _ = output(capsys)
    assert out.startswith("確認: 完了（実行 0001、1 ページ、完全な生成）\n")
    assert (
        "  宣言で許容した警告:\n    p.1 element_overlap label / frame（理由: ラベルを枠に重ねて示す、宣言: slides.py:42）"
        in out
    )
    assert out.splitlines()[-2] == f"次の操作: slidekit export --build {run}"


def test_export_carry_and_clean_commands(tmp_path):
    first = fixture_run(tmp_path)
    root = first.parent.parent.parent
    assert main(["export", "--build", str(first)]) == 1
    assert not (root / "presentation.pptx").exists()
    complete(first)
    assert main(["export", "--build", "1", "--project", str(root)]) == 0
    assert (root / "presentation.pptx").read_bytes() == (first / "deck.pptx").read_bytes()
    second = fixture_run(tmp_path, "0002")
    assert main(["review", "carry", "--from", str(first), "--to", str(second)]) == 0
    assert sorted(record["kind"] for record in review_records(second)) == ["content", "visual"]
    third = fixture_run(tmp_path, "0003")
    # 自動の引き継ぎ：前の実行から、証拠が一致する確認を写す。--to latest は完成物にできる最新の実行。
    assert main(["review", "carry", "--from", "auto", "--to", "latest", "--project", str(root)]) == 0
    assert sorted(record["kind"] for record in review_records(third)) == ["content", "visual"]
    # 条件のない整理は誤り。反映した実行は残し、それより前の実行だけを候補にする。--apply がなければ消さない。
    with pytest.raises(SystemExit) as raised:
        main(["clean", "--project", str(root)])
    assert raised.value.code == 2
    assert main(["clean", "--project", str(root), "--before", "0003"]) == 0
    assert second.exists()
    assert main(["clean", "--project", str(root), "--before", "0003", "--apply"]) == 0
    assert first.exists() and not second.exists() and third.exists()
    assert main(["clean", "--project", str(root), "--keep", "0", "--json"]) == 0


# ---------------------------------------------------------------- 確認記録の操作


def test_review_commands_record_accept_and_revoke(tmp_path):
    run = fixture_run(tmp_path, warning=True)
    build = ["--build", str(run)]
    assert (
        main(["review", "record", *build, "--kind", "visual", "--slides", "intro", *REVIEWER, "--note", "表示を確認"])
        == 0
    )
    assert main(["review", "record", *build, "--kind", "overview", *REVIEWER, "--note", "一覧を確認"]) == 0
    assert (
        main(["review", "record", *build, "--kind", "content", "--slides", "intro", *REVIEWER, "--note", "内容を確認"])
        == 0
    )
    assert main(["review", "accept", *build, "--diagnostic", "d_overlap", *REVIEWER, "--reason", "意図した重なり"]) == 0
    assert [
        (record["kind"], record.get("slide_id"), record.get("renderer"), record.get("variant"))
        for record in review_records(run)
    ] == [
        ("visual", "intro", "powerpoint", "standard"),
        ("overview", None, "powerpoint", "standard"),
        ("content", "intro", None, None),
        ("accept", "intro", None, None),
    ]
    assert main(["check", *build, "--final"]) == 0
    visual = review_records(run)[0]["id"]
    assert main(["review", "revoke", *build, "--record", visual, *REVIEWER, "--reason", "表示を見直す"]) == 0
    assert [(record["kind"], record.get("record_id")) for record in review_records(run)][-1] == ("revoke", visual)
    assert main(["check", *build, "--final"]) == 1


def test_review_record_passes_page_renderer_and_variant(tmp_path):
    run = fixture_run(tmp_path, pages=("intro", "method"), fallback=("method",))
    command = ["review", "record", "--build", str(run), *REVIEWER, "--note", "表示を確認"]
    # 確認対象にない表示ソフト・互換表示が要らないページ・ないページは、記録せずに失敗する。
    assert main([*command, "--kind", "visual", "--slides", "missing"]) == 2
    assert main([*command, "--kind", "visual", "--slides", "intro", "--variant", "fallback"]) == 2
    assert main([*command, "--kind", "visual", "--slides", "intro", "--renderer", "libreoffice"]) == 2
    assert review_records(run) == []
    # 種類のカンマ区切り、範囲、--variant all。番号は元のページ番号。
    assert main([*command, "--kind", "visual,content", "--slides", "1-2", "--variant", "all", "--json"]) == 0
    assert sorted((record["kind"], record["slide_id"], record.get("variant")) for record in review_records(run)) == [
        ("content", "intro", None),
        ("content", "method", None),
        ("visual", "intro", "standard"),
        ("visual", "method", "fallback"),
        ("visual", "method", "standard"),
    ]
    assert len({record["batch"] for record in review_records(run)}) == 1


def test_record_requires_slides_and_rejects_latest(tmp_path, capsys):
    run = fixture_run(tmp_path)
    root = run.parent.parent.parent
    command = ["review", "record", "--kind", "visual", *REVIEWER, "--note", "表示を確認"]
    assert main([*command, "--build", str(run)]) == 2
    assert "--slides を指定する" in output(capsys)[0]
    assert main([*command, "--build", "latest", "--project", str(root), "--slides", "all"]) == 2
    assert "latest は使えない" in output(capsys)[0]
    assert review_records(run) == []
    assert main([*command, "--build", "1", "--project", str(root), "--slides", "all"]) == 0
    out, _ = output(capsys)
    assert out.splitlines()[0].startswith("確認の記録: 記録した（1 件、b_")
    assert "  表示（powerpoint・標準） 1 件（intro）" in out.splitlines()


def test_review_status_next_commands_need_real_reviewer(tmp_path, capsys):
    # review status が示す記録のコマンドは、確認者と確認内容が置き場所のままでは記録できない。書き換えると記録できる。
    run = fixture_run(tmp_path)
    assert main(["review", "status", "--build", str(run), "--json"]) == 1
    status = json.loads(output(capsys)[0])
    # 先頭は未確認をまとめて記録するコマンド（--pending）。続けて組ごとのコマンドを示す。
    assert status["status"] == "incomplete" and len(status["next"]) == 4
    assert "--pending --status-id " + status["status_id"] in status["next"][0]
    for command in status["next"][1:]:
        tokens = shlex.split(command)
        assert tokens[:3] == ["slidekit", "review", "record"]
        assert main(tokens[1:]) == 2
        replaced = [
            {records.REVIEWER_PLACEHOLDER: "検査担当", records.NOTE_PLACEHOLDER: "画像を開いて確かめた"}.get(
                token, token
            )
            for token in tokens[1:]
        ]
        assert main(replaced) == 0
    capsys.readouterr()
    assert main(["review", "status", "--build", str(run)]) == 0
    assert output(capsys)[0].splitlines()[0] == "確認: 完了（実行 0001、1 ページ、完全な生成）"


def test_accept_by_code_and_revoke_batch(tmp_path, capsys):
    run = fixture_run(tmp_path, pages=("intro", "method"))
    set_diagnostics(
        run,
        [
            warning("d_a", "intro", "box", "label", slide_number=1),
            warning("d_b", "method", "box", "label", slide_number=2),
            warning("d_c", "method", "note", code="math_not_marked", slide_number=2),
        ],
    )
    build = ["--build", str(run)]
    # --diagnostic と --code は同時に指定できない。
    with pytest.raises(SystemExit):
        main(
            [
                "review",
                "accept",
                *build,
                "--diagnostic",
                "d_a",
                "--code",
                "element_overlap",
                *REVIEWER,
                "--reason",
                "理由",
            ]
        )
    capsys.readouterr()
    assert (
        main(
            [
                "review",
                "accept",
                *build,
                "--code",
                "element_overlap",
                "--slides",
                "method",
                *REVIEWER,
                "--reason",
                "意図した重なり",
            ]
        )
        == 0
    )
    lines = output(capsys)[0].splitlines()
    assert lines[0].startswith("警告の許容: 許容した（1 件、b_")
    assert lines[1:3] == [f"実行: {run}", "  ページ method element_overlap box / label"]
    assert [record.get("diagnostic_id") for record in review_records(run)] == ["d_b"]
    assert main(["review", "accept", *build, "--diagnostic", "d_a,d_c", *REVIEWER, "--reason", "確かめた"]) == 0
    capsys.readouterr()
    batch = review_records(run)[-1]["batch"]
    assert main(["review", "revoke", *build, "--record", batch, *REVIEWER, "--reason", "見直す"]) == 0
    out, _ = output(capsys)
    assert out.splitlines()[0] == f"記録の取り消し: 取り消した（2 件、{review_records(run)[-1]['batch']}）"
    revoked = {record["record_id"] for record in review_records(run) if record["kind"] == "revoke"}
    assert revoked == {record["id"] for record in review_records(run) if record.get("batch") == batch}


def test_record_outside_runs_fails_without_lock(tmp_path):
    run = fixture_run(tmp_path)
    copied = tmp_path / "elsewhere" / "0001"
    copied.parent.mkdir()
    run.rename(copied)
    command = [
        "review",
        "record",
        "--build",
        str(copied),
        "--kind",
        "content",
        "--slides",
        "intro",
        *REVIEWER,
        "--note",
        "内容を確認",
    ]
    assert main(command) == 2
    assert not list((tmp_path / "elsewhere").glob(f"*{records.LOCK_NAME}"))
    assert not (tmp_path / records.LOCK_NAME).exists() and review_records(copied) == []


# ---------------------------------------------------------------- 2026-11 の追加


def test_slides_accepts_values_separated_by_spaces(deck, capsys, monkeypatch):
    # --slides 1 2 は --slides 1,2 と同じ指定にする（「unrecognized arguments: 2」で止めない）。
    seen = {}

    def layout(path, *, slides=None, progress=None):
        seen["slides"] = slides
        return {"command": "layout", "status": "passed", "exit_code": 0}

    monkeypatch.setattr(pipeline, "layout", layout)
    assert main(["layout", "--project", str(deck), "--slides", "1", "2-3", "intro"]) == 0
    assert seen["slides"] == "1,2-3,intro"
    with pytest.raises(SystemExit):
        main(["layout", "--project", str(deck), "--slides", "1", "5-3"])
    assert "--slides" in capsys.readouterr().err


def test_check_defaults_to_latest(tmp_path, capsys):
    run = fixture_run(tmp_path)
    root = run.parent.parent.parent
    assert main(["check", "--project", str(root), "--json"]) == 0
    assert json.loads(output(capsys)[0])["build"] == str(run)


def test_review_record_pending_from_command(tmp_path, capsys):
    run = fixture_run(tmp_path, pages=("intro", "method"))
    assert main(["review", "status", "--build", str(run), "--json"]) == 1
    status = json.loads(output(capsys)[0])
    command = shlex.split(status["next"][0])
    assert command[:6] == ["slidekit", "review", "record", "--build", str(run), "--pending"]
    # --pending は --kind・--slides と組み合わせない。--status-id が要る。
    with pytest.raises(SystemExit):
        main(["review", "record", "--build", str(run), "--pending", "--kind", "visual", *REVIEWER, "--note", "見た"])
    with pytest.raises(SystemExit):
        main(["review", "record", "--build", str(run), "--pending", *REVIEWER, "--note", "見た"])
    capsys.readouterr()
    arguments = ["review", "record", "--build", str(run), "--pending", "--status-id", status["status_id"], *REVIEWER]
    assert main([*arguments, "--note", "画像と本文を確認", "--note-content", "依頼と照合"]) == 0
    out = output(capsys)[0].splitlines()
    assert "残り: 表示 0、一覧画像 0、内容 0、警告 0" in out
    assert out[-2] == f"次の操作: slidekit export --build {run}"
    assert {record["note"] for record in review_records(run)} == {"画像と本文を確認", "依頼と照合"}


def test_next_operations_follow_invocation(monkeypatch, tmp_path):
    # python -m slidekit で起動したときは、その Python の絶対パスと -m slidekit で次の操作を書く。
    monkeypatch.setattr(cli.sys, "argv", [str(tmp_path / "slidekit" / "__main__.py")])
    monkeypatch.setattr(cli.sys, "executable", "/opt/env/bin/../bin/python")
    program = cli.invocation()
    assert program == "/opt/env/bin/python -m slidekit"
    text = "次の操作: slidekit build --project deck.toml\n  slidekit の内部で発生"
    assert cli._with_program(text, program) == (
        "次の操作: /opt/env/bin/python -m slidekit build --project deck.toml\n  slidekit の内部で発生"
    )
    # PATH にない slidekit を絶対パスで起動したときは、そのパスにする。
    script = tmp_path / "venv" / "bin" / "slidekit"
    monkeypatch.setattr(cli.sys, "argv", [str(script)])
    monkeypatch.setattr(cli.shutil, "which", lambda name: None)
    assert cli.invocation() == str(script)
    monkeypatch.setattr(cli.shutil, "which", lambda name: str(script))
    assert cli.invocation() == "slidekit"


def test_layout_wraps_lists_wrapped_lines(deck, capsys, monkeypatch):
    result = {
        "command": "layout",
        "status": "passed",
        "exit_code": 0,
        "layout": {
            "slides": [
                {
                    "id": "intro",
                    "number": 1,
                    "elements": [
                        {
                            "id": "lead",
                            "inner_pt": [0, 0, 100, 50],
                            "lines": [
                                {"text": "届いたた", "width_pt": 80.0, "wrap": "mid_phrase"},
                                {"text": "め、次へ", "width_pt": 80.0, "wrap": None},
                            ],
                        },
                        {"id": "one", "lines": [{"text": "一行", "width_pt": 40.0, "wrap": None}]},
                    ],
                }
            ]
        },
    }
    monkeypatch.setattr(pipeline, "layout", lambda path, **options: result)
    assert main(["layout", "--project", str(deck), "--wraps"]) == 0
    lines = output(capsys)[0].splitlines()
    start = lines.index("自動で折り返した要素:")
    assert lines[start + 1 : start + 4] == [
        "  p.1 lead（内側の幅 100.0 pt）",
        "    「届いたた」 80.0 pt ← 文節の途中",
        "    「め、次へ」 80.0 pt",
    ]
    assert main(["layout", "--project", str(deck)]) == 0
    assert "自動で折り返した要素:" not in output(capsys)[0]


def test_measure_command_shows_lines(deck, capsys):
    assert main(["measure", "--project", str(deck), "--width", "60", "AAAA BBBB CCCC", "A"]) == 0
    lines = output(capsys)[0].splitlines()
    assert lines[0] == "計測: 2 件（幅 60.0 pt）"
    assert lines[1].startswith("  1: ") and "1 行に収める枠の幅" in lines[1]
    assert lines[-1] == "結果: 成功（終了コード 0）"
    assert main(["measure", "--project", str(deck), "--json", "--width", "content", "AB"]) == 0
    result = json.loads(output(capsys)[0])
    assert result["width_pt"] > 60 and result["results"][0]["lines"][0]["text"] == "AB"
