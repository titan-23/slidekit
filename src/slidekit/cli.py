"""資料を生成し、確認して固定パスへ反映するコマンド。

既定の出力は要約（`report.summarize`）で、`--json` を付けると結果の辞書を JSON で出す。どちらも標準出力へ 1 回だけ出す。
進捗・資料コードの print・トレースバックは標準エラーへ出す。`--quiet` は進捗だけを止める（トレースバックは出す）。
終了コードは、成功 0、品質上の未解決事項 1、入力・環境・変換の失敗と中断 2 である。
"""

from __future__ import annotations

import argparse
from contextlib import redirect_stdout
import json
import math
import os
from pathlib import Path
import re
import shlex
import shutil
import sys
import traceback

from . import __version__

# ---------------------------------------------------------------- 引数の値


def _positive_integer(text: str) -> int:
    try:
        value = int(text)
    except ValueError:
        value = 0
    if value < 1:
        raise argparse.ArgumentTypeError(f"1 以上の整数を指定してください: {text}")
    return value


def _non_negative_integer(text: str) -> int:
    try:
        value = int(text)
    except ValueError:
        value = -1
    if value < 0:
        raise argparse.ArgumentTypeError(f"0 以上の整数を指定してください: {text}")
    return value


def _seconds(text: str) -> float:
    try:
        value = float(text)
    except ValueError:
        value = math.nan
    if not math.isfinite(value) or value < 0:
        raise argparse.ArgumentTypeError(f"0 以上の秒数を指定してください: {text}")
    return value


def _slides(text: str) -> str:
    """部分生成・部分の検査のページの指定。書式の誤りは構築の前に引数の誤りとして示す。"""
    from .selection import parse

    try:
        parse(text)
    except (TypeError, ValueError) as error:
        raise argparse.ArgumentTypeError(str(error)) from None
    return text


class _JoinedSlides(argparse.Action):
    """--slides の値。空白で区切った複数の値（--slides 1 2）もカンマ区切りの 1 つの指定にまとめる。"""

    def __init__(self, *args, check: bool = True, **kwargs):
        self.check = check
        super().__init__(*args, **kwargs)

    def __call__(self, parser, namespace, values, option_string=None):
        parts = [value.strip(",") for value in values if value.strip(",")]
        text = ",".join(parts)
        if self.check and text not in ("all", "pending"):
            try:
                _slides(text)
            except argparse.ArgumentTypeError as error:
                parser.error(f"{option_string}: {error}")
        setattr(namespace, self.dest, text)


def _slides_option(child: argparse.ArgumentParser, help: str, *, check: bool = True) -> None:
    child.add_argument(
        "--slides", nargs="+", action=_JoinedSlides, check=check, metavar="SPEC", help=help + "。空白で区切ってもよい"
    )


def _measure_width(text: str) -> float | str:
    if text == "content":
        return text
    try:
        value = float(text)
    except ValueError:
        value = math.nan
    if not math.isfinite(value) or value <= 0:
        raise argparse.ArgumentTypeError(f"正の幅（pt）か content を指定してください: {text}")
    return value


def _carry_source(text: str) -> str:
    """build の --carry-from。auto・none・latest・実行の番号・既存の実行ディレクトリのほかは、生成の前に引数の誤りにする。"""
    if text in ("auto", "none", "latest") or (text.isascii() and text.isdigit()) or (text and Path(text).is_dir()):
        return text
    raise argparse.ArgumentTypeError(
        f"auto、none、latest、実行の番号、実行ディレクトリのパスのどれかを指定してください: {text}"
    )


# ---------------------------------------------------------------- 引数


_SLIDES_HELP = "ページ（番号・範囲・識別子・id:識別子のカンマ区切り。例: 3,5-7,intro）"
_BUILD_HELP = "実行（実行ディレクトリのパス、番号、latest）"


def _project(child: argparse.ArgumentParser, *, required: bool, help: str | None = None) -> None:
    child.add_argument(
        "--project", required=required, type=Path, metavar="P", help=help or "deck.toml またはそのディレクトリ"
    )


def _run_project(child: argparse.ArgumentParser) -> None:
    _project(
        child,
        required=False,
        help="実行の番号と latest を探す資料（deck.toml かそのディレクトリ。既定はカレントの deck.toml）",
    )


def parser() -> argparse.ArgumentParser:
    """コマンドの引数。すべての下位コマンドが --json と --quiet を受け付ける。"""
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--json", action="store_true", help="要約の代わりに、結果の辞書を JSON で出す")
    common.add_argument("--quiet", action="store_true", help="進捗を出さない（エラーの詳細は標準エラーへ出す）")
    common.add_argument(
        "--verbose", action="store_true", help="要約に、許容した警告と生成の参考の一覧も出す（既定は件数だけ）"
    )
    command = argparse.ArgumentParser(prog="slidekit", description="編集可能な PowerPoint 資料の生成と確認")
    command.add_argument("--version", action="version", version=f"slidekit {__version__}", help="版を示して終わる")
    commands = command.add_subparsers(dest="command", required=True, metavar="コマンド")

    def add(group, name: str, description: str, *, leaf: bool = True) -> argparse.ArgumentParser:
        # --json と --quiet は操作を選ぶ最後のコマンドにだけ付ける（途中のコマンドにも付けると、後の既定値で上書きされる）。
        return group.add_parser(name, parents=[common] if leaf else [], help=description, description=description)

    child = add(commands, "layout", "配置と記述規則だけを検査する（表示ソフトを使わず、ファイルを書かない）")
    _project(child, required=True)
    _slides_option(child, "検査する" + _SLIDES_HELP)
    child.add_argument(
        "--wraps", action="store_true", help="自動で折り返した要素の、計測した行の文字列と幅を要約に出す"
    )

    child = add(
        commands, "measure", "文字列を資料の書式で計測し、折り返した行と寸法を示す（要素を作らず、ファイルを書かない）"
    )
    _project(child, required=True)
    child.add_argument("text", nargs="+", help="計測する文字列（複数を並べてよい）")
    child.add_argument(
        "--width", type=_measure_width, metavar="PT", help="枠の幅（pt）か content（本文領域の幅）。省くと折り返さない"
    )
    child.add_argument(
        "--kind", default="text", choices=("text", "box", "equation"), help="要素の種類（既定 text。equation は式）"
    )
    child.add_argument("--role", help="役割（既定はテーマの default_role）")
    child.add_argument("--size", type=float, metavar="PT", help="文字サイズ（pt。役割の値を替える）")
    child.add_argument(
        "--page-kind", default="content", help="ページ種別（既定 content。--width content の本文領域に使う）"
    )
    child.add_argument("--markup", action="store_true", help="$…$ を数式、**…** を強調として読む（markup.parse）")

    child = add(commands, "doctor", "配置の検査に加え、出力先・実行環境・表示ソフトを事前に確かめる")
    _project(child, required=True)
    child.add_argument("--wait", type=_seconds, metavar="秒", help="別の変換が終わるのを待つ上限（秒）")

    child = add(commands, "build", "新しい実行を生成する")
    _project(child, required=True)
    _slides_option(child, "生成する" + _SLIDES_HELP + "。完成物にはできない")
    child.add_argument("--draft", action="store_true", help="下書き（互換画像を省く。完成物にはできない）")
    child.add_argument(
        "--stop-on-error",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="配置と入力の検査でエラーがあれば変換の前に止める（既定は下書きでだけ止める）",
    )
    child.add_argument("--jobs", type=_positive_integer, metavar="N", help="PowerPoint を使わない処理の同時プロセス数")
    child.add_argument("--wait", type=_seconds, metavar="秒", help="別の変換が終わるのを待つ上限（秒）")
    child.add_argument("--keep-intermediate", action="store_true", help="互換表示の確認用複製と最終形の PPTX を残す")
    child.add_argument(
        "--carry-from",
        default="auto",
        type=_carry_source,
        metavar="auto|none|RUN",
        help="生成の後に確認を引き継ぐ元（既定 auto：同じ実行保存先の前の実行から。none は引き継がない。"
        "RUN は実行のパスか番号か、生成の前の最新の実行を指す latest）",
    )

    child = add(commands, "check", "生成記録と成果物を検査する")
    child.add_argument("--build", default="latest", metavar="RUN", help="検査する" + _BUILD_HELP + "。既定は latest")
    _run_project(child)
    child.add_argument("--final", action="store_true", help="内容確認と表示確認も必須にする（最終検査）")

    child = add(commands, "export", "確認を終えた完成物を固定パスへ反映する")
    child.add_argument("--build", required=True, metavar="RUN", help="反映する" + _BUILD_HELP)
    _run_project(child)

    child = add(commands, "clean", "古い実行の削除候補を示し、--apply で削除する（条件を 1 つ以上指定する）")
    _project(child, required=True)
    child.add_argument("--before", metavar="RUN", help="この実行の識別子より前の実行")
    child.add_argument(
        "--keep", type=_non_negative_integer, metavar="N", help="条件に当たる実行のうち、新しい N 件を残す"
    )
    child.add_argument("--failed", action="store_true", help="失敗・停止・中断した実行")
    child.add_argument("--drafts", action="store_true", help="下書きと部分生成の実行")
    child.add_argument("--apply", action="store_true", help="候補を実際に削除する")

    review = add(commands, "review", "表示確認・内容確認・警告への対応を記録する", leaf=False)
    actions = review.add_subparsers(dest="action", required=True, metavar="操作")
    child = add(actions, "status", "残りの確認・未対応の警告と、次の操作を示す")
    child.add_argument("--build", default="latest", metavar="RUN", help=_BUILD_HELP + "。既定は latest")
    _run_project(child)
    child = add(actions, "record", "画像や本文を開いて確かめた確認を記録する")
    child.add_argument("--build", required=True, metavar="RUN", help="記録する実行（パスか番号。latest は使えない）")
    _run_project(child)
    child.add_argument("--kind", metavar="K[,K]", help="確認の種類（visual、overview、content のカンマ区切り）")
    child.add_argument(
        "--pending",
        action="store_true",
        help="review status が示した未確認の確認をすべて記録する（--status-id が必要。--kind・--slides は使わない）",
    )
    _slides_option(
        child,
        "確認したページ（all、pending、番号・範囲・識別子のカンマ区切り）。visual と content では必須",
        check=False,
    )
    child.add_argument(
        "--variant", default="standard", choices=("standard", "fallback", "all"), help="表示方法（既定 standard）"
    )
    child.add_argument(
        "--status-id", metavar="ID", help="review status が示した照合用の値（--slides pending のときに必須）"
    )
    child.add_argument("--reviewer", required=True, help="実際に確認した人")
    child.add_argument("--note", help="確かめた内容（--pending で種類ごとの確認内容をすべて書くときは省いてよい）")
    for name, label in (("visual", "表示"), ("overview", "一覧画像"), ("content", "内容")):
        child.add_argument(
            f"--note-{name}", metavar="NOTE", help=f"--pending のときの{label}の確認内容（--note より優先）"
        )
    child = add(actions, "sheet", "ページの画像を前の実行・互換表示と横に並べた比較画像を作る（確認の記録はしない）")
    child.add_argument("--build", default="latest", metavar="RUN", help=_BUILD_HELP + "。既定は latest")
    _run_project(child)
    _slides_option(
        child, "並べるページ（pending（既定。表示の確認が未完了のページ）、all、番号・範囲・識別子）", check=False
    )
    child.add_argument(
        "--variant", default="all", choices=("standard", "fallback", "all"), help="並べる表示方法（既定 all）"
    )
    child.add_argument(
        "--compare",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="前の実行の画像を並べ、差のある領域に枠を描く（既定は描く）",
    )
    child = add(actions, "accept", "警告を理由付きで許容する（エラーは許容できない）")
    child.add_argument("--build", required=True, metavar="RUN", help="記録する実行（パスか番号。latest は使えない）")
    _run_project(child)
    target = child.add_mutually_exclusive_group(required=True)
    target.add_argument("--diagnostic", metavar="ID[,ID]", help="許容する警告の識別子")
    target.add_argument("--code", metavar="CODE", help="許容する警告の種類（--slides と --element で絞れる）")
    _slides_option(child, "--code で許容するページ（番号・範囲・識別子。all と pending は使えない）", check=False)
    child.add_argument("--element", metavar="ID", help="--code で許容する要素か、組の相手の要素の識別子")
    child.add_argument("--reviewer", required=True, help="許容を判断した人")
    child.add_argument("--reason", required=True, help="許容する理由")
    child = add(actions, "revoke", "確認または許容の記録を取り消す")
    child.add_argument("--build", required=True, metavar="RUN", help="記録のある実行（パスか番号。latest は使えない）")
    _run_project(child)
    child.add_argument(
        "--record",
        required=True,
        metavar="ID[,ID]",
        help="取り消す記録の識別子（r_…）か、まとめて作った記録の batch（b_…）",
    )
    child.add_argument("--reviewer", required=True, help="取り消しを判断した人")
    child.add_argument("--reason", required=True, help="取り消す理由")
    child = add(actions, "carry", "前の実行の確認と許容を、証拠が一致する対象へ引き継ぐ")
    child.add_argument(
        "--from", dest="source", required=True, metavar="auto|RUN", help="引き継ぎ元（auto は前の実行から自動で選ぶ）"
    )
    child.add_argument("--to", dest="target", required=True, metavar="RUN|latest", help="引き継ぎ先")
    _run_project(child)

    renderer = add(commands, "renderer", "表示ソフトの排他と残存記録を調べ、解除する", leaf=False)
    actions = renderer.add_subparsers(dest="action", required=True, metavar="操作")
    for name, description in (
        ("status", "排他・残存記録・使用中の処理と、PowerPoint の処理（slidekit が起動したものか）を示す"),
        ("clear", "安全を確かめてから残存記録を解除する"),
    ):
        child = add(actions, name, description)
        _project(child, required=False, help="実行環境を読む資料（省くと共通の実行環境だけを読む）")
        if name == "clear":
            child.add_argument(
                "--force",
                action="store_true",
                help="PowerPoint が起動しているときも解除する（PowerPoint が検査用の複製を開いていないことを確かめてから使う）",
            )
            child.add_argument(
                "--stop-owned",
                action="store_true",
                help="slidekit が起動したと記録（PID と開始時刻）で確かめることができ、ウィンドウのない PowerPoint だけを"
                "終了してから解除する（ほかの PowerPoint は終了しない）",
            )
    return command


# ---------------------------------------------------------------- 実行


def command_name(args: argparse.Namespace) -> str:
    """要約と結果に使うコマンドの名前（"build"、"review record" など）。"""
    action = getattr(args, "action", None)
    return f"{args.command} {action}" if action else args.command


def _run(args: argparse.Namespace, progress) -> dict:
    from . import pipeline, records, runs

    if args.command == "layout":
        return pipeline.layout(args.project, slides=args.slides, progress=progress)
    if args.command == "measure":
        return pipeline.measure(
            args.project,
            args.text,
            width=args.width,
            kind=args.kind,
            role=args.role,
            size=args.size,
            page_kind=args.page_kind,
            markup=args.markup,
        )
    if args.command == "doctor":
        return pipeline.doctor(args.project, wait_seconds=args.wait, progress=progress)
    if args.command == "build":
        return pipeline.build(
            args.project,
            slides=args.slides,
            draft=args.draft,
            stop_on_error=args.stop_on_error,
            jobs=args.jobs,
            wait_seconds=args.wait,
            keep_intermediate=args.keep_intermediate,
            carry_from=args.carry_from,
            progress=progress,
        )
    if args.command == "check":
        return records.check_build(records.resolve_run(args.build, project=args.project), args.final)
    if args.command == "export":
        return records.export_build(records.resolve_run(args.build, project=args.project))
    if args.command == "clean":
        from .config import load_project

        return runs.clean_runs(
            load_project(args.project),
            before=args.before,
            keep=args.keep,
            failed=args.failed,
            drafts=args.drafts,
            apply=args.apply,
        )
    if args.command == "renderer":
        from . import renderers
        from .config import load_common_environment, load_project

        environment = load_project(args.project).environment if args.project is not None else load_common_environment()
        if args.action == "status":
            return renderers.renderer_status(environment)
        return renderers.clear_pending(environment, force=args.force, stop_owned=args.stop_owned)
    if args.action == "status":
        return records.review_status(args.build, project=args.project)
    if args.action == "record" and args.pending:
        return records.record_pending(
            args.build,
            args.reviewer,
            args.note,
            status_id=args.status_id,
            notes={"visual": args.note_visual, "overview": args.note_overview, "content": args.note_content},
            project=args.project,
        )
    if args.action == "record":
        return records.record_review(
            args.build,
            args.kind,
            args.slides,
            args.reviewer,
            args.note,
            args.variant,
            status_id=args.status_id,
            project=args.project,
        )
    if args.action == "sheet":
        return records.review_sheet(
            args.build, args.slides or "pending", variant=args.variant, compare=args.compare, project=args.project
        )
    if args.action == "accept":
        return records.accept_warnings(
            args.build,
            reviewer=args.reviewer,
            reason=args.reason,
            diagnostic_ids=args.diagnostic,
            code=args.code,
            slides=args.slides,
            element=args.element,
            project=args.project,
        )
    if args.action == "revoke":
        return records.revoke_review(args.build, args.record, args.reviewer, args.reason, project=args.project)
    return records.carry_reviews(args.source, args.target, project=args.project)


def _explained(error: BaseException) -> bool:
    """利用者の誤りを説明する例外か（確認記録の操作の拒否、入力と設定の誤り）。トレースバックと場所は示さない。"""
    from .errors import InputError
    from .records import RecordError

    return isinstance(error, (RecordError, InputError))


def _failure(name: str, error: Exception, progress) -> dict:
    """最上位で捕まえた例外の結果。説明のための例外でなければ、トレースバックを標準エラーへ出し、場所を診断に入れる。"""
    code = getattr(error, "code", None)
    result = {
        "command": name,
        "status": "failed",
        "exception": type(error).__name__,
        "message": str(error) or type(error).__name__,
        "code": code if isinstance(code, str) else None,
        "exit_code": 2,
    }
    details = getattr(error, "details", None)
    if isinstance(details, dict) and isinstance(details.get("problems"), list):
        result["problems"] = details["problems"]
    if not _explained(error):
        from .errors import failure_diagnostic

        progress.error("".join(traceback.format_exception(error)))
        result["issues"] = [failure_diagnostic(error, stage="command")]
    return result


def _interrupted(name: str, error: KeyboardInterrupt) -> dict:
    """利用者の中断。生成の中断なら、生成が付けた結果（実行ディレクトリと記録の状態）を使う。"""
    result = getattr(error, "result", None)
    if isinstance(result, dict):
        return {**result, "status": "failed", "message": "処理を中断しました", "exit_code": 2}
    return {
        "command": name,
        "status": "failed",
        "exception": "KeyboardInterrupt",
        "message": "処理を中断しました",
        "exit_code": 2,
    }


_COMMAND_MENTION = re.compile(
    r"(?<![\w./-])slidekit(?= (?:layout|measure|doctor|build|check|export|clean|review|renderer)\b)"
)


def invocation() -> str:
    """実行中の slidekit を呼び出したときの書き方（要約の「次の操作」のコマンドに使う）。

    `python -m slidekit` で起動したときは、その Python の絶対パスと `-m slidekit` にする。コマンドの slidekit が
    PATH で見つかる同じファイルなら `slidekit`、見つからなければ起動したファイルの絶対パスにする。
    """
    argv0 = sys.argv[0] if sys.argv else ""
    if not argv0:
        return "slidekit"
    path = Path(argv0)
    if path.name == "__main__.py":
        # 仮想環境の python は実体へのリンクなので、リンクをたどらずに正規化する（たどると仮想環境の外になる）。
        return f"{shlex.quote(os.path.normpath(os.path.abspath(sys.executable)))} -m slidekit"
    found = shutil.which("slidekit")
    try:
        if found and Path(found).resolve() == path.resolve():
            return "slidekit"
    except OSError:
        pass
    if path.name.startswith("slidekit") and (path.is_absolute() or os.sep in argv0):
        return shlex.quote(os.path.normpath(os.path.abspath(argv0)))
    return "slidekit"


def _with_program(text: str, program: str) -> str:
    """文中の slidekit のコマンドの書き方を、起動の書き方に置き換える。"""
    return _COMMAND_MENTION.sub(lambda _match: program, text) if isinstance(text, str) else text


def main(argv=None) -> int:
    command = parser()
    args = command.parse_args(argv)
    if args.command == "clean" and args.before is None and args.keep is None and not args.failed and not args.drafts:
        command.error("clean には --before、--keep、--failed、--drafts のいずれかを指定してください")
    if args.command == "review" and args.action == "record":
        if args.pending and (args.kind is not None or args.slides is not None):
            command.error("review record の --pending は --kind・--slides と組み合わせない")
        if args.pending and args.status_id is None:
            command.error("review record --pending には review status が示す --status-id が必要です")
        if not args.pending and args.kind is None:
            command.error("review record には --kind か --pending を指定してください")
        if not args.pending and args.note is None:
            command.error("review record には --note を指定してください")
        if not args.pending and any((args.note_visual, args.note_overview, args.note_content)):
            command.error("--note-visual・--note-overview・--note-content は --pending と組み合わせる")
    from .progress import Progress
    from .report import summarize

    name = command_name(args)
    progress = Progress(None if args.quiet else sys.stderr)
    try:
        # 資料コードの print は標準エラーへ回し、標準出力には結果だけを出す。
        with redirect_stdout(sys.stderr):
            result = _run(args, progress)
    except KeyboardInterrupt as error:
        result = _interrupted(name, error)
    except Exception as error:
        result = _failure(name, error, progress)
    program = invocation() if argv is None else "slidekit"
    if args.json:
        if program != "slidekit" and isinstance(result, dict) and isinstance(result.get("next"), list):
            result = {**result, "next": [_with_program(item, program) for item in result["next"]]}
        print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
    else:
        text = summarize(name, result, verbose=args.verbose, wraps=getattr(args, "wraps", False))
        print(_with_program(text, program) if program != "slidekit" else text)
    code = result.get("exit_code", 0) if isinstance(result, dict) else 2
    return code if isinstance(code, int) and not isinstance(code, bool) else 2


if __name__ == "__main__":
    sys.exit(main())
