"""生成記録の検査、確認記録、完成物の反映を扱う。

記録の形式は 2（`SCHEMA_VERSION`）である。形式 1（slidekit 0.1）の実行には、確認の記録・許容・取り消し・引き継ぎ・
検査・反映・残りの確認の一覧を行わず、`LegacyRecordError` にする（再生成してから確認する）。形式 1 の読み取りは、整理
（`runs.protected_runs`）と完成物の退避に要る採用記録・反映の記録と、写しの引き継ぎ元だけに限り、
このモジュールの非公開の関数（`_read_legacy_export_record`・`_legacy_run_references`）に置く。

確認記録は、記録した実行の生成記録と成果物のハッシュ値に結び付ける（binding）。引き継いだ記録（写し）は
人が記録した元の記録だけを指し、表示の確認の写しは元の記録の実行の画像との比較（render_check）を持つ。
記録の読み書きは資料単位のロック（実行保存先の `.slidekit-export.lock`）の中で行う。runs/ の外へ複製した実行には
記録せず、検査だけをロックなしで行う。資料のディレクトリを実行保存先ごと移しても、同じ実行保存先の実行は同じ資料として扱い、
確認を引き継ぐことができる。
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Iterator, Mapping
from contextlib import contextmanager, nullcontext
from datetime import datetime, timezone
from functools import cached_property
import json
import math
import os
from pathlib import Path
import re
import shlex
import shutil
import time
from typing import Any
import uuid

from . import runs, selection
from .evidence import NOISE_PIXELS, RENDER_CHECK_METHOD, difference_boxes, render_difference
from .files import file_hash, json_digest, stage_copy, write_json

SCHEMA_VERSION = 2
# 資料単位のロックを待つ秒数の既定値（_lock の wait_seconds が None のとき）。
LOCK_WAIT_SECONDS = 30
LOCK_NAME = ".slidekit-export.lock"
CHECK_FILE = "check.json"
FINAL_CHECK_FILE = "final_check.json"
# review status が示すコマンドの、確認者・確認内容・理由の置き場所。記録と許容はこの値と < > で囲んだ値を拒否する
# （示したコマンドを画像を見ずにそのまま実行しても、確認を記録しない）。
REVIEWER_PLACEHOLDER = "<確認者>"
NOTE_PLACEHOLDER = "<確認した内容>"
REASON_PLACEHOLDER = "<理由>"

_LEGACY_VERSIONS = (1,)
_LEGACY_RUN = "古い形式（schema_version 1）の実行である。現在の slidekit で再生成してから確認する"
_LOCK_POLL_SECONDS = 0.2
_RUN_NUMBER = re.compile(r"[0-9]+")
_CORE_ARTIFACTS = frozenset({"deck.pptx", "deck.pdf", "deck.json", "notes.md", "diagnostics.json"})
_REVIEW_KINDS = frozenset({"visual", "overview", "content"})
_RECORD_KINDS = _REVIEW_KINDS | {"accept"}
# 種類・表示方法の並び（review status の組の順、記録の順）。
_KIND_ORDER = ("visual", "overview", "content")
_VARIANTS = ("standard", "fallback")
# 写しに、元の記録から引き継がない項目。
_NOT_INHERITED = frozenset({"id", "created_at", "binding", "source", "via", "render_check", "batch"})
# 入力の変化の知らせに並べるパスの数。
_NOTICE_PATHS = 5
_PLACEHOLDERS = frozenset({REVIEWER_PLACEHOLDER, NOTE_PLACEHOLDER, REASON_PLACEHOLDER})
_ENCLOSED = re.compile(r"[<＜].*[>＞]", re.S)
# ページの指定で、識別子と紛れる語と数字の項目（範囲の両端）。
_KEYWORDS = ("all", "pending")
_NUMBER_ITEM = re.compile(r"([0-9]+)(?:\s*-\s*([0-9]+))?")
_SLIDES_REQUIRED = "--slides を指定する（all、番号、範囲、識別子。例: 3,5-7,intro）"
_DIFFERENT_RENDER = "表示が引き継ぎ元と大きく異なる"
_LEGACY_SKIPPED = "古い形式（schema_version 1）"
_ALLOW_WARNING_HINT = "意図した例外なら、資料コードで slide.allow_warning() を宣言して再生成する"


class RecordError(ValueError):
    """記録が破損しているか、操作の前提を満たしていない。code は診断の識別子（既定 record_error）である。"""

    code = "record_error"

    def __init__(self, message: str, *, code: str | None = None) -> None:
        super().__init__(message)
        if code is not None:
            self.code = code


class LegacyRecordError(RecordError):
    """古い形式（schema_version 1）の実行か記録である。確認は再生成した実行で行う。"""

    code = "legacy_run"


class _LockBusy(RecordError):
    """資料単位のロックが、待ち時間を過ぎても使用中だった（別の処理が実行中である）。"""


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _new_id() -> str:
    return f"r_{uuid.uuid4().hex[:16]}"


def _new_batch() -> str:
    return f"b_{uuid.uuid4().hex[:16]}"


# 記録の読み取り


def _version(value: dict) -> int | None:
    version = value.get("schema_version")
    # 真偽値は整数として扱わない（True == 1 を形式 1 と読まない）。
    return version if type(version) is int else None


def _load_json(path: Path) -> dict:
    """JSON の辞書を読む。形式の番号は確かめない。"""
    try:
        value = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise RecordError(f"記録を読めない: {path}: {error}") from error
    if not isinstance(value, dict):
        raise RecordError(f"記録の形式が不正である: {path}")
    return value


def _read_json(path: Path) -> dict:
    """実行ディレクトリの形式 2 の記録を読む。形式 1 は LegacyRecordError、ほかの形式は RecordError。"""
    value = _load_json(path)
    version = _version(value)
    if version == SCHEMA_VERSION:
        return value
    if version in _LEGACY_VERSIONS:
        raise LegacyRecordError(f"{_LEGACY_RUN}: {Path(path).parent}")
    raise RecordError(f"対応していない記録形式: {path}")


def read_record(path: str | Path, *, allow_legacy: bool = False) -> dict:
    """採用記録（deliverables.json）と反映の記録（exports の transaction.json・pending.json）を読む。

    形式 2 はそのまま返す。形式 1 は allow_legacy が真なら `_read_legacy_export_record` で読み（構造は同じ）、
    偽なら LegacyRecordError にする。ほかの形式と、読み込みに失敗した記録は RecordError。
    """
    path = Path(path)
    value = _load_json(path)
    version = _version(value)
    if version == SCHEMA_VERSION:
        return value
    if version in _LEGACY_VERSIONS:
        if allow_legacy:
            return _read_legacy_export_record(path, value)
        raise LegacyRecordError(f"古い形式（schema_version 1）の記録である: {path}")
    raise RecordError(f"対応していない記録形式: {path}")


def _read_legacy_export_record(path: Path, value: dict | None = None) -> dict:
    """形式 1 の採用記録・反映の記録を読む。構造は形式 2 と同じなので、そのまま返す。形式 1 でなければ RecordError。

    value は読み込み済みの内容（読み直さないため）。
    """
    value = _load_json(path) if value is None else value
    if _version(value) not in _LEGACY_VERSIONS:
        raise RecordError(f"形式 1 の記録ではない: {path}")
    return value


def _copy_sources(review: dict) -> list[str]:
    """確認記録の写しが指す引き継ぎ元の実行（記録の値のまま、重複なし）。"""
    sources: list[str] = []
    for record in review["records"]:
        source = record.get("source")
        if isinstance(source, dict) and isinstance(source.get("run"), str) and source["run"] not in sources:
            sources.append(source["run"])
    return sources


def _loose_status(run: Path) -> tuple[dict | None, str | None]:
    """生成記録を形式を問わずに読み、（記録、status）を返す。読めなければ両方 None。"""
    manifest, _ = runs._read_manifest(run)
    if manifest is None:
        return None, None
    status = manifest.get("status")
    return manifest, status if isinstance(status, str) else None


def _legacy_sources(review: dict, run: Path) -> list[str]:
    """形式 1 の確認記録の写しが指す引き継ぎ元（形式 1 の写しは、写した時点の 1 つ前の実行を指す）。形式 1 でなければ RecordError。"""
    if _version(review) not in _LEGACY_VERSIONS:
        raise RecordError(f"形式 1 の確認記録ではない: {run}")
    return _copy_sources(_check_review(review, run))


def _legacy_run_references(run: Path, review: dict | None = None) -> dict:
    """形式 1 の実行の状態と、確認記録の写しが指す引き継ぎ元。review は読み込み済みの確認記録（読み直さないため）。

    戻り値：{"status": str | None, "sources": [str, …]}。確認記録が形式 1 でなければ RecordError。
    """
    run = Path(run)
    review = _load_json(run / "review.json") if review is None else review
    _, status = _loose_status(run)
    return {"status": status, "sources": _legacy_sources(review, run)}


def _review_sources(review: dict, run: Path) -> list[str]:
    """確認記録（形式 2 か 1）の写しが指す引き継ぎ元の実行（_moved_sources）。ほかの形式は RecordError。"""
    if _version(review) not in (SCHEMA_VERSION, *_LEGACY_VERSIONS):
        raise RecordError(f"対応していない記録形式: {run / 'review.json'}")
    return _moved_sources(_check_review(review, run), run)


def _moved_sources(review: dict, run: Path) -> list[str]:
    """確認記録の写しが指す引き継ぎ元の実行（重複なし）。資料を実行保存先ごと移した後の値は、今の場所に読み替える。

    移していない値と相対パスの値は記録の値のまま返す（相対パスは資料のディレクトリから読むので、移しても今の場所を指す）。
    読み替えは `_moved_run` で判断する。review は形を確かめた確認記録（形式 2 か 1）である。
    """
    run = Path(run).resolve()
    cache: dict[Path, _Run] = {}
    result: list[str] = []
    for record in review["records"]:
        source = record.get("source")
        if not isinstance(source, dict) or not isinstance(source.get("run"), str):
            continue
        path = Path(source["run"])
        moved = _moved_run(path.resolve(), run, source.get("record_id"), cache) if path.is_absolute() else None
        text = str(moved) if moved is not None else source["run"]
        if text not in result:
            result.append(text)
    return result


def _run_sources(run: Path) -> list[str]:
    """確認記録の写しが指す引き継ぎ元の実行。確認記録がなければ空。runs.protected_runs が使う。

    値は記録の値のままで、資料を実行保存先ごと移した後の値だけを今の場所に読み替える（移した後も引き継ぎ元を保護するため）。
    runs は生成記録を自分で読むので、ここでは確認記録だけを読む（読み替えが要るときは、同じ名前の実行の確認記録も読む）。
    """
    path = Path(run) / "review.json"
    return _review_sources(_load_json(path), Path(run)) if path.exists() else []


def run_references(run: str | Path) -> dict:
    """実行の状態と、その確認記録の写しが指す引き継ぎ元の実行。

    戻り値：{"status", "exportable", "sources": [引き継ぎ元の実行, …]}。形式 2 の写しは人が記録した
    元の記録の実行を指すので、引き継ぎを重ねても途中の実行は入らない。形式 1 の実行は `_legacy_run_references` で読み、
    写しは 1 つ前の実行を指す（完成物にはできない）。生成記録は形式を問わずに読み、読めなければ status を None にする。
    sources は記録の値のままで、資料を実行保存先ごと移した後の値だけを今の場所に読み替える。
    確認記録がなければ sources は空。確認記録を読めなければ RecordError。
    """
    run = Path(run)
    path = run / "review.json"
    review = _load_json(path) if path.exists() else None
    if review is not None and _version(review) in _LEGACY_VERSIONS:
        references = _legacy_run_references(run, review)
        return {**references, "sources": _moved_sources(review, run), "exportable": False}
    manifest, status = _loose_status(run)
    exportable = bool(
        manifest is not None and _version(manifest) == SCHEMA_VERSION and manifest.get("exportable") is True
    )
    sources = [] if review is None else _review_sources(review, run)
    return {"status": status, "exportable": exportable, "sources": sources}


# 資料単位のロック


def _try_lock(handle) -> bool:
    """ロックの取得を 1 回だけ試み、取得したら真、別の処理が持っていれば偽を返す。handle はバッファを介さずに開いたファイル。"""
    if os.name == "nt":
        import msvcrt

        # LockFile で固定した範囲は、ほかのプロセスからも同じプロセスの別のハンドルからも読むことができないので、読まない。
        # 空かどうかは大きさで調べ、固定する先頭の 1 バイトを用意する。msvcrt.locking は現在の位置から固定する。
        if os.fstat(handle.fileno()).st_size == 0:
            handle.write(b"0")
        handle.seek(0)
        try:
            msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError:
            return False
        return True
    import fcntl

    try:
        fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError:
        return False
    return True


def _unlock(handle) -> None:
    """Windows では、閉じる前にロックを解く（閉じた後に OS が解くまでの時間は資源の状況による）。POSIX では閉じれば OS が解く。"""
    if os.name != "nt":
        return
    import msvcrt

    try:
        handle.seek(0)
        msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
    except OSError:
        # 解けなくても、閉じれば OS が解く。
        pass


@contextmanager
def _lock(path: Path, *, wait_seconds: float | None = None) -> Iterator[None]:
    """停止後にも残るファイルを使い、OS が解放する排他制御を行う。path はロックのファイルで、ディレクトリがなければ作る。

    使用中なら wait_seconds（None なら LOCK_WAIT_SECONDS）秒まで、0.2 秒ごとに取り直す。その間ずっと使用中なら RecordError
    （_LockBusy）。資料単位のロックは `_output_lock` で取る。
    """
    limit = LOCK_WAIT_SECONDS if wait_seconds is None else wait_seconds
    if isinstance(limit, bool) or not isinstance(limit, (int, float)) or not math.isfinite(limit) or limit < 0:
        raise ValueError(f"ロックの待ち時間には 0 以上の秒数を指定する: {limit!r}")
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    # バッファを介さずに開く（seek がファイル記述子の位置を必ず動かすようにする）。
    handle = path.open("a+b", buffering=0)
    locked = False
    try:
        deadline = time.monotonic() + limit
        while not _try_lock(handle):
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                waited = f"（{limit:g} 秒待った）" if limit else ""
                raise _LockBusy(f"別の処理が実行中である{waited}: {path}")
            time.sleep(min(_LOCK_POLL_SECONDS, remaining))
        locked = True
        yield
    finally:
        if locked:
            _unlock(handle)
        # close は正常終了でも異常終了でも OS のロックを解放する。
        handle.close()


@contextmanager
def _output_lock(output: Path, *, wait_seconds: float | None = None) -> Iterator[None]:
    """資料単位のロック。ロックのファイルは実行保存先（output）に置き、資料のディレクトリには置かない。

    確認・許容・取り消し・引き継ぎ・検査・反映・整理は、同じ実行保存先の実行と exports/ だけを読み書きするので、
    実行保存先ごとに排他にする。実行保存先がなければ作る。ロックのファイルを初めて置くときは、runs・exports だけを含む
    （slidekit が作った）実行保存先に .gitignore を置く（runs.new_run と同じ規則。先にロックのファイルがあると、
    生成が .gitignore を置かなくなるため）。
    """
    output = Path(output)
    output.mkdir(parents=True, exist_ok=True)
    if not os.path.lexists(output / LOCK_NAME):
        runs._write_gitignore(output)
    with _lock(output / LOCK_NAME, wait_seconds=wait_seconds):
        yield


# 実行の指定と場所


def _run_path(value: Path) -> Path:
    run = Path(value).resolve()
    if not run.is_dir():
        raise RecordError(f"実行ディレクトリがない: {run}")
    return run


def _project_runs(project: Any) -> tuple[Any, Path]:
    """資料（deck.toml かそのディレクトリか Project。None ならカレントの deck.toml）と、その実行保存先の runs。"""
    if project is None:
        project = Path.cwd() / "deck.toml"
        if not project.is_file():
            raise RecordError("実行の番号や latest を使うには --project で資料を指定する")
    try:
        output = runs._output_dir(project)
    except (OSError, ValueError) as error:
        raise RecordError(f"資料の設定を読めない: {project}: {error}") from error
    return project, output / "runs"


def resolve_run(value: str | Path | int, *, project: Any = None, allow_latest: bool = True) -> Path:
    """実行の指定（パス、番号、latest）を実行ディレクトリにする。

    既存のディレクトリはそのまま（解決したパス）。latest は資料の完成物にできる最新の実行（runs.latest_run）、
    数字だけの指定（12、0012）は資料の実行保存先の runs/0012 である。資料は project（deck.toml かそのディレクトリ）で、
    省略するとカレントディレクトリの deck.toml を使う。資料の設定は output_dir だけを読む（load_project は呼ばない）。
    確認の記録・許容・取り消しは allow_latest=False で呼ぶ（画像を見た後の別の生成で、見ていない実行に記録しないため）。
    """
    text = str(value)
    if text and Path(text).is_dir():
        return Path(text).resolve()
    if text == "latest":
        if not allow_latest:
            raise RecordError(
                "確認の記録・許容・取り消しには実行を明示する（latest は使えない）。review status が示す番号を指定する"
            )
        project, runs_dir = _project_runs(project)
        found = runs.latest_run(project)
        if found is None:
            raise RecordError(f"完成物にできる実行がない: {runs_dir}")
        return found
    if _RUN_NUMBER.fullmatch(text):
        _, runs_dir = _project_runs(project)
        run = runs_dir / f"{int(text):04d}"
        if not run.is_dir():
            raise RecordError(f"指定した実行がない: {run}")
        return run.resolve()
    raise RecordError(f"指定した実行がない: {Path(text)}")


def _artifact_path(run: Path, name: str) -> Path:
    relative = Path(name)
    result = (run / relative).resolve()
    if relative.is_absolute() or not result.is_relative_to(run) or result == run:
        raise RecordError(f"成果物のパスが実行ディレクトリの外を指している: {name}")
    return result


def _project_root(manifest: dict) -> Path:
    return Path(manifest["project_root"]).resolve()


def _record_path(value: str, root: Path) -> Path:
    """記録に書いたパス。相対パスは資料のディレクトリから読む（絶対パスはそのまま）。"""
    path = Path(value)
    return (path if path.is_absolute() else root / path).resolve()


def _relative(path: Path, root: Path) -> str:
    """資料のディレクトリからの相対パス（区切りは /）。相対にできなければ（Windows の別のドライブなど）絶対パス。"""
    try:
        return Path(os.path.relpath(Path(path).resolve(), root)).as_posix()
    except ValueError:
        return str(path)


def _shown(path: Path) -> str:
    """コマンドに書く実行や資料のパス。カレントディレクトリの中なら相対パス、そうでなければ絶対パス。"""
    path = Path(path)
    try:
        text = path.relative_to(Path.cwd().resolve()).as_posix() or "."
    except (OSError, ValueError):
        text = str(path)
    return shlex.quote(text)


def _run_output(run: Path) -> Path | None:
    """実行の今の実行保存先（runs/ の親）。実行が runs/ の下になければ（別の場所へ複製した実行など）None。"""
    run = Path(run)
    return run.parent.parent if run.parent.name == "runs" else None


def _recording_output(run: Path) -> Path:
    """記録の操作（確認・許容・取り消し・引き継ぎ）をする実行の実行保存先。資料単位のロックはここに置く。

    runs/ の外の実行は RecordError にする。関係のない場所にロックのファイルを作らないためで、runs/ の外へ複製した実行の
    記録は、反映にも、実行保存先の実行への引き継ぎにも使うことができない。
    """
    output = _run_output(run)
    if output is None:
        raise RecordError(
            f"実行保存先の runs/ にない実行には記録できない（別の場所へ複製した実行は、記録しても反映も引き継ぎもできない）: {run}"
        )
    return output


def _recorded_output(manifest: Any) -> Path | None:
    """生成記録の実行保存先（生成したときの場所）。記録にないか読めなければ None。"""
    if not isinstance(manifest, dict):
        return None
    config = manifest.get("config")
    value = config.get("output_dir") if isinstance(config, dict) else None
    if not isinstance(value, str) or not value:
        return None
    path = Path(value)
    if not path.is_absolute():
        root = manifest.get("project_root")
        if not isinstance(root, str) or not root:
            return None
        path = Path(root) / path
    return path.resolve()


def _project_place(manifest: dict) -> str | None:
    """実行保存先から見た資料（deck.toml）の場所（区切りは /）。求められなければ None。"""
    output = _recorded_output(manifest)
    project = manifest.get("project")
    if output is None or not isinstance(project, str) or not project:
        return None
    try:
        return Path(os.path.relpath(Path(project).resolve(), output)).as_posix()
    except ValueError:
        return None


def _same_project(first: dict, second: dict) -> bool:
    """同じ実行保存先にある 2 つの実行の生成記録が、同じ資料のものか。

    資料の場所が同じか、実行保存先から見た資料の場所が同じなら同じ資料とする。資料のディレクトリを実行保存先ごと移すと、
    移す前の実行の資料の場所（絶対パス）は変わるが、実行保存先から見た場所は変わらない。
    """
    if first.get("project") == second.get("project"):
        return True
    place = _project_place(first)
    return place is not None and place == _project_place(second)


def _current_project(run: Path, manifest: dict) -> Path:
    """資料（deck.toml）の今の場所。実行保存先から見た資料の場所を、実行の今の実行保存先に当てはめる。"""
    place = _project_place(manifest)
    output = _run_output(run)
    if place is None or output is None:
        return Path(manifest["project"])
    return Path(os.path.normpath(output / place))


def _moved_run(recorded: Path, run: Path, record_id: Any, cache: dict) -> Path | None:
    """資料を実行保存先ごと移した後の、写しの引き継ぎ元の実行の今の場所。移したと判断できなければ None。

    recorded は写しに書いた引き継ぎ元（写したときの絶対パス）、run は写しを持つ実行、record_id は写しが指す元の記録の
    識別子、cache は実行の記録の読み込み（_state）である。写しは同じ runs/ の実行を指すので、recorded が run と同じ runs/ に
    なければ、runs/ ごと移した（か複製した）ものとみなし、run と同じ runs/ にある同じ名前の実行が record_id の記録を
    持つなら、その実行を今の場所とする。識別子は乱数なので、移した回数や、移した後に作った写しかによらず判断できる。
    別の場所の実行にだけある記録は指さない。
    """
    if recorded.parent == run.parent:
        return None
    sibling = run.parent / recorded.name
    if sibling == run or not isinstance(record_id, str) or record_id not in _state(cache, sibling).record_ids:
        return None
    return sibling


def _source_run(source: dict, run: Path, cache: dict) -> Path:
    """写しの source が指す実行の今の場所（資料を実行保存先ごと移した後は、移した先の実行）。"""
    recorded = _record_path(source["run"], _project_root(_state(cache, run).manifest))
    return _moved_run(recorded, run, source.get("record_id"), cache) or recorded


# 実行の記録


def _missing_manifest(run: Path) -> str:
    """生成記録がない実行の説明。生成中の印から、生成中か中断したかを示す。"""
    state = runs.marker_state(run)
    where = f"PID {state['pid']}、段階 {state['stage']}"
    if state["state"] == "running":
        return f"生成中の実行である（{where}）。生成が終わってから操作する: {run}"
    if state["state"] == "interrupted":
        return f"生成が中断された（{where}）。再生成する: {run}"
    if state["state"] == "unknown":
        return f"生成記録がない（{state['reason']}）: {run}"
    return f"生成記録がない: {run}"


def _manifest(run: Path) -> dict:
    """形式 2 の生成記録を読み、記録の操作が使う項目の形を確かめる。"""
    path = run / "manifest.json"
    if not path.exists():
        raise RecordError(_missing_manifest(run))
    manifest = _read_json(path)
    for name, expected in (
        ("artifacts", dict),
        ("input_files", dict),
        ("slides", list),
        ("required_reviews", list),
        ("deliverables", dict),
        ("config", dict),
    ):
        if not isinstance(manifest.get(name), expected):
            raise RecordError(f"生成記録の {name} が不正である: {run}")
    if not isinstance(manifest.get("project"), str) or not isinstance(manifest.get("project_root"), str):
        raise RecordError(f"生成記録に資料の場所がない: {run}")
    if not isinstance(manifest.get("status"), str):
        raise RecordError(f"生成記録の status が不正である: {run}")
    if type(manifest.get("evidence_version")) is not int:
        raise RecordError(f"生成記録の evidence_version が不正である: {run}")
    if manifest.get("mode") not in ("full", "draft"):
        raise RecordError(f"生成記録の mode が不正である: {run}")
    selection_value = manifest.get("selection")
    if selection_value is not None and not isinstance(selection_value, dict):
        raise RecordError(f"生成記録の selection が不正である: {run}")
    exportable = manifest.get("exportable")
    # 完成物にできるのは、全ページの通常の生成だけである。
    if not isinstance(exportable, bool) or exportable and (manifest["mode"] != "full" or selection_value is not None):
        raise RecordError(f"生成記録の exportable が不正である: {run}")
    for slide in manifest["slides"]:
        if not isinstance(slide, dict) or not isinstance(slide.get("id"), str) or not slide["id"]:
            raise RecordError(f"生成記録のページが不正である: {run}")
    for name, digest in manifest["artifacts"].items():
        if not isinstance(name, str) or not isinstance(digest, str):
            raise RecordError(f"成果物の記録が不正である: {run}")
        _artifact_path(run, name)
    for requirement in manifest["required_reviews"]:
        if not isinstance(requirement, dict) or requirement.get("kind") not in _REVIEW_KINDS:
            raise RecordError(f"確認対象の記録が不正である: {run}")
        if not requirement.get("evidence_hash"):
            raise RecordError(f"確認対象の証拠がない: {run}")
    return manifest


def _check_review(review: dict, run: Path) -> dict:
    if not isinstance(review.get("records"), list):
        raise RecordError(f"確認履歴が不正である: {run}")
    identities = set()
    for record in review["records"]:
        if not isinstance(record, dict) or not isinstance(record.get("id"), str):
            raise RecordError(f"確認履歴の識別子が不正である: {run}")
        if record["id"] in identities:
            raise RecordError(f"確認履歴の識別子が重複している: {run}")
        identities.add(record["id"])
    return review


def _review(run: Path) -> dict:
    return _check_review(_read_json(run / "review.json"), run)


def _diagnostics(run: Path) -> list[dict]:
    data = _read_json(run / "diagnostics.json")
    if not isinstance(data.get("issues"), list):
        raise RecordError(f"診断の一覧が不正である: {run}")
    identities = set()
    for issue in data["issues"]:
        if not isinstance(issue, dict) or not issue.get("diagnostic_id") or not issue.get("evidence_hash"):
            raise RecordError(f"診断の識別子または証拠がない: {run}")
        if issue["diagnostic_id"] in identities:
            raise RecordError(f"診断の識別子が重複している: {run}")
        identities.add(issue["diagnostic_id"])
    return data["issues"]


def _issue(code: str, message: str, **details: Any) -> dict:
    return {"code": code, "severity": "error", "message": message, **details}


def _artifact_issues(run: Path, manifest: dict) -> list[dict]:
    issues = []
    if manifest["status"] == "built":
        for missing in sorted(_CORE_ARTIFACTS - set(manifest["artifacts"])):
            issues.append(_issue("artifact_unrecorded", "必須成果物が生成記録にない", path=missing))
    for name, expected in manifest["artifacts"].items():
        try:
            actual = file_hash(_artifact_path(run, name))
        except OSError:
            actual = None
        if actual != expected:
            issues.append(
                _issue("artifact_changed", "成果物が生成時と一致しない", path=name, expected=expected, actual=actual)
            )
    return issues


def _input_kind(name: str, manifest: dict | None) -> str:
    """変わった入力の種類（設定・テーマ・Python・入力文書など）。要約で、どの入力が変わったかを示すのに使う。"""
    if not manifest:
        return "入力"
    config = manifest.get("config") if isinstance(manifest.get("config"), dict) else {}
    inputs = config.get("inputs") if isinstance(config.get("inputs"), dict) else {}
    path = Path(name)
    if name == manifest.get("project") or name in (config.get("extends") or []):
        return "設定"
    if name == config.get("theme"):
        return "テーマ"
    if name == config.get("rules"):
        return "記述規則"
    if name in (inputs.get("python") or []):
        return "Python"
    for document in inputs.get("documents") or []:
        if isinstance(document, dict) and document.get("path") == name:
            return f"入力文書 {document.get('id')}"
    if name in (inputs.get("bibliography") or []):
        return "文献"
    template = config.get("template")
    if isinstance(template, dict) and template.get("path") == name:
        return "雛形"
    environment = manifest.get("environment") if isinstance(manifest.get("environment"), dict) else {}
    if name in (environment.get("config_files") or []):
        return "実行環境"
    fonts = ((environment.get("fonts") or {}).get("fonts") or {}) if isinstance(environment.get("fonts"), dict) else {}
    for face in fonts.values():
        if isinstance(face, dict) and name in (face.get("path"), (face.get("bold") or {}).get("path")):
            return "フォント"
    if ".dist-info" in name:
        return "パッケージ"
    if path.parent == Path(__file__).resolve().parent or Path(__file__).resolve().parent in path.parents:
        return "slidekit"
    for directory in inputs.get("asset_dirs") or []:
        if Path(directory) in path.parents:
            return "素材"
    project = Path(str(manifest.get("project") or "")).parent
    if path.name == "brief.md" and path.parent == project:
        return "依頼内容"
    if path.suffix == ".py":
        return "Python"
    return "その他"


def _input_issues(inputs: Mapping[str, str], manifest: dict | None = None) -> list[dict]:
    issues = []
    for name, expected in inputs.items():
        try:
            actual = file_hash(Path(name))
        except OSError:
            actual = None
        if actual != expected:
            issues.append(
                _issue(
                    "input_changed",
                    "参照入力が生成時と一致しないため、再生成が必要である",
                    path=name,
                    input_kind=_input_kind(name, manifest),
                    expected=expected,
                    actual=actual,
                )
            )
    return issues


def _project_inputs(manifest: dict) -> dict[str, str]:
    """資料のディレクトリと参照範囲（inputs.read_roots）の中の入力。"""
    roots = [Path(manifest["project_root"])]
    inputs = manifest["config"].get("inputs")
    if isinstance(inputs, dict) and isinstance(inputs.get("read_roots"), list):
        roots.extend(Path(root) for root in inputs["read_roots"] if isinstance(root, str))
    return {
        name: digest
        for name, digest in manifest["input_files"].items()
        if any(Path(name).is_relative_to(root) for root in roots)
    }


def _notices(manifest: dict) -> list[dict]:
    """記録はできるが、この実行を反映できないことの知らせ。

    資料の入力（資料のディレクトリと参照範囲の中のファイル）だけを照合する。フォントやパッケージなどの環境のファイルは、
    記録の操作のたびにはハッシュし直さず、check と export が照合する。
    """
    changed = _input_issues(_project_inputs(manifest), manifest)
    if not changed:
        return []
    return [
        {
            "code": "input_changed",
            "message": "入力が生成後に変わっている。この実行は反映できない。記録は再生成後、証拠が一致すれば引き継がれる",
            "paths": [issue["path"] for issue in changed[:_NOTICE_PATHS]],
            "kinds": sorted({issue["input_kind"] for issue in changed}),
            "count": len(changed),
        }
    ]


class _Run:
    """1 回の操作の中で、実行の記録を 1 回だけ読み、ハッシュ値を 1 回だけ求める。"""

    def __init__(self, path: Path):
        self.path = path

    @cached_property
    def manifest(self) -> dict:
        return _manifest(self.path)

    @cached_property
    def review(self) -> dict:
        return _review(self.path)

    @cached_property
    def review_hash(self) -> str:
        return file_hash(self.path / "review.json")

    @cached_property
    def record_ids(self) -> frozenset[str]:
        """確認記録の識別子（形式 2 か 1）。確認記録がないか読めなければ空。写しの引き継ぎ元の読み替え（_moved_run）が使う。"""
        try:
            try:
                review = self.review
            except LegacyRecordError:
                review = _check_review(_load_json(self.path / "review.json"), self.path)
        except RecordError:
            return frozenset()
        return frozenset(record["id"] for record in review["records"])

    @cached_property
    def diagnostics(self) -> list[dict]:
        return _diagnostics(self.path)

    @cached_property
    def artifact_issues(self) -> list[dict]:
        return _artifact_issues(self.path, self.manifest)

    @cached_property
    def input_issues(self) -> list[dict]:
        return _input_issues(self.manifest["input_files"], self.manifest)

    @cached_property
    def binding(self) -> dict:
        return _binding(self.path, self.manifest)

    def with_fresh_reviews(self) -> _Run:
        """確認記録だけを読み直す写し。生成記録・診断・成果物と入力の照合の結果・binding は使い回す。"""
        fresh = _Run(self.path)
        fresh.__dict__.update(
            {
                key: value
                for key, value in self.__dict__.items()
                if key not in {"path", "review", "review_hash", "record_ids"}
            }
        )
        return fresh


def _state(cache: dict, run: Path) -> _Run:
    state = cache.get(run)
    if state is None:
        state = cache[run] = _Run(run)
    return state


def _assert_recordable(state: _Run) -> None:
    """確認・許容・引き継ぎの対象にできる実行か確かめる：成果物が生成記録と一致し、生成が完了し、完成物にできる。

    入力の変化では拒否しない（記録は成果物に結び付くので、_notices で知らせるだけにする）。
    """
    issues = state.artifact_issues
    if issues:
        # 整合検査と同じ識別子（artifact_changed・artifact_unrecorded）で知らせる。
        first = issues[0]
        raise RecordError(f"{first['message']}: {first.get('path', state.path)}", code=first["code"])
    if state.manifest["status"] != "built":
        raise RecordError(f"生成が完了していない: {state.path}", code="build_incomplete")
    if not state.manifest["exportable"]:
        raise RecordError(
            f"下書き・部分生成の実行は、確認記録の対象外である: {state.path}", code="build_not_exportable"
        )


def _target_key(value: dict) -> tuple:
    return tuple(value.get(key) for key in ("kind", "slide_id", "renderer", "variant"))


def _diagnostic_key(value: dict) -> tuple:
    return tuple(value.get(key) for key in ("code", "slide_id", "element_id", "related_element_id", "scope"))


def _binding(run: Path, manifest: dict) -> dict:
    return {"manifest_hash": file_hash(run / "manifest.json"), "artifacts_hash": json_digest(manifest["artifacts"])}


def _record_matches(record: dict, manifest: dict) -> bool:
    return any(
        _target_key(record) == _target_key(target) and record.get("evidence_hash") == target["evidence_hash"]
        for target in manifest["required_reviews"]
    )


def _warning_for(record: dict, diagnostics: list[dict]) -> dict | None:
    """許容の記録と、識別子・対象・証拠が一致する警告。"""
    return next(
        (
            issue
            for issue in diagnostics
            if issue["diagnostic_id"] == record.get("diagnostic_id")
            and issue.get("severity") == "warning"
            and _diagnostic_key(issue) == _diagnostic_key(record)
            and issue["evidence_hash"] == record.get("evidence_hash")
        ),
        None,
    )


def _declared(warning: dict) -> bool:
    """資料コードで許容を宣言した警告（validators.apply_declarations が declared を付ける）。"""
    declared = warning.get("declared")
    return isinstance(declared, dict) and isinstance(declared.get("reason"), str) and bool(declared["reason"].strip())


def _record_problem(state: _Run, record: dict) -> str | None:
    """記録そのものがその実行で有効か確かめる（写しの引き継ぎ元は確かめない）。問題があれば理由を返す。"""
    manifest = state.manifest
    if state.artifact_issues or manifest["status"] != "built":
        return "確認した成果物が生成記録と一致しない"
    if not manifest["exportable"]:
        return "下書き・部分生成の実行の記録である"
    if any(item.get("kind") == "revoke" and item.get("record_id") == record["id"] for item in state.review["records"]):
        return "確認記録が取り消されている"
    if record.get("kind") not in _RECORD_KINDS:
        return "確認記録の種類が不正である"
    if not all(
        isinstance(record.get(key), str) and record[key].strip()
        for key in ("reviewer", "note", "created_at", "evidence_hash")
    ):
        return "確認者、確認内容、時刻または証拠が欠けている"
    if record.get("binding") != state.binding:
        return "確認記録が現在の生成記録を対象にしていない"
    if record["kind"] == "accept":
        if _warning_for(record, state.diagnostics) is None:
            return "許容した警告の証拠が一致しない"
    elif not _record_matches(record, manifest):
        return "確認対象の証拠が一致しない"
    return None


def _visual(manifest: dict, slide_id: str, key: str) -> dict | None:
    """生成記録のページの表示の画像（path と pixel_hash を持つ）。key は "表示ソフト:表示方法"。なければ None。"""
    slide = next((item for item in manifest["slides"] if item["id"] == slide_id), None)
    visuals = slide.get("visuals") if slide is not None else None
    visual = visuals.get(key) if isinstance(visuals, dict) else None
    if isinstance(visual, dict) and isinstance(visual.get("path"), str) and isinstance(visual.get("pixel_hash"), str):
        return visual
    return None


def _footer(manifest: dict, slide_id: str) -> dict | None:
    slide = next((item for item in manifest["slides"] if item["id"] == slide_id), None)
    parts = slide.get("page_parts") if slide is not None else None
    footer = parts.get("footer") if isinstance(parts, dict) else None
    return footer if isinstance(footer, dict) else None


def _same_body(first: dict, second: dict, target: dict) -> bool:
    """2 つの実行の表示の証拠のうち、ページ番号の図形を除いた証拠（visuals の body_hash）が一致するか。"""
    key = f"{target.get('renderer')}:{target.get('variant')}"
    before, after = _visual(first, target.get("slide_id"), key), _visual(second, target.get("slide_id"), key)
    body = (before or {}).get("body_hash")
    return isinstance(body, str) and bool(body) and body == (after or {}).get("body_hash")


def _footer_masks(first: dict, second: dict, slide_id: str) -> list[list[float]]:
    """2 つの実行でページ番号の図形の文字か領域が違うときだけ、両方の領域（pt の [x, y, w, h]）を返す。"""
    footers = [_footer(manifest, slide_id) for manifest in (first, second)]
    if (
        footers[0] is None
        or footers[1] is None
        or (footers[0].get("text"), footers[0].get("area")) == (footers[1].get("text"), footers[1].get("area"))
    ):
        return []
    masks: list[list[float]] = []
    for footer in footers:
        area = footer.get("area")
        if (
            isinstance(area, list)
            and len(area) == 4
            and all(
                isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)
                for value in area
            )
            and min(area[2:]) >= 0
            and [float(value) for value in area] not in masks
        ):
            masks.append([float(value) for value in area])
    return masks


def _render_check(origin: _Run, target: _Run, record: dict) -> dict:
    """表示の確認の写しの照合（render_check）：元の記録の実行の画像と、引き継ぎ先の画像の比較。

    画素のハッシュ値が同じなら比べずに強い差を 0 とする。ページ番号の文字だけが違うときは、番号の領域を除いて比べる。
    """
    key = f"{record.get('renderer')}:{record.get('variant')}"
    before, after = _visual(origin.manifest, record["slide_id"], key), _visual(target.manifest, record["slide_id"], key)
    if before is None or after is None:
        raise RecordError(f"表示の画像が生成記録にない: {record['slide_id']}（{key}）")
    masks: list[list[float]] = []
    if before["pixel_hash"] == after["pixel_hash"]:
        strong = 0
    else:
        masks = _footer_masks(origin.manifest, target.manifest, record["slide_id"])
        strong = render_difference(
            _artifact_path(origin.path, before["path"]), _artifact_path(target.path, after["path"]), masks=masks
        )["strong_pixels"]
    return {
        "method": RENDER_CHECK_METHOD,
        "source_pixel_hash": before["pixel_hash"],
        "target_pixel_hash": after["pixel_hash"],
        "strong_pixels": strong,
        "masked": masks,
    }


def _render_check_problem(record: dict, origin: _Run, target: _Run) -> str | None:
    """写しの render_check が現在の方法と、元の記録の実行と引き継ぎ先の画像に合うか確かめる（画像は読み直さない）。"""
    check = record.get("render_check")
    if not isinstance(check, dict):
        return "表示の照合の記録がない"
    if check.get("method") != RENDER_CHECK_METHOD:
        return "表示の照合の方法が現在と異なる"
    key = f"{record.get('renderer')}:{record.get('variant')}"
    before, after = _visual(origin.manifest, record["slide_id"], key), _visual(target.manifest, record["slide_id"], key)
    if (
        before is None
        or after is None
        or check.get("source_pixel_hash") != before["pixel_hash"]
        or check.get("target_pixel_hash") != after["pixel_hash"]
    ):
        return "表示の照合の記録が画像と一致しない"
    strong = check.get("strong_pixels")
    if type(strong) is not int or not 0 <= strong <= NOISE_PIXELS:
        return _DIFFERENT_RENDER
    return None


def _valid_copy(run: Path, state: _Run, record: dict, cache: dict) -> tuple[bool, str]:
    """写しが指す元の記録と、その実行を確かめる。途中の実行（記録を見つけた実行）は使わない。"""
    source = record["source"]
    if not isinstance(source, dict) or not all(
        isinstance(source.get(key), str) and source[key] for key in ("run", "record_id", "record_hash")
    ):
        return False, "引き継ぎ元の記録の指定が不正である"
    origin = _source_run(source, run, cache)
    if origin == run:
        return False, "引き継ぎ元が同じ実行である"
    if origin.parent != run.parent:
        return False, "引き継ぎ元が同じ実行保存先にない"
    origin_state = _state(cache, origin)
    if origin_state.manifest["evidence_version"] != state.manifest["evidence_version"]:
        return False, "引き継ぎ元と証拠の形式が異なる"
    original = next((item for item in origin_state.review["records"] if item["id"] == source["record_id"]), None)
    if original is None or json_digest(original) != source["record_hash"]:
        return False, "引き継ぎ元の記録が欠落または変更されている"
    if original.get("source") is not None:
        return False, "引き継ぎ元が人の記録した元の記録でない"
    if original.get("kind") == "overview":
        return False, "一覧画像の確認は引き継げない"
    same_target = (
        _diagnostic_key(original) == _diagnostic_key(record)
        if record["kind"] == "accept"
        else _target_key(original) == _target_key(record)
    )
    check = record.get("render_check") if isinstance(record.get("render_check"), dict) else {}
    footer_only = (
        record["kind"] == "visual"
        and check.get("footer_only") is True
        and check.get("source_evidence_hash") == original.get("evidence_hash")
        and bool(check.get("masked"))
        and _same_body(origin_state.manifest, state.manifest, record)
    )
    if (
        original.get("kind") != record["kind"]
        or not same_target
        or (original.get("evidence_hash") != record.get("evidence_hash") and not footer_only)
    ):
        return False, "引き継ぎ元と確認対象が一致しない"
    problem = _record_problem(origin_state, original)
    if problem:
        return False, f"引き継ぎ元: {problem}"
    if record["kind"] == "visual":
        problem = _render_check_problem(record, origin_state, state)
        if problem:
            return False, problem
    return True, ""


def _valid_record(run: Path, record: dict, cache: dict) -> tuple[bool, str]:
    """記録がこの実行で有効か確かめる。写しは、人が記録した元の記録とその実行も確かめる。"""
    try:
        state = _state(cache, run)
        problem = _record_problem(state, record)
        if problem:
            return False, problem
        if record.get("source") is None:
            return True, ""
        return _valid_copy(run, state, record, cache)
    except (OSError, KeyError, TypeError, ValueError) as error:
        return False, f"確認履歴を検証できない: {error}"


def _original(run: Path, record: dict, cache: dict) -> tuple[Path, dict]:
    """記録が写しなら、元の記録の実行と元の記録。写しでなければ、その実行と記録そのもの（_valid_record の後に呼ぶ）。"""
    source = record.get("source")
    if source is None:
        return run, record
    origin = _source_run(source, run, cache)
    original = next(item for item in _state(cache, origin).review["records"] if item["id"] == source["record_id"])
    return origin, original


def _origin_record(run: Path, record: dict, cache: dict) -> tuple[_Run, dict] | None:
    """記録の元の記録（写しなら source が指す人の記録、写しでなければ記録そのもの）とその実行。たどれなければ None。

    元の記録が run と同じ runs/ にあり、写しに書いた記録のハッシュ値と一致し、それ自身が写しでないことだけを確かめる。
    元の記録がその実行で有効かは、使う側が _valid_record で確かめる。
    """
    source = record.get("source")
    if source is None:
        return _state(cache, run), record
    if not isinstance(source, dict) or not all(
        isinstance(source.get(key), str) and source[key] for key in ("run", "record_id", "record_hash")
    ):
        return None
    try:
        origin = _source_run(source, run, cache)
        if origin == run or origin.parent != run.parent:
            return None
        origin_state = _state(cache, origin)
        # 使う側が読む生成記録も、ここで読み込みに成功するかを確かめる（形式 1 などの実行は元にしない）。
        origin_state.manifest
        original = next((item for item in origin_state.review["records"] if item["id"] == source["record_id"]), None)
    except (OSError, KeyError, TypeError, ValueError):
        return None
    if original is None or original.get("source") is not None or json_digest(original) != source["record_hash"]:
        return None
    return origin_state, original


def _copy_of(original: dict, origin: Path, via: Path, binding: dict, check: dict | None) -> dict:
    """元の記録の写し。人が記録した元の記録を指し、記録を見つけた実行を via に残す。"""
    copy = {key: value for key, value in original.items() if key not in _NOT_INHERITED}
    copy.update(
        id=_new_id(),
        created_at=_now(),
        binding=binding,
        source={"run": str(origin), "record_id": original["id"], "record_hash": json_digest(original)},
        via=str(via),
    )
    if check is not None:
        copy["render_check"] = check
    return copy


# 確認の状況


def _numbers(manifest: dict) -> dict[str, int]:
    """ページの識別子から元のページ番号への辞書（番号がなければ位置）。"""
    result = {}
    for position, slide in enumerate(manifest["slides"], 1):
        number = slide.get("number")
        result[slide["id"]] = number if type(number) is int else position
    return result


class _Progress:
    """最終検査と同じ基準で求めた、確認記録の有効性と、確認・警告の対応の状況。ファイルは書かない。

    reviews は（確認対象、それを満たす有効な記録の列）の列、warnings は（警告、"declared"・"accepted"・"pending"、
    許容する有効な記録の列）の列、missing は満たされていない確認対象、unaccepted は未対応の警告である。
    """

    def __init__(self, run: Path, cache: dict):
        state = _state(cache, run)
        manifest = state.manifest
        self.valid: list[dict] = []
        self.invalid: list[dict] = []
        for record in state.review["records"]:
            if record.get("kind") == "revoke":
                continue
            ok, reason = _valid_record(run, record, cache)
            if ok:
                self.valid.append(record)
            else:
                self.invalid.append({"record_id": record["id"], "reason": reason})
        self.reviews = [
            (
                required,
                [
                    record
                    for record in self.valid
                    if record["kind"] != "accept"
                    and _target_key(record) == _target_key(required)
                    and record.get("evidence_hash") == required["evidence_hash"]
                ],
            )
            for required in manifest["required_reviews"]
        ]
        self.warnings: list[tuple[dict, str, list[dict]]] = []
        for warning in state.diagnostics:
            if warning.get("severity") != "warning":
                continue
            if _declared(warning):
                self.warnings.append((warning, "declared", []))
                continue
            matching = [
                record
                for record in self.valid
                if record["kind"] == "accept"
                and record.get("diagnostic_id") == warning["diagnostic_id"]
                and record.get("evidence_hash") == warning["evidence_hash"]
                and _diagnostic_key(record) == _diagnostic_key(warning)
            ]
            self.warnings.append((warning, "accepted" if matching else "pending", matching))
        self.missing = [required for required, matching in self.reviews if not matching]
        self.unaccepted = [warning for warning, status, _ in self.warnings if status == "pending"]
        names: set[str] = set()
        carried = renumbered = 0
        for _, matching in self.reviews:
            names.update(record["reviewer"] for record in matching)
            # 人の記録がなく、写しだけで満たした確認の件数を求める。番号の領域を除いて比べた写しは、
            # 番号だけが変わったページのものである。
            if matching and all(record.get("source") is not None for record in matching):
                carried += 1
                if any((record.get("render_check") or {}).get("masked") for record in matching):
                    renumbered += 1
        for _, _, matching in self.warnings:
            names.update(record["reviewer"] for record in matching)
        self.reviewers = sorted(names)
        statuses = Counter(status for _, status, _ in self.warnings)
        self.counts = {
            "reviews": {
                "required": len(self.reviews),
                "satisfied": len(self.reviews) - len(self.missing),
                "carried": carried,
                "renumbered": renumbered,
            },
            "warnings": {
                "total": len(self.warnings),
                "accepted": statuses["accepted"],
                "declared": statuses["declared"],
                "unaccepted": statuses["pending"],
            },
        }

    def pending(self, kind: str, renderer: str | None, variant: str | None) -> list[dict]:
        """種類・表示ソフト・表示方法の組の、満たされていない確認対象（生成記録の順）。"""
        return [
            required
            for required in self.missing
            if (required.get("kind"), required.get("renderer"), required.get("variant")) == (kind, renderer, variant)
        ]


def _status_id(state: _Run, kind: str, renderer: str | None, variant: str | None, pending: list[dict]) -> str:
    """review status が組ごとに示す照合用の値。未確認のページとその証拠が変わると変わる。"""
    value = [
        state.manifest.get("run_id") or state.path.name,
        state.binding["manifest_hash"],
        kind,
        renderer,
        variant,
        [[item.get("slide_id"), item["evidence_hash"]] for item in pending],
    ]
    return "s_" + json_digest(value)[:16]


def _primary(manifest: dict) -> str:
    """主表示ソフト（設定の render.primary。なければ確認対象から求める）。"""
    render = manifest["config"].get("render")
    primary = render.get("primary") if isinstance(render, dict) else None
    if isinstance(primary, str) and primary:
        return primary
    found = [
        item.get("renderer")
        for item in manifest["required_reviews"]
        if item.get("kind") == "visual" and isinstance(item.get("renderer"), str)
    ]
    return "powerpoint" if not found or "powerpoint" in found else found[0]


def _renderers(manifest: dict, kind: str) -> list[str]:
    """種類の確認対象にある表示ソフト（主表示ソフトを先に、ほかは名前の順）。"""
    primary = _primary(manifest)
    found = {
        item["renderer"]
        for item in manifest["required_reviews"]
        if item.get("kind") == kind and isinstance(item.get("renderer"), str)
    }
    return sorted(found, key=lambda name: (name != primary, name))


def _label(kind: str, renderer: str | None, variant: str | None) -> str:
    if kind == "visual":
        return f"表示（{renderer}・{'互換' if variant == 'fallback' else '標準'}）"
    if kind == "overview":
        return f"一覧画像（{renderer}）"
    return "内容"


def _declared_entry(warning: dict, numbers: Mapping[str, int]) -> dict:
    keys = ("diagnostic_id", "code", "slide_id", "element_id", "related_element_id")
    number = warning.get("slide_number")
    return {
        **{key: warning.get(key) for key in keys},
        "slide_number": number if type(number) is int else numbers.get(warning.get("slide_id")),
        "reason": warning["declared"]["reason"],
        "location": warning["declared"].get("location"),
    }


# 検査


def _evaluate(run: Path, cache: dict, *, final: bool) -> dict:
    """生成記録・成果物・入力・診断と、最終検査なら確認記録を照合した結果。ファイルは書かない。

    check_build・export_build が使う（review_status も同じ _Progress で確認記録を照合する）。cache は 1 回の操作の中で、
    実行ごとの記録とハッシュ値を使い回すための辞書である。最終でない検査は確認記録を照合しないので、summary の件数のうち
    記録で決まるもの（satisfied・carried・renumbered・accepted・unaccepted・reviewers）は None にする。終了コードは、
    確認を加えても通らない問題（成果物・入力の不一致、生成の未完了、完成物にできない実行）があれば 2、ほかの問題があれば 1 である。
    宣言で許容した警告は declared_warnings に、ページ番号・種類・要素・理由・宣言の場所とともに並べる。
    """
    state = _state(cache, run)
    manifest = state.manifest
    blocking = _blocking(state, final=final)
    diagnostics = state.diagnostics
    errors = [issue for issue in diagnostics if issue.get("severity") == "error"]
    warnings = [issue for issue in diagnostics if issue.get("severity") == "warning"]
    declared = [warning for warning in warnings if _declared(warning)]
    issues = [*blocking, *errors]
    reviews: dict[str, Any] = {
        "required": len(manifest["required_reviews"]),
        "satisfied": None,
        "carried": None,
        "renumbered": None,
    }
    counts: dict[str, Any] = {"total": len(warnings), "accepted": None, "declared": len(declared), "unaccepted": None}
    reviewers = None
    invalid_records: list[dict] = []
    if final:
        progress = _Progress(run, cache)
        invalid_records = progress.invalid
        page_numbers = _numbers(manifest)
        issues.extend(
            _issue(
                "review_missing",
                "必要な確認が完了していない",
                target=required,
                **(
                    {"slide_number": page_numbers[required["slide_id"]]}
                    if required.get("slide_id") in page_numbers
                    else {}
                ),
            )
            for required in progress.missing
        )
        issues.extend(
            _issue("warning_unaccepted", "警告が未対応である", diagnostic_id=warning["diagnostic_id"])
            for warning in progress.unaccepted
        )
        reviews, counts, reviewers = progress.counts["reviews"], progress.counts["warnings"], progress.reviewers
    numbers = _numbers(manifest)
    return {
        "command": "check",
        "schema_version": SCHEMA_VERSION,
        "build": str(run),
        "run_id": manifest.get("run_id"),
        "status": "failed" if issues else "passed",
        "final": bool(final),
        "checked_at": _now(),
        "manifest_hash": state.binding["manifest_hash"],
        "review_hash": state.review_hash,
        "issues": issues,
        "warnings": warnings,
        "declared_warnings": [_declared_entry(warning, numbers) for warning in declared],
        "invalid_records": invalid_records,
        "summary": {"reviews": reviews, "warnings": counts, "errors": len(errors), "reviewers": reviewers},
        "exit_code": 2 if blocking else (1 if issues else 0),
    }


def _blocking(state: _Run, *, final: bool) -> list[dict]:
    """確認を加えても通らない問題：成果物・入力の不一致、生成の未完了、（最終検査では）完成物にできない実行。"""
    blocking = [*state.artifact_issues, *state.input_issues]
    if state.manifest["status"] != "built":
        blocking.append(_issue("build_incomplete", "生成が完了していない"))
    if final and not state.manifest["exportable"]:
        blocking.append(
            _issue("build_not_exportable", "下書き・部分生成の実行は完成物にできない。全ページを通常の生成で作り直す")
        )
    return blocking


def check_build(run: Path, final: bool = False) -> dict:
    """保存済みの成果物を検査し、結果を check.json（最終検査は final_check.json）へ保存する。

    runs/ の外の実行（別の場所へ複製したもの）は、生成・整理・反映のどれも扱わないので、ロックを取らずに検査する
    （関係のない場所にロックのファイルを作らない）。
    """
    run = _run_path(run)
    output = _run_output(run)
    with _output_lock(output) if output is not None else nullcontext():
        result = _evaluate(run, {}, final=final)
        write_json(run / (FINAL_CHECK_FILE if final else CHECK_FILE), result)
    return result


# 確認の記録


def _new_record(kind: str, reviewer: str, note: str, binding: dict | None, **fields: Any) -> dict:
    if not isinstance(reviewer, str) or not reviewer.strip():
        raise RecordError("確認者を指定する必要がある")
    if not isinstance(note, str) or not note.strip():
        raise RecordError("確認内容または理由を指定する必要がある")
    return {
        "id": _new_id(),
        "kind": kind,
        "reviewer": reviewer,
        "note": note,
        "created_at": _now(),
        "binding": binding,
        **fields,
    }


def _written(value: Any, label: str, instruction: str) -> None:
    """確認者・確認内容・理由の値を確かめる。空の値と、review status が示した置き場所のままの値（< > で囲んだ値）を拒否する。"""
    if not isinstance(value, str) or not value.strip():
        raise RecordError(f"{label}を指定する必要がある")
    text = value.strip()
    if text in _PLACEHOLDERS or _ENCLOSED.fullmatch(text):
        raise RecordError(f"{label}が置き場所のままである: {text}。{instruction}")


def _kinds(kind: Any) -> list[str]:
    """確認の種類（"visual"、"visual,content"、またはその列）を、重複のない列にする。"""
    if isinstance(kind, str):
        values = kind.split(",")
    elif isinstance(kind, (list, tuple)):
        values = list(kind)
    else:
        raise RecordError(f"確認の種類が不正である: {kind}")
    result: list[str] = []
    for value in values:
        name = value.strip() if isinstance(value, str) else value
        if name == "":
            continue
        if name not in _REVIEW_KINDS:
            raise RecordError(f"確認の種類が不正である: {value}")
        if name not in result:
            result.append(name)
    if not result:
        raise RecordError("確認の種類を指定する必要がある")
    return result


def _identifiers(value: Any, message: str) -> list[str]:
    """カンマ区切りの識別子か、識別子の列を、重複のない列にする。"""
    values = value.split(",") if isinstance(value, str) else list(value) if isinstance(value, (list, tuple)) else None
    if values is None:
        raise RecordError(message)
    result: list[str] = []
    for item in values:
        if not isinstance(item, str):
            raise RecordError(message)
        if item.strip() and item.strip() not in result:
            result.append(item.strip())
    if not result:
        raise RecordError(message)
    return result


def _check_spec(spec: str, manifest: dict) -> None:
    """ページの指定が識別子と紛れないか確かめる。

    数字だけの項目（範囲の両端を含む）と all・pending が既存のページの識別子と同じならエラーにする（selection は数字を常に番号と
    読むので、識別子のつもりで別のページに記録しないため）。all・pending はほかの項目と組み合わせない。
    """
    ids = {slide["id"] for slide in manifest["slides"]}
    items = [part.strip() for part in spec.split(",") if part.strip()]
    if len(items) > 1 and any(item in _KEYWORDS for item in items):
        raise RecordError("all と pending は、ほかの項目と組み合わせない")
    for item in items:
        if item in _KEYWORDS and item in ids:
            raise RecordError(
                f"指定の語と同じ名前のページがある: {item}。そのページは id:{item} と書き、全ページは番号の範囲で書く"
            )
        match = _NUMBER_ITEM.fullmatch(item)
        for value in match.groups() if match else ():
            # "03" は番号 3 と読むので、識別子 "3" とも比べる。
            for name in () if value is None else (value, str(int(value))):
                if name in ids:
                    raise RecordError(f"番号と同じ名前のページがある: {name}。識別子なら id:{name} と書く")


def _select_slides(spec: str, manifest: dict) -> list[str]:
    """番号・範囲・識別子・id:識別子の指定（selection の書式）を、ページの識別子の列（資料の順）にする。

    番号は元のページ番号である。all と pending は呼び出し側が扱う。識別子と紛れる指定は _check_spec でエラーにし、
    selection の誤りは RecordError にする。
    """
    text = spec.strip()
    _check_spec(text, manifest)
    try:
        return [slide["id"] for slide in selection.select(manifest["slides"], text)]
    except (TypeError, ValueError) as error:
        raise RecordError(str(error)) from error


def _groups_for(manifest: dict, kinds: list[str], renderer: str | None, variant: str) -> list[tuple]:
    """記録する（種類、表示ソフト、表示方法）の組。renderer の None は主表示ソフト、all は確認対象にある全表示ソフトである。"""
    groups = []
    for kind in kinds:
        if kind == "content":
            groups.append(("content", None, None))
            continue
        if kind == "overview" and variant == "fallback":
            raise RecordError("一覧画像の確認は標準表示だけである。--variant fallback と組み合わせない")
        renderers = _renderers(manifest, kind) if renderer == "all" else [renderer or _primary(manifest)]
        variants = ["standard"] if kind == "overview" else (list(_VARIANTS) if variant == "all" else [variant])
        groups.extend((kind, name, value) for name in renderers for value in variants)
    return groups


def _not_needed(kind: str, groups: list[tuple], missing: list[str], numbers: Mapping[str, int]) -> str:
    pages = selection.format_numbers(numbers[slide_id] for slide_id in missing)
    if kind == "visual" and all(group[2] == "fallback" for group in groups):
        return (
            f"互換表示の確認が要らないページが含まれている: {pages}（数式がない）。all を使うと必要なページだけを選ぶ"
        )
    labels = "、".join(_label(*group) for group in groups)
    return f"{labels}の確認が要らないページが含まれている: {pages}。all を使うと必要なページだけを選ぶ"


def _record_targets(
    run: Path, cache: dict, kinds: list[str], slides: Any, renderer: str | None, variant: str, status_id: str | None
) -> list[dict]:
    """記録する確認対象（生成記録の required_reviews の項目）の列。ロックの中で呼ぶ。"""
    state = _state(cache, run)
    manifest = state.manifest
    page_kinds = [kind for kind in kinds if kind != "overview"]
    text = slides.strip() if isinstance(slides, str) else None
    if not page_kinds:
        if slides is not None and text != "all":
            raise RecordError("一覧画像の確認には --slides を指定しない（指定するなら all）")
    elif slides is None or text == "" or (isinstance(slides, (list, tuple)) and not slides):
        raise RecordError(_SLIDES_REQUIRED)
    if status_id is not None and text != "pending":
        raise RecordError("--status-id は --slides pending と組み合わせる")
    if page_kinds and text in _KEYWORDS:
        _check_spec(text, manifest)
    groups = _groups_for(manifest, kinds, renderer, variant)
    if page_kinds and text == "pending":
        if len(kinds) != 1 or renderer == "all" or variant == "all":
            raise RecordError("pending は種類・表示ソフト・表示方法を一つずつ指定する")
        if not isinstance(status_id, str) or not status_id:
            raise RecordError("pending には review status が示す --status-id が必要である")
        pending = _Progress(run, cache).pending(*groups[0])
        if not pending:
            raise RecordError("未確認のページがない")
        if _status_id(state, *groups[0], pending) != status_id:
            raise RecordError("未確認のページが review status の後に変わった。review status をやり直す")
        return pending
    required = {_target_key(item): item for item in manifest["required_reviews"]}
    numbers = _numbers(manifest)
    selected: list[str] | None = None
    if isinstance(slides, (list, tuple)):
        # 列の要素は識別子として扱う（番号として読まない）。
        wanted = {str(item) for item in slides}
        unknown = sorted(wanted - set(numbers))
        if unknown:
            raise RecordError(f"未知のページが指定された: {unknown}")
        selected = [slide_id for slide_id in numbers if slide_id in wanted]
    elif page_kinds and text != "all":
        selected = _select_slides(text, manifest)
    result = []
    for kind in kinds:
        kind_groups = [group for group in groups if group[0] == kind]
        if kind == "overview":
            result.extend(
                required[key]
                for key in ((kind, None, name, value) for _, name, value in kind_groups)
                if key in required
            )
            continue
        pages = list(numbers) if selected is None else selected
        found: set[str] = set()
        for _, name, value in kind_groups:
            for slide_id in pages:
                item = required.get((kind, slide_id, name, value))
                if item is not None:
                    result.append(item)
                    found.add(slide_id)
        missing = [slide_id for slide_id in pages if slide_id not in found]
        if selected is not None and missing:
            raise RecordError(_not_needed(kind, kind_groups, missing, numbers))
    if not result:
        labels = "、".join(_label(*group) for group in groups)
        raise RecordError(f"指定した確認対象が生成記録にない: {labels}")
    return result


def _target_name(target: dict) -> str:
    if target["kind"] == "content":
        return "content"
    if target["kind"] == "overview":
        return f"overview:{target.get('renderer')}"
    return f"visual:{target.get('renderer')}:{target.get('variant')}"


def record_review(
    run: Any,
    kind: Any,
    slides: Any,
    reviewer: str,
    note: str,
    renderer: str | None = None,
    variant: str = "standard",
    *,
    status_id: str | None = None,
    project: Any = None,
) -> dict:
    """表示・一覧画像・内容の確認を記録する。

    run は実行のパスか番号（latest は指定できない。project は番号を探す資料）。kind は "visual"、"visual,content" のような
    カンマ区切りか、その列。slides は表示・内容の確認では必須で、all（選んだ確認が要る全ページ）、pending（review status が示した
    未確認のページ。status_id が必須で、種類・表示ソフト・表示方法を一つずつ指定する）、番号・範囲・識別子・id:識別子のカンマ区切り、
    または識別子の列である。一覧画像だけなら None か all。renderer は表示ソフト（None は主表示ソフト、all は確認対象にある全表示ソフト）、
    variant は standard・fallback・all である。明示したページにその確認が要らなければエラーにし、all の表示ソフト・表示方法では
    要らない組を飛ばす。確認者と確認内容は、review status が示した置き場所（< > で囲んだ値）のままでは記録しない。
    1 回の呼び出しで作る記録には同じ batch を付ける。

    成果物が生成記録と一致する、完了した全ページの通常の生成にだけ記録できる。入力が生成後に変わっていても記録はでき、
    戻り値の notices で知らせる（その実行は反映できない。記録は再生成後、証拠が一致すれば引き継がれる）。
    """
    kinds = _kinds(kind)
    _written(reviewer, "確認者", "実際に確認した人の名前を書く")
    _written(note, "確認内容", "画像や本文を開いて確かめた内容を書く")
    if variant not in (*_VARIANTS, "all"):
        raise RecordError(f"表示方法が不正である（standard、fallback、all）: {variant}")
    if renderer is not None and (not isinstance(renderer, str) or not renderer.strip()):
        raise RecordError(f"表示ソフトが不正である: {renderer!r}")
    run = _run_path(resolve_run(run, project=project, allow_latest=False))
    with _output_lock(_recording_output(run)):
        cache: dict[Path, _Run] = {}
        state = _state(cache, run)
        _assert_recordable(state)
        notices = _notices(state.manifest)
        targets = _record_targets(run, cache, kinds, slides, renderer, variant, status_id)
        batch = _new_batch()
        records = [
            _new_record(
                target["kind"],
                reviewer,
                note,
                state.binding,
                batch=batch,
                **{key: target.get(key) for key in ("slide_id", "renderer", "variant", "evidence_hash")},
            )
            for target in targets
        ]
        review = state.review
        review["records"].extend(records)
        write_json(run / "review.json", review)
        remaining = _remaining(run, cache)
    return _recorded(run, batch, records, targets, notices, remaining)


def _recorded(run: Path, batch: str, records: list[dict], targets: list[dict], notices: list, remaining: dict) -> dict:
    """確認の記録の結果。記録した対象と、残っている確認・未対応の警告の件数を示す。"""
    grouped: dict[str, list] = {}
    for target in targets:
        grouped.setdefault(_target_name(target), []).append(target.get("slide_id"))
    return {
        "command": "review record",
        "status": "recorded",
        "exit_code": 0,
        "build": str(run),
        "batch": batch,
        "records": records,
        "targets": grouped,
        "remaining": remaining,
        "notices": notices,
    }


def record_pending(
    run: Any,
    reviewer: str,
    note: str | None,
    *,
    status_id: str,
    notes: dict | None = None,
    project: Any = None,
) -> dict:
    """review status が示した未確認の確認（表示・一覧画像・内容のすべての組）を、1 回でまとめて記録する。

    status_id は review status の結果の status_id（未確認の確認とその証拠から作る値）で、review status の後に未確認の対象が
    変わっていれば記録しない。notes は種類（visual・overview・content）ごとの確認内容で、ない種類には note を使う。
    """
    _written(reviewer, "確認者", "実際に確認した人の名前を書く")
    notes = {key: value for key, value in (notes or {}).items() if value is not None}
    for key, value in notes.items():
        if key not in _REVIEW_KINDS:
            raise RecordError(f"確認の種類が不正である（visual、overview、content）: {key}")
        _written(value, "確認内容", "画像や本文を開いて確かめた内容を書く")
    if note is not None:
        _written(note, "確認内容", "画像や本文を開いて確かめた内容を書く")
    if not isinstance(status_id, str) or not status_id:
        raise RecordError("--pending には review status が示す --status-id が必要である")
    run = _run_path(resolve_run(run, project=project, allow_latest=False))
    with _output_lock(_recording_output(run)):
        cache: dict[Path, _Run] = {}
        state = _state(cache, run)
        _assert_recordable(state)
        notices = _notices(state.manifest)
        targets = _Progress(run, cache).missing
        if not targets:
            raise RecordError("未確認の確認がない")
        if _status_id(state, "all", None, None, targets) != status_id:
            raise RecordError("未確認の確認が review status の後に変わった。review status をやり直す")
        lacking = sorted({target["kind"] for target in targets if (notes.get(target["kind"]) or note) is None})
        if lacking:
            raise RecordError("確認内容がない種類がある: " + "、".join(lacking) + "（--note か --note-<種類> で書く）")
        batch = _new_batch()
        records = [
            _new_record(
                target["kind"],
                reviewer,
                notes.get(target["kind"]) or note,
                state.binding,
                batch=batch,
                **{key: target.get(key) for key in ("slide_id", "renderer", "variant", "evidence_hash")},
            )
            for target in targets
        ]
        review = state.review
        review["records"].extend(records)
        write_json(run / "review.json", review)
        remaining = _remaining(run, cache)
    return _recorded(run, batch, records, targets, notices, remaining)


def accept_warnings(
    run: Any,
    *,
    reviewer: str,
    reason: str,
    diagnostic_ids: Any = None,
    code: str | None = None,
    slides: str | None = None,
    element: str | None = None,
    project: Any = None,
) -> dict:
    """警告を理由付きで許容する（エラーは許容できない）。記録できる実行と notices は record_review と同じ。

    diagnostic_ids（識別子のカンマ区切りか列）か code（警告の種類）のどちらか一方を指定する。code では、slides（ページの指定。
    all と pending は指定できない。指定するとページのない警告を除く）と element（要素か組の相手の識別子）で絞る。
    資料コードで許容を宣言した警告と、有効な許容がすでにある警告は記録せず、skipped に理由を入れる。対象が 0 件ならエラー。
    警告 1 件につき 1 記録を作り、同じ batch を付ける。確認者と理由は置き場所（< > で囲んだ値）のままでは記録しない。
    """
    _written(reviewer, "確認者", "実際に確認した人の名前を書く")
    _written(reason, "理由", "警告を許容する理由を書く")
    if (diagnostic_ids is None) == (code is None):
        raise RecordError("警告の識別子か種類のどちらか一方を指定する")
    if diagnostic_ids is not None and (slides is not None or element is not None):
        raise RecordError("--slides と --element は --code と組み合わせる")
    identifiers = None if diagnostic_ids is None else _identifiers(diagnostic_ids, "警告の識別子を指定する必要がある")
    if code is not None and (not isinstance(code, str) or not code.strip()):
        raise RecordError("警告の種類を指定する必要がある")
    run = _run_path(resolve_run(run, project=project, allow_latest=False))
    with _output_lock(_recording_output(run)):
        cache: dict[Path, _Run] = {}
        state = _state(cache, run)
        _assert_recordable(state)
        manifest = state.manifest
        notices = _notices(manifest)
        diagnostics = state.diagnostics
        if identifiers is not None:
            known = {issue["diagnostic_id"]: issue for issue in diagnostics}
            wrong = [
                identifier
                for identifier in identifiers
                if identifier not in known or known[identifier].get("severity") != "warning"
            ]
            if wrong:
                raise RecordError(f"許容できる警告の識別子ではない: {', '.join(wrong)}")
            chosen = [known[identifier] for identifier in identifiers]
        else:
            chosen = [
                issue
                for issue in diagnostics
                if issue.get("severity") == "warning" and issue.get("code") == code.strip()
            ]
            if slides is not None:
                text = slides.strip() if isinstance(slides, str) else ""
                if not text or text in _KEYWORDS:
                    raise RecordError(
                        "警告の許容の --slides には番号・範囲・識別子を指定する（all と pending は使えない）"
                    )
                pages = set(_select_slides(text, manifest))
                chosen = [issue for issue in chosen if issue.get("slide_id") in pages]
            if element is not None:
                chosen = [
                    issue for issue in chosen if element in (issue.get("element_id"), issue.get("related_element_id"))
                ]
        accepted = {
            warning["diagnostic_id"] for warning, status, _ in _Progress(run, cache).warnings if status == "accepted"
        }
        batch = _new_batch()
        records, skipped = [], []
        for warning in chosen:
            if _declared(warning):
                skipped.append({"diagnostic_id": warning["diagnostic_id"], "reason": "資料コードで許容している"})
            elif warning["diagnostic_id"] in accepted:
                skipped.append({"diagnostic_id": warning["diagnostic_id"], "reason": "すでに許容している"})
            else:
                records.append(
                    _new_record(
                        "accept",
                        reviewer,
                        reason,
                        state.binding,
                        batch=batch,
                        **{
                            key: warning.get(key)
                            for key in (
                                "diagnostic_id",
                                "code",
                                "slide_id",
                                "element_id",
                                "related_element_id",
                                "scope",
                                "evidence_hash",
                            )
                        },
                    )
                )
        if not records:
            detail = f"（{len(skipped)} 件は許容済みか、資料コードで許容している）" if skipped else ""
            raise RecordError(f"条件に合う未対応の警告がない{detail}")
        review = state.review
        review["records"].extend(records)
        write_json(run / "review.json", review)
    return {
        "command": "review accept",
        "status": "accepted",
        "exit_code": 0,
        "build": str(run),
        "batch": batch,
        "records": records,
        "skipped": skipped,
        "notices": notices,
    }


def accept_warning(run: Any, diagnostic_id: str, reviewer: str, reason: str) -> dict:
    """警告を 1 件、理由付きで許容する（accept_warnings の識別子 1 件の指定）。record は作った記録である。"""
    result = accept_warnings(run, reviewer=reviewer, reason=reason, diagnostic_ids=[diagnostic_id])
    return {**result, "record": result["records"][0]}


def revoke_review(run: Any, record_id: Any, reviewer: str, reason: str, *, project: Any = None) -> dict:
    """確認または許容の記録を取り消す。成果物や入力の状態によらず受け付ける。

    record_id は記録の識別子（r_…）か batch（b_…）、またはそれらのカンマ区切りか列である。batch はその 1 回のコマンドで
    作った記録すべてを指す。写しを取り消すと、元の記録の実行にも取り消し（cascade に写しの実行と識別子を書く）を加えるので、
    ほかの実行にある同じ元の記録の写しも無効になる。元の記録の実行を先に書く（途中で失敗しても、元の記録が取り消された
    安全な状態になる）。元の記録の実行や元の記録がなければ写しだけを取り消し、notices で知らせる。
    """
    identifiers = _identifiers(record_id, "取り消す確認または許容の識別子を指定する必要がある")
    run = _run_path(resolve_run(run, project=project, allow_latest=False))
    with _output_lock(_recording_output(run)):
        cache: dict[Path, _Run] = {}
        state = _state(cache, run)
        review = state.review
        chosen: list[dict] = []
        unknown = []
        for identifier in identifiers:
            found = [
                record
                for record in review["records"]
                if record.get("kind") != "revoke"
                and (record["id"] == identifier or (identifier.startswith("b_") and record.get("batch") == identifier))
            ]
            if not found:
                unknown.append(identifier)
            chosen.extend(record for record in found if record not in chosen)
        if unknown:
            raise RecordError(
                f"取り消す確認または許容の識別子を指定する必要がある（見つからない: {', '.join(unknown)}）"
            )
        batch = _new_batch()
        # 元資料が更新された後でも、古い確認を取り消すことができるようにする。
        revokes = [
            _new_record("revoke", reviewer, reason, state.binding, record_id=record["id"], batch=batch)
            for record in chosen
        ]
        cascaded, notices, changed = [], [], {}
        for record in chosen:
            source = record.get("source")
            if (
                not isinstance(source, dict)
                or not isinstance(source.get("run"), str)
                or not isinstance(source.get("record_id"), str)
            ):
                continue
            missing = {
                "message": "元の記録の実行がないため、写しだけを取り消した",
                "record_id": record["id"],
                "source_run": source["run"],
            }
            try:
                origin = _source_run(source, run, cache)
                if origin == run or origin.parent != run.parent:
                    notices.append(missing)
                    continue
                origin_state = _state(cache, origin)
                origin_review = origin_state.review
            except (OSError, RecordError):
                notices.append(missing)
                continue
            original = next((item for item in origin_review["records"] if item["id"] == source["record_id"]), None)
            if original is None:
                notices.append({**missing, "message": "元の記録がないため、写しだけを取り消した"})
                continue
            if any(
                item.get("kind") == "revoke" and item.get("record_id") == original["id"]
                for item in origin_review["records"]
            ):
                continue
            try:
                binding = origin_state.binding
            except (OSError, RecordError):
                binding = None
            cascade = _new_record(
                "revoke",
                reviewer,
                reason,
                binding,
                record_id=original["id"],
                batch=batch,
                cascade={"run": str(run), "record_id": record["id"]},
            )
            origin_review["records"].append(cascade)
            changed[origin] = origin_review
            cascaded.append({"run": str(origin), "record_id": original["id"], "revoke_id": cascade["id"]})
        for origin, value in changed.items():
            write_json(origin / "review.json", value)
        review["records"].extend(revokes)
        write_json(run / "review.json", review)
    return {
        "command": "review revoke",
        "status": "revoked",
        "exit_code": 0,
        "build": str(run),
        "batch": batch,
        "record": revokes[0],
        "records": revokes,
        "revoked": [record["id"] for record in chosen],
        "cascaded": cascaded,
        "notices": notices,
    }


# 引き継ぎ


def carry_reviews(source: Any, target: Any, *, project: Any = None) -> dict:
    """source の実行の確認記録のうち、target でも確認対象と証拠が一致するものを写す（手動の引き継ぎ）。

    source が "auto" なら carry_reviews_auto を呼ぶ。実行はパスか番号（target には latest も指定できる。project は番号を探す資料）。
    写しは、人が記録した元の記録を指す（source の記録が写しなら、その元の記録）。via には記録を見つけた実行を残す。
    元の記録が target にある記録（逆向きの引き継ぎ）と一覧画像の確認は写さない。表示の確認は元の記録の実行の画像と
    target の画像を 1 画素のずれを許して比べ、大きな差があれば写さない。
    両方の実行が同じ資料の同じ実行保存先にある、証拠の形式が同じ、完成物にできる実行である必要がある。資料のディレクトリを
    実行保存先ごと移した後も、同じ実行保存先の実行で、実行保存先から見た資料の場所が同じなら、同じ資料として引き継ぐ。
    """
    if isinstance(source, str) and source == "auto":
        return carry_reviews_auto(target, project=project)
    source = _run_path(resolve_run(source, project=project))
    target = _run_path(resolve_run(target, project=project))
    if source == target:
        raise RecordError("同じ実行へ確認を引き継ぐことはできない")
    # clean が引き継ぎ元を削除する間に、新しい参照を作らない。
    with _output_lock(_recording_output(target)):
        cache: dict[Path, _Run] = {}
        old, new = _state(cache, source), _state(cache, target)
        if source.parent != target.parent or not _same_project(old.manifest, new.manifest):
            raise RecordError("確認を引き継ぐ実行は、同じ資料の同じ実行保存先に置く必要がある")
        if old.manifest["evidence_version"] != new.manifest["evidence_version"]:
            raise RecordError(
                f"証拠の形式が異なる実行の間では引き継げない: {source} は形式 {old.manifest['evidence_version']}、"
                f"{target} は形式 {new.manifest['evidence_version']}"
            )
        _assert_recordable(old)
        _assert_recordable(new)
        notices = _notices(new.manifest)
        current = new.review
        carried = {
            (str(_source_run(item["source"], target, cache)), item["source"].get("record_id"))
            for item in current["records"]
            if isinstance(item.get("source"), dict) and isinstance(item["source"].get("run"), str)
        }
        added, skipped, renumbered = [], [], []

        def skip(record: dict, reason: str) -> None:
            skipped.append({"source_record_id": record["id"], "reason": reason})

        for record in old.review["records"]:
            kind = record.get("kind")
            if kind == "revoke":
                continue
            if kind == "overview":
                skip(record, "一覧画像の確認は実行ごとに必要である")
                continue
            valid, reason = _valid_record(source, record, cache)
            if not valid:
                skip(record, reason)
                continue
            if kind == "accept" and _warning_for(record, new.diagnostics) is None:
                skip(record, "警告の対象または証拠が一致しない")
                continue
            if kind != "accept" and not _record_matches(record, new.manifest):
                skip(record, "ページの確認対象または証拠が一致しない")
                continue
            origin, original = _original(source, record, cache)
            if origin == target:
                # 逆向きの引き継ぎで、引き継ぎ先が元の記録を持っている。写すと、自分の実行を指す無効な写しになる。
                skip(record, "元の記録が引き継ぎ先にある")
                continue
            reference = (str(origin), original["id"])
            if reference in carried:
                skip(record, "すでに引き継いでいる")
                continue
            check = None
            if kind == "visual":
                try:
                    check = _render_check(_state(cache, origin), new, record)
                except (OSError, ValueError) as error:
                    skip(record, f"表示の画像を比べられない: {error}")
                    continue
                strong = check["strong_pixels"]
                if strong is None or strong > NOISE_PIXELS:
                    detail = "画像の寸法が違う" if strong is None else f"1 画素のずれを超える差が {strong} 画素"
                    skip(record, f"{_DIFFERENT_RENDER}（{detail}）")
                    continue
            copy = _copy_of(original, origin, source, new.binding, check)
            # 数式のページは標準表示と互換表示の 2 つの確認を引き継ぐので、ページは 1 回だけ入れる。
            if check is not None and check["masked"] and record["slide_id"] not in renumbered:
                renumbered.append(record["slide_id"])
            current["records"].append(copy)
            carried.add(reference)
            added.append(copy)
        current.setdefault("carry_attempts", []).append(
            {
                "created_at": _now(),
                "mode": "manual",
                "source": str(source),
                "records": [item["id"] for item in added],
                "skipped": skipped,
            }
        )
        write_json(target / "review.json", current)
    return {
        "command": "review carry",
        "status": "carried",
        "exit_code": 0,
        "build": str(target),
        "records": added,
        "skipped": skipped,
        "renumbered": renumbered,
        "notices": notices,
    }


def _newest_first(path: Path) -> tuple:
    """実行を番号の大きい順に並べる鍵（番号でない名前の実行は最後に、名前の順）。"""
    number = runs._is_number(path.name)
    return (not number, -int(path.name) if number else 0, path.name)


def carry_reviews_auto(target: Any, *, wait_seconds: float | None = None, project: Any = None) -> dict:
    """同じ実行保存先の前の実行から、target でも確認対象と証拠が一致する確認と許容を自動で引き継ぐ。

    候補は、target 以外の、形式 2・生成が完了した・完成物にできる・同じ資料・同じ証拠の形式の実行で、記録があるもの（番号の大きい順）。
    形式 1 と証拠の形式が違う実行は飛ばし、件数を skipped_runs に入れる。各候補の記録を元の記録（写しなら source が指す人の記録）に
    まとめ、target のまだ満たされていない確認対象（一覧画像を除く）と未対応の警告（資料コードで宣言した警告を除く）ごとに、
    鍵と証拠が一致する元の記録を新しい実行の順に探す。元の記録がその実行で有効で、表示の確認なら画像の比較の条件を満たす
    最初のものを写す。採用しなかった理由は、引き継ぎに至らなかった確認対象ごとに 1 件として、理由ごとの件数だけを残す。
    資料単位のロックが wait_seconds（None なら LOCK_WAIT_SECONDS）秒を過ぎても使用中なら、例外にせず status が skipped の結果を返す。
    target は形式 2・生成が完了した・完成物にできる実行である必要がある（そうでなければ RecordError）。
    """
    target = _run_path(resolve_run(target, project=project))
    output = _recording_output(target)
    try:
        with _output_lock(output, wait_seconds=wait_seconds):
            return _carry_auto(target)
    except _LockBusy:
        return {
            "command": "review carry",
            "status": "skipped",
            "exit_code": 0,
            "build": str(target),
            "message": f"別の処理が実行中のため引き継がなかった。後で slidekit review carry --from auto --to {_shown(target)} を実行する",
        }


def _carry_auto(target: Path) -> dict:
    cache: dict[Path, _Run] = {}
    new = _state(cache, target)
    _assert_recordable(new)
    notices = _notices(new.manifest)
    version = new.manifest["evidence_version"]
    skipped_runs: Counter[str] = Counter()
    candidates: list[_Run] = []
    for path in sorted(target.parent.iterdir(), key=_newest_first):
        if path == target or path.is_symlink() or not path.is_dir() or not (path / "manifest.json").is_file():
            continue
        state = _state(cache, path)
        try:
            manifest = state.manifest
            if manifest["evidence_version"] != version:
                skipped_runs["証拠の形式が異なる"] += 1
                continue
            if manifest["status"] != "built" or not manifest["exportable"] or not _same_project(manifest, new.manifest):
                continue
            if state.review["records"]:
                candidates.append(state)
        except LegacyRecordError:
            skipped_runs[_LEGACY_SKIPPED] += 1
        except RecordError:
            skipped_runs["記録を読めない"] += 1
    # 候補の記録を元の記録へまとめる（同じ元の記録は 1 つにし、記録を見つけた最も新しい実行を via にする）。
    # 元の記録はその実行の新しい順（同じ実行では後の記録から）に探す。
    originals: list[tuple[_Run, dict, Path]] = []
    seen: set[tuple[Path, str]] = set()
    for state in candidates:
        for record in reversed(state.review["records"]):
            if record.get("kind") not in _RECORD_KINDS or record["kind"] == "overview":
                continue
            found = _origin_record(state.path, record, cache)
            if found is None or (found[0].path, found[1]["id"]) in seen:
                continue
            seen.add((found[0].path, found[1]["id"]))
            originals.append((*found, state.path))
    originals.sort(key=lambda item: _newest_first(item[0].path))
    progress = _Progress(target, cache)
    added: list[dict] = []
    sources: Counter[str] = Counter()
    renumbered: list[str] = []
    renumbered_from: Counter[str] = Counter()
    skipped: Counter[str] = Counter()

    def matches(origin: _Run, original: dict, item: dict, warning: bool) -> str | None:
        """元の記録が確認対象に合うか。証拠が同じなら "exact"、表示の確認でページ番号の図形を除いた証拠だけが
        同じなら "footer"（番号の枠・書式・文字だけが変わったページ）、合わなければ None。"""
        if warning:
            return "exact" if original.get("kind") == "accept" and _warning_for(original, [item]) is not None else None
        if original.get("kind") != item["kind"] or _target_key(original) != _target_key(item):
            return None
        if original.get("evidence_hash") == item["evidence_hash"]:
            return "exact"
        if item["kind"] == "visual" and _same_body(origin.manifest, new.manifest, item):
            return "footer"
        return None

    def attempt(origin: _Run, original: dict, footer: bool) -> tuple[str, dict | None]:
        """元の記録を写せるか確かめ、（写せない理由、表示の照合）を返す。写せるなら理由は空である。"""
        if origin.path == target:
            return "元の記録が引き継ぎ先にある", None
        if origin.manifest["evidence_version"] != version or not _same_project(origin.manifest, new.manifest):
            return "引き継ぎ元が別の資料か、証拠の形式が異なる", None
        valid, reason = _valid_record(origin.path, original, cache)
        if not valid or original["kind"] != "visual":
            return reason, None
        try:
            check = _render_check(origin, new, original)
        except (OSError, ValueError):
            return "表示の画像を比べられない", None
        strong = check["strong_pixels"]
        if footer and not check["masked"]:
            # 証拠はページ番号の図形だけが違うが、番号の領域を除いて比べていない（番号の図形の書式の変化など）。
            return _DIFFERENT_RENDER, check
        return ("" if strong is not None and strong <= NOISE_PIXELS else _DIFFERENT_RENDER), check

    def adopt(item: dict, warning: bool) -> dict | None:
        """確認対象か警告に合う元の記録を新しい順に試し、最初に写せたものの写しを返す。なければ最初の理由の件数を 1 増やす。"""
        first_reason = None
        for origin, original, via in originals:
            mode = matches(origin, original, item, warning)
            if mode is None:
                continue
            reason, check = attempt(origin, original, mode == "footer")
            if not reason:
                sources[str(origin.path)] += 1
                if check is not None and check["masked"] and original["slide_id"] not in renumbered:
                    renumbered.append(original["slide_id"])
                    renumbered_from[origin.path.name] += 1
                if mode == "footer":
                    # 番号の図形を除いた証拠と、番号の領域を除いた画像が一致した写し。確認対象の証拠に結び付け、
                    # 元の記録の証拠を照合の記録に残す（_valid_copy が同じ条件で確かめ直す）。
                    check = {**check, "footer_only": True, "source_evidence_hash": original["evidence_hash"]}
                    copy = _copy_of(original, origin.path, via, new.binding, check)
                    copy["evidence_hash"] = item["evidence_hash"]
                    return copy
                return _copy_of(original, origin.path, via, new.binding, check)
            first_reason = first_reason or reason
        if first_reason:
            skipped[first_reason] += 1
        return None

    remaining: Counter[str] = Counter()
    pending_reasons: list[dict] = []
    for required in progress.missing:
        copy = adopt(required, False) if required["kind"] != "overview" else None
        if copy is None:
            remaining[required["kind"]] += 1
            if required["kind"] != "overview":
                pending_reasons.append(_pending_reason(required, originals, new.manifest))
        else:
            added.append(copy)
    for warning in progress.unaccepted:
        copy = adopt(warning, True)
        if copy is None:
            remaining["warnings"] += 1
        else:
            added.append(copy)
    # 表示環境の違いは、確認の記録がある最も新しい実行と比べる（引き継げなかった理由を示すため）。
    compared = candidates[0] if candidates else None
    changes = _environment_changes(compared.manifest, new.manifest) if compared is not None else []
    current = new.review
    current["records"].extend(added)
    current.setdefault("carry_attempts", []).append(
        {
            "created_at": _now(),
            "mode": "auto",
            "sources": dict(sources),
            "records": [item["id"] for item in added],
            "skipped": dict(skipped),
            "skipped_runs": dict(skipped_runs),
            "pending_reasons": pending_reasons,
        }
    )
    write_json(target / "review.json", current)
    return {
        "command": "review carry",
        "status": "carried",
        "exit_code": 0,
        "build": str(target),
        "mode": "auto",
        "records": added,
        "sources": dict(sources),
        "renumbered": renumbered,
        "renumbered_from": dict(renumbered_from),
        "skipped": dict(skipped),
        "skipped_runs": dict(skipped_runs),
        "remaining": {name: remaining[name] for name in ("visual", "overview", "content", "warnings")},
        "pending_reasons": pending_reasons,
        "compared_run": compared.path.name if compared is not None else None,
        "environment_changes": changes,
        "notices": notices,
    }


def _slide_entry(manifest: dict, slide_id: Any) -> dict:
    return next((slide for slide in manifest.get("slides") or [] if slide.get("id") == slide_id), {})


def _pending_reason(required: dict, originals: list, manifest: dict) -> dict:
    """引き継げなかった確認対象の理由。同じ確認対象の最も新しい元の記録の実行と、生成記録の内訳で比べる。

    表示は環境の証拠（environment_hash）が違えば表示環境の変更、同じならページの表示の変更とする。
    内容は content_parts で本文と参照する元資料・文献項目を比べる（内訳のない古い実行では「内容の変更」）。
    """
    entry = {key: required.get(key) for key in ("kind", "slide_id", "renderer", "variant")}
    key = _target_key(required)
    prior = next(
        (
            origin
            for origin, original, _via in originals
            if original.get("kind") == required["kind"] and _target_key(original) == key
        ),
        None,
    )
    if prior is None:
        return {**entry, "reason": "前の確認の記録がない"}
    entry["compared_run"] = prior.path.name
    old, new = _slide_entry(prior.manifest, required.get("slide_id")), _slide_entry(manifest, required.get("slide_id"))
    if required["kind"] == "visual":
        name = f"{required.get('renderer')}:{required.get('variant')}"
        before = ((old.get("visuals") or {}).get(name) or {}).get("environment_hash")
        after = ((new.get("visuals") or {}).get(name) or {}).get("environment_hash")
        return {**entry, "reason": "表示環境の変更" if before and after and before != after else "ページの表示の変更"}
    before, after = old.get("content_parts"), new.get("content_parts")
    if not isinstance(before, dict) or not isinstance(after, dict):
        return {**entry, "reason": "内容の変更"}
    if before.get("page") != after.get("page"):
        return {**entry, "reason": "本文の変更"}
    documents = sorted(
        str(name)
        for name in set(before.get("documents") or {}) | set(after.get("documents") or {})
        if (before.get("documents") or {}).get(name) != (after.get("documents") or {}).get(name)
    )
    if documents:
        return {**entry, "reason": "元資料の変更", "documents": documents}
    if before.get("bibliography") != after.get("bibliography"):
        return {**entry, "reason": "文献項目の変更"}
    return {**entry, "reason": "内容の変更"}


def _environment_changes(before: dict, after: dict) -> list[dict]:
    """2 つの実行の生成記録の環境（表示ソフト・パッケージ・Python・フォント・計測の版）の違い。"""
    old, new = before.get("environment") or {}, after.get("environment") or {}
    changes = []

    def compare(item: str, a: Any, b: Any) -> None:
        if a != b:
            changes.append({"item": item, "before": a, "after": b})

    renderers = (old.get("renderers") or {}, new.get("renderers") or {})
    for name in sorted(set(renderers[0]) | set(renderers[1])):
        compare(
            f"表示ソフト {name}",
            (renderers[0].get(name) or {}).get("version"),
            (renderers[1].get(name) or {}).get("version"),
        )
    packages = (old.get("packages") or {}, new.get("packages") or {})
    for name in sorted(set(packages[0]) | set(packages[1])):
        compare(name, packages[0].get(name), packages[1].get(name))
    compare("Python", old.get("python"), new.get("python"))
    fonts = (old.get("fonts") or {}, new.get("fonts") or {})
    compare("計測の版", fonts[0].get("measure_version"), fonts[1].get("measure_version"))
    compare("数式の計測の版", fonts[0].get("math_engine"), fonts[1].get("math_engine"))
    faces = (fonts[0].get("fonts") or {}, fonts[1].get("fonts") or {})
    for family in sorted(set(faces[0]) | set(faces[1])):
        a, b = faces[0].get(family) or {}, faces[1].get(family) or {}
        if (a.get("sha256"), (a.get("bold") or {}).get("sha256")) != (
            b.get("sha256"),
            (b.get("bold") or {}).get("sha256"),
        ):
            changes.append({"item": f"フォント {family}", "before": a.get("version"), "after": b.get("version")})
    return changes


# 残りの確認の一覧


def _blocking_groups(issues: list[dict]) -> list[dict]:
    """確認を加えても通らない問題を、識別子ごとにまとめる（パスは最初の 5 件）。"""
    groups: dict[str, dict] = {}
    for issue in issues:
        group = groups.setdefault(
            issue["code"], {"code": issue["code"], "message": issue["message"], "paths": [], "count": 0}
        )
        group["count"] += 1
        if issue.get("path") is not None and len(group["paths"]) < _NOTICE_PATHS:
            group["paths"].append(issue["path"])
        if issue.get("input_kind"):
            kinds = group.setdefault("kinds", {})
            entry = kinds.setdefault(issue["input_kind"], {"count": 0, "first": issue.get("path")})
            entry["count"] += 1
    return list(groups.values())


def _review_groups(state: _Run, progress: _Progress, numbers: Mapping[str, int]) -> list[dict]:
    """確認対象を（種類、表示ソフト、表示方法）の組にまとめ、件数と未確認のページを求める。

    未確認のページには、最後の自動の引き継ぎが記録した引き継げなかった理由（reason・documents）を添える。
    """
    primary = _primary(state.manifest)
    attempts = state.review.get("carry_attempts") or []
    last = attempts[-1] if attempts and isinstance(attempts[-1], dict) else {}
    reasons = {
        (item.get("kind"), item.get("slide_id"), item.get("renderer"), item.get("variant")): item
        for item in last.get("pending_reasons") or []
        if isinstance(item, dict)
    }
    groups: dict[tuple, dict] = {}
    for required, matching in progress.reviews:
        kind, renderer, variant = required["kind"], required.get("renderer"), required.get("variant")
        group = groups.setdefault(
            (kind, renderer, variant),
            {
                "kind": kind,
                "renderer": renderer,
                "variant": variant,
                "required": 0,
                "satisfied": 0,
                "carried": 0,
                "renumbered": 0,
                "pending": [],
            },
        )
        group["required"] += 1
        if not matching:
            slide_id = required.get("slide_id")
            item = {"slide_id": slide_id, "number": numbers.get(slide_id) if slide_id else None}
            known = reasons.get((kind, slide_id, renderer, variant))
            if known is not None:
                item["reason"] = known.get("reason")
                if known.get("documents"):
                    item["documents"] = known["documents"]
            group["pending"].append(item)
            continue
        group["satisfied"] += 1
        if all(record.get("source") is not None for record in matching):
            group["carried"] += 1
            if any((record.get("render_check") or {}).get("masked") for record in matching):
                group["renumbered"] += 1
    result = []
    for (kind, renderer, variant), group in sorted(
        groups.items(),
        key=lambda item: (
            _KIND_ORDER.index(item[0][0]),
            item[0][1] != primary,
            item[0][1] or "",
            item[0][2] != "standard",
        ),
    ):
        pages = kind != "overview" and bool(group["pending"])
        group["slides"] = (
            selection.format_numbers(item["number"] for item in group["pending"] if item["number"] is not None)
            if pages
            else None
        )
        group["status_id"] = (
            _status_id(state, kind, renderer, variant, progress.pending(kind, renderer, variant)) if pages else None
        )
        result.append(group)
    return result


def _record_command(run: Path, group: dict, primary: str) -> str:
    """未確認の組を記録するコマンド。確認者と確認内容は置き場所で示す（画像を見て書き換えてから実行する）。"""
    parts = ["slidekit review record --build", _shown(run), "--kind", group["kind"]]
    if group["kind"] != "content":
        if group["renderer"] != primary:
            parts += ["--renderer", group["renderer"]]
        if group["variant"] != "standard":
            parts += ["--variant", group["variant"]]
    if group["kind"] != "overview":
        parts += ["--slides", group["slides"]]
    parts += ["--reviewer", f'"{REVIEWER_PLACEHOLDER}"', "--note", f'"{NOTE_PLACEHOLDER}"']
    return " ".join(parts)


def _pending_command(run: Path, status_id: str) -> str:
    """未確認の確認をすべてまとめて記録するコマンド（review record --pending）。"""
    return (
        f"slidekit review record --build {_shown(run)} --pending --status-id {status_id} "
        f'--reviewer "{REVIEWER_PLACEHOLDER}" --note "{NOTE_PLACEHOLDER}"'
    )


def _remaining(run: Path, cache: dict) -> dict:
    """記録の後に残っている確認と未対応の警告の件数（種類ごと）。成果物と入力はハッシュし直さず、確認記録だけを読み直す。"""
    progress = _Progress(run, {path: state.with_fresh_reviews() for path, state in cache.items()})
    counts = {"visual": 0, "overview": 0, "content": 0}
    for required in progress.missing:
        if required.get("kind") in counts:
            counts[required["kind"]] += 1
    return {**counts, "warnings": len(progress.unaccepted)}


def review_status(run: Any = "latest", *, project: Any = None) -> dict:
    """残りの確認・未対応の警告と、次の操作を示す。ファイルを書かず、ロックも取らない。

    run は実行のパス、番号、latest（完成物にできる最新の実行）。project は番号と latest を探す資料である。
    確認記録は最終検査と同じ基準で照合する。reviews は（種類、表示ソフト、表示方法）の組ごとの件数・未確認のページ・範囲表記・
    照合用の値（status_id。review record --slides pending に渡す）、warnings は警告の件数と未対応の警告、declared_warnings は
    資料コードで宣言して通した警告（ページ番号・種類・要素・理由・宣言の場所）である。
    status は、成果物・入力の不一致、生成の未完了、完成物にできない実行があれば blocked（終了コード 2）、エラーの診断があれば
    blocked（1）、未確認・未対応の警告があれば incomplete（1）、どれもなければ complete（0）である。
    next の記録と許容のコマンドは、確認者・確認内容・理由を置き場所（<確認者> など）で示す。置き場所のままでは記録しない。
    形式 1 の実行は LegacyRecordError にする。
    """
    run = _run_path(resolve_run(run, project=project))
    cache: dict[Path, _Run] = {}
    state = _state(cache, run)
    manifest = state.manifest
    numbers = _numbers(manifest)
    blocking = _blocking_groups(_blocking(state, final=True))
    diagnostics = state.diagnostics

    def page(issue: dict) -> int | None:
        number = issue.get("slide_number")
        return number if type(number) is int else numbers.get(issue.get("slide_id"))

    errors = [
        {
            "diagnostic_id": issue["diagnostic_id"],
            "code": issue.get("code"),
            "slide_id": issue.get("slide_id"),
            "slide_number": page(issue),
            "element_id": issue.get("element_id"),
            "message": issue.get("message"),
        }
        for issue in diagnostics
        if issue.get("severity") == "error"
    ]
    progress = _Progress(run, cache)
    groups = _review_groups(state, progress, numbers)
    pending_warnings = [
        {
            "diagnostic_id": warning["diagnostic_id"],
            "code": warning.get("code"),
            "slide_id": warning.get("slide_id"),
            "slide_number": page(warning),
            "element_id": warning.get("element_id"),
            "related_element_id": warning.get("related_element_id"),
            "message": warning.get("message"),
        }
        for warning in progress.unaccepted
    ]
    declared = [_declared_entry(warning, numbers) for warning, status, _ in progress.warnings if status == "declared"]
    if blocking:
        status, exit_code = "blocked", 2
    elif errors:
        status, exit_code = "blocked", 1
    elif any(group["pending"] for group in groups) or pending_warnings:
        status, exit_code = "incomplete", 1
    else:
        status, exit_code = "complete", 0
    if status == "blocked":
        operations = [f"slidekit build --project {_shown(_current_project(run, manifest))}"]
    elif status == "complete":
        operations = [f"slidekit export --build {_shown(run)}"]
    else:
        primary = _primary(manifest)
        operations = [_record_command(run, group, primary) for group in groups if group["pending"]]
        if len(operations) > 1:
            # 画像と本文をすべて見た後に、未確認の確認をまとめて 1 回で記録するコマンドを先に示す。
            operations.insert(0, _pending_command(run, _status_id(state, "all", None, None, progress.missing)))
        operations += [
            f"slidekit review accept --build {_shown(run)} --diagnostic {warning['diagnostic_id']} "
            f'--reviewer "{REVIEWER_PLACEHOLDER}" --reason "{REASON_PLACEHOLDER}"'
            for warning in pending_warnings
        ]
        if pending_warnings:
            operations.append(_ALLOW_WARNING_HINT)
    counts = progress.counts["warnings"]
    return {
        "command": "review status",
        "status": status,
        "exit_code": exit_code,
        "build": str(run),
        "run_id": manifest.get("run_id") or run.name,
        "mode": manifest["mode"],
        "exportable": manifest["exportable"],
        "slides": len(manifest["slides"]),
        "blocking": blocking,
        "errors": errors,
        "reviews": groups,
        "warnings": {
            "total": counts["total"],
            "accepted": counts["accepted"],
            "declared": counts["declared"],
            "pending": pending_warnings,
        },
        "declared_warnings": declared,
        "invalid_records": progress.invalid,
        "summary": {**progress.counts, "errors": len(errors), "reviewers": progress.reviewers},
        "status_id": _status_id(state, "all", None, None, progress.missing) if progress.missing else None,
        "next": operations,
    }


# 完成物の反映


def _output_root(run: Path, manifest: dict) -> Path:
    """実行ディレクトリの位置（runs/ の親）から出力先を求める。

    実行ディレクトリが生成記録の出力先の runs/ にない場合（別の場所へ複製した場合など）は RecordError にする。
    複製した実行の export が、元の資料の exports/ に反映の履歴を書かないようにするためである。
    """
    output = run.parent.parent
    recorded = _recorded_output(manifest)
    if recorded is None:
        raise RecordError(f"生成記録に出力先がない: {run}")
    if run.parent.name != "runs" or recorded != output:
        raise RecordError(
            f"実行ディレクトリが生成記録の出力先の runs/ にない（別の場所へ複製した実行は反映できない）: {run}"
            f"（生成記録の出力先: {recorded}）"
        )
    return output


def _same_or_contained(a: Path, b: Path) -> bool:
    return a == b or a.is_relative_to(b) or b.is_relative_to(a)


def _deliverable_paths(run: Path, manifest: dict, output: Path) -> dict[str, Path]:
    names = ("pptx", "pdf", "notes", "record")
    if set(manifest["deliverables"]) != set(names):
        raise RecordError("完成物の保存先には pptx、pdf、notes、record の四つを指定する")
    paths = {}
    root = _project_root(manifest)
    forbidden = [Path(name).resolve() for name in manifest["input_files"]]
    forbidden.extend([run.parent, output / "runs", output / "exports"])
    inputs = manifest["config"].get("inputs")
    for name in inputs.get("asset_dirs", []) if isinstance(inputs, dict) else []:
        forbidden.append((root / name).resolve())
    for name in names:
        raw = Path(manifest["deliverables"][name])
        if not raw.is_absolute():
            raise RecordError(f"完成物の保存先が絶対パスではない: {raw}")
        path = raw.resolve()
        if path.is_dir() or any(_same_or_contained(path, other) for other in forbidden):
            raise RecordError(f"完成物の保存先が入力または実行保存先と重複している: {path}")
        for other in [*forbidden, *paths.values()]:
            if _same_or_contained(path, other) or (path.exists() and other.exists() and path.samefile(other)):
                raise RecordError(f"完成物の保存先がほかのファイルと重複している: {path}")
        paths[name] = path
    return paths


def _restore_transaction(transaction: dict) -> None:
    # 復旧を始める前に旧版をすべて確かめる。
    for item in transaction["files"]:
        destination = Path(item["destination"])
        backup = Path(item["backup"]) if item.get("backup") else None
        if backup and (not backup.is_file() or file_hash(backup) != item["old_hash"]):
            raise RecordError(f"復旧用の旧版が破損している: {backup}")
        if destination.exists() and file_hash(destination) not in {item["new_hash"], item.get("old_hash")}:
            raise RecordError(f"中断後に完成物が変更されているため、自動復旧できない: {destination}")
    for item in transaction["files"]:
        destination = Path(item["destination"])
        if item.get("backup"):
            if not destination.exists() or file_hash(destination) != item["old_hash"]:
                temporary = stage_copy(Path(item["backup"]), destination)
                try:
                    os.replace(temporary, destination)
                finally:
                    temporary.unlink(missing_ok=True)
        else:
            destination.unlink(missing_ok=True)
        Path(item["temporary"]).unlink(missing_ok=True)


def _recover_pending(output: Path, project: str) -> None:
    """前回の反映が中断して残した pending.json を復旧する（形式 1 の記録も読む）。"""
    pending = output / "exports" / "pending.json"
    if not pending.exists():
        return
    transaction = read_record(pending, allow_legacy=True)
    if transaction.get("project") != project:
        raise RecordError("未完了の反映記録が別の資料を指している")
    if transaction.get("status") != "committed":
        _restore_transaction(transaction)
        transaction["status"] = "recovered"
        transaction["recovered_at"] = _now()
    write_json(Path(transaction["history"]) / "transaction.json", transaction)
    pending.unlink()


def _previous_run(path: Path, root: Path) -> str | None:
    """採用中の実行（採用記録の run）を、資料のディレクトリからの相対パスで返す。採用記録の相対パスは資料のディレクトリから読む。

    採用記録がないか、その読み込みに失敗した（slidekit を使わずに置いた完成物など）ときは None にする。旧版は内容を変えずに退避する。
    """
    if not path.exists():
        return None
    try:
        value = read_record(path, allow_legacy=True).get("run")
    except RecordError:
        return None
    return _relative(_record_path(value, root), root) if isinstance(value, str) and value else None


def _stage(transaction: dict, history: Path, key: str, source: Path, destination: Path, expected: str) -> None:
    """反映する 1 ファイルを置き換え先の隣へ複製し、置き換え先の旧版を退避して、反映の記録に加える。"""
    temporary = stage_copy(source, destination)
    if file_hash(temporary) != expected:
        temporary.unlink(missing_ok=True)
        raise RecordError(f"反映用の複製が元と一致しない: {destination}")
    backup = history / f"previous-{key}" if destination.exists() else None
    old_hash = file_hash(destination) if backup else None
    transaction["files"].append(
        {
            "destination": str(destination),
            "temporary": str(temporary),
            "backup": str(backup) if backup else None,
            "old_hash": old_hash,
            "new_hash": expected,
        }
    )
    if backup:
        shutil.copyfile(destination, backup)
        if file_hash(backup) != old_hash:
            raise RecordError(f"旧版の退避内容が一致しない: {destination}")


def _assert_unchanged(run: Path, manifest: dict) -> None:
    """成果物と入力を読み直し、生成記録と一致することを確かめる。"""
    issues = [*_artifact_issues(run, manifest), *_input_issues(manifest["input_files"], manifest)]
    if issues:
        first = issues[0]
        raise RecordError(f"{first['message']}: {first.get('path', run)}")


def export_build(run: Path) -> dict:
    """最終検査に通った実行を固定パスへ反映し、旧版を退避する。

    反映の履歴（transaction.json）・最終検査の結果の複製・採用記録の原本は、実行ディレクトリの位置から求めた出力先の
    exports/{識別子}/ に置く。採用記録（deliverables.json）の run・project・artifacts と、反映の履歴の run・previous_run は、
    資料のディレクトリからの相対パスで書く。最終検査は最初に全体を行い、反映の直前に確認記録だけを確かめ直す（成果物と入力は、その間に読み直して
    照合する）。どちらの結果も final_check.json に保存し、反映の直前の結果を退避先に複製して採用記録から参照する。
    """
    run = _run_path(run)
    manifest = _manifest(run)
    root = _project_root(manifest)
    output = _output_root(run, manifest)
    with _output_lock(output):
        _recover_pending(output, manifest["project"])
        cache: dict[Path, _Run] = {}
        checked = _evaluate(run, cache, final=True)
        write_json(run / FINAL_CHECK_FILE, checked)
        if checked["status"] != "passed":
            # 足りない確認と未対応の警告を、記録と許容のコマンドとともに示す（review status と同じ次の操作）。
            try:
                operations = review_status(run)["next"]
            except RecordError:
                operations = []
            return {
                "command": "export",
                "status": "failed",
                "exit_code": checked["exit_code"],
                "check": checked,
                "final_check": str(run / FINAL_CHECK_FILE),
                "next": operations,
            }
        manifest = cache[run].manifest
        paths = _deliverable_paths(run, manifest, output)
        identifier = f"{datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S')}-{uuid.uuid4().hex[:8]}"
        history = output / "exports" / identifier
        history.mkdir(parents=True)
        mapping = {"pptx": "deck.pptx", "pdf": "deck.pdf", "notes": "notes.md"}
        # 反映した実行と退避した旧版の実行は、採用記録と同じく資料のディレクトリからの相対パスで書く（資料を out/ ごと移しても、
        # 整理はこの記録でこれらの実行を保護できる）。
        transaction = {
            "schema_version": SCHEMA_VERSION,
            "id": identifier,
            "project": manifest["project"],
            "run": _relative(run, root),
            "previous_run": _previous_run(paths["record"], root),
            "history": str(history),
            "status": "prepared",
            "created_at": _now(),
            "files": [],
        }
        pending = output / "exports" / "pending.json"
        try:
            for key, name in mapping.items():
                _stage(transaction, history, key, run / name, paths[key], manifest["artifacts"][name])
            # 複製の間に、成果物・入力・生成記録・確認記録が変わっていないことを確かめる。
            _assert_unchanged(run, manifest)
            if file_hash(run / "manifest.json") != checked["manifest_hash"]:
                raise RecordError("反映の準備中に生成記録が変更された")
            if file_hash(run / "review.json") != checked["review_hash"]:
                raise RecordError("反映の準備中に確認記録が変更された")
            # 別の実行で取り消された確認も、採用記録を書く直前に確かめる。成果物と入力はハッシュし直さない。
            final = _evaluate(run, {path: state.with_fresh_reviews() for path, state in cache.items()}, final=True)
            write_json(run / FINAL_CHECK_FILE, final)
            if final["status"] != "passed" or final["review_hash"] != checked["review_hash"]:
                raise RecordError("反映中に入力または確認の状態が変更された")
            shutil.copyfile(run / FINAL_CHECK_FILE, history / FINAL_CHECK_FILE)
            adoption = {
                "schema_version": SCHEMA_VERSION,
                "run": _relative(run, root),
                "run_id": manifest.get("run_id", run.name),
                "project": _relative(Path(manifest["project"]), root),
                "exported_at": _now(),
                "manifest_hash": checked["manifest_hash"],
                "review_hash": checked["review_hash"],
                "artifacts": {
                    _relative(paths[key], root): manifest["artifacts"][name] for key, name in mapping.items()
                },
                "final_check": {
                    "checked_at": final["checked_at"],
                    "sha256": file_hash(history / FINAL_CHECK_FILE),
                    "summary": final["summary"],
                },
            }
            adoption_source = history / "adoption.json"
            write_json(adoption_source, adoption)
            _stage(transaction, history, "record", adoption_source, paths["record"], file_hash(adoption_source))
            write_json(history / "transaction.json", transaction)
            write_json(pending, transaction)
            for item in transaction["files"]:
                os.replace(item["temporary"], item["destination"])
            transaction["status"] = "committed"
            transaction["completed_at"] = _now()
            write_json(pending, transaction)
            write_json(history / "transaction.json", transaction)
            pending.unlink()
        except Exception:
            if pending.exists():
                _restore_transaction(transaction)
                transaction["status"] = "rolled_back"
                write_json(history / "transaction.json", transaction)
                pending.unlink()
            else:
                for item in transaction["files"]:
                    Path(item["temporary"]).unlink(missing_ok=True)
            raise
    return {
        "command": "export",
        "status": "exported",
        "exit_code": 0,
        "run": str(run),
        "record": str(paths["record"]),
        "deliverables": {key: str(path) for key, path in paths.items()},
        "history": str(history),
        "final_check": str(run / FINAL_CHECK_FILE),
    }


# 整理


def clean_runs(project: Any, before: str, apply: bool = False) -> dict:
    """0.1 の呼び方の互換：before より前の実行を整理する（runs.clean_runs）。

    戻り値の candidates は 0.1 の形（候補のパスの文字列の列）にする。候補ごとの理由と大きさは runs.clean_runs を直接呼ぶと得られる。
    """
    try:
        result = runs.clean_runs(project, before=before, apply=apply)
    except RecordError:
        raise
    except ValueError as error:
        raise RecordError(str(error)) from error
    return {**result, "candidates": result["candidate_paths"]}


# 前の実行との比較と比較画像


def _previous_build(run: Path, cache: dict) -> _Run | None:
    """run より前の、同じ資料の完了した実行のうち最も新しいもの（下書き・部分生成も含む）。"""
    current = _state(cache, run)
    for path in sorted(run.parent.iterdir(), key=_newest_first):
        if path == run or path.is_symlink() or not path.is_dir() or not (path / "manifest.json").is_file():
            continue
        if _newest_first(path) <= _newest_first(run):
            continue
        try:
            state = _state(cache, path)
            manifest = state.manifest
            if manifest["status"] == "built" and _same_project(manifest, current.manifest):
                return state
        except RecordError:
            continue
    return None


def _page_change(origin: _Run, target: _Run, slide_id: str, key: str, *, boxes: bool = True) -> tuple[str, list]:
    """1 ページの標準表示の画像の比較：（"unchanged"・"renumbered"・"changed"・"missing"、差のある領域）。

    boxes が偽なら、差のある領域を求めない（比べた結果だけが要るとき）。
    """
    before, after = _visual(origin.manifest, slide_id, key), _visual(target.manifest, slide_id, key)
    if before is None or after is None:
        return "missing", []
    if before["pixel_hash"] == after["pixel_hash"]:
        return "unchanged", []
    source, image = _artifact_path(origin.path, before["path"]), _artifact_path(target.path, after["path"])
    masks = _footer_masks(origin.manifest, target.manifest, slide_id)
    strong = render_difference(source, image, masks=masks)["strong_pixels"]
    if strong is not None and strong <= NOISE_PIXELS:
        return ("renumbered" if masks else "unchanged"), []
    return "changed", difference_boxes(source, image, masks=masks) if boxes else []


def compare_previous(run: Any, *, project: Any = None) -> dict | None:
    """前の実行（_previous_build）と、主表示ソフトの標準表示のページ画像を比べる。前の実行がなければ None。

    戻り値の unchanged・renumbered（ページ番号の領域を除くと同じ）・changed・added（前の実行にないページ）は
    ページの識別子の列、removed は前の実行だけにあったページの識別子、numbers は今の実行のページ番号である。
    """
    run = _run_path(resolve_run(run, project=project))
    cache: dict[Path, _Run] = {}
    target = _state(cache, run)
    origin = _previous_build(run, cache)
    if origin is None:
        return None
    key = f"{_primary(target.manifest)}:standard"
    result: dict[str, Any] = {
        "run": origin.path.name,
        "unchanged": [],
        "renumbered": [],
        "changed": [],
        "added": [],
        "removed": [],
    }
    previous = {slide["id"] for slide in origin.manifest["slides"]}
    for slide in target.manifest["slides"]:
        if _visual(target.manifest, slide["id"], key) is None:
            continue
        if slide["id"] not in previous:
            result["added"].append(slide["id"])
            continue
        try:
            change, _boxes = _page_change(origin, target, slide["id"], key, boxes=False)
        except (OSError, ValueError):
            change = "changed"
        result["changed" if change == "missing" else change].append(slide["id"])
    current = {slide["id"] for slide in target.manifest["slides"]}
    result["removed"] = [slide["id"] for slide in origin.manifest["slides"] if slide["id"] not in current]
    result["numbers"] = {slide_id: number for slide_id, number in _numbers(target.manifest).items()}
    return result


def review_sheet(
    run: Any = "latest",
    slides: str = "pending",
    *,
    variant: str = "all",
    compare: bool = True,
    project: Any = None,
) -> dict:
    """ページの画像を横に並べた比較画像を、実行ディレクトリの sheets/ に作る（確認の記録は書かない）。

    slides は pending（表示の確認が未完了のページ）、all、番号・範囲・識別子の指定。variant は standard・fallback・all。
    compare が真なら、前の実行（_previous_build）の標準表示の画像を左に並べ、差のある領域を今の画像に枠で示す。
    """
    if variant not in (*_VARIANTS, "all"):
        raise RecordError(f"表示方法が不正である（standard、fallback、all）: {variant}")
    run = _run_path(resolve_run(run, project=project))
    cache: dict[Path, _Run] = {}
    state = _state(cache, run)
    manifest = state.manifest
    primary = _primary(manifest)
    text = (slides or "pending").strip()
    if text == "pending":
        wanted = {
            required.get("slide_id") for required in _Progress(run, cache).missing if required.get("kind") == "visual"
        }
        selected = [slide["id"] for slide in manifest["slides"] if slide["id"] in wanted]
    elif text == "all":
        selected = [slide["id"] for slide in manifest["slides"]]
    else:
        selected = _select_slides(text, manifest)
    origin = _previous_build(run, cache) if compare else None
    numbers = _numbers(manifest)
    variants = list(_VARIANTS) if variant == "all" else [variant]
    columns = ([f"previous {origin.path.name}"] if origin is not None else []) + [
        "standard" if name == "standard" else "fallback" for name in variants
    ]
    rows = []
    changes: dict[str, str] = {}
    for slide_id in selected:
        cells: dict[str, dict] = {}
        boxes: list = []
        if origin is not None:
            before = _visual(origin.manifest, slide_id, f"{primary}:standard")
            if before is not None:
                cells[columns[0]] = {"path": str(_artifact_path(origin.path, before["path"]))}
                try:
                    changes[slide_id], boxes = _page_change(origin, state, slide_id, f"{primary}:standard")
                except (OSError, ValueError):
                    changes[slide_id] = "changed"
            else:
                changes[slide_id] = "added"
        for name in variants:
            visual = _visual(manifest, slide_id, f"{primary}:{name}")
            if visual is not None:
                cells[name] = {
                    "path": str(_artifact_path(run, visual["path"])),
                    "boxes": boxes if name == "standard" else [],
                }
        if not any(name in cells for name in variants):
            continue
        number = numbers.get(slide_id)
        label = (f"p.{number} " if number is not None else "") + slide_id
        if slide_id in changes:
            label += f" ({changes[slide_id]})"
        rows.append({"slide_id": slide_id, "label": label, "cells": cells})
    if not rows:
        raise RecordError("比較画像に並べるページがない（--slides を確かめる）")
    # どのページにも画像がない列（数式のない資料の互換表示など）は並べない。
    columns = [column for column in columns if any(column in row["cells"] for row in rows)]
    from .sheets import write_sheets

    prefix = "pending" if text == "pending" else "sheet"
    paths = write_sheets(rows, columns, run / "sheets", prefix)
    return {
        "command": "review sheet",
        "status": "written",
        "exit_code": 0,
        "build": str(run),
        "compared_run": origin.path.name if origin is not None else None,
        "slides": [row["slide_id"] for row in rows],
        "changes": changes,
        "sheets": [str(path) for path in paths],
    }
