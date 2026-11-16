"""構築、検査、ローカル変換を一つの実行記録へまとめる。

- `layout()`：構築・素材の準備・配置と記述規則の検査だけを行う。表示ソフトを使わず、ファイルを書かない。
- `doctor()`：`layout()` の検査に加え、出力先・実行環境・表示ソフトの事前確認を行う。
- `build()`：新しい実行ディレクトリに、PPTX・PDF・ページ画像・生成記録を作り、前の実行から確認を自動で引き継ぐ。

互換画像を付ける資料の生成は、中間 PPTX → 最終形の PPTX（互換画像だけ仮の画像）→ PowerPoint で 1 回変換（標準表示）
→ その PDF から互換画像を切り出して差し替え → 最終 PPTX → 構造の一致の確認 → 互換表示の変換（数式のページだけ）の順に行う。
標準表示の確認・PDF の検査・画像化は、最終 PPTX と同じスライドの XML を持つ最終形の PPTX を変換した PDF で行う。
PPTX の検査と、PDF の検査・画像化は、すべての変換の後に順に行う（PowerPoint の排他は最後の変換の後に解放する）。
"""

from __future__ import annotations

import contextlib
from datetime import datetime, timezone
import copy
from functools import partial
import importlib.metadata
import importlib.machinery
import importlib.util
import os
from pathlib import Path
import re
import shutil
import signal
import socket
import sys
import tempfile
import threading
import traceback
import unicodedata
import uuid

from . import fonts as font_files
from . import records, runs, selection
from .assets import prepare_assets
from .config import load_project
from .errors import failure_diagnostic
from .evidence import (
    EVIDENCE_VERSION,
    RENDER_SCALE,
    content_digest,
    content_parts,
    deck_digest,
    display_environment,
    overview_digest,
    pixel_digest,
    pptx_pages,
    render_page_image,
    visual_body_digest,
    visual_digest,
)
from .fallback import compare_package_structure, fallback_required, finalize_math, prepare_fallback, prepare_final_form
from .files import file_hash, write_json
from .measure import MEASURE_VERSION, ElementLayout, layout_record, layout_scene
from .model import Deck
from .parallel import Workers, resolve_jobs
from .pdf_validation import validate_pdf
from .pptx_validation import validate_pptx
from .progress import Progress
from .renderers import RendererError, open_session
from .report import count_issues, sort_issues
from .validators import apply_declarations, make_diagnostic, merge_diagnostics, skip_role_checks, validate_scene
from .writer import write_pptx

# 成果物の一覧に入れないファイル：生成中の印と、生成の後に書き換わる記録（確認記録と検査の結果）。
_RECORD_FILES = frozenset({".building", "manifest.json", "review.json", "check.json", "final_check.json"})
# layout.json（配置の検査の段階で書く予測）の形式。形を変えたら上げる。
LAYOUT_RECORD_VERSION = 1
# 確認用の複製・最終形の PPTX（--keep-intermediate でなければ消す）。
_INTERMEDIATE_FILES = ("intermediate/fallback.pptx", "intermediate/final_form.pptx")
# 表示環境の証拠に入れる変換の結果の項目。
_ACTUAL_KEYS = frozenset({"renderer", "version", "build", "executable", "windows_user_sid"})
# 生成記録に版を記録するパッケージ。
_PACKAGES = ("python-pptx", "Pillow", "fonttools", "pypdfium2", "pypdf", "lxml", "CairoSVG", "bibtexparser")
# 実行の番号の指定（records.resolve_run と同じ書式）。
_RUN_NUMBER = re.compile(r"[0-9]+")
# ページの識別子を画像の名前にするときに %XX にする文字（Windows のファイル名で禁止されている文字と %）。
_UNSAFE_NAME = re.compile(r'[<>:"\\|?*%]')


def _library_inputs(project):
    for source in sorted(Path(__file__).parent.rglob("*")):
        if source.is_file() and source.suffix in {".py", ".ps1", ".toml"}:
            project.input_files[str(source.resolve())] = file_hash(source)


class _FreshSourceLoader(importlib.machinery.SourceFileLoader):
    def get_code(self, fullname):
        # 同じ秒・同じ長さの編集でも古い .pyc を使用しない。
        return self.source_to_code(self.get_data(self.path), self.path)


class _ProjectFinder:
    def __init__(self, roots):
        self.roots = roots

    def find_spec(self, fullname, path=None, target=None):
        spec = importlib.machinery.PathFinder.find_spec(fullname, path, target)
        if spec and spec.origin and spec.origin.endswith(".py"):
            source = Path(spec.origin).resolve()
            if any(source.is_relative_to(root) for root in self.roots):
                spec.loader = _FreshSourceLoader(fullname, str(source))
                return spec
        return None


def construct(project):
    """資料のコードを呼ぶ。任意の Python のコードを隔離する仕組みではない。

    文献ファイルは `project.bibliography` で 1 回だけ読み、キーの一覧（`_bibliography_keys`）と文献項目の証拠
    （`_bibliography_entries`）を設定に入れる。生成関数を呼ぶ前に `project.collect_errors` を真にし、作成 API が
    続行できる入力の誤りを `scene["issues"]` に集めるようにする。`sys.path` の先頭には資料のディレクトリと
    入口のファイルのディレクトリ（異なる場合）をこの順に入れる（継承した設定の入口が隣の補助を読み込むことができるように）。
    部分生成・部分の検査では、生成の流れが `project.selection` に --slides の指定を入れてから呼ぶ。生成関数は
    `context.selection` と `context.selected()` で選んだページを調べ、選んだページを含む章だけを作ってよい。
    """
    _library_inputs(project)
    # context.selected() が偽を返した識別子（選んだ章だけを構築したかの判定に使う）。構築ごとに記録し直す。
    project.__dict__.pop("_unselected", None)
    bibliography = project.bibliography
    project.config["_bibliography_keys"] = sorted(bibliography)
    project.config["_bibliography_entries"] = {key: bibliography.digest(key) for key in bibliography}
    filename, function = project.config["entrypoint"].rsplit(":", 1)
    path = project.resolve_input(filename)
    name = "_slidekit_deck_" + uuid.uuid4().hex
    spec = importlib.util.spec_from_file_location(name, path, loader=_FreshSourceLoader(name, str(path)))
    if spec is None or spec.loader is None:
        raise ValueError("生成コードを読み込めません")
    module = importlib.util.module_from_spec(spec)
    previous_path = list(sys.path)
    previous_modules = dict(sys.modules)
    library_root = Path(__file__).parent.resolve()
    roots = project.read_roots
    sources = {}

    def source_of(imported):
        """モジュールのファイルの（解決したパス、参照範囲の中か、資料のモジュールか）。1 回の構築の中で覚える。"""
        source = getattr(imported, "__file__", None)
        if not source or not isinstance(source, (str, os.PathLike)):
            return None
        if source not in sources:
            resolved = Path(source).resolve()
            inside = any(resolved.is_relative_to(root) for root in roots)
            sources[source] = (resolved, inside, inside and not resolved.is_relative_to(library_root))
        return sources[source]

    def local_module(imported):
        found = source_of(imported)
        return found is not None and found[2]

    displaced = {key: imported for key, imported in previous_modules.items() if local_module(imported)}

    def loaded_here(imported_name):
        # 構築の間に読み込まれたか、退避した資料のモジュールの名前。ほかの名前は構築の前と同じなので調べない。
        return imported_name not in previous_modules or imported_name in displaced

    for key in displaced:
        sys.modules.pop(key, None)
    finder = _ProjectFinder(roots)
    sys.meta_path.insert(0, finder)
    directories = [str(project.root)]
    if str(path.parent) not in directories:
        directories.append(str(path.parent))
    sys.path[:0] = directories
    sys.modules[name] = module
    project.collect_errors = True
    try:
        spec.loader.exec_module(module)
        deck = getattr(module, function)(project)
        if not isinstance(deck, Deck):
            raise TypeError("生成関数は Deck を返してください")
        scene = deck.to_dict()
        # 選択（--slides）で資料のコードが章を 1 つも作らなかったときは、ページ 0 件の誤りにしない。
        # 選んだページの絞り込み（_select）が、指定の識別子の誤り（候補は作らなかったページの識別子）として示す。
        if not scene["slides"] and not (project.selection is not None and project.__dict__.get("_unselected")):
            raise ValueError("少なくとも 1 ページを作成してください")
        # 資料から読み込まれた補助 Python も再確認の対象にする。
        for imported_name, imported in tuple(sys.modules.items()):
            if not loaded_here(imported_name):
                continue
            found = source_of(imported)
            if found is not None and found[1] and found[0].suffix == ".py":
                project.resolve_input(found[0])
        return scene
    finally:
        sys.path[:] = previous_path
        sys.meta_path.remove(finder)
        for imported_name, imported in tuple(sys.modules.items()):
            if loaded_here(imported_name) and local_module(imported):
                sys.modules.pop(imported_name, None)
        sys.modules.update(displaced)
        sys.modules.pop(name, None)


def _environment(project, registry):
    """変換に渡す実行環境（計測で使った書体だけを持つ）と、生成記録の環境の記録を返す。

    計測で使った書体と、版を記録するパッケージの記録のファイルを、入力の一覧に加える。
    """
    inventory = registry.inventory()
    runtime = copy.deepcopy(project.environment)
    runtime["fonts"] = {}
    for family, face in inventory["fonts"].items():
        spec = {"path": face["path"], "index": face["index"]}
        project.input_files[face["path"]] = face["sha256"]
        if face.get("bold"):
            spec.update(bold_path=face["bold"]["path"], bold_index=face["bold"]["index"])
            project.input_files[face["bold"]["path"]] = face["bold"]["sha256"]
        runtime["fonts"][family] = spec
    versions = {}
    for package in _PACKAGES:
        try:
            distribution = importlib.metadata.distribution(package)
            versions[package] = distribution.version
            for item in distribution.files or []:
                if item.name in {"METADATA", "RECORD"} and ".dist-info" in str(item):
                    source = Path(distribution.locate_file(item)).resolve()
                    project.input_files[str(source)] = file_hash(source)
        except importlib.metadata.PackageNotFoundError:
            pass
    fingerprint = {
        "python": sys.version.split()[0],
        "packages": versions,
        "fonts": inventory,
        "config_files": runtime.get("config_files", []),
        "renderers": {},
    }
    return runtime, fingerprint


def _font_directories(project, *, detect: bool) -> list[str] | None:
    """記録する既定フォントのフォルダー。確認の証拠には入れない。

    設定がなければ自動探索の結果を使うが、detect が偽なら探さない（WSL では PowerShell を新たに起動するため）。
    PowerPoint の変換セッションを開いた後は Windows への問い合わせを記憶しているので、探しても起動は増えない。
    """
    if "font_directories" not in project.environment and not detect:
        return None
    return [str(folder) for folder in font_files.font_directories(project.environment)]


def _holder(project, command: str, run: Path | None) -> dict:
    """変換を待つ側に見せる、この処理の情報。"""
    return {
        "project": str(project.path),
        "run": str(run) if run is not None else None,
        "command": command,
        "host": socket.gethostname(),
        "python_pid": os.getpid(),
    }


def _event_text(event: dict) -> str | None:
    """変換セッションの通知を進捗の 1 行にする。

    待っている相手（使用中の記録）は、Windows の作業プロセスの PID（pid）と、その処理を始めた側の情報（holder：
    資料・実行・Python の PID）を持つ。WSL の利用者が ps で探すことができるように、Python の PID を優先して示し、
    なければ Windows の作業プロセスの PID と判断できる形で示す。
    """
    if event.get("event") != "waiting":
        return None
    record = event.get("holder") if isinstance(event.get("holder"), dict) else {}
    inner = record.get("holder") if isinstance(record.get("holder"), dict) else {}
    holder = {**record, **inner}
    details = []
    if holder.get("run"):
        details.append(f"実行 {holder['run']}")
    elif holder.get("project"):
        details.append(f"資料 {holder['project']}")
    if holder.get("python_pid"):
        details.append(f"PID {holder['python_pid']}")
    elif record.get("pid"):
        details.append(f"Windows の作業プロセスの PID {record['pid']}")
    waited = event.get("waited_seconds")
    if isinstance(waited, (int, float)) and not isinstance(waited, bool):
        details.append(f"経過 {waited:.0f} 秒")
    return "変換待ち: 別の処理が PowerPoint を使用中" + (f"（{'、'.join(details)}）" if details else "")


def _on_event(progress):
    def notify(event):
        text = _event_text(event)
        if text:
            progress.event(text)

    return notify


class _SlidesError(selection.SelectionError):
    """選んだページの絞り込み（`_select`）での --slides の指定の誤り。資料のコードが出した `SelectionError` と区別する。"""


def _selection_diagnostic(error: BaseException, stage: str) -> dict:
    """--slides の指定の誤りの診断。利用者の入力の誤りなので、場所とトレースバックを持たない。"""
    message = str(error)
    return make_diagnostic(
        "invalid_selection", message, scope=stage, observations={"exception": "SelectionError", "message": message}
    )


def _select(scene: dict, spec: str | None, project=None) -> dict:
    """部分生成・部分の検査のページに絞った場面。作成 API の誤りも選んだページのものだけを残す。

    場面の selection の construct は、資料のコードが全ページを作ったなら "all"、選んだ章だけを作った
    （`context.selected()` が偽を返した識別子のページを作らなかった）なら "selected" である。全ページを作ったときの
    ページ番号は元のままである。選んだ章だけを作ったときは、作ったページの中の番号が元のページ番号と違うので、
    selection の numbers と total_slides を None にし、診断・予測・生成記録・deck.json・ノートの見出し・ページ画像の名前では
    ページを識別子だけで示す（`_unnumbered`）。ページ番号の飾りは作ったページの中の番号のままである。
    範囲外の番号とない識別子は `_SlidesError` にする（呼び出し側が、トレースバックのない入力の誤りとして示す）。
    """
    if spec is None:
        return scene
    unselected = set(getattr(project, "_unselected", None) or ())
    try:
        # 作らなかったページの識別子も、ない識別子の候補にする（章を 1 つも作らなかった打ち間違いでも候補を示す）。
        chosen = selection.select(scene["slides"], spec, known=unselected)
    except selection.SelectionError as error:
        raise _SlidesError(str(error)) from None
    kept = {page["id"] for page in chosen}
    described = selection.describe(chosen, spec, len(scene["slides"]))
    skipped = unselected - {page.get("id") for page in scene["slides"]}
    if skipped:
        described.update(numbers=None, total_slides=None, construct="selected")
    else:
        described["construct"] = "all"
    result = {**scene, "slides": chosen, "selection": described}
    if "issues" in scene:
        result["issues"] = [
            issue for issue in scene["issues"] if issue.get("slide_id") is None or issue.get("slide_id") in kept
        ]
    return result


def _unnumbered(scene: dict | None) -> bool:
    """資料のコードが選んだ章だけを作り、ページ番号が元のページ番号でない場面か。"""
    chosen = (scene or {}).get("selection")
    return isinstance(chosen, dict) and chosen.get("construct") == "selected"


def _recorded_scene(scene: dict) -> dict:
    """deck.json に書く場面。選んだ章だけを作った部分生成では、ページの番号（作ったページの中の番号）を None にする。"""
    if not _unnumbered(scene):
        return scene
    return {**scene, "slides": [{**page, "number": None} for page in scene["slides"]]}


def _check_layout(project, scene):
    """配置と記述規則の検査。layout・doctor・build が共有する。

    `project.fonts` で全要素を 1 回だけ計測し、その結果（`measure.layout_scene`）を検査・`layout.json`・PDF の照合で使い回す。
    作成 API が集めた続行できる入力の誤り（`scene["issues"]`）も合わせる。戻り値は（診断、計測の結果）。
    """
    fonts = project.fonts
    layouts = layout_scene(scene, fonts)
    issues = [*scene.get("issues", []), *validate_scene(scene, project, fonts, layouts=layouts)]
    return issues, layouts


def _layout_document(scene: dict, layouts) -> dict:
    """layout.json（予測）の内容。計測できなかった要素は、診断が示すので載せない。"""
    slides = []
    for slide in scene.get("slides", []):
        elements = []
        for element in slide.get("elements", []):
            value = (layouts or {}).get((slide.get("id"), element.get("id")))
            if isinstance(value, ElementLayout):
                elements.append(layout_record(value))
        number = None if _unnumbered(scene) else slide.get("number")
        slides.append({"id": slide.get("id"), "number": number, "elements": elements})
    return {"schema_version": LAYOUT_RECORD_VERSION, "measure_version": MEASURE_VERSION, "slides": slides}


def _finish_issues(issues: list[dict], scene: dict | None) -> list[dict]:
    """同じ診断をまとめ、資料のコードの許容の宣言を当て、ページ順に並べる。

    資料のコードが選んだ章だけを作った場面では、診断のページ番号（作ったページの中の番号）を除き、識別子だけで示す。
    """
    issues = merge_diagnostics(issues)
    if scene is not None:
        issues = apply_declarations(skip_role_checks(issues, scene), scene)
    issues = sort_issues(issues, scene["slides"] if scene is not None else None)
    if _unnumbered(scene):
        issues = [{key: value for key, value in issue.items() if key != "slide_number"} for issue in issues]
    return issues


def _close_fonts(project) -> None:
    """構築・検査・書き出しで共有した書体を、最後に 1 回だけ閉じる（作っていなければ何もしない）。"""
    registry = project.__dict__.get("_font_registry")
    close = getattr(registry, "close", None)
    if close is not None:
        close()


def _traceback_text(error: BaseException) -> str:
    return "".join(traceback.format_exception(error))


def _inspect(project, spec, progress, where: dict):
    """構築し、選んだページに絞り、一時フォルダーで素材を準備して配置を検査する（layout・doctor が共有する）。

    資料のディレクトリと出力先へは書かない。where["stage"] に進んだ段階を書く（失敗の診断に使う）。
    戻り値は（場面、診断、計測の結果）。
    """
    where["stage"] = "construct"
    with progress.stage("construct"):
        scene = _select(construct(project), spec, project)
    with tempfile.TemporaryDirectory(prefix="slidekit-layout-") as temporary:
        where["stage"] = "assets"
        with progress.stage("assets"):
            prepare_assets(scene, project, Path(temporary))
        where["stage"] = "layout"
        with progress.stage("layout"):
            issues, layouts = _check_layout(project, scene)
    return scene, issues, layouts


def layout(path, *, slides=None, progress=None) -> dict:
    """構築・素材の準備・配置と記述規則の検査だけを行う。表示ソフトを使わず、ファイルを書かない。

    slides（`selection` の書式）を指定すると、検査と報告を選んだページに限る。指定は生成関数に `context.selection` として
    渡し、資料のコードは選んだページを含む章だけを作ってよい（`context.selected()`。ほかの章の書きかけの誤りで止まらない）。
    戻り値の `layout` は layout.json と同じ形の予測である。構築の例外は失敗の診断にして終了コード 2 で返す。
    """
    progress = Progress(None) if progress is None else progress
    if slides is not None:
        selection.parse(slides)
    project = load_project(path)
    project.selection = slides
    scene, issues, layouts, exit_code = None, [], None, 0
    where = {"stage": "construct"}
    try:
        scene, issues, layouts = _inspect(project, slides, progress, where)
    except _SlidesError as error:
        issues.append(_selection_diagnostic(error, where["stage"]))
        exit_code = 2
    except Exception as error:
        progress.error(_traceback_text(error))
        issues.append(failure_diagnostic(error, stage=where["stage"], project=project))
        exit_code = 2
    finally:
        _close_fonts(project)
    issues = _finish_issues(issues, scene)
    counts = count_issues(issues)
    if not exit_code and counts["error"]:
        exit_code = 1
    return {
        "command": "layout",
        "status": "failed" if exit_code else "passed",
        "project": str(project.path),
        "slides": len(scene["slides"]) if scene else 0,
        "selection": scene.get("selection") if scene else None,
        "fallback_required": bool(scene) and fallback_required(scene, project),
        "counts": counts,
        "issues": issues,
        "layout": _layout_document(scene, layouts) if scene is not None and layouts is not None else None,
        "timings": progress.timings,
        "elapsed_seconds": progress.elapsed,
        "exit_code": exit_code,
    }


def measure(path, texts, *, width=None, kind="text", role=None, size=None, page_kind="content", markup=False) -> dict:
    """文字列を資料の書式で計測する（`slide.measure()` と同じ計算）。要素を作らず、ファイルを書かない。

    width は枠の幅（pt）か "content"（page_kind のページの本文領域の幅）で、None なら折り返さない。
    size は役割の文字サイズを替える。markup が真なら `markup.parse` で $…$ を数式として読む。
    結果の各項目は、行（文字列・幅・折り返しの種類）、枠の寸法、1 行に収めるのに要る枠の幅を持つ。
    """
    from .markup import parse

    project = load_project(path)
    try:
        slide = Deck(project).slide(id="measure", kind=page_kind)
        frame = slide.content.w if width == "content" else width
        arguments = {} if size is None else {"size_pt": size}
        results = []
        for text in texts:
            content = parse(text) if markup and kind != "equation" else text
            layout = slide.measure(content, kind=kind, width=frame, role=role, **arguments)
            insets = layout.insets
            results.append(
                {
                    "text": text,
                    "lines": [
                        {"text": line.text, "width_pt": round(line.width, 2), "wrap": line.wrap}
                        for line in layout.lines
                    ],
                    "frame_width_pt": round(layout.width, 2),
                    "frame_height_pt": round(layout.height, 2),
                    "available_width_pt": (
                        None if layout.available_width is None else round(layout.available_width, 2)
                    ),
                    "natural_width_pt": round(layout.natural_width, 2),
                    "single_line_frame_width_pt": round(layout.natural_width + insets.left + insets.right, 2),
                }
            )
    finally:
        _close_fonts(project)
    return {
        "command": "measure",
        "status": "passed",
        "project": str(project.path),
        "kind": kind,
        "role": role,
        "size_pt": size,
        "width_pt": None if frame is None else round(frame, 2),
        "results": results,
        "exit_code": 0,
    }


def doctor(path, *, wait_seconds=None, progress=None) -> dict:
    """`layout()` の検査に加え、出力先・実行環境・表示ソフトの事前確認を行う。資料のディレクトリと出力先へは書かない。

    出力先に書き込むことができなければ例外にする。表示ソフトの事前確認が失敗しても例外にせず、失敗の診断を配置の診断と合わせて
    終了コード 2 で返す。結果の `environment` に既定フォントのフォルダー（`font_directories`）を入れる。
    """
    progress = Progress(None) if progress is None else progress
    project = load_project(path)
    _check_output_locations(project)
    scene, issues, environment, report, exit_code = None, [], {}, None, 0
    where = {"stage": "construct"}
    try:
        scene, issues, _ = _inspect(project, None, progress, where)
        where["stage"] = "environment"
        with progress.stage("environment"):
            runtime, environment = _environment(project, project.fonts)
        where["stage"] = "doctor"
        with progress.stage("doctor"):
            try:
                with open_session(
                    runtime,
                    wait_seconds=wait_seconds,
                    holder=_holder(project, "doctor", None),
                    on_event=_on_event(progress),
                ) as session:
                    detected = session.doctor()
            except RendererError as error:
                issues.append(failure_diagnostic(error, stage="doctor", project=project))
                report = {"status": "failed", "version": None, "message": str(error)}
                exit_code = 2
            else:
                environment["renderers"]["powerpoint"] = detected
                report = {"status": "passed", "version": detected.get("version"), "message": None}
        # 自動探索は、事前確認が通り Windows への問い合わせを記憶しているときだけ行う。
        # 事前確認に失敗したセッションでは記憶がないことがあり、探すと PowerShell を新たに起動するため。
        environment["font_directories"] = _font_directories(project, detect=report["status"] == "passed")
    except Exception as error:
        progress.error(_traceback_text(error))
        issues.append(failure_diagnostic(error, stage=where["stage"], project=project))
        exit_code = 2
    finally:
        _close_fonts(project)
    issues = _finish_issues(issues, scene)
    counts = count_issues(issues)
    if not exit_code and counts["error"]:
        exit_code = 1
    return {
        "command": "doctor",
        "status": "failed" if exit_code else "passed",
        "project": str(project.path),
        "slides": len(scene["slides"]) if scene else 0,
        "fallback_required": bool(scene) and fallback_required(scene, project),
        "environment": environment,
        "powerpoint": report,
        "counts": counts,
        "issues": issues,
        "timings": progress.timings,
        "elapsed_seconds": progress.elapsed,
        "exit_code": exit_code,
    }


def _check_output_locations(project):
    """出力先に近い既存ディレクトリで書き込み可否を確かめる。"""
    directories = {project.output_dir, *(Path(value).parent for value in project.config["deliverables"].values())}
    checked = set()
    for directory in directories:
        while not directory.exists():
            directory = directory.parent
        if not directory.is_dir():
            raise ValueError(f"出力先の親がディレクトリではありません: {directory}")
        if directory in checked:
            continue
        try:
            with tempfile.TemporaryFile(prefix=".slidekit-doctor-", dir=directory) as probe:
                probe.write(b"slidekit")
                probe.flush()
        except OSError as error:
            raise OSError(f"出力先へ書き込めません: {directory}: {error}") from error
        checked.add(directory)


def _image_name(label) -> str:
    """ページ画像の名前。元のページ番号なら slide-012.png、ページの識別子なら slide-{識別子}.png にする。

    識別子のうち Windows のファイル名で禁止されている文字と % は %XX にする（別の識別子と同じ名前にならない）。
    """
    if isinstance(label, int) and not isinstance(label, bool):
        return f"slide-{label:03d}.png"
    return "slide-" + _UNSAFE_NAME.sub(lambda match: f"%{ord(match.group()):02X}", str(label)) + ".png"


def _render_images(pdf, directory, labels, workers=None, *, overview=True):
    """PDF の各ページを画像にし、overview が真なら一覧画像も作る。

    labels は PDF の各ページの元のページ番号（資料のコードが選んだ章だけを作った部分生成ではページの識別子）で、
    画像の名前と一覧画像の見出しに使う（数が PDF のページ数と違えば誤り）。
    ページ画像は workers（`parallel.Workers`）でページごとに並列に作る。互換表示では一覧画像と縮小画像を作らない。
    画素のハッシュ値はページ画像・一覧画像とも `evidence.pixel_digest` の式で求める。
    """
    from PIL import Image, ImageDraw

    from .pdf import open_pdf

    with open_pdf(pdf) as document:
        count = len(document)
    if not count:
        raise ValueError("PDF にページがありません")
    if count != len(labels):
        raise ValueError("変換後のページ数が異なります")
    function = render_page_image if overview else partial(render_page_image, thumbnail=False)
    arguments = [
        (str(pdf), index, str(directory / _image_name(label)), RENDER_SCALE) for index, label in enumerate(labels)
    ]
    results = workers.map(function, arguments) if workers is not None else [function(*item) for item in arguments]
    pictures = [{"path": Path(result["path"]), "pixel_hash": result["pixel_hash"]} for result in results]
    if not overview:
        return pictures, []
    thumbnails = [
        (label, Image.frombytes("RGB", tuple(result["thumbnail_size"]), result["thumbnail"]))
        for label, result in zip(labels, results)
    ]
    contacts = []
    for start in range(0, len(thumbnails), 12):
        batch = thumbnails[start : start + 12]
        columns = 2 if len(batch) > 1 else 1
        canvas = Image.new("RGB", (columns * 500 + 20, ((len(batch) + columns - 1) // columns) * 306 + 20), "#e9edf1")
        draw = ImageDraw.Draw(canvas)
        for i, (label, thumbnail) in enumerate(batch):
            x, y = 20 + i % columns * 500, 20 + i // columns * 306
            canvas.paste(thumbnail, (x, y))
            draw.text((x, y + 275), str(label), fill="#223344")
        path = directory / f"contact-{start // 12 + 1:03d}.png"
        canvas.save(path)
        contacts.append({"path": path, "pixel_hash": pixel_digest(canvas.width, canvas.height, canvas.tobytes())})
    return pictures, contacts


def _declared_fonts(project):
    """表示環境の証拠に入れるフォント：実行環境の書体（宣言した書体と、テーマの書体として自動で見つけた書体）ごとの内容の証拠。

    計測で使った書体の一覧（`registry.inventory()`）は使わない。使う書体が増えただけで全ページの表示の証拠が変わらないようにする。
    ハッシュ値は、計測で使った書体なら生成記録の入力の値を使い、使っていない書体だけここで求める（入力には加えない）。
    """

    def digest(path):
        value = project.input_files.get(str(path))
        if value is None:
            try:
                value = file_hash(path)
            except OSError:
                value = "missing"
        return value

    result = {}
    for family, font in sorted(project.environment.get("fonts", {}).items()):
        if not isinstance(font, dict) or not font.get("path"):
            continue
        bold = font.get("bold_path")
        result[family] = {
            "sha256": digest(font["path"]),
            "index": font.get("index", 0),
            "bold_sha256": digest(bold) if bold else None,
            "bold_index": font.get("bold_index", 0) if bold else None,
        }
    return result


def _link(source: Path, target: Path) -> None:
    """target を source のハードリンクにする。できなければ複製する（PDF を二重に保存しない）。"""
    try:
        os.link(source, target)
    except OSError:
        shutil.copyfile(source, target)


def _wide(character: str) -> bool:
    return unicodedata.east_asian_width(character) in ("W", "F")


def _title_text(title) -> str:
    """題名を文字列にする。断片の列（作成 API の題名）なら、数式の断片は "$…$" として含める。"""
    if isinstance(title, dict):
        title = [title]
    if isinstance(title, (list, tuple)):
        parts = []
        for fragment in title:
            if not isinstance(fragment, dict):
                parts.append(str(fragment))
            elif fragment.get("kind") == "math":
                parts.append(f"${fragment.get('text', '')}$")
            else:
                parts.append(str(fragment.get("text", "")))
        return "".join(parts)
    return "" if title is None else str(title)


def _heading_text(title) -> str:
    """ノートの見出し。題名の各行の前後の空白を除いてつなぐ。

    前の行の末尾と次の行の先頭がともに全角の文字（east_asian_width が W・F）なら何も挟まず、それ以外は半角スペース 1 つでつなぐ。
    """
    result = ""
    for line in _title_text(title).splitlines():
        line = line.strip()
        if not line:
            continue
        if result and not (_wide(result[-1]) and _wide(line[0])):
            result += " "
        result += line
    return result


def notes_markdown(scene: dict) -> str:
    """ノート（notes.md）。各ページを「## {元のページ番号}. {見出し}」とノートで並べ、最後に改行を 1 つ付ける。題名が空なら番号だけにする。

    資料のコードが選んだ章だけを作った部分生成では、作ったページの中の番号は元のページ番号でないので、番号の代わりに
    ページの識別子を示す（「## {識別子}. {見出し}」）。
    """
    unnumbered = _unnumbered(scene)
    sections = []
    for position, page in enumerate(scene.get("slides", []), 1):
        heading = _heading_text(page.get("title", ""))
        number = page.get("id") if unnumbered else page.get("number", position)
        head = f"## {number}. {heading}" if heading else f"## {number}."
        sections.append(f"{head}\n\n{page.get('notes') or ''}")
    return "\n\n".join(sections) + "\n"


class _Build:
    """1 回の生成の状態と段階。`build()` が作り、`execute()` の後（例外でも）に `record()` で記録を書く。"""

    def __init__(self, project, run, progress, *, spec, draft, stop_on_error, jobs, wait_seconds, keep_intermediate):
        self.project, self.run, self.progress = project, run, progress
        self.spec, self.draft, self.stop_on_error = spec, draft, stop_on_error
        self.jobs, self.wait_seconds, self.keep_intermediate = jobs, wait_seconds, keep_intermediate
        self.stage, self.status = "construct", "failed"
        self.scene = None
        self.issues: list[dict] = []
        self.operations: dict = {"parallel": {"jobs": jobs}}
        self.environment: dict = {}
        self.runtime = None
        self.layouts = None
        self.slides: list[dict] = []
        self.requirements: list[dict] = []
        self.deck_hash = None
        self.fallback = False
        self.session = None
        self.images: dict = {}
        self.workers = None
        self.declared: dict = {}
        # 検査の段階の名前から結果。PDF の検査と画像化は変換の順に入る（集約がこの順に読む）。
        self.results: dict = {}
        # 成果物の相対パスから、途中で求めた SHA-256（同じファイルを求め直さない）。
        self.hashes: dict[str, str] = {}

    # 段階

    @contextlib.contextmanager
    def step(self, name: str):
        """段階を始める。生成中の印の段階を書き換え、進捗と所要時間を記録する。"""
        self.stage = name
        runs.update_marker(self.run, stage=name)
        with self.progress.stage(name):
            yield

    def checked(self, name: str, function, *args):
        """検査の段階を行い、結果を名前で残す（集約が使う）。"""
        with self.step(name):
            value = function(*args)
        self.results[name] = value
        return value

    # 生成の流れ

    def execute(self) -> None:
        project, run = self.project, self.run
        with self.step("construct"):
            self.scene = scene = _select(construct(project), self.spec, project)
        with self.step("assets"):
            self.operations["assets"] = prepare_assets(scene, project, run / "assets")
        with self.step("layout"):
            issues, self.layouts = _check_layout(project, scene)
            self.issues.extend(issues)
            # 変換の前に、構築結果と配置の診断を書く（中断した実行でも読むことができる）。
            write_json(run / "deck.json", _recorded_scene(scene))
            (run / "notes.md").write_text(notes_markdown(scene), encoding="utf-8")
            write_json(run / "layout.json", _layout_document(scene, self.layouts))
            self.write_diagnostics(provisional=True)
            # 内容の証拠は素材の準備の後の場面で求める（画像は元素材の内容の証拠を持ち、実行ごとの変換先を使わない）。
            self.deck_hash = deck_digest(scene, project)
            # 選んだ章だけを作った部分生成では、作ったページの中の番号を記録しない（元のページ番号でないため）。
            unnumbered = _unnumbered(scene)
            self.slides = [
                {
                    "id": page["id"],
                    "number": None if unnumbered else page.get("number", number),
                    "content_hash": content_digest(page, project),
                    "content_parts": content_parts(page, project),
                    "page_parts": None,
                    "visuals": {},
                }
                for number, page in enumerate(scene["slides"], 1)
            ]
            self.requirements = [
                {
                    "kind": "content",
                    "slide_id": page["id"],
                    "renderer": None,
                    "variant": None,
                    "evidence_hash": page["content_hash"],
                }
                for page in self.slides
            ]
        self.fallback = not self.draft and fallback_required(scene, project)
        deck = run / "deck.pptx"
        written = run / "intermediate/native.pptx" if self.fallback else deck
        with self.step("pptx"):
            # 書き出しと検査には、計測と同じ書体（project.fonts）を渡す（書体の区切りと中央揃えの補正が一致する）。
            self.operations["write"] = write_pptx(scene, project, written, fonts=project.fonts)
            self.hashes[written.relative_to(run).as_posix()] = self.operations["write"]["sha256"]
            if self.fallback:
                # 変換環境に問題があっても、生成済みの編集可能な資料は残す（互換画像を作ったら置き換える）。
                shutil.copyfile(written, deck)
                self.hashes["deck.pptx"] = self.operations["write"]["sha256"]
        with self.step("environment"):
            # 書き出しまでに使った書体で、変換に渡す実行環境と生成記録の環境を作る。
            self.runtime, self.environment = _environment(project, project.fonts)
            self.declared = _declared_fonts(project)
            # 既定フォントのフォルダーは設定の値か None。PowerPoint の変換セッションを開いたら、自動探索の結果で上書きする。
            self.environment["font_directories"] = _font_directories(project, detect=False)
        if self.stop_on_error and any(issue.get("severity") == "error" for issue in self.issues):
            # 変換の前に止める。互換画像はまだないので、その欠落は報告しない。
            with self.step("pptx_check"):
                issues, parts = self.check_pptx(fallback_expected=False)
            self.issues.extend(issues)
            self.set_page_parts(parts)
            self.status, self.stage = "stopped", "layout"
            return
        with Workers(self.jobs) as workers, contextlib.ExitStack() as stack:
            self.workers = workers
            try:
                self.convert(stack)
            finally:
                # powershell_launches は、そのセッションが起動した PowerShell の数である（作業プロセスと、セッションの中で
                # 初めて行った Windows への問い合わせ）。フォントの自動探索などで先に問い合わせた生成では、プロセス全体で
                # 2 回起動しても 1 になる。
                self.operations["sessions"] = {} if self.session is None else {"powerpoint": self.session.operations}
        with self.step("collect"):
            self.collect()
        self.status, self.stage = "built", "complete"

    def convert(self, stack) -> None:
        """事前確認と変換の後に、PPTX の検査と PDF の検査・画像化を順に行う。

        PowerPoint の変換は 1 つのセッション（排他を保持する）で行い、最後の変換の後に閉じる。検査と画像化は、すべての変換の後に、
        変換の順に行う（変換を待つ間に別のスレッドで進める方式は、短縮が 2 秒に満たなかったので採らない）。
        """
        project, run = self.project, self.run
        with self.step("doctor"):
            self.session = stack.enter_context(
                open_session(
                    self.runtime,
                    wait_seconds=self.wait_seconds,
                    holder=_holder(project, "build", run),
                    on_event=_on_event(self.progress),
                )
            )
            self.environment["renderers"]["powerpoint"] = self.session.doctor()
            # セッションを開くことができたので Windows への問い合わせを記憶している（探しても PowerShell の起動は増えない）。
            self.environment["font_directories"] = _font_directories(project, detect=True)
        if self.fallback:
            rendered = self.convert_with_fallback()
        else:
            pdf = run / "renders/powerpoint/standard/deck.pdf"
            result, date = self.render("standard", run / "deck.pptx", pdf)
            # 最後の変換の後に排他を解放する（以後の検査・画像化は排他の外で行う）。
            self.close_session()
            rendered = [("standard", pdf, result, date, None)]
        self.checked("pptx_check", self.check_pptx)
        for variant, pdf, result, date, kept in rendered:
            self.checked(f"pdf:powerpoint:{variant}", self.process_pdf, variant, pdf, result, date, kept)

    def convert_with_fallback(self) -> list[tuple]:
        """互換画像を付ける資料の変換（標準表示 1 回、互換表示 1 回）。

        検査する PDF の列（表示方法、PDF、変換の結果、生成した日、確認用複製に残したページ）を返す。
        """
        project, run = self.project, self.run
        deck, native = run / "deck.pptx", run / "intermediate/native.pptx"
        final_form = run / "intermediate/final_form.pptx"
        scale = project.config["math"].get("fallback_scale", 4)
        with self.step("final_form"):
            # 入力のハッシュ値は書き出しと変換で求めた値を渡す（同じファイルを読み直さない）。
            prepared = prepare_final_form(
                native, final_form, scale=scale, source_sha256=self.hashes.get("intermediate/native.pptx")
            )
            self.hashes["intermediate/final_form.pptx"] = prepared["sha256"]
        standard = run / "renders/powerpoint/standard/deck.pdf"
        standard_result, standard_date = self.render("standard", final_form, standard)
        with self.step("math_fallback"):
            self.hashes.pop("deck.pptx", None)
            finalized = finalize_math(
                final_form,
                standard,
                deck,
                scale=scale,
                workers=self.workers,
                source_pptx_sha256=prepared["sha256"],
                source_pdf_sha256=standard_result.get("output_sha256"),
            )
            self.hashes["deck.pptx"] = finalized["sha256"]
            self.operations["math_fallback"] = finalized
            # 最終 PPTX と、標準表示の確認に変換した最終形の PPTX が、互換画像の中身を除いて同一であることを確かめる。
            structure = compare_package_structure(final_form, deck)
            self.operations["final_form"] = {
                **{key: value for key, value in prepared.items() if key not in ("output", "shapes")},
                "structure": structure,
                "kept": self.keep_intermediate,
            }
            if not structure["matched"]:
                self.issues.append(
                    make_diagnostic(
                        "final_form_mismatch",
                        "最終 PPTX と、標準表示の確認に変換した最終形の PPTX のスライドの XML か関係が一致しない",
                        scope="build",
                        observations={
                            "differences": structure["differences"][:20],
                            "count": len(structure["differences"]),
                        },
                    )
                )
            self.issues.extend(self.clipped_crops(finalized["crops"]))
        clone = run / "intermediate/fallback.pptx"
        with self.step("fallback_clone"):
            self.operations["fallback_clone"] = {
                **prepare_fallback(deck, clone, math_only=True, source_sha256=finalized["sha256"]),
                "kept": self.keep_intermediate,
            }
            self.hashes["intermediate/fallback.pptx"] = self.operations["fallback_clone"]["sha256"]
        pdf = run / "renders/powerpoint/fallback/deck.pdf"
        result, date = self.render("fallback", clone, pdf)
        self.close_session()
        return [
            ("standard", standard, standard_result, standard_date, None),
            ("fallback", pdf, result, date, self.operations["fallback_clone"]["kept_slides"]),
        ]

    def close_session(self) -> None:
        """変換セッションを閉じる。自分で起動した PowerPoint の終了を待つ時間（数秒）を、所要時間の段階に入れる。"""
        with self.step("close:powerpoint"):
            self.session.close()

    def render(self, variant: str, source: Path, pdf: Path) -> tuple[dict, str]:
        """1 回の変換。戻り値は（変換の結果、生成した日）。日付のフィールドを表示するページの証拠に入れる（UTC の日付）。"""
        name = f"render:powerpoint:{variant}"
        with self.step(name):
            pdf.parent.mkdir(parents=True, exist_ok=True)
            result = self.session.render(source, pdf, label=variant)
            self.operations[name] = result
            if result.get("output_sha256"):
                self.hashes[pdf.relative_to(self.run).as_posix()] = result["output_sha256"]
        return result, datetime.now(timezone.utc).date().isoformat()

    def check_pptx(self, *, fallback_expected=None) -> tuple[list[dict], dict]:
        """最終 PPTX の検査と、表示の証拠に使うページの部品（最終 PPTX から 1 回だけ求める）。"""
        deck = self.run / "deck.pptx"
        if fallback_expected is None and self.draft:
            fallback_expected = False
        issues = validate_pptx(
            deck, self.scene, self.project, fonts=self.project.fonts, fallback_expected=fallback_expected
        )
        parts = pptx_pages(deck)
        if list(parts) != [page["id"] for page in self.scene["slides"]]:
            raise ValueError(f"最終 PPTX のページが構築結果と一致しない: {', '.join(parts)}")
        return issues, parts

    def clipped_crops(self, crops: list[dict]) -> list[dict]:
        """互換画像の切り出しが範囲で切り詰められた数式の警告。

        超えた部分に範囲が掛かるほかの要素（`nearby_elements`）があれば、観測値に加え、文言でその要素を示す。
        ない警告の観測値は変えない（許容の記録の証拠を保つ）。
        """
        pages = {page["id"]: page for page in self.scene["slides"]}
        issues = []
        for crop in crops:
            if not crop.get("clipped"):
                continue
            page = pages.get(crop.get("slide_id")) or {}
            element = next(
                (item for item in page.get("elements", []) if item.get("id") == crop.get("element_id")), None
            )
            observed = {
                key: crop.get(key) for key in ("rectangle_pt", "frame_pt", "crop_pt", "margin_pt", "outline_pt")
            }
            nearby = list(crop.get("nearby_elements") or [])
            message = "数式の描画が互換画像の範囲を超えている。文字枠を広げる"
            if nearby:
                observed["nearby_elements"] = nearby
                message = (
                    f"数式の描画が互換画像の範囲を超えている。超えた部分に要素 {', '.join(nearby)} の範囲が掛かる。"
                    "超えた描画がその要素のものなら数式は欠けない。互換表示を確かめる"
                )
            issues.append(
                make_diagnostic(
                    "fallback_crop_clipped",
                    message,
                    scope="pptx",
                    slide_id=crop.get("slide_id"),
                    element_id=crop.get("element_id"),
                    slide_number=page.get("number"),
                    observations=observed,
                    evidence={"element": element, "crop_pt": crop.get("crop_pt"), "margin_pt": crop.get("margin_pt")},
                )
            )
        return issues

    def process_pdf(self, variant: str, pdf: Path, render: dict, render_date: str, kept=None) -> dict:
        """変換した PDF の検査と画像化。PPTX の検査（ページの部品）の後に行う。

        互換表示の確認用複製は数式のページだけなので、検査には複製のページだけを複製の順に並べた場面を渡し、
        表示の証拠の番号は複製の中の位置にする。画像の名前は元のページ番号（選んだ章だけを作った部分生成ではページの識別子）である。
        """
        _, parts = self.results["pptx_check"]
        project, scene = self.project, self.scene
        if kept is not None:
            positions = {item["id"]: item["number"] for item in kept}
            pages = {page["id"]: page for page in scene["slides"]}
            checked = {**scene, "slides": [pages[item["id"]] for item in kept]}
        else:
            positions = {slide_id: item.number for slide_id, item in parts.items()}
            checked = scene
        actual = {key: value for key, value in render.items() if key in _ACTUAL_KEYS}
        environment_hash = display_environment(
            self.environment["renderers"].get("powerpoint"),
            actual,
            self.declared,
            self.environment["packages"].get("pypdfium2"),
            scale=RENDER_SCALE,
        )
        evidence = {
            slide_id: visual_digest(parts[slide_id], variant, environment_hash, render_date=render_date, number=number)
            for slide_id, number in positions.items()
        }
        bodies = {
            slide_id: visual_body_digest(
                parts[slide_id], variant, environment_hash, render_date=render_date, number=number
            )
            for slide_id, number in positions.items()
        }
        record: dict = {}
        issues = validate_pdf(
            pdf,
            checked,
            project,
            variant,
            fonts=project.fonts,
            layouts=self.layouts,
            page_evidence=evidence,
            layout_record=record,
        )
        write_json(pdf.parent / "layout.json", record)
        order = sorted(positions, key=positions.get)
        unnumbered = _unnumbered(scene)
        labels = {page["id"]: page["id"] if unnumbered else page["number"] for page in self.slides}
        pictures, contacts = _render_images(
            pdf, pdf.parent, [labels[slide_id] for slide_id in order], self.workers, overview=variant == "standard"
        )
        return {
            "issues": issues,
            "environment_hash": environment_hash,
            "evidence": evidence,
            "bodies": bodies,
            "pictures": dict(zip(order, pictures)),
            "contacts": contacts,
        }

    def set_page_parts(self, parts) -> None:
        for page in self.slides:
            page["page_parts"] = parts[page["id"]].to_json()

    def collect(self) -> None:
        """検査の結果、表示の証拠、確認対象を、変換の順にまとめる。"""
        run = self.run
        issues, parts = self.results["pptx_check"]
        self.issues.extend(issues)
        self.set_page_parts(parts)
        for name, value in self.results.items():
            if not name.startswith("pdf:"):
                continue
            key = name.removeprefix("pdf:")
            variant = key.rpartition(":")[2]
            self.issues.extend(value["issues"])
            for page in self.slides:
                picture = value["pictures"].get(page["id"])
                if picture is None:
                    continue
                evidence_hash = value["evidence"][page["id"]]
                page["visuals"][key] = {
                    "path": picture["path"].relative_to(run).as_posix(),
                    "pixel_hash": picture["pixel_hash"],
                    "environment_hash": value["environment_hash"],
                    "evidence_hash": evidence_hash,
                }
                body = value.get("bodies", {}).get(page["id"])
                if body is not None:
                    # ページ番号の図形を除いた証拠（番号の枠・書式・文字だけが変わったページの引き継ぎに使う）。
                    page["visuals"][key]["body_hash"] = body
                self.requirements.append(
                    {
                        "kind": "visual",
                        "slide_id": page["id"],
                        "renderer": "powerpoint",
                        "variant": variant,
                        "evidence_hash": evidence_hash,
                    }
                )
            if variant == "standard":
                self.requirements.append(
                    {
                        "kind": "overview",
                        "slide_id": None,
                        "renderer": "powerpoint",
                        "variant": variant,
                        "evidence_hash": overview_digest(
                            [contact["pixel_hash"] for contact in value["contacts"]],
                            value["environment_hash"],
                            self.deck_hash,
                        ),
                    }
                )
                shown = value["contacts"]
            else:
                shown = [value["pictures"][slide_id] for slide_id in value["pictures"]]
            self.images[key] = [item["path"].relative_to(run).as_posix() for item in shown]
        standard = run / "renders/powerpoint/standard/deck.pdf"
        if standard.is_file():
            _link(standard, run / "deck.pdf")
            name = standard.relative_to(run).as_posix()
            if name in self.hashes:
                self.hashes["deck.pdf"] = self.hashes[name]

    # 失敗と記録

    def failed(self, error: BaseException) -> None:
        if isinstance(error, _SlidesError):
            # --slides の指定の誤りは利用者の入力の誤りなので、トレースバックを書かず、文言だけの診断にする。
            self.issues.append(_selection_diagnostic(error, self.stage))
            self.status = "failed"
            return
        trace = _traceback_text(error)
        path = self.run / "traceback.txt"
        try:
            path.write_text(trace, encoding="utf-8")
        except OSError:
            path = None
        self.progress.error(trace)
        self.issues.append(failure_diagnostic(error, stage=self.stage, project=self.project, traceback_path=path))
        self.status = "failed"

    def interrupted(self) -> None:
        self.issues.append(make_diagnostic("build_interrupted", "利用者が生成を中断しました", scope=self.stage))
        self.status = "failed"

    def write_diagnostics(self, *, provisional: bool = False) -> list[dict]:
        issues = _finish_issues(self.issues, self.scene)
        value = {"schema_version": records.SCHEMA_VERSION, "issues": issues}
        if provisional:
            # 変換の前に書く暫定の診断。生成の終わりに書き直す。
            value["provisional"] = True
        write_json(self.run / "diagnostics.json", value)
        return issues

    def record(self) -> None:
        """診断・確認記録・生成記録を書き、生成中の印を消す。例外で止まった実行でも行う。"""
        try:
            self._write_records()
        finally:
            # 記録を書くことができなくても、生成中の印は残さない（中断の判定は生成記録がないことで分かる）。
            runs.remove_marker(self.run)

    def _write_records(self) -> None:
        run, project = self.run, self.project
        with self.progress.stage("record"):
            if not self.keep_intermediate:
                for name in _INTERMEDIATE_FILES:
                    (run / name).unlink(missing_ok=True)
            self.issues = self.write_diagnostics()
            write_json(run / "review.json", {"schema_version": records.SCHEMA_VERSION, "records": []})
            artifacts = {}
            for file in sorted(run.rglob("*")):
                if file.is_file() and file.name not in _RECORD_FILES:
                    name = file.relative_to(run).as_posix()
                    artifacts[name] = self.hashes.get(name) or file_hash(file)
            scene = self.scene or {}
            manifest = {
                "schema_version": records.SCHEMA_VERSION,
                "evidence_version": EVIDENCE_VERSION,
                "run_id": run.name,
                "project": str(project.path),
                "project_root": str(project.root),
                "created_at": datetime.now(timezone.utc).isoformat(),
                "status": self.status,
                "stage": self.stage,
                "mode": self.mode,
                "selection": scene.get("selection"),
                "exportable": self.exportable,
                "timings": self.progress.timings,
                "elapsed_seconds": self.progress.elapsed,
                "input_files": project.input_files,
                "config": project.config,
                "theme": project.theme,
                "rules": project.rules,
                "environment": self.environment,
                "artifacts": artifacts,
                "deck_content_hash": self.deck_hash,
                "slides": self.slides,
                "required_reviews": self.requirements,
                "deliverables": project.config["deliverables"],
                "operations": self.operations,
                "presentation_seconds": sum(page.get("seconds") or 0 for page in scene.get("slides", [])),
            }
            write_json(run / "manifest.json", manifest)

    @property
    def mode(self) -> str:
        return "draft" if self.draft else "full"

    @property
    def exportable(self) -> bool:
        # 完成物にできるのは、全ページの通常の生成だけである。
        return not self.draft and (self.scene or {}).get("selection") is None

    def result(self) -> dict:
        counts = count_issues(self.issues)
        exit_code = 2 if self.status == "failed" else 1 if self.status == "stopped" or counts["error"] else 0
        scene = self.scene or {}
        return {
            "command": "build",
            "status": self.status,
            "build": str(self.run),
            "run_id": self.run.name,
            "project": str(self.project.path),
            "stage": self.stage,
            "mode": self.mode,
            "selection": scene.get("selection"),
            "exportable": self.exportable,
            "slides": len(scene.get("slides", [])),
            "fallback": self.fallback,
            "counts": counts,
            "issues": self.issues,
            "timings": self.progress.timings,
            "elapsed_seconds": self.progress.elapsed,
            "images": self.images,
            "exit_code": exit_code,
        }


def _carry_source(project, carry_from) -> tuple:
    """`build` の carry_from を生成の前に確かめ、（引き継ぎ元、生成の後に示す引き継ぎの失敗の文言）を返す。

    auto・none・None・実行の番号・実行ディレクトリのパスはそのまま使う。latest は生成した実行自身を指さないように、
    生成の前の完成物にできる最新の実行に固定する（なければ生成は行い、引き継ぎを失敗として示す）。
    それ以外の書式の誤りは実行ディレクトリを作る前に ValueError にする。
    """
    if carry_from is None or (isinstance(carry_from, str) and carry_from in ("auto", "none")):
        return carry_from, None
    text = str(carry_from)
    if text == "latest":
        try:
            return records.resolve_run("latest", project=project), None
        except records.RecordError as error:
            return None, str(error)
    if text and (_RUN_NUMBER.fullmatch(text) or Path(text).is_dir()):
        return carry_from, None
    raise ValueError(f"carry_from は auto、none、latest、実行の番号、実行ディレクトリのパスのどれかにする: {text}")


def _carry(project, state, carry_from, failure=None) -> dict | None:
    """生成の後の確認の引き継ぎ（`build` の carry_from）。引き継がないときは None を返す。

    完了した全ページの通常の生成だけが対象である（下書き・部分生成・停止・失敗の実行には引き継がない）。
    引き継ぎの失敗（ロックを取得できない、指定した実行が不正など。failure は生成の前に分かった失敗の文言）は生成の失敗にせず、
    status が skipped か failed の結果にする。成功したときは、引き継いだ件数（count）と、引き継ぎ元・番号の領域を除いて
    比べたページ（重複なし。標準表示と互換表示の両方を引き継いだページも 1 回）・引き継がなかった理由ごとの件数・
    残りの確認（自動の引き継ぎだけ）を返す。
    """
    if state.status != "built" or not state.exportable:
        return None
    if failure is not None:
        return {"status": "failed", "mode": "manual", "message": failure}
    if carry_from is None or carry_from == "none":
        return None
    automatic = isinstance(carry_from, str) and carry_from == "auto"
    try:
        if automatic:
            result = records.carry_reviews_auto(state.run)
        else:
            result = records.carry_reviews(carry_from, state.run, project=project)
    except Exception as error:
        return {
            "status": "failed",
            "mode": "auto" if automatic else "manual",
            "message": str(error) or type(error).__name__,
        }
    if result.get("status") != "carried":
        return {
            "status": result.get("status"),
            "mode": "auto" if automatic else "manual",
            "message": result.get("message"),
        }
    copied = result.get("records") or []
    if automatic:
        sources, skipped, remaining = result.get("sources") or {}, result.get("skipped") or {}, result.get("remaining")
    else:
        sources = {str(records.resolve_run(carry_from, project=project)): len(copied)} if copied else {}
        skipped = {}
        for item in result.get("skipped") or []:
            skipped[item["reason"]] = skipped.get(item["reason"], 0) + 1
        remaining = None
    # 引き継ぎは記録ごとにページを挙げる（数式のページは標準表示と互換表示の 2 件）ので、ページの重複を除く。
    renumbered = list(dict.fromkeys(result.get("renumbered") or []))
    return {
        "status": "carried",
        "mode": "auto" if automatic else "manual",
        "count": len(copied),
        "sources": sources,
        "renumbered": renumbered,
        "skipped": skipped,
        "skipped_runs": result.get("skipped_runs") or {},
        "remaining": remaining,
        "renumbered_from": result.get("renumbered_from") or {},
        "pending_reasons": result.get("pending_reasons") or [],
        "compared_run": result.get("compared_run"),
        "environment_changes": result.get("environment_changes") or [],
        "notices": result.get("notices") or [],
    }


# 生成の間、Ctrl+C（SIGINT）と同じ後片付けを通す終了の信号。Windows には SIGHUP がない。
_INTERRUPT_SIGNALS = tuple(getattr(signal, name) for name in ("SIGTERM", "SIGHUP") if hasattr(signal, name))


@contextlib.contextmanager
def _signals_as_interrupt():
    """SIGTERM と SIGHUP を KeyboardInterrupt にし、抜けるときに元の処理へ戻す。

    既定の処理（その場で終了）のままの信号だけを差し替える（呼び出し元が設定した処理は変えない）。信号の処理は主スレッドでしか
    設定できないので、ほかのスレッドでは何もしない。2 回目以降の信号は、変換の作業プロセスの切り離しと記録の書き込みを
    最後まで行うために無視する。
    """
    if threading.current_thread() is not threading.main_thread():
        yield
        return
    received = []

    def interrupt(number, frame):
        if not received:
            received.append(number)
            raise KeyboardInterrupt

    previous = {}
    for number in _INTERRUPT_SIGNALS:
        if signal.getsignal(number) is signal.SIG_DFL:
            previous[number] = signal.signal(number, interrupt)
    try:
        yield
    finally:
        for number, handler in previous.items():
            signal.signal(number, handler)


@_signals_as_interrupt()
def build(
    path,
    *,
    slides=None,
    draft=False,
    stop_on_error=None,
    jobs=None,
    wait_seconds=None,
    keep_intermediate=False,
    carry_from="auto",
    progress=None,
) -> dict:
    """新しい実行ディレクトリに資料を生成する。

    - slides：部分生成のページ（`selection` の書式）。選んだページだけを書き出す。完成物にできない。指定は生成関数に
      `context.selection` として渡し、資料のコードは選んだページを含む章だけを作ってよい（`context.selected()`）。
    - draft：下書き。互換画像を省く。完成物にできない。
    - stop_on_error：配置・記述規則・入力の検査でエラーがあれば、PPTX の検査の後に変換の前で止める（status "stopped"）。
      None なら下書きでだけ有効。
    - jobs：PowerPoint を使わない処理の同時プロセス数（`parallel.resolve_jobs`）。
    - wait_seconds：別の変換が終わるのを待つ上限（秒）。None なら実行環境の設定。
    - keep_intermediate：互換表示の確認用複製と最終形の PPTX を残す。
    - carry_from：生成の後に確認と許容を引き継ぐ元。"auto"（既定）は同じ実行保存先の前の実行から自動で
      （`records.carry_reviews_auto`）、実行のパス・番号はその実行から、"latest" は生成の前の完成物にできる最新の実行から、
      "none" は引き継がない。完了した全ページの通常の生成だけで行い、結果を戻り値の carry に入れる。引き継ぎの失敗は
      生成の失敗にせず、終了コードも変えない。書式の誤りは実行ディレクトリを作る前に ValueError にする。

    戻り値の status は生成記録と同じ（built・stopped・failed）で、品質の結果は counts と exit_code で示す。
    利用者の中断（KeyboardInterrupt）でも記録を書き、診断 build_interrupted を加え、その時点の戻り値を例外の属性 result に
    付けてから送り直す（コマンドが実行ディレクトリを示すことができるように）。生成の間は、主スレッドで呼んだとき、
    既定の処理のままの SIGTERM と SIGHUP も中断として扱う（`_signals_as_interrupt`）。変換中なら、作業プロセスを切り離して
    残存記録を残し、作業プロセスが変換の後に複製を閉じて記録を消し、自分で起動した PowerPoint を終了する。
    """
    progress = Progress(None) if progress is None else progress
    if slides is not None:
        # 書式の誤りは実行ディレクトリを作る前に知らせる。
        selection.parse(slides)
    project = load_project(path)
    project.selection = slides
    source, carry_failure = _carry_source(project, carry_from)
    jobs = resolve_jobs(jobs, project.environment)
    stop_on_error = bool(draft) if stop_on_error is None else bool(stop_on_error)
    options = {
        "draft": bool(draft),
        "slides": slides,
        "jobs": jobs,
        "stop_on_error": stop_on_error,
        "wait_seconds": wait_seconds,
        "keep_intermediate": bool(keep_intermediate),
        "carry_from": carry_from if isinstance(carry_from, str) or carry_from is None else str(carry_from),
    }
    run = runs.new_run(project, command="build", options=options)
    state = _Build(
        project,
        run,
        progress,
        spec=slides,
        draft=bool(draft),
        stop_on_error=stop_on_error,
        jobs=jobs,
        wait_seconds=wait_seconds,
        keep_intermediate=bool(keep_intermediate),
    )
    try:
        try:
            state.execute()
        except KeyboardInterrupt:
            state.interrupted()
            raise
        except Exception as error:
            state.failed(error)
        finally:
            _close_fonts(project)
            state.record()
        result = state.result()
        # 確認の引き継ぎは生成記録を書いて生成中の印を消した後に行う。
        carried = _carry(project, state, source, carry_failure)
        if carried is not None:
            result["carry"] = carried
        if result.get("status") == "built":
            # 前の実行とのページ画像の比較（変わっていないページ・番号だけ変わったページ・変わったページ）。
            try:
                result["comparison"] = records.compare_previous(state.run)
            except (records.RecordError, OSError, ValueError) as error:
                result["comparison"] = {"error": str(error)}
    except KeyboardInterrupt as interrupt:
        interrupt.result = state.result()
        raise
    return result
