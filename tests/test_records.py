"""確認記録・引き継ぎ・最終検査と、完成物を反映・復旧する処理の検証。

生成記録は合成し、証拠は生成と同じ evidence の関数で求める。ページ画像は小さな合成画像にする。
"""

from collections import Counter
import json
import os
from pathlib import Path
import re
import shlex
import shutil
import stat
import sys
import threading
import time
from types import SimpleNamespace

from PIL import Image, ImageDraw
import pytest

from conftest import DECK
from slidekit import records, runs
from slidekit.evidence import (
    EVIDENCE_VERSION,
    NOISE_PIXELS,
    RENDER_CHECK_METHOD,
    RENDER_SCALE,
    PageParts,
    content_digest,
    deck_digest,
    display_environment,
    overview_digest,
    pixel_digest,
    render_difference,
    visual_digest,
)
from slidekit.files import file_hash, json_digest

# 合成のページは 160×90 pt。番号の図形の領域（pt の [x, y, w, h]）。
PAGE_SIZE = (160, 90)
FOOTER_AREA = [128.0, 74.0, 24.0, 12.0]
RENDER_DATE = "2026-10-12"
SELECTION = {"spec": "1", "slide_ids": ["intro"], "numbers": [1], "total_slides": 2}
# 内容の証拠と資料全体の証拠を求めるための、最小の資料。
EVIDENCE_PROJECT = SimpleNamespace(config={"metadata": {"title": "確認記録の検査"}}, input_files={}, brief_path=None)
WARNING = {
    "diagnostic_id": "d_overlap",
    "code": "element_overlap",
    "slide_id": "intro",
    "element_id": "box",
    "related_element_id": "label",
    "scope": "layout",
    "severity": "warning",
    "message": "図形が重なる",
    "observations": {"area": 12},
    "evidence_hash": "evidence",
}
INPUT_NOTICE = "入力が生成後に変わっている。この実行は反映できない。記録は再生成後、証拠が一致すれば引き継がれる"


def save(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")


def page_image(path, *, shift=0, shape=True, number=1):
    """倍率 RENDER_SCALE で画像にしたページを模す。暗い長方形の図形 1 つと、番号の領域の中の番号ごとに長さの違う線を描く。"""
    image = Image.new("RGB", (PAGE_SIZE[0] * RENDER_SCALE, PAGE_SIZE[1] * RENDER_SCALE), "white")
    draw = ImageDraw.Draw(image)
    if shape:
        draw.rectangle([40 + shift, 40, 200 + shift, 120], fill="#20252B")
    x, y = (value * RENDER_SCALE for value in FOOTER_AREA[:2])
    draw.rectangle([x + 4, y + 4, x + 4 + 6 * number, y + 10], fill="#606B78")
    path.parent.mkdir(parents=True, exist_ok=True)
    image.save(path)
    return pixel_digest(image.width, image.height, image.tobytes())


def fixture_run(
    tmp_path,
    name="0001",
    *,
    warning=False,
    content="same",
    parts="same",
    environment="fonts-v1",
    payload="same",
    pages=("intro",),
    number=1,
    shift=0,
    shape=True,
    schema=None,
    mode="full",
    selection=None,
    evidence_version=EVIDENCE_VERSION,
    fallback=(),
    media="same",
    renderers=("powerpoint",),
):
    """合成した生成記録を持つ実行を tmp_path/project/out/runs/{name} に作る。

    content はページの内容（ノート）、parts は表示を決める PPTX の部品、environment は表示環境、payload は成果物の中身を変える。
    shift と shape は画像だけを変える（証拠は変えない）。number は最初のページの番号で、番号の図形の文字と画像に出る。
    schema=1 は slidekit 0.1 の形式の実行にする。mode と selection で下書き・部分生成の実行にする。
    fallback は数式のあるページ（PowerPoint の互換表示の確認が要る）、media は互換画像の中身（部品の証拠の full だけを変える）、
    renderers は表示ソフト（最初が主表示ソフト）である。
    """
    project = tmp_path / "project"
    project.mkdir(parents=True, exist_ok=True)
    source = project / "slides.py"
    if not source.exists():
        source.write_text("元資料", encoding="utf-8")
    config = project / "deck.toml"
    if not config.exists():
        config.write_text(DECK, encoding="utf-8")
    run = project / "out" / "runs" / name
    run.mkdir(parents=True)
    version = records.SCHEMA_VERSION if schema is None else schema
    save(run / "diagnostics.json", {"schema_version": version, "issues": [WARNING] if warning else []})
    for filename in ("deck.pptx", "deck.pdf", "deck.json", "notes.md"):
        (run / filename).write_text(f"{filename}:{payload}", encoding="utf-8")
    environments = {
        renderer: display_environment(
            renderer, {"version": environment}, {"renderer": renderer, "version": environment}, {}, "fixture"
        )
        for renderer in renderers
    }
    contacts = {
        renderer: page_image(
            run / f"renders/{renderer}/standard/contact-001.png", shift=shift, shape=shape, number=number
        )
        for renderer in renderers
    }
    slides, contents, visuals = [], [], []
    clone = 0
    for index, slide_id in enumerate(pages):
        page_number = number + index
        full = ["full", slide_id, parts] + ([] if media == "same" else [media])
        page = PageParts(
            slide_id,
            index + 1,
            json_digest(["choice", slide_id, parts]),
            json_digest(full),
            {"area": FOOTER_AREA, "text": str(page_number)},
            (),
        )
        shown = {}
        views = [(renderer, "standard") for renderer in renderers] + (
            [("powerpoint", "fallback")] if slide_id in fallback else []
        )
        for renderer, variant in views:
            path = f"renders/{renderer}/{variant}/slide-{page_number:03d}.png"
            pixel_hash = page_image(run / path, shift=shift, shape=shape, number=page_number)
            if variant == "fallback":
                # 互換表示の確認用複製は数式のページだけなので、証拠の位置は複製の中の位置である。
                clone += 1
            evidence = visual_digest(
                page,
                renderer,
                variant,
                environments[renderer],
                render_date=RENDER_DATE,
                number=clone if variant == "fallback" else None,
            )
            shown[f"{renderer}:{variant}"] = {
                "path": path,
                "pixel_hash": pixel_hash,
                "environment_hash": environments[renderer],
                "evidence_hash": evidence,
            }
            visuals.append(
                {
                    "kind": "visual",
                    "slide_id": slide_id,
                    "renderer": renderer,
                    "variant": variant,
                    "evidence_hash": evidence,
                }
            )
        content_hash = content_digest(
            {"id": slide_id, "title": slide_id, "notes": content, "elements": []}, EVIDENCE_PROJECT
        )
        slides.append(
            {
                "id": slide_id,
                "number": page_number,
                "content_hash": content_hash,
                "page_parts": page.to_json(),
                "visuals": shown,
            }
        )
        contents.append(
            {"kind": "content", "slide_id": slide_id, "renderer": None, "variant": None, "evidence_hash": content_hash}
        )
    deck = deck_digest(
        {"slides": [{"id": slide_id, "title": slide_id} for slide_id in pages], "size": list(PAGE_SIZE)},
        EVIDENCE_PROJECT,
    )
    overviews = [
        {
            "kind": "overview",
            "slide_id": None,
            "renderer": renderer,
            "variant": "standard",
            "evidence_hash": overview_digest([contacts[renderer]], environments[renderer], deck),
        }
        for renderer in renderers
    ]
    manifest = {
        "schema_version": version,
        "evidence_version": evidence_version,
        "run_id": name,
        "project": str(config),
        "project_root": str(project),
        "status": "built",
        "stage": "complete",
        "mode": mode,
        "selection": selection,
        "exportable": mode == "full" and selection is None,
        "input_files": {str(source): file_hash(source), str(config): file_hash(config)},
        "config": {
            "output_dir": str(project / "out"),
            "inputs": {"read_roots": [str(project)], "asset_dirs": [str(project / "assets")]},
            "render": {"primary": renderers[0], "required": list(renderers)},
        },
        "environment": {},
        "artifacts": {str(path.relative_to(run)): file_hash(path) for path in sorted(run.rglob("*")) if path.is_file()},
        "deck_content_hash": deck,
        "slides": slides,
        "required_reviews": [*contents, *visuals, *overviews],
        "deliverables": {
            "pptx": str(project / "presentation.pptx"),
            "pdf": str(project / "presentation.pdf"),
            "notes": str(project / "speaker_notes.md"),
            "record": str(project / "deliverables.json"),
        },
    }
    if version == 1:
        # 0.1 の生成記録には、証拠の形式・生成の種類・ページの部品と表示の証拠がない。
        for key in ("evidence_version", "mode", "selection", "exportable", "deck_content_hash"):
            del manifest[key]
        for slide in manifest["slides"]:
            del slide["page_parts"]
            for visual in slide["visuals"].values():
                del visual["evidence_hash"]
    save(run / "manifest.json", manifest)
    save(run / "review.json", {"schema_version": version, "records": []})
    return run


def complete(run, warning=False):
    """全ページの表示・一覧画像・内容を確認し、warning が真なら警告も許容する。"""
    result = {}
    for kind in ("visual", "overview", "content"):
        result[kind] = records.record_review(run, kind, "all", "確認者", "表示と内容を確認")
    if warning:
        result["warning"] = records.accept_warning(run, "d_overlap", "確認者", "図の意図に沿う重なりである")
    return result


def project_for(run):
    """整理（runs.clean_runs）に渡す資料。生成記録は形式を問わずに読む。"""
    manifest = json.loads((run / "manifest.json").read_text(encoding="utf-8"))
    return SimpleNamespace(
        root=Path(manifest["project_root"]),
        output_dir=run.parent.parent,
        config={"deliverables": manifest["deliverables"]},
    )


def rewrite_manifest(run, mutate):
    manifest = records._manifest(run)
    mutate(manifest)
    save(run / "manifest.json", manifest)


def visual_of(run, slide_id="intro"):
    slide = next(item for item in records._manifest(run)["slides"] if item["id"] == slide_id)
    return slide["visuals"]["powerpoint:standard"]


def carried_visual(result):
    return next(record for record in result["records"] if record["kind"] == "visual")


def set_diagnostics(run, issues):
    """診断の一覧を書き換え、生成記録の成果物のハッシュ値を合わせる。"""
    save(run / "diagnostics.json", {"schema_version": records.SCHEMA_VERSION, "issues": issues})
    rewrite_manifest(
        run, lambda value: value["artifacts"].update({"diagnostics.json": file_hash(run / "diagnostics.json")})
    )


def warning(identifier, slide_id, element_id, related=None, *, code="element_overlap", **extra):
    """警告の診断（証拠は識別子ごとに別の値）。"""
    return {
        **WARNING,
        "diagnostic_id": identifier,
        "code": code,
        "slide_id": slide_id,
        "element_id": element_id,
        "related_element_id": related,
        "evidence_hash": f"evidence-{identifier}",
        **extra,
    }


def direct_record(run, kind="content", slide_id="intro"):
    """記録の操作を通さずに、その実行に正しく結び付いた人の記録を書く（下書き・部分生成の実行の記録を再現する）。"""
    manifest = records._manifest(run)
    target = next(
        item for item in manifest["required_reviews"] if item["kind"] == kind and item["slide_id"] == slide_id
    )
    review = records._review(run)
    record = {
        "id": f"r_direct{len(review['records']):010d}",
        "kind": kind,
        "slide_id": slide_id,
        "renderer": target["renderer"],
        "variant": target["variant"],
        "evidence_hash": target["evidence_hash"],
        "reviewer": "確認者",
        "note": "内容を確認",
        "created_at": "2026-10-12T00:00:00+00:00",
        "binding": records._binding(run, manifest),
    }
    review["records"].append(record)
    save(run / "review.json", review)
    return record


def run_command(command, **replacements):
    """review status が示したコマンドを、引数の値のまま records の関数に渡す。

    replacements は書き換える引数（reviewer・note・reason など）。コマンドの解釈そのものは CLI の試験が確かめる。
    """
    tokens = shlex.split(command)
    assert tokens[:2] == ["slidekit", "review"] and len(tokens[3:]) % 2 == 0, command
    options = dict(zip(tokens[3::2], tokens[4::2]))
    options.update({f"--{key.replace('_', '-')}": value for key, value in replacements.items()})
    if tokens[2] == "record":
        return records.record_review(
            options["--build"],
            options["--kind"],
            options.get("--slides"),
            options["--reviewer"],
            options["--note"],
            options.get("--renderer"),
            options.get("--variant", "standard"),
            status_id=options.get("--status-id"),
        )
    assert tokens[2] == "accept", command
    return records.accept_warnings(
        options["--build"],
        reviewer=options["--reviewer"],
        reason=options["--reason"],
        diagnostic_ids=options.get("--diagnostic"),
        code=options.get("--code"),
        slides=options.get("--slides"),
        element=options.get("--element"),
    )


# 確認と最終検査


def test_final_requires_reviews_and_warning_acceptance(tmp_path):
    run = fixture_run(tmp_path, warning=True)
    assert records.check_build(run)["status"] == "passed"
    result = records.check_build(run, final=True)
    assert result["exit_code"] == 1
    assert len([item for item in result["issues"] if item["code"] == "review_missing"]) == 3
    assert any(item["code"] == "warning_unaccepted" for item in result["issues"])
    complete(run, warning=True)
    assert records.check_build(run, final=True)["status"] == "passed"


@pytest.mark.parametrize("filename, expected", [("deck.pptx", "artifact_changed"), ("../slides.py", "input_changed")])
def test_changed_inputs_or_artifacts_block_record_check_and_export(tmp_path, filename, expected):
    run = fixture_run(tmp_path)
    complete(run)
    root = run.parent.parent.parent
    path = root / "slides.py" if filename.startswith("..") else run / filename
    path.write_text("変更", encoding="utf-8")
    result = records.check_build(run, final=True)
    assert result["exit_code"] == 2
    assert any(issue["code"] == expected for issue in result["issues"])
    if expected == "artifact_changed":
        # 確認記録は成果物に結び付くので、成果物が変わった実行には記録できない。識別子は整合検査と同じである。
        with pytest.raises(records.RecordError, match="成果物") as caught:
            records.record_review(run, "content", "all", "確認者", "内容確認")
        assert caught.value.code == "artifact_changed"
    else:
        # 入力の変化では記録でき、その実行を反映できないことを知らせる。
        recorded = records.record_review(run, "content", "all", "確認者", "内容確認")
        assert [record["kind"] for record in recorded["records"]] == ["content"]
        assert recorded["notices"] == [
            {"code": "input_changed", "message": INPUT_NOTICE, "paths": [str(path)], "count": 1}
        ]
    assert records.export_build(run)["status"] == "failed"
    assert not (root / "presentation.pptx").exists()


def test_final_check_survives_later_check(tmp_path):
    run = fixture_run(tmp_path)
    complete(run)
    assert records.check_build(run, final=True)["status"] == "passed"
    saved = (run / "final_check.json").read_bytes()
    assert records._read_json(run / "final_check.json")["final"] is True
    assert not (run / "check.json").exists()
    # 最終でない検査は check.json に書き、最終検査の結果を上書きしない。
    records.check_build(run)
    assert (run / "final_check.json").read_bytes() == saved
    assert records._read_json(run / "check.json")["final"] is False


def test_check_summary_counts_reviews_warnings_and_declarations(tmp_path):
    run = fixture_run(tmp_path, warning=True)
    declared = {
        **WARNING,
        "diagnostic_id": "d_declared",
        "element_id": "note",
        "declared": {"reason": "意図した重なり", "location": "slides.py:12"},
    }
    save(run / "diagnostics.json", {"schema_version": records.SCHEMA_VERSION, "issues": [WARNING, declared]})
    rewrite_manifest(
        run, lambda value: value["artifacts"].update({"diagnostics.json": file_hash(run / "diagnostics.json")})
    )
    plain = records.check_build(run)
    # 最終でない検査は確認記録を照合しないので、記録で決まる件数は None である。
    assert plain["summary"] == {
        "reviews": {"required": 3, "satisfied": None, "carried": None, "renumbered": None},
        "warnings": {"total": 2, "accepted": None, "declared": 1, "unaccepted": None},
        "errors": 0,
        "reviewers": None,
    }
    complete(run, warning=True)
    final = records.check_build(run, final=True)
    # 資料コードで宣言した警告は、許容の記録なしで最終検査を通る。
    assert final["status"] == "passed"
    # 宣言で通した警告は、ページ・種類・要素・理由・宣言の場所とともに並べる。
    assert final["declared_warnings"] == [
        {
            "diagnostic_id": "d_declared",
            "code": "element_overlap",
            "slide_id": "intro",
            "slide_number": 1,
            "element_id": "note",
            "related_element_id": "label",
            "reason": "意図した重なり",
            "location": "slides.py:12",
        }
    ]
    assert final["summary"] == {
        "reviews": {"required": 3, "satisfied": 3, "carried": 0, "renumbered": 0},
        "warnings": {"total": 2, "accepted": 1, "declared": 1, "unaccepted": 0},
        "errors": 0,
        "reviewers": ["確認者"],
    }


def test_draft_and_partial_runs_cannot_be_reviewed_or_exported(tmp_path):
    full = fixture_run(tmp_path, warning=True)
    complete(full, warning=True)
    root = full.parent.parent.parent
    for name, changes in (("0002", {"mode": "draft"}), ("0003", {"selection": SELECTION})):
        run = fixture_run(tmp_path, name, warning=True, **changes)
        assert records._manifest(run)["exportable"] is False
        operations = (
            lambda: records.record_review(run, "content", "all", "確認者", "内容を確認"),
            lambda: records.accept_warning(run, "d_overlap", "確認者", "意図した重なり"),
            lambda: records.carry_reviews(full, run),
            lambda: records.carry_reviews(run, full),
        )
        for operation in operations:
            with pytest.raises(records.RecordError, match="下書き・部分生成の実行は、確認記録の対象外である") as caught:
                operation()
            assert caught.value.code == "build_not_exportable"
        assert records._review(run)["records"] == []
        # 最終でない検査は通常どおり行い、最終検査と反映は完成物にできない実行として止める。
        assert records.check_build(run)["status"] == "passed"
        final = records.check_build(run, final=True)
        assert final["exit_code"] == 2 and "build_not_exportable" in {issue["code"] for issue in final["issues"]}
        assert records.export_build(run)["status"] == "failed"
    assert not (root / "presentation.pptx").exists()
    # 完成物にできない形の生成記録（下書きなのに exportable が真）は読まない。
    rewrite_manifest(full, lambda value: value.update(mode="draft"))
    with pytest.raises(records.RecordError, match="exportable が不正"):
        records.check_build(full)


def test_unfinished_run_is_explained_by_marker(tmp_path):
    project = SimpleNamespace(path=tmp_path / "deck.toml", output_dir=tmp_path / "out")
    running = runs.new_run(project, command="build", options={})
    with pytest.raises(records.RecordError, match=r"生成中の実行である（PID \d+、段階 construct）"):
        records.check_build(running)
    interrupted = runs.new_run(project, command="build", options={})
    runs.update_marker(interrupted, stage="pptx")
    marker = json.loads((interrupted / ".building").read_text(encoding="utf-8"))
    save(interrupted / ".building", {**marker, "pid": 999999999})
    with pytest.raises(records.RecordError, match=r"生成が中断された（PID 999999999、段階 pptx）"):
        records.check_build(interrupted)
    empty = tmp_path / "out" / "runs" / "0009"
    empty.mkdir()
    with pytest.raises(records.RecordError, match="生成記録がない"):
        records.export_build(empty)


def test_resolve_run_accepts_paths_numbers_and_latest(tmp_path, monkeypatch):
    first = fixture_run(tmp_path)
    draft = fixture_run(tmp_path, "0002", mode="draft")
    root = first.parent.parent.parent
    assert records.resolve_run(first) == records.resolve_run(str(first)) == first
    for value in ("1", "0001", 1):
        assert records.resolve_run(value, project=root / "deck.toml") == first
    assert records.resolve_run("2", project=root) == draft
    # latest は完成物にできる最新の実行で、下書きは選ばない。確認の記録には使うことができない。
    assert records.resolve_run("latest", project=root) == first
    with pytest.raises(records.RecordError, match="latest は使えない"):
        records.resolve_run("latest", project=root, allow_latest=False)
    assert records.resolve_run("1", project=root, allow_latest=False) == first
    for value in ("7", tmp_path / "missing", "run-1"):
        with pytest.raises(records.RecordError, match="指定した実行がない"):
            records.resolve_run(value, project=root)
    monkeypatch.chdir(root)
    assert records.resolve_run("latest") == first
    monkeypatch.chdir(tmp_path)
    with pytest.raises(records.RecordError, match="--project で資料を指定する"):
        records.resolve_run("1")
    other = tmp_path / "other"
    other.mkdir()
    (other / "deck.toml").write_text(DECK, encoding="utf-8")
    with pytest.raises(records.RecordError, match="完成物にできる実行がない"):
        records.resolve_run("latest", project=other)


def test_invalid_record_target_is_rejected_and_revocation_preserves_history(tmp_path):
    run = fixture_run(tmp_path)
    with pytest.raises(records.RecordError, match="未知のページ"):
        records.record_review(run, "content", ["missing"], "確認者", "内容を確認")
    with pytest.raises(records.RecordError, match="警告"):
        records.accept_warning(run, "missing", "確認者", "許容する")
    original = records.record_review(run, "content", "all", "確認者", "内容を確認")["records"][0]
    (run.parent.parent.parent / "slides.py").write_text("変更", encoding="utf-8")
    records.revoke_review(run, original["id"], "確認者", "再確認する")
    review = records._review(run)
    assert len(review["records"]) == 2
    assert review["records"][0] == original


# 引き継ぎ


def test_carry_matches_evidence_but_never_overview(tmp_path):
    old = fixture_run(tmp_path, warning=True)
    complete(old, warning=True)
    new = fixture_run(tmp_path, "0002", warning=True, payload="different-package-bytes")
    result = records.carry_reviews(old, new)
    assert {record["kind"] for record in result["records"]} == {"visual", "content", "accept"}
    assert [item["reason"] for item in result["skipped"]] == ["一覧画像の確認は実行ごとに必要である"]
    # 写しは人が記録した元の記録を指す。
    assert {(record["source"]["run"], record["via"]) for record in result["records"]} == {(str(old), str(old))}
    checked = records.check_build(new, final=True)
    assert [(issue["code"], issue["target"]["kind"]) for issue in checked["issues"]] == [("review_missing", "overview")]
    records.record_review(new, "overview", "all", "確認者", "ページ順を確認")
    assert records.check_build(new, final=True)["status"] == "passed"
    assert records.carry_reviews(old, new)["records"] == []


def test_note_change_invalidates_content_but_not_identical_visual(tmp_path):
    old = fixture_run(tmp_path)
    complete(old)
    source = old.parent.parent.parent / "slides.py"
    source.write_text("ノートだけを変更", encoding="utf-8")
    new = fixture_run(tmp_path, "0002", content="changed-notes")
    result = records.carry_reviews(old, new)
    assert [record["kind"] for record in result["records"]] == ["visual"]
    assert any("証拠" in item["reason"] for item in result["skipped"])


def test_font_change_invalidates_visual_carry(tmp_path):
    old = fixture_run(tmp_path)
    complete(old)
    new = fixture_run(tmp_path, "0002", environment="different-font-environment")
    result = records.carry_reviews(old, new)
    assert [record["kind"] for record in result["records"]] == ["content"]


def test_visual_carry_uses_parts_not_pixels(tmp_path):
    old = fixture_run(tmp_path)
    complete(old)
    new = fixture_run(tmp_path, "0002", shift=1)
    before, after = visual_of(old), visual_of(new)
    assert before["evidence_hash"] == after["evidence_hash"] and before["pixel_hash"] != after["pixel_hash"]
    check = carried_visual(records.carry_reviews(old, new))["render_check"]
    assert check == {
        "method": RENDER_CHECK_METHOD,
        "source_pixel_hash": before["pixel_hash"],
        "target_pixel_hash": after["pixel_hash"],
        "strong_pixels": check["strong_pixels"],
        "masked": [],
    }
    assert 0 <= check["strong_pixels"] <= NOISE_PIXELS
    records.record_review(new, "overview", "all", "確認者", "全体確認")
    assert records.check_build(new, final=True)["status"] == "passed"


def test_visual_carry_rejects_large_render_difference(tmp_path):
    old = fixture_run(tmp_path)
    complete(old)
    # 部品の証拠は同じでも、図形が描かれなかった表示は引き継がない。
    new = fixture_run(tmp_path, "0002", shape=False)
    assert visual_of(old)["evidence_hash"] == visual_of(new)["evidence_hash"]
    result = records.carry_reviews(old, new)
    assert [record["kind"] for record in result["records"]] == ["content"]
    reasons = [item["reason"] for item in result["skipped"]]
    assert any(
        reason.startswith("表示が引き継ぎ元と大きく異なる（1 画素のずれを超える差が ") for reason in reasons
    ), reasons


def test_renumbered_page_carries_with_footer_mask(tmp_path):
    old = fixture_run(tmp_path)
    complete(old)
    # 先頭にページを加えて番号が 1 から 2 になった（桁数が同じなので部品の証拠は同じ）。
    new = fixture_run(tmp_path, "0002", number=2)
    before, after = visual_of(old), visual_of(new)
    assert before["evidence_hash"] == after["evidence_hash"]
    # 番号の領域を除かなければ、番号の違いを大きな差とみなす。
    unmasked = render_difference(old / before["path"], new / after["path"])["strong_pixels"]
    assert unmasked > NOISE_PIXELS
    result = records.carry_reviews(old, new)
    check = carried_visual(result)["render_check"]
    assert (check["masked"], check["strong_pixels"]) == ([FOOTER_AREA], 0)
    assert result["renumbered"] == ["intro"]
    records.record_review(new, "overview", "all", "確認者", "全体確認")
    checked = records.check_build(new, final=True)
    assert checked["status"] == "passed"
    assert checked["summary"]["reviews"] == {"required": 3, "satisfied": 3, "carried": 2, "renumbered": 1}
    # 自動の引き継ぎも、元の記録の実行の画像と、番号の領域を除いて比べる。
    third = fixture_run(tmp_path, "0003", number=2)
    auto = records.carry_reviews_auto(third)
    assert auto["renumbered"] == ["intro"]
    assert (carried_visual(auto)["render_check"]["masked"], carried_visual(auto)["source"]["run"]) == (
        [FOOTER_AREA],
        str(old),
    )


def test_renumbered_formula_page_is_listed_once(tmp_path):
    # 数式のページは標準表示と互換表示の両方の確認を番号の領域を除いて引き継ぐが、renumbered にはページを 1 回だけ入れる。
    old = fixture_run(tmp_path, fallback=("intro",))
    complete(old)
    records.record_review(old, "visual", "all", "確認者", "互換表示を確認", variant="fallback")
    for name, carry in (("0002", lambda run: records.carry_reviews(old, run)), ("0003", records.carry_reviews_auto)):
        result = carry(fixture_run(tmp_path, name, number=2, fallback=("intro",)))
        visuals = [record for record in result["records"] if record["kind"] == "visual"]
        assert sorted(record["variant"] for record in visuals) == ["fallback", "standard"]
        assert all(record["render_check"]["masked"] == [FOOTER_AREA] for record in visuals)
        assert result["renumbered"] == ["intro"]


@pytest.mark.parametrize(
    "change, reason",
    [
        ("removed", "表示の照合の記録がない"),
        ({"method": RENDER_CHECK_METHOD + 1}, "表示の照合の方法が現在と異なる"),
        ({"source_pixel_hash": "0" * 64}, "表示の照合の記録が画像と一致しない"),
        ({"target_pixel_hash": "0" * 64}, "表示の照合の記録が画像と一致しない"),
        ({"strong_pixels": NOISE_PIXELS + 1}, "表示が引き継ぎ元と大きく異なる"),
    ],
)
def test_tampered_render_check_is_rejected(tmp_path, change, reason):
    old = fixture_run(tmp_path)
    complete(old)
    new = fixture_run(tmp_path, "0002", shift=1)
    records.carry_reviews(old, new)
    records.record_review(new, "overview", "all", "確認者", "全体確認")
    review = records._review(new)
    copy = next(item for item in review["records"] if item["kind"] == "visual")
    if change == "removed":
        del copy["render_check"]
    else:
        copy["render_check"].update(change)
    save(new / "review.json", review)
    result = records.check_build(new, final=True)
    assert result["invalid_records"] == [{"record_id": copy["id"], "reason": reason}]
    assert [(issue["code"], issue["target"]["kind"]) for issue in result["issues"]] == [("review_missing", "visual")]


def test_carried_record_points_to_original_and_middle_run_can_be_cleaned(tmp_path):
    first = fixture_run(tmp_path, warning=True)
    original = complete(first, warning=True)
    second = fixture_run(tmp_path, "0002", warning=True)
    records.carry_reviews(first, second)
    third = fixture_run(tmp_path, "0003", warning=True)
    carried = records.carry_reviews(second, third)["records"]
    # 写しの写しを作らず、人が記録した元の記録（first）を指す。記録を見つけた実行は via に残る。
    assert {(record["source"]["run"], record["via"]) for record in carried} == {(str(first), str(second))}
    assert {record["source"]["record_id"] for record in carried} == {
        original["visual"]["records"][0]["id"],
        original["content"]["records"][0]["id"],
        original["warning"]["record"]["id"],
    }
    records.record_review(third, "overview", "all", "確認者", "全体確認")
    # 途中の実行は引き継ぎ元として保護されないので、整理で消すことができる。消しても third は最終検査に通る。
    removed = runs.clean_runs(project_for(third), before="0003", apply=True)["removed"]
    assert removed == [str(second)] and first.exists()
    assert records.check_build(third, final=True)["status"] == "passed"


def test_render_check_compares_with_original_run_image(tmp_path):
    first = fixture_run(tmp_path)
    complete(first)
    # 途中の実行の画像だけが 1 画素ずれている。
    second = fixture_run(tmp_path, "0002", shift=1)
    records.carry_reviews(first, second)
    third = fixture_run(tmp_path, "0003")
    assert visual_of(second)["pixel_hash"] != visual_of(first)["pixel_hash"] == visual_of(third)["pixel_hash"]
    # 写しの照合は記録を見つけた実行（second）でなく、元の記録の実行（first）の画像と比べる。
    check = carried_visual(records.carry_reviews(second, third))["render_check"]
    assert check == {
        "method": RENDER_CHECK_METHOD,
        "source_pixel_hash": visual_of(first)["pixel_hash"],
        "target_pixel_hash": visual_of(third)["pixel_hash"],
        "strong_pixels": 0,
        "masked": [],
    }
    records.record_review(third, "overview", "all", "確認者", "全体確認")
    assert records.check_build(third, final=True)["status"] == "passed"


def test_reverse_carry_skips_records_whose_original_is_in_target(tmp_path):
    first = fixture_run(tmp_path, warning=True)
    complete(first, warning=True)
    second = fixture_run(tmp_path, "0002", warning=True)
    records.carry_reviews(first, second)
    records.record_review(second, "overview", "all", "確認者", "全体確認")
    before = records._review(first)["records"]
    # second の写しの元の記録は first にある。逆向きに引き継いでも、first 自身を指す写しを作らない。
    result = records.carry_reviews(second, first)
    assert result["records"] == []
    assert Counter(item["reason"] for item in result["skipped"]) == {
        "元の記録が引き継ぎ先にある": 3,
        "一覧画像の確認は実行ごとに必要である": 1,
    }
    assert records._review(first)["records"] == before
    checked = records.check_build(first, final=True)
    assert (checked["status"], checked["invalid_records"]) == ("passed", [])
    # first は自分を引き継ぎ元として参照しないので、second を整理で消せば、first は引き継ぎ元として保護されない。
    assert records.run_references(first)["sources"] == []
    assert runs.clean_runs(project_for(first), before="0003", apply=True)["removed"] == [str(second)]
    assert first.resolve() not in runs.protected_runs(project_for(first), first.parent, first.parent.parent)


def test_copy_must_point_to_human_record(tmp_path):
    first = fixture_run(tmp_path)
    complete(first)
    second = fixture_run(tmp_path, "0002")
    copy = next(item for item in records.carry_reviews(first, second)["records"] if item["kind"] == "content")
    third = fixture_run(tmp_path, "0003")
    records.carry_reviews(first, third)
    # 写しを元の記録として指す写しは無効にする（取り消しの波及と整理の保護が元の記録を前提にするため）。
    review = records._review(third)
    forged = {
        **copy,
        "id": "r_forged",
        "binding": review["records"][-1]["binding"],
        "source": {"run": str(second), "record_id": copy["id"], "record_hash": json_digest(copy)},
    }
    review["records"].append(forged)
    save(third / "review.json", review)
    result = records.check_build(third, final=True)
    assert result["invalid_records"] == [{"record_id": "r_forged", "reason": "引き継ぎ元が人の記録した元の記録でない"}]


def test_carry_rejects_different_evidence_version(tmp_path):
    old = fixture_run(tmp_path)
    complete(old)
    new = fixture_run(tmp_path, "0002", evidence_version=EVIDENCE_VERSION + 1)
    with pytest.raises(records.RecordError, match="証拠の形式が異なる実行の間では引き継げない") as caught:
        records.carry_reviews(old, new)
    assert f"{old} は形式 {EVIDENCE_VERSION}、{new} は形式 {EVIDENCE_VERSION + 1}" in str(caught.value)
    assert records._review(new)["records"] == []


def test_manual_carry_requires_same_project_and_runs_directory(tmp_path):
    old = fixture_run(tmp_path)
    complete(old)
    with pytest.raises(records.RecordError, match="同じ実行へ確認を引き継ぐことはできない"):
        records.carry_reviews(old, old)
    # 同じ実行保存先にある別の資料の実行と、同じ資料の実行を別の実行保存先へ複製したもの。
    other = fixture_run(tmp_path, "0002")
    rewrite_manifest(other, lambda value: value.update(project=str(tmp_path / "other" / "deck.toml")))
    elsewhere = tmp_path / "elsewhere" / "out" / "runs" / "0003"
    shutil.copytree(fixture_run(tmp_path, "0003"), elsewhere)
    for target in (other, elsewhere):
        with pytest.raises(records.RecordError, match="同じ資料の同じ実行保存先に置く必要がある"):
            records.carry_reviews(old, target)
        assert records._review(target)["records"] == []


def test_warning_same_id_with_worse_measurement_requires_new_acceptance(tmp_path):
    old = fixture_run(tmp_path, warning=True)
    complete(old, warning=True)
    new = fixture_run(tmp_path, "0002", warning=True)
    diagnostics = records._read_json(new / "diagnostics.json")
    diagnostics["issues"][0]["observations"]["area"] = 120
    diagnostics["issues"][0]["evidence_hash"] = "worse-overlap"
    save(new / "diagnostics.json", diagnostics)
    rewrite_manifest(
        new, lambda value: value["artifacts"].update({"diagnostics.json": file_hash(new / "diagnostics.json")})
    )
    result = records.carry_reviews(old, new)
    assert "accept" not in {item["kind"] for item in result["records"]}


def test_revocation_invalidates_multihop_carry_even_after_input_changed(tmp_path):
    first = fixture_run(tmp_path, warning=True)
    original = complete(first, warning=True)
    second = fixture_run(tmp_path, "0002", warning=True)
    records.carry_reviews(first, second)
    records.record_review(second, "overview", "all", "確認者", "全体確認")
    third = fixture_run(tmp_path, "0003", warning=True)
    carried = records.carry_reviews(second, third)["records"]
    # 引き継ぎを重ねても、写しは人が記録した元の記録を指す。
    assert {record["source"]["run"] for record in carried} == {str(first)}
    records.record_review(third, "overview", "all", "確認者", "全体確認")
    assert records.check_build(third, final=True)["status"] == "passed"
    # 入力が変わった後でも、取り消しは受け付ける。
    (first.parent.parent.parent / "slides.py").write_text("変更", encoding="utf-8")
    records.revoke_review(first, original["visual"]["records"][0]["id"], "確認者", "確認を見直す")
    result = records.check_build(third, final=True)
    assert result["status"] == "failed"
    assert any(item["code"] == "review_missing" and item["target"]["kind"] == "visual" for item in result["issues"])
    assert any(item["reason"] == "引き継ぎ元: 確認記録が取り消されている" for item in result["invalid_records"])


def test_revoking_warning_acceptance_invalidates_carry(tmp_path):
    old = fixture_run(tmp_path, warning=True)
    original = complete(old, warning=True)
    new = fixture_run(tmp_path, "0002", warning=True)
    records.carry_reviews(old, new)
    records.record_review(new, "overview", "all", "確認者", "全体確認")
    records.revoke_review(old, original["warning"]["record"]["id"], "確認者", "重なりを修正する")
    result = records.check_build(new, final=True)
    assert any(item["code"] == "warning_unaccepted" for item in result["issues"])


def test_carry_rejects_corrupted_source_and_reports_changed_target_inputs(tmp_path):
    old = fixture_run(tmp_path)
    complete(old)
    new = fixture_run(tmp_path, "0002")
    (old / "deck.pptx").write_text("破損", encoding="utf-8")
    with pytest.raises(records.RecordError, match="成果物"):
        records.carry_reviews(old, new)
    (old / "deck.pptx").write_text("deck.pptx:same", encoding="utf-8")
    # 引き継ぎ先の入力の変化では拒否せず、その実行を反映できないことを知らせる。
    source = old.parent.parent.parent / "slides.py"
    source.write_text("変更", encoding="utf-8")
    result = records.carry_reviews(old, new)
    assert {record["kind"] for record in result["records"]} == {"visual", "content"}
    assert result["notices"] == [{"code": "input_changed", "message": INPUT_NOTICE, "paths": [str(source)], "count": 1}]


@pytest.mark.parametrize("change", ["edited", "removed"])
def test_carried_record_requires_unchanged_source(tmp_path, change):
    old = fixture_run(tmp_path)
    complete(old)
    new = fixture_run(tmp_path, "0002")
    records.carry_reviews(old, new)
    records.record_review(new, "overview", "all", "確認者", "全体確認")
    assert records.check_build(new, final=True)["status"] == "passed"
    review = records._review(old)
    visual = next(item for item in review["records"] if item["kind"] == "visual")
    if change == "edited":
        visual["note"] = "後から書き換えた確認内容"
    else:
        review["records"].remove(visual)
    save(old / "review.json", review)
    result = records.check_build(new, final=True)
    assert result["exit_code"] == 1
    assert [item["reason"] for item in result["invalid_records"]] == ["引き継ぎ元の記録が欠落または変更されている"]
    assert [(issue["code"], issue["target"]["kind"]) for issue in result["issues"]] == [("review_missing", "visual")]


def forge_copy(target, source, kind, **changes):
    """引き継ぎ元の記録を正しく指す写しを、確認記録へ直接書き加える（改ざんの再現）。"""
    original = next(item for item in records._review(source)["records"] if item["kind"] == kind)
    review = records._review(target)
    # 対象の実行で正規に作った記録の binding を写す。
    copy = {
        **original,
        **changes,
        "id": "r_forged",
        "binding": review["records"][-1]["binding"],
        "source": {"run": str(source), "record_id": original["id"], "record_hash": json_digest(original)},
    }
    review["records"].append(copy)
    save(target / "review.json", review)
    return copy


def test_forged_copy_of_overview_review_is_rejected(tmp_path):
    old = fixture_run(tmp_path)
    complete(old)
    new = fixture_run(tmp_path, "0002")
    records.carry_reviews(old, new)
    forge_copy(new, old, "overview")
    result = records.check_build(new, final=True)
    assert result["invalid_records"] == [{"record_id": "r_forged", "reason": "一覧画像の確認は引き継げない"}]
    assert [(issue["code"], issue["target"]["kind"]) for issue in result["issues"]] == [("review_missing", "overview")]


def test_forged_copy_with_new_evidence_is_rejected(tmp_path):
    old = fixture_run(tmp_path)
    complete(old)
    new = fixture_run(tmp_path, "0002", parts="changed")
    records.carry_reviews(old, new)
    records.record_review(new, "overview", "all", "確認者", "全体確認")
    # 表示が変わった実行の証拠に書き換えた写しは、元の確認と対象が一致しない。
    forge_copy(new, old, "visual", evidence_hash=visual_of(new)["evidence_hash"])
    result = records.check_build(new, final=True)
    assert result["invalid_records"] == [{"record_id": "r_forged", "reason": "引き継ぎ元と確認対象が一致しない"}]
    assert [(issue["code"], issue["target"]["kind"]) for issue in result["issues"]] == [("review_missing", "visual")]


@pytest.mark.parametrize(
    "case, reason",
    [
        ("same_run", "引き継ぎ元が同じ実行である"),
        ("other_runs_directory", "引き継ぎ元が同じ実行保存先にない"),
        ("evidence_version", "引き継ぎ元と証拠の形式が異なる"),
        ("draft", "引き継ぎ元: 下書き・部分生成の実行の記録である"),
    ],
)
def test_forged_copy_source_run_is_rejected(tmp_path, case, reason):
    """写しの引き継ぎ元は、同じ実行保存先にある別の、証拠の形式が同じ、完成物にできる実行に限る。"""
    run = fixture_run(tmp_path)
    records.record_review(run, "overview", "all", "確認者", "全体確認")
    if case == "same_run":
        source = run
        records.record_review(run, "content", "all", "確認者", "内容を確認")
    elif case == "other_runs_directory":
        # 内容が同じで、同じ名前の実行が同じ実行保存先にあっても、別の場所へ複製した実行にだけある記録は指すことができない。
        original = fixture_run(tmp_path, "0002")
        source = tmp_path / "elsewhere" / "out" / "runs" / "0002"
        shutil.copytree(original, source)
        complete(source)
    elif case == "evidence_version":
        source = fixture_run(tmp_path, "0002", evidence_version=EVIDENCE_VERSION + 1)
        complete(source)
    else:
        # 下書きには記録できないので、正しく結び付いた記録を直接書く。
        source = fixture_run(tmp_path, "0002", mode="draft")
        manifest = records._manifest(source)
        target = next(item for item in manifest["required_reviews"] if item["kind"] == "content")
        save(
            source / "review.json",
            {
                "schema_version": records.SCHEMA_VERSION,
                "records": [
                    {
                        "id": "r_draft",
                        "kind": "content",
                        "slide_id": "intro",
                        "renderer": None,
                        "variant": None,
                        "evidence_hash": target["evidence_hash"],
                        "reviewer": "確認者",
                        "note": "内容を確認",
                        "created_at": "2026-10-12T00:00:00+00:00",
                        "binding": records._binding(source, manifest),
                    }
                ],
            },
        )
    forge_copy(run, source, "content")
    result = records.check_build(run, final=True)
    assert result["invalid_records"] == [{"record_id": "r_forged", "reason": reason}]


# 記録の形と改ざん


@pytest.mark.parametrize("name", ["manifest.json", "review.json", "diagnostics.json"])
def test_unknown_schema_cannot_pass(tmp_path, name):
    run = fixture_run(tmp_path)
    value = json.loads((run / name).read_text(encoding="utf-8"))
    value["schema_version"] = 20
    save(run / name, value)
    with pytest.raises(records.RecordError, match="対応していない記録形式") as caught:
        records.check_build(run, final=True)
    assert not isinstance(caught.value, records.LegacyRecordError)


def test_artifact_cannot_escape_run(tmp_path):
    run = fixture_run(tmp_path)
    rewrite_manifest(run, lambda value: value["artifacts"].update({"../outside.txt": "hash"}))
    with pytest.raises(records.RecordError, match="外を指して"):
        records.check_build(run)


def test_manifest_without_core_artifact_cannot_pass(tmp_path):
    run = fixture_run(tmp_path)
    rewrite_manifest(run, lambda value: value["artifacts"].pop("deck.pptx"))
    result = records.check_build(run)
    assert result["exit_code"] == 2
    assert [(issue["code"], issue["path"]) for issue in result["issues"]] == [("artifact_unrecorded", "deck.pptx")]
    with pytest.raises(records.RecordError, match="必須成果物") as caught:
        records.record_review(run, "content", "all", "確認者", "内容を確認")
    assert caught.value.code == "artifact_unrecorded"
    # 完了していない生成にも記録できない（整合検査と同じ build_incomplete）。
    failed = fixture_run(tmp_path, "0002")
    rewrite_manifest(failed, lambda value: value.update(status="failed", stage="pptx"))
    with pytest.raises(records.RecordError, match="生成が完了していない") as caught:
        records.record_review(failed, "content", "all", "確認者", "内容を確認")
    assert caught.value.code == "build_incomplete"


def test_rewritten_manifest_invalidates_existing_records(tmp_path):
    run = fixture_run(tmp_path, warning=True)
    complete(run, warning=True)
    root = run.parent.parent.parent
    rewrite_manifest(run, lambda value: value["deliverables"].update({"pptx": str(root / "other.pptx")}))
    result = records.check_build(run, final=True)
    assert result["exit_code"] == 1
    assert [item["reason"] for item in result["invalid_records"]] == ["確認記録が現在の生成記録を対象にしていない"] * 4
    assert sorted(issue["code"] for issue in result["issues"]) == ["review_missing"] * 3 + ["warning_unaccepted"]
    assert records.export_build(run)["status"] == "failed"
    assert not (root / "other.pptx").exists()


@pytest.mark.parametrize(
    "kind, field, reason, code",
    [
        ("visual", "evidence_hash", "確認対象の証拠が一致しない", "review_missing"),
        ("visual", "slide_id", "確認対象の証拠が一致しない", "review_missing"),
        ("accept", "evidence_hash", "許容した警告の証拠が一致しない", "warning_unaccepted"),
        ("accept", "element_id", "許容した警告の証拠が一致しない", "warning_unaccepted"),
    ],
)
def test_tampered_review_record_is_rejected(tmp_path, kind, field, reason, code):
    run = fixture_run(tmp_path, warning=True)
    complete(run, warning=True)
    review = records._review(run)
    record = next(item for item in review["records"] if item["kind"] == kind)
    record[field] = "書き換えた値"
    save(run / "review.json", review)
    result = records.check_build(run, final=True)
    assert result["exit_code"] == 1
    assert result["invalid_records"] == [{"record_id": record["id"], "reason": reason}]
    assert [issue["code"] for issue in result["issues"]] == [code]
    assert records.export_build(run)["status"] == "failed"


# 反映


@pytest.mark.parametrize("collision", ["input", "artifact", "sibling_run", "duplicate", "asset", "hardlink"])
def test_export_rejects_destination_collisions(tmp_path, collision):
    run = fixture_run(tmp_path)
    root = run.parent.parent.parent
    manifest = records._manifest(run)
    if collision == "input":
        destination = root / "slides.py"
    elif collision == "artifact":
        destination = run / "deck.pptx"
    elif collision == "sibling_run":
        destination = run.parent / "0002" / "deck.pptx"
    elif collision == "duplicate":
        destination = Path(manifest["deliverables"]["pdf"])
    elif collision == "hardlink":
        destination = root / "input-link.pptx"
        destination.hardlink_to(root / "slides.py")
    else:
        destination = root / "assets" / "original.pptx"
    rewrite_manifest(run, lambda value: value["deliverables"].update({"pptx": str(destination)}))
    complete(run)
    with pytest.raises(records.RecordError, match="重複"):
        records.export_build(run)


def test_export_saves_consistent_files_and_old_version(tmp_path):
    first = fixture_run(tmp_path, payload="first")
    complete(first)
    assert records.export_build(first)["status"] == "exported"
    second = fixture_run(tmp_path, "0002", payload="second")
    complete(second)
    second_result = records.export_build(second)
    root = second.parent.parent.parent
    adoption = records.read_record(Path(second_result["record"]))
    assert adoption["run"] == "out/runs/0002"
    assert adoption["review_hash"] == file_hash(second / "review.json")
    for path, digest in adoption["artifacts"].items():
        assert file_hash(root / path) == digest
    history = Path(second_result["history"])
    assert (history / "previous-pptx").read_bytes() == (first / "deck.pptx").read_bytes()
    assert records.read_record(history / "transaction.json")["previous_run"] == "out/runs/0001"


def test_export_keeps_final_check_in_history_and_adoption(tmp_path):
    old = fixture_run(tmp_path, warning=True)
    complete(old, warning=True)
    new = fixture_run(tmp_path, "0002", warning=True)
    records.carry_reviews(old, new)
    records.record_review(new, "overview", "all", "別の確認者", "一覧を確認")
    result = records.export_build(new)
    assert result["final_check"] == str(new / "final_check.json")
    history = Path(result["history"])
    copied = history / "final_check.json"
    assert copied.read_bytes() == (new / "final_check.json").read_bytes()
    final = records._read_json(copied)
    assert (final["final"], final["status"]) == (True, "passed")
    adoption = records.read_record(Path(result["record"]))
    assert adoption["final_check"]["sha256"] == file_hash(copied)
    assert adoption["final_check"]["checked_at"] == final["checked_at"]
    assert (
        adoption["final_check"]["summary"]
        == final["summary"]
        == {
            "reviews": {"required": 3, "satisfied": 3, "carried": 2, "renumbered": 0},
            "warnings": {"total": 1, "accepted": 1, "declared": 0, "unaccepted": 0},
            "errors": 0,
            "reviewers": ["別の確認者", "確認者"],
        }
    )


def test_adoption_record_uses_project_relative_paths(tmp_path):
    first = fixture_run(tmp_path, payload="first")
    complete(first)
    result = records.export_build(first)
    root = first.parent.parent.parent
    record = Path(result["record"])
    # 公開する完成物に、資料の場所の絶対パスを書かない。
    assert str(tmp_path) not in record.read_text(encoding="utf-8")
    adoption = records.read_record(record)
    assert (adoption["run"], adoption["project"]) == ("out/runs/0001", "deck.toml")
    assert set(adoption["artifacts"]) == {"presentation.pptx", "presentation.pdf", "speaker_notes.md"}
    for name, digest in adoption["artifacts"].items():
        assert file_hash(root / name) == digest
    # 相対パスの採用記録も次の反映と整理の保護で資料のディレクトリから読む。
    second = fixture_run(tmp_path, "0002", payload="second")
    complete(second)
    history = Path(records.export_build(second)["history"])
    transaction = records.read_record(history / "transaction.json")
    assert (transaction["run"], transaction["previous_run"]) == ("out/runs/0002", "out/runs/0001")
    protected = runs.protected_runs(project_for(second), second.parent, second.parent.parent)
    assert protected[second.resolve()][0] == "現在採用している実行"


@pytest.mark.skipif(os.name == "nt", reason="POSIX の権限を確かめる")
def test_exported_files_follow_umask(tmp_path):
    previous = os.umask(0o022)
    try:
        first = fixture_run(tmp_path, payload="first")
        complete(first)
        records.export_build(first)
        second = fixture_run(tmp_path, "0002", payload="second")
        complete(second)
        result = records.export_build(second)
    finally:
        os.umask(previous)
    # 反映した完成物と退避した旧版の権限は、一時ファイルの 600 でなく umask にしたがう。
    paths = [Path(path) for path in result["deliverables"].values()] + sorted(
        Path(result["history"]).glob("previous-*")
    )
    assert len(paths) == 8
    assert {path.name: stat.S_IMODE(path.stat().st_mode) for path in paths} == {path.name: 0o644 for path in paths}


def test_export_rechecks_reviews_before_writing_adoption(tmp_path, monkeypatch):
    old = fixture_run(tmp_path)
    original = complete(old)
    new = fixture_run(tmp_path, "0002")
    records.carry_reviews(old, new)
    records.record_review(new, "overview", "all", "確認者", "全体確認")
    unchanged = records._assert_unchanged

    def revoke_elsewhere(run, manifest):
        unchanged(run, manifest)
        # 完成物を複製している間に、引き継ぎ元の実行の確認記録へ取り消しが書かれた場合を再現する。
        review = records._review(old)
        review["records"].append(
            {"id": "r_revokedoutside", "kind": "revoke", "record_id": original["visual"]["records"][0]["id"]}
        )
        save(old / "review.json", review)

    monkeypatch.setattr(records, "_assert_unchanged", revoke_elsewhere)
    with pytest.raises(records.RecordError, match="反映中に入力または確認の状態が変更された"):
        records.export_build(new)
    root = new.parent.parent.parent
    assert not [path for path in root.iterdir() if path.is_file() and path.name not in {"slides.py", "deck.toml"}]
    # 反映の直前の最終検査の結果も final_check.json に残す。
    final = records._read_json(new / "final_check.json")
    assert final["status"] == "failed"
    assert final["invalid_records"][0]["reason"] == "引き継ぎ元: 確認記録が取り消されている"


@pytest.mark.parametrize("place", ["other_output", "outside_runs"])
def test_copied_run_cannot_export_into_original_project(tmp_path, place):
    run = fixture_run(tmp_path)
    complete(run)
    root = run.parent.parent.parent
    # 別の出力先の runs/ か、元の出力先の runs/ でない場所へ複製した実行。どちらの生成記録も元の資料の出力先を指す。
    parent = tmp_path / "elsewhere" / "out" / "runs" if place == "other_output" else root / "out" / "backup"
    copy = parent / run.name
    shutil.copytree(run, copy)
    # 反映の履歴を元の資料の exports/ に書かない。
    with pytest.raises(records.RecordError, match="別の場所へ複製した実行は反映できない"):
        records.export_build(copy)
    assert not (root / "out" / "exports").exists()
    assert not (root / "presentation.pptx").exists() and not (root / "deliverables.json").exists()
    assert records.export_build(run)["status"] == "exported"


def test_export_failure_restores_all_old_files(tmp_path, monkeypatch):
    first = fixture_run(tmp_path, payload="first")
    complete(first)
    previous = records.export_build(first)
    old_files = {name: Path(path).read_bytes() for name, path in previous["deliverables"].items()}
    second = fixture_run(tmp_path, "0002", payload="second")
    complete(second)
    original_replace = records.os.replace
    failed = False

    def fail_pdf(source, destination):
        nonlocal failed
        if Path(destination) == Path(previous["deliverables"]["pdf"]) and not failed:
            failed = True
            raise OSError("容量不足を再現")
        return original_replace(source, destination)

    monkeypatch.setattr(records.os, "replace", fail_pdf)
    with pytest.raises(OSError, match="容量不足"):
        records.export_build(second)
    for name, path in previous["deliverables"].items():
        assert Path(path).read_bytes() == old_files[name]
    assert not (second.parent.parent / "exports" / "pending.json").exists()


def interrupt_second_export(tmp_path, monkeypatch):
    """1 回目を反映した後、2 回目の反映を PDF の置き換えで中断する。PPTX だけが新しい版になる。"""
    first = fixture_run(tmp_path, payload="first")
    complete(first)
    deliverables = records.export_build(first)["deliverables"]
    second = fixture_run(tmp_path, "0002", payload="second")
    complete(second)
    original_replace = records.os.replace

    def interrupt_pdf(source, destination):
        if Path(destination) == Path(deliverables["pdf"]):
            raise KeyboardInterrupt("中断を再現")
        return original_replace(source, destination)

    monkeypatch.setattr(records.os, "replace", interrupt_pdf)
    with pytest.raises(KeyboardInterrupt):
        records.export_build(second)
    monkeypatch.setattr(records.os, "replace", original_replace)
    pending = second.parent.parent / "exports" / "pending.json"
    assert pending.exists()
    assert Path(deliverables["pptx"]).read_bytes() == (second / "deck.pptx").read_bytes()
    return first, second, deliverables, pending


def test_next_export_recovers_interrupted_transaction(tmp_path, monkeypatch):
    first, second, deliverables, pending = interrupt_second_export(tmp_path, monkeypatch)
    interrupted = records.read_record(pending)
    # 最終検査が失敗する場合も、中断した反映の復旧を先に行う。
    (second / "deck.pptx").write_text("破損", encoding="utf-8")
    result = records.export_build(second)
    assert result["status"] == "failed"
    assert not pending.exists()
    assert Path(deliverables["pptx"]).read_bytes() == (first / "deck.pptx").read_bytes()
    assert records.read_record(Path(interrupted["history"]) / "transaction.json")["status"] == "recovered"


@pytest.mark.parametrize("damage", ["changed", "missing"])
def test_recovery_refuses_damaged_backup(tmp_path, monkeypatch, damage):
    _, second, deliverables, pending = interrupt_second_export(tmp_path, monkeypatch)
    transaction = records.read_record(pending)
    backup = Path(next(item["backup"] for item in transaction["files"] if item["destination"] == deliverables["pptx"]))
    if damage == "changed":
        backup.write_text("破損した旧版", encoding="utf-8")
    else:
        backup.unlink()
    state = {name: Path(path).read_bytes() for name, path in deliverables.items()}
    with pytest.raises(records.RecordError, match="復旧用の旧版が破損している"):
        records.export_build(second)
    # 復旧も反映も行わず、手で確かめることができるように未完了の記録を残す。
    assert pending.exists()
    assert {name: Path(path).read_bytes() for name, path in deliverables.items()} == state


def test_recovery_refuses_deliverable_changed_after_interruption(tmp_path, monkeypatch):
    _, second, deliverables, pending = interrupt_second_export(tmp_path, monkeypatch)
    edited = Path(deliverables["pdf"])
    edited.write_text("中断後に編集した内容", encoding="utf-8")
    with pytest.raises(records.RecordError, match="中断後に完成物が変更されている"):
        records.export_build(second)
    assert pending.exists()
    assert edited.read_text(encoding="utf-8") == "中断後に編集した内容"
    # すべての完成物を検査してから戻すので、置き換え済みの PPTX も戻さない。
    assert Path(deliverables["pptx"]).read_bytes() == (second / "deck.pptx").read_bytes()


def test_record_and_export_hash_each_file_once(tmp_path, monkeypatch):
    run = fixture_run(tmp_path, pages=("intro", "method", "summary"))
    manifest = records._manifest(run)
    artifacts = {(run / name).resolve() for name in manifest["artifacts"]}
    inputs = {Path(name).resolve() for name in manifest["input_files"]}
    counts = Counter()
    hash_file = records.file_hash

    def counted(path):
        counts[Path(path).resolve()] += 1
        return hash_file(path)

    monkeypatch.setattr(records, "file_hash", counted)
    assert len(records.record_review(run, "visual", "all", "確認者", "表示を確認")["records"]) == 3
    # 記録の binding に使う生成記録のハッシュ値は、記録の数によらず 1 回だけ求める。
    assert counts[(run / "manifest.json").resolve()] == 1
    assert {path: counts[path] for path in artifacts | inputs} == {path: 1 for path in artifacts | inputs}
    complete(run)
    counts.clear()
    assert records.export_build(run)["status"] == "exported"
    # 反映の直前の最終検査は確認記録だけを確かめ直す。成果物と入力は、最初の検査と複製の後の照合の 2 回だけ読む。
    assert {path: counts[path] for path in artifacts | inputs} == {path: 2 for path in artifacts | inputs}
    assert counts[(run / "manifest.json").resolve()] == 2


# 残りの確認の一覧とまとめての記録


def test_review_status_groups_pending_and_suggests_commands(tmp_path, monkeypatch):
    pages = tuple(f"p{number}" for number in range(1, 8))
    run = fixture_run(tmp_path, pages=pages, fallback=("p5",))
    root = run.parent.parent.parent
    declared = warning(
        "d_declared", "p4", "label", "frame", declared={"reason": "ラベルを枠に重ねて示す", "location": "slides.py:42"}
    )
    set_diagnostics(run, [warning("d_overlap", "p2", "box", "label"), declared])
    monkeypatch.chdir(root)
    records.record_review(run, "visual", "1,2,4", "確認者", "表示を確認")
    records.record_review(run, "content", "all", "確認者", "内容を確認")
    before = records._review(run)["records"]
    status = records.review_status("latest", project=root)
    assert records._review(run)["records"] == before
    assert (status["status"], status["exit_code"], status["build"], status["run_id"]) == (
        "incomplete",
        1,
        str(run),
        "0001",
    )
    groups = [
        (group["kind"], group["renderer"], group["variant"], group["required"], group["satisfied"], group["slides"])
        for group in status["reviews"]
    ]
    assert groups == [
        ("visual", "powerpoint", "standard", 7, 3, "3,5-7"),
        ("visual", "powerpoint", "fallback", 1, 0, "5"),
        ("overview", "powerpoint", "standard", 1, 0, None),
        ("content", None, None, 7, 7, None),
    ]
    visual = status["reviews"][0]
    assert visual["pending"] == [{"slide_id": f"p{number}", "number": number} for number in (3, 5, 6, 7)]
    assert re.fullmatch(r"s_[0-9a-f]{16}", visual["status_id"]) and status["reviews"][2]["status_id"] is None
    assert status["reviews"][2]["pending"] == [{"slide_id": None, "number": None}]
    assert status["warnings"] == {
        "total": 2,
        "accepted": 0,
        "declared": 1,
        "pending": [
            {
                "diagnostic_id": "d_overlap",
                "code": "element_overlap",
                "slide_id": "p2",
                "slide_number": 2,
                "element_id": "box",
                "related_element_id": "label",
                "message": "図形が重なる",
            }
        ],
    }
    # 宣言で通した警告を、ページ・種類・要素・理由・宣言の場所で並べる。
    assert status["declared_warnings"] == [
        {
            "diagnostic_id": "d_declared",
            "code": "element_overlap",
            "slide_id": "p4",
            "slide_number": 4,
            "element_id": "label",
            "related_element_id": "frame",
            "reason": "ラベルを枠に重ねて示す",
            "location": "slides.py:42",
        }
    ]
    placeholders = f'--reviewer "{records.REVIEWER_PLACEHOLDER}" --note "{records.NOTE_PLACEHOLDER}"'
    assert status["next"] == [
        f"slidekit review record --build out/runs/0001 --kind visual --slides 3,5-7 {placeholders}",
        f"slidekit review record --build out/runs/0001 --kind visual --variant fallback --slides 5 {placeholders}",
        f"slidekit review record --build out/runs/0001 --kind overview {placeholders}",
        f'slidekit review accept --build out/runs/0001 --diagnostic d_overlap --reviewer "{records.REVIEWER_PLACEHOLDER}" '
        f'--reason "{records.REASON_PLACEHOLDER}"',
        "意図した例外なら、資料コードで slide.allow_warning() を宣言して再生成する",
    ]
    # 示したコマンドをそのまま渡すと、置き場所のままの確認者・確認内容・理由として拒否する（画像を見ずに記録しない）。
    for command in status["next"][:4]:
        with pytest.raises(records.RecordError, match="置き場所のままである"):
            run_command(command)
    assert records._review(run)["records"] == before
    # 置き場所を書き換えると記録できる。
    for command in status["next"][:3]:
        run_command(command, reviewer="山田", note="画像を開いて文字と図を確認した")
    run_command(status["next"][3], reviewer="山田", reason="ラベルを図形に重ねる意図どおりの配置である")
    done = records.review_status(run)
    assert (done["status"], done["exit_code"], done["next"]) == (
        "complete",
        0,
        ["slidekit export --build out/runs/0001"],
    )
    assert done["summary"] == {
        "reviews": {"required": 16, "satisfied": 16, "carried": 0, "renumbered": 0},
        "warnings": {"total": 2, "accepted": 1, "declared": 1, "unaccepted": 0},
        "errors": 0,
        "reviewers": ["山田", "確認者"],
    }
    # エラーの診断がある実行（終了コード 1）と、成果物が変わった実行（2）は blocked で、再生成を示す。
    broken = fixture_run(tmp_path, "0002", pages=pages)
    set_diagnostics(
        broken,
        [
            {
                **warning("d_error", "p3", "body"),
                "severity": "error",
                "code": "text_overflow",
                "message": "文字が枠からはみ出す",
            }
        ],
    )
    errors = records.review_status("2", project=root)
    assert (errors["status"], errors["exit_code"], errors["next"]) == (
        "blocked",
        1,
        ["slidekit build --project deck.toml"],
    )
    assert errors["errors"] == [
        {
            "diagnostic_id": "d_error",
            "code": "text_overflow",
            "slide_id": "p3",
            "slide_number": 3,
            "element_id": "body",
            "message": "文字が枠からはみ出す",
        }
    ]
    (run / "deck.pdf").write_text("変更", encoding="utf-8")
    changed = records.review_status(run)
    assert (changed["status"], changed["exit_code"], changed["next"]) == (
        "blocked",
        2,
        ["slidekit build --project deck.toml"],
    )
    assert [(item["code"], item["paths"], item["count"]) for item in changed["blocking"]] == [
        ("artifact_changed", ["deck.pdf"], 1)
    ]


def test_placeholder_reviewer_note_and_reason_are_rejected(tmp_path):
    run = fixture_run(tmp_path, warning=True)
    for reviewer, note in (
        (records.REVIEWER_PLACEHOLDER, "表示を確認"),
        ("確認者", records.NOTE_PLACEHOLDER),
        ("<山田>", "表示を確認"),
        ("確認者", " ＜確認した内容＞ "),
    ):
        with pytest.raises(records.RecordError, match="置き場所のままである"):
            records.record_review(run, "visual", "all", reviewer, note)
    for reviewer, reason in ((records.REVIEWER_PLACEHOLDER, "意図した重なり"), ("確認者", records.REASON_PLACEHOLDER)):
        with pytest.raises(records.RecordError, match="置き場所のままである"):
            records.accept_warning(run, "d_overlap", reviewer, reason)
    assert records._review(run)["records"] == []
    # 山括弧を含むだけの値は受け付ける。
    assert len(records.record_review(run, "visual", "all", "確認者", "矢印 <- の向きを確認")["records"]) == 1


def test_record_multiple_kinds_and_all_variants(tmp_path):
    pages = ["intro", "formula", "summary"]
    run = fixture_run(tmp_path, pages=tuple(pages), fallback=("formula",), renderers=("powerpoint", "libreoffice"))
    first = records.record_review(run, "visual,content", "all", "確認者", "表示と内容を確認", variant="all")
    assert re.fullmatch(r"b_[0-9a-f]{16}", first["batch"])
    assert {record["batch"] for record in first["records"]} == {first["batch"]}
    # 主表示ソフトの標準表示と、互換表示が要るページの互換表示だけを記録する。
    assert first["targets"] == {
        "visual:powerpoint:standard": pages,
        "visual:powerpoint:fallback": ["formula"],
        "content": pages,
    }
    second = records.record_review(
        run, ["visual", "overview"], "2", "確認者", "全表示ソフトで確認", renderer="all", variant="all"
    )
    assert second["batch"] != first["batch"]
    assert second["targets"] == {
        "visual:powerpoint:standard": ["formula"],
        "visual:powerpoint:fallback": ["formula"],
        "visual:libreoffice:standard": ["formula"],
        "overview:powerpoint": [None],
        "overview:libreoffice": [None],
    }
    pending = [
        (group["kind"], group["renderer"], group["variant"], group["slides"])
        for group in records.review_status(run)["reviews"]
        if group["pending"]
    ]
    assert pending == [("visual", "libreoffice", "standard", "1,3")]
    records.record_review(run, "visual", "all", "確認者", "LibreOffice の表示を確認", renderer="libreoffice")
    assert records.check_build(run, final=True)["status"] == "passed"
    with pytest.raises(records.RecordError, match="一覧画像の確認は標準表示だけである"):
        records.record_review(run, "overview", None, "確認者", "一覧を確認", variant="fallback")


def test_record_requires_slides_and_rejects_latest(tmp_path):
    run = fixture_run(tmp_path, warning=True)
    root = run.parent.parent.parent
    for kind, slides in (("visual", None), ("visual,content", " "), ("content", [])):
        with pytest.raises(
            records.RecordError, match=re.escape("--slides を指定する（all、番号、範囲、識別子。例: 3,5-7,intro）")
        ):
            records.record_review(run, kind, slides, "確認者", "確認した")
    with pytest.raises(records.RecordError, match="一覧画像の確認には --slides を指定しない"):
        records.record_review(run, "overview", "1", "確認者", "一覧を確認")
    operations = (
        lambda: records.record_review("latest", "content", "all", "確認者", "内容を確認", project=root),
        lambda: records.accept_warnings(
            "latest", reviewer="確認者", reason="意図した重なり", code="element_overlap", project=root
        ),
        lambda: records.revoke_review("latest", "r_0000000000000000", "確認者", "見直す", project=root),
    )
    for operation in operations:
        with pytest.raises(records.RecordError, match="latest は使えない"):
            operation()
    assert records._review(run)["records"] == []
    # 番号は明示なので受け付ける。一覧画像だけなら --slides は要らない。
    numbered = records.record_review("1", "content", "all", "確認者", "内容を確認", project=root)
    assert (numbered["build"], [record["kind"] for record in numbered["records"]]) == (str(run), ["content"])
    assert [
        record["kind"] for record in records.record_review(run, "overview", None, "確認者", "一覧を確認")["records"]
    ] == ["overview"]


def test_fallback_all_selects_only_math_pages(tmp_path):
    run = fixture_run(tmp_path, pages=("intro", "formula", "summary"), fallback=("formula",))
    result = records.record_review(run, "visual", "all", "確認者", "互換表示を確認", variant="fallback")
    assert [(record["slide_id"], record["variant"]) for record in result["records"]] == [("formula", "fallback")]


def test_explicit_page_without_fallback_is_rejected(tmp_path):
    run = fixture_run(tmp_path, pages=("intro", "formula", "summary"), fallback=("formula",))
    with pytest.raises(
        records.RecordError,
        match=re.escape(
            "互換表示の確認が要らないページが含まれている: 1,3（数式がない）。all を使うと必要なページだけを選ぶ"
        ),
    ):
        records.record_review(run, "visual", "1-3", "確認者", "互換表示を確認", variant="fallback")
    # 確認対象にない表示ソフトのページも記録しない。
    with pytest.raises(
        records.RecordError, match=re.escape("表示（libreoffice・標準）の確認が要らないページが含まれている: 2")
    ):
        records.record_review(run, "visual", "2", "確認者", "表示を確認", renderer="libreoffice")
    assert records._review(run)["records"] == []
    # --variant all は明示したページの標準表示と、そのうち互換表示が要るページの互換表示を記録する。
    result = records.record_review(run, "visual", "1,2", "確認者", "表示を確認", variant="all")
    assert [(record["slide_id"], record["variant"]) for record in result["records"]] == [
        ("intro", "standard"),
        ("formula", "standard"),
        ("formula", "fallback"),
    ]


def test_numbers_ranges_and_ids_select_pages(tmp_path):
    run = fixture_run(tmp_path, pages=("a", "b", "c", "d", "e"))
    for spec, expected in (
        ("1,3-4", ["a", "c", "d"]),
        (" e , id:b ", ["b", "e"]),
        ("2-2,b", ["b"]),
        (["c", "a"], ["a", "c"]),
    ):
        result = records.record_review(run, "content", spec, "確認者", "内容を確認")
        assert [record["slide_id"] for record in result["records"]] == expected, spec
    for spec, message in (
        ("9", "ページ番号が範囲外"),
        ("x", "ページの識別子がありません: x"),
        ("3-", "解釈できません"),
        ("all,1", "ほかの項目と組み合わせない"),
        (["3"], "未知のページ"),
    ):
        with pytest.raises(records.RecordError, match=message):
            records.record_review(run, "content", spec, "確認者", "内容を確認")


def test_numeric_slide_id_is_ambiguous(tmp_path):
    run = fixture_run(tmp_path, pages=("intro", "3", "summary"))
    for spec in ("3", "1-3", "03"):
        with pytest.raises(
            records.RecordError, match=re.escape("番号と同じ名前のページがある: 3。識別子なら id:3 と書く")
        ):
            records.record_review(run, "content", spec, "確認者", "内容を確認")
    assert [
        record["slide_id"] for record in records.record_review(run, "content", "id:3", "確認者", "内容")["records"]
    ] == ["3"]
    assert [
        record["slide_id"] for record in records.record_review(run, "content", "1-2", "確認者", "内容")["records"]
    ] == ["intro", "3"]
    # all・pending という識別子のページも同じ規則（id: を付ける）。
    named = fixture_run(tmp_path / "named", pages=("all", "pending"))
    for spec in ("all", "pending"):
        with pytest.raises(
            records.RecordError, match=f"指定の語と同じ名前のページがある: {spec}。そのページは id:{spec} と書き"
        ):
            records.record_review(
                named, "content", spec, "確認者", "内容を確認", status_id="s_0" if spec == "pending" else None
            )
    result = records.record_review(named, "content", "id:all,id:pending", "確認者", "内容を確認")
    assert [record["slide_id"] for record in result["records"]] == ["all", "pending"]


def test_pending_requires_matching_status_id(tmp_path):
    first = fixture_run(tmp_path, pages=("a", "b", "c"))
    complete(first)
    second = fixture_run(tmp_path, "0002", pages=("a", "b", "c"))
    records.carry_reviews_auto(second)
    copy = next(
        record
        for record in records._review(second)["records"]
        if record["kind"] == "visual" and record["slide_id"] == "c"
    )
    records.revoke_review(second, copy["id"], "確認者", "表示を見直す")
    visual = records.review_status(second)["reviews"][0]
    assert (visual["slides"], visual["pending"]) == ("3", [{"slide_id": "c", "number": 3}])
    stale = visual["status_id"]
    # review status の後に、元の記録の取り消しで未確認のページが増えた。見た時点の集まりと違うので記録しない。
    original = next(
        record
        for record in records._review(first)["records"]
        if record["kind"] == "visual" and record["slide_id"] == "b"
    )
    records.revoke_review(first, original["id"], "確認者", "表示を見直す")
    with pytest.raises(
        records.RecordError, match="未確認のページが review status の後に変わった。review status をやり直す"
    ):
        records.record_review(second, "visual", "pending", "確認者", "表示を確認", status_id=stale)
    fresh = records.review_status(second)["reviews"][0]
    assert fresh["slides"] == "2-3" and fresh["status_id"] != stale
    status_id = fresh["status_id"]
    for kind, renderer, variant, value, message in (
        ("visual", None, "standard", None, "pending には review status が示す --status-id が必要である"),
        ("visual,content", None, "standard", status_id, "pending は種類・表示ソフト・表示方法を一つずつ指定する"),
        ("visual", None, "all", status_id, "pending は種類・表示ソフト・表示方法を一つずつ指定する"),
        ("visual", "all", "standard", status_id, "pending は種類・表示ソフト・表示方法を一つずつ指定する"),
        ("overview", None, "standard", status_id, "一覧画像の確認には --slides を指定しない"),
        ("content", None, "standard", status_id, "未確認のページがない"),
    ):
        with pytest.raises(records.RecordError, match=message):
            records.record_review(second, kind, "pending", "確認者", "表示を確認", renderer, variant, status_id=value)
    with pytest.raises(records.RecordError, match="--status-id は --slides pending と組み合わせる"):
        records.record_review(second, "visual", "2-3", "確認者", "表示を確認", status_id=status_id)
    result = records.record_review(second, "visual", "pending", "確認者", "表示を確認", status_id=status_id)
    assert [record["slide_id"] for record in result["records"]] == ["b", "c"]
    with pytest.raises(records.RecordError, match="未確認のページがない"):
        records.record_review(second, "visual", "pending", "確認者", "表示を確認", status_id=status_id)


# 自動の引き継ぎと取り消しの波及


def test_auto_carry_finds_newest_valid_original_across_runs(tmp_path):
    first = fixture_run(tmp_path, warning=True)
    original = complete(first, warning=True)
    # 0002 でページを変えた。警告は同じなので、許容だけを引き継ぐ。
    second = fixture_run(tmp_path, "0002", warning=True, parts="changed", content="changed")
    result = records.carry_reviews_auto(second)
    assert [(record["kind"], record["source"]["run"]) for record in result["records"]] == [("accept", str(first))]
    assert result["remaining"] == {"visual": 1, "overview": 1, "content": 1, "warnings": 0}
    records.record_review(second, "visual,content", "all", "確認者", "変えたページを確認")
    # 0003 で元に戻すと、0002 の記録でなく 0001 の元の記録を引き継ぐ。記録を見つけた実行は via に残る。
    third = fixture_run(tmp_path, "0003", warning=True)
    result = records.carry_reviews_auto(third)
    assert sorted((record["kind"], record["source"]["run"], record["via"]) for record in result["records"]) == [
        ("accept", str(first), str(second)),
        ("content", str(first), str(first)),
        ("visual", str(first), str(first)),
    ]
    assert (result["status"], result["sources"], result["skipped"], result["skipped_runs"]) == (
        "carried",
        {str(first): 3},
        {},
        {},
    )
    assert result["remaining"] == {"visual": 0, "overview": 1, "content": 0, "warnings": 0}
    attempt = records._review(third)["carry_attempts"][-1]
    assert attempt == {
        "created_at": attempt["created_at"],
        "mode": "auto",
        "sources": {str(first): 3},
        "records": [record["id"] for record in result["records"]],
        "skipped": {},
        "skipped_runs": {},
    }
    # 新しい実行に同じ証拠の人の記録があれば、そちらを元にする。取り消すと、その前の有効な元の記録を使う。
    newer = records.record_review(third, "content", "all", "別の確認者", "内容を読み直した")["records"][0]
    fourth = fixture_run(tmp_path, "0004", warning=True)
    content = next(record for record in records.carry_reviews_auto(fourth)["records"] if record["kind"] == "content")
    assert (content["source"]["run"], content["source"]["record_id"]) == (str(third), newer["id"])
    records.revoke_review(third, newer["id"], "別の確認者", "読み直しを取り消す")
    fifth = fixture_run(tmp_path, "0005", warning=True)
    content = next(record for record in records.carry_reviews_auto(fifth)["records"] if record["kind"] == "content")
    assert (content["source"]["run"], content["source"]["record_id"]) == (
        str(first),
        original["content"]["records"][0]["id"],
    )
    records.record_review(fifth, "overview", None, "確認者", "一覧を確認")
    assert records.check_build(fifth, final=True)["status"] == "passed"


def test_auto_carry_counts_skip_reasons(tmp_path):
    old = fixture_run(tmp_path)
    complete(old)
    # 部品の証拠は同じでも、図形が描かれなかった表示は引き継がず、理由ごとの件数だけを残す。
    new = fixture_run(tmp_path, "0002", shape=False)
    result = records.carry_reviews_auto(new)
    assert [record["kind"] for record in result["records"]] == ["content"]
    assert result["skipped"] == {"表示が引き継ぎ元と大きく異なる": 1}
    assert records._review(new)["carry_attempts"][-1]["skipped"] == {"表示が引き継ぎ元と大きく異なる": 1}
    assert result["remaining"] == {"visual": 1, "overview": 1, "content": 0, "warnings": 0}


def test_fallback_media_change_keeps_standard_review(tmp_path):
    """標準表示の確認の証拠は互換画像のメディアの中身を含まない。

    互換画像だけが変わった実行では、標準表示の確認を引き継ぎ、互換表示の確認はやり直す。
    """
    pages = ("intro", "formula")
    first = fixture_run(tmp_path, pages=pages, fallback=("formula",))
    records.record_review(first, "visual,content", "all", "確認者", "全表示を確認", variant="all")
    second = fixture_run(tmp_path, "0002", pages=pages, fallback=("formula",), media="recropped")
    views = [records._manifest(run)["slides"][1]["visuals"] for run in (first, second)]
    assert views[0]["powerpoint:standard"]["evidence_hash"] == views[1]["powerpoint:standard"]["evidence_hash"]
    assert views[0]["powerpoint:fallback"]["evidence_hash"] != views[1]["powerpoint:fallback"]["evidence_hash"]
    result = records.carry_reviews_auto(second)
    assert sorted((record["kind"], record["slide_id"], record["variant"]) for record in result["records"]) == [
        ("content", "formula", None),
        ("content", "intro", None),
        ("visual", "formula", "standard"),
        ("visual", "intro", "standard"),
    ]
    assert result["remaining"] == {"visual": 1, "overview": 1, "content": 0, "warnings": 0}
    fallback = next(group for group in records.review_status(second)["reviews"] if group["variant"] == "fallback")
    assert fallback["slides"] == "2"


def test_auto_carry_skips_legacy_draft_and_partial_runs(tmp_path):
    fixture_run(tmp_path, "0001", schema=1)
    full = fixture_run(tmp_path, "0002")
    complete(full)
    # 下書き・部分生成の実行の記録は、証拠が同じでも元にしない。証拠の形式が違う実行と形式 1 の実行は件数を集計する。
    for name, changes in (("0003", {"mode": "draft"}), ("0004", {"selection": SELECTION})):
        direct_record(fixture_run(tmp_path, name, **changes))
    complete(fixture_run(tmp_path, "0005", evidence_version=EVIDENCE_VERSION + 1))
    target = fixture_run(tmp_path, "0006")
    result = records.carry_reviews_auto(target)
    assert {(record["kind"], record["source"]["run"]) for record in result["records"]} == {
        ("visual", str(full)),
        ("content", str(full)),
    }
    assert result["sources"] == {str(full): 2}
    assert result["skipped_runs"] == {"古い形式（schema_version 1）": 1, "証拠の形式が異なる": 1}
    # 候補にしなかった実行の件数は、確認記録の引き継ぎの記録にも残す。
    assert records._review(target)["carry_attempts"][-1]["skipped_runs"] == result["skipped_runs"]


def test_auto_carry_ignores_other_project_runs(tmp_path):
    """同じ実行保存先にある別の資料の実行は候補にせず、同じ資料の実行の写しが指していても元にしない。"""
    same = fixture_run(tmp_path, warning=True)
    complete(same)
    # 0002 は別の資料の実行で、証拠は同じ。警告も許容している（同じ資料の実行は許容していない）。
    other = fixture_run(tmp_path, "0002", warning=True)
    rewrite_manifest(other, lambda value: value.update(project=str(tmp_path / "other" / "deck.toml")))
    complete(other, warning=True)
    # 0003 は同じ資料の実行で、別の資料の実行の記録を指す写しを持つ（改ざんの再現）。
    forged = fixture_run(tmp_path, "0003", warning=True)
    records.record_review(forged, "overview", None, "確認者", "一覧を確認")
    forge_copy(forged, other, "visual")
    target = fixture_run(tmp_path, "0004", warning=True)
    result = records.carry_reviews_auto(target)
    assert {(record["kind"], record["source"]["run"]) for record in result["records"]} == {
        ("visual", str(same)),
        ("content", str(same)),
    }
    # 別の資料の実行は候補にしないので、その許容は不採用の理由にも含めない。
    assert (result["sources"], result["skipped"], result["skipped_runs"]) == ({str(same): 2}, {}, {})
    assert result["remaining"] == {"visual": 0, "overview": 1, "content": 0, "warnings": 1}


def test_auto_carry_skipped_when_locked(tmp_path, monkeypatch):
    first = fixture_run(tmp_path)
    complete(first)
    second = fixture_run(tmp_path, "0002")
    monkeypatch.setattr(records, "LOCK_WAIT_SECONDS", 0)
    # 示すコマンドの実行は、カレントディレクトリからの相対パスで書く。
    monkeypatch.chdir(second.parent.parent.parent)
    message = (
        "別の処理が実行中のため引き継がなかった。後で slidekit review carry --from auto --to out/runs/0002 を実行する"
    )
    with records._lock(second.parent.parent / records.LOCK_NAME):
        for result in (records.carry_reviews_auto(second), records.carry_reviews("auto", second)):
            assert (result["status"], result["exit_code"], result["message"]) == ("skipped", 0, message)
    assert records._review(second)["records"] == []
    assert len(records.carry_reviews_auto(second, wait_seconds=0)["records"]) == 2


def test_revoking_copy_revokes_original_everywhere(tmp_path):
    first = fixture_run(tmp_path, warning=True)
    original = complete(first, warning=True)
    second = fixture_run(tmp_path, "0002", warning=True)
    records.carry_reviews_auto(second)
    third = fixture_run(tmp_path, "0003", warning=True)
    records.carry_reviews_auto(third)
    records.record_review(third, "overview", None, "確認者", "一覧を確認")
    assert records.check_build(third, final=True)["status"] == "passed"
    copy = next(record for record in records._review(second)["records"] if record["kind"] == "visual")
    result = records.revoke_review(second, copy["id"], "確認者", "表示を見直す")
    assert (result["revoked"], result["notices"]) == ([copy["id"]], [])
    # 元の記録の実行にも取り消しを加え、どの写しから波及したかを cascade に残す。
    cascade = records._review(first)["records"][-1]
    assert (cascade["kind"], cascade["record_id"], cascade["cascade"]) == (
        "revoke",
        original["visual"]["records"][0]["id"],
        {"run": str(second), "record_id": copy["id"]},
    )
    assert result["cascaded"] == [{"run": str(first), "record_id": cascade["record_id"], "revoke_id": cascade["id"]}]
    # 0003 の写しも同じ元の記録を指すので無効になる。
    checked = records.check_build(third, final=True)
    assert [(issue["code"], issue["target"]["kind"]) for issue in checked["issues"]] == [("review_missing", "visual")]
    assert [item["reason"] for item in checked["invalid_records"]] == ["引き継ぎ元: 確認記録が取り消されている"]
    # 元の記録の実行がなければ、写しだけを取り消して知らせる。
    content = next(record for record in records._review(third)["records"] if record["kind"] == "content")
    shutil.rmtree(first)
    result = records.revoke_review(third, content["id"], "確認者", "内容を見直す")
    assert result["cascaded"] == []
    assert [notice["message"] for notice in result["notices"]] == ["元の記録の実行がないため、写しだけを取り消した"]
    assert records._review(third)["records"][-1]["record_id"] == content["id"]


# 警告の許容と取り消しのまとめ


def test_accept_by_code_slides_and_element(tmp_path):
    run = fixture_run(tmp_path, pages=("a", "b", "c"))
    set_diagnostics(
        run,
        [
            warning("d1", "a", "box", "label"),
            warning("d2", "b", "box", "label"),
            warning("d3", "b", "note", "frame"),
            warning("d4", "b", "frame", "box"),
            warning("d5", "b", "box", "caption", declared={"reason": "見出しに重ねる", "location": "slides.py:8"}),
            warning("d6", "b", "box", code="text_padding_overlap"),
            warning("d7", "b", "box", "label2"),
        ],
    )
    records.accept_warning(run, "d2", "確認者", "意図した重なり")
    result = records.accept_warnings(
        run, reviewer="確認者", reason="ラベルを枠に重ねる", code="element_overlap", slides="2", element="box"
    )
    # 種類・ページ・要素（組の相手を含む）で絞った警告だけに、同じ batch の記録を作る。
    assert [record["diagnostic_id"] for record in result["records"]] == ["d4", "d7"]
    assert {record["batch"] for record in result["records"]} == {result["batch"]}
    assert result["skipped"] == [
        {"diagnostic_id": "d2", "reason": "すでに許容している"},
        {"diagnostic_id": "d5", "reason": "資料コードで許容している"},
    ]
    rest = records.accept_warnings(run, reviewer="確認者", reason="残りの重なりも意図どおり", code="element_overlap")
    assert [record["diagnostic_id"] for record in rest["records"]] == ["d1", "d3"]
    final = records.check_build(run, final=True)
    assert [issue["diagnostic_id"] for issue in final["issues"] if issue["code"] == "warning_unaccepted"] == ["d6"]


def test_accept_without_match_or_with_both_selectors_fails(tmp_path):
    run = fixture_run(tmp_path, warning=True)
    set_diagnostics(
        run, [WARNING, {**warning("d_error", "intro", "body"), "severity": "error", "code": "text_overflow"}]
    )
    for arguments, message in (
        ({"diagnostic_ids": ["d_overlap"], "code": "element_overlap"}, "警告の識別子か種類のどちらか一方を指定する"),
        ({}, "警告の識別子か種類のどちらか一方を指定する"),
        ({"code": "notes_style"}, "条件に合う未対応の警告がない"),
        ({"code": "element_overlap", "element": "missing"}, "条件に合う未対応の警告がない"),
        ({"code": "element_overlap", "slides": "all"}, "all と pending は使えない"),
        ({"diagnostic_ids": "missing,d_error"}, "許容できる警告の識別子ではない: missing, d_error"),
        ({"diagnostic_ids": ["d_overlap"], "element": "box"}, "--slides と --element は --code と組み合わせる"),
    ):
        with pytest.raises(records.RecordError, match=re.escape(message)):
            records.accept_warnings(run, reviewer="確認者", reason="意図した重なり", **arguments)
    assert records._review(run)["records"] == []
    records.accept_warnings(run, reviewer="確認者", reason="意図した重なり", diagnostic_ids="d_overlap")
    with pytest.raises(
        records.RecordError,
        match=re.escape("条件に合う未対応の警告がない（1 件は許容済みか、資料コードで許容している）"),
    ):
        records.accept_warnings(run, reviewer="確認者", reason="もう一度", code="element_overlap")


def test_revoke_batch_revokes_all_records(tmp_path):
    run = fixture_run(tmp_path, pages=("a", "b"))
    recorded = records.record_review(run, "visual,content", "all", "確認者", "表示と内容を確認")
    records.record_review(run, "overview", None, "確認者", "一覧を確認")
    assert records.check_build(run, final=True)["status"] == "passed"
    result = records.revoke_review(run, recorded["batch"], "確認者", "まとめて見直す")
    assert result["revoked"] == [record["id"] for record in recorded["records"]]
    assert len(result["records"]) == 4 and {record["batch"] for record in result["records"]} == {result["batch"]}
    checked = records.check_build(run, final=True)
    assert sorted((issue["target"]["kind"], issue["target"]["slide_id"]) for issue in checked["issues"]) == [
        ("content", "a"),
        ("content", "b"),
        ("visual", "a"),
        ("visual", "b"),
    ]
    with pytest.raises(records.RecordError, match="見つからない: b_0000000000000000"):
        records.revoke_review(run, f"{recorded['records'][0]['id']},b_0000000000000000", "確認者", "見直す")


def test_declared_warning_passes_final_without_record(tmp_path):
    run = fixture_run(tmp_path)
    set_diagnostics(
        run,
        [
            warning(
                "d_declared",
                "intro",
                "label",
                "frame",
                declared={"reason": "ラベルを枠に重ねて示す", "location": "slides.py:42"},
            )
        ],
    )
    records.record_review(run, "visual,overview,content", "all", "確認者", "表示と内容を確認")
    final = records.check_build(run, final=True)
    assert final["status"] == "passed"
    assert final["summary"]["warnings"] == {"total": 1, "accepted": 0, "declared": 1, "unaccepted": 0}
    entry = {
        "diagnostic_id": "d_declared",
        "code": "element_overlap",
        "slide_id": "intro",
        "slide_number": 1,
        "element_id": "label",
        "related_element_id": "frame",
        "reason": "ラベルを枠に重ねて示す",
        "location": "slides.py:42",
    }
    assert final["declared_warnings"] == [entry]
    status = records.review_status(run)
    assert (status["status"], status["declared_warnings"], status["warnings"]["pending"]) == ("complete", [entry], [])
    assert not [record for record in records._review(run)["records"] if record["kind"] == "accept"]
    with pytest.raises(records.RecordError, match="条件に合う未対応の警告がない"):
        records.accept_warning(run, "d_declared", "確認者", "重ねて許容する")


# 資料の移動とロックの場所


def test_carry_after_moving_project_with_output_directory(tmp_path):
    first = fixture_run(tmp_path, warning=True)
    complete(first, warning=True)
    second = fixture_run(tmp_path, "0002", warning=True)
    records.carry_reviews_auto(second)
    records.record_review(second, "overview", None, "確認者", "一覧を確認")
    # 資料のディレクトリを out/ ごと別の場所へ移す。
    (tmp_path / "moved").mkdir()
    shutil.move(tmp_path / "project", tmp_path / "moved" / "project")
    root = tmp_path / "moved" / "project"
    first, second = root / "out" / "runs" / "0001", root / "out" / "runs" / "0002"
    # 移す前に作った写しも、移した先の元の記録を指すものとして有効で、整理はその元の記録の実行を保護する。
    checked = records.check_build(second, final=True)
    assert checked["invalid_records"] == [] and {issue["code"] for issue in checked["issues"]} == {"input_changed"}
    assert records.run_references(second)["sources"] == [str(first)]
    project = SimpleNamespace(
        root=root, output_dir=root / "out", config={"deliverables": {"record": str(root / "deliverables.json")}}
    )
    assert runs.protected_runs(project, second.parent, second.parent.parent)[first] == ["確認記録の引き継ぎ元"]
    # 移動前の実行から、手動でも自動でも引き継ぐことができる。
    third = fixture_run(tmp_path / "moved", "0003", warning=True)
    manual = records.carry_reviews(first, third)
    assert {(record["kind"], record["source"]["run"]) for record in manual["records"]} == {
        ("visual", str(first)),
        ("content", str(first)),
        ("accept", str(first)),
    }
    fourth = fixture_run(tmp_path / "moved", "0004", warning=True)
    auto = records.carry_reviews_auto(fourth)
    assert (auto["sources"], auto["skipped_runs"]) == ({str(first): 3}, {})
    records.record_review(fourth, "overview", None, "確認者", "一覧を確認")
    assert records.export_build(fourth)["status"] == "exported"
    # 移す前の場所は作り直さない。
    assert not (tmp_path / "project").exists()


def test_export_history_protects_runs_after_moving_project(tmp_path):
    # 反映の履歴は実行を資料のディレクトリからの相対パスで書くので、資料を out/ ごと移した後も、反映した実行と退避した旧版を
    # 整理が保護する（以前は絶対パスで書き、移した後は保護が外れた）。
    first = fixture_run(tmp_path, payload="first")
    complete(first)
    records.export_build(first)
    second = fixture_run(tmp_path, "0002", payload="second")
    complete(second)
    records.export_build(second)
    root = move_project(tmp_path / "project", tmp_path / "moved" / "project")
    output = root / "out"
    project = SimpleNamespace(
        root=root, output_dir=output, config={"deliverables": {"record": str(root / "deliverables.json")}}
    )
    protected = runs.protected_runs(project, output / "runs", output)
    # 理由の順は反映の履歴の識別子（秒までの時刻と乱数）の順なので、同じ秒の反映では入れ替わる。
    assert {path.name: sorted(reasons) for path, reasons in protected.items()} == {
        "0001": sorted(["完成物として反映した実行", "旧版として退避した実行"]),
        "0002": sorted(["現在採用している実行", "完成物として反映した実行"]),
    }
    assert all(
        str(tmp_path) not in json.dumps({key: value[key] for key in ("run", "previous_run")})
        for value in map(records.read_record, sorted(output.glob("exports/*/transaction.json")))
    )


def move_project(source, destination):
    """資料のディレクトリを out/ ごと移す。"""
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(source, destination)
    return destination


def test_carry_after_moving_project_twice(tmp_path):
    """移した後に作った写しも、もう一度移した後に有効である（引き継ぎ元は元の記録の識別子で同じ実行と判断する）。"""
    complete(fixture_run(tmp_path, warning=True), warning=True)
    fixture_run(tmp_path, "0002", warning=True)
    # 1 回目の移動の後に、移す前に生成した実行どうしで引き継ぐ。写しは 1 回目の移動先の実行を指す。
    runs_dir = move_project(tmp_path / "project", tmp_path / "first" / "project") / "out" / "runs"
    carried = records.carry_reviews(runs_dir / "0001", runs_dir / "0002")
    assert {record["source"]["run"] for record in carried["records"]} == {str(runs_dir / "0001")}
    records.record_review(runs_dir / "0002", "overview", None, "確認者", "一覧を確認")
    root = move_project(tmp_path / "first" / "project", tmp_path / "second" / "project")
    first, second = root / "out" / "runs" / "0001", root / "out" / "runs" / "0002"
    checked = records.check_build(second, final=True)
    assert checked["invalid_records"] == [] and {issue["code"] for issue in checked["issues"]} == {"input_changed"}
    assert records.run_references(second)["sources"] == [str(first)]
    project = SimpleNamespace(
        root=root, output_dir=root / "out", config={"deliverables": {"record": str(root / "deliverables.json")}}
    )
    assert runs.protected_runs(project, second.parent, second.parent.parent)[first] == ["確認記録の引き継ぎ元"]
    # 写しの取り消しは移した先の元の記録の実行に波及する。
    copy = next(record for record in records._review(second)["records"] if record["kind"] == "content")
    result = records.revoke_review(second, copy["id"], "確認者", "内容を見直す")
    assert ([item["run"] for item in result["cascaded"]], result["notices"]) == ([str(first)], [])


def test_legacy_copy_sources_follow_moved_project(tmp_path):
    """形式 1 の写し（1 つ前の実行を指す）の引き継ぎ元も、資料を out/ ごと移した後の場所に読み替える（整理で保護するため）。"""
    old = [fixture_run(tmp_path, f"{number:04d}", schema=1) for number in (1, 2)]
    legacy_copy(old[0], old[0].parent / "0000", identifier="r_0000000000000000")
    legacy_copy(old[1], old[0])
    assert records.run_references(old[1])["sources"] == [str(old[0])]
    runs_dir = move_project(tmp_path / "project", tmp_path / "moved" / "project") / "out" / "runs"
    assert records.run_references(runs_dir / "0002") == {
        "status": "built",
        "exportable": False,
        "sources": [str(runs_dir / "0001")],
    }
    # 同じ名前の実行があっても、指す記録を持たなければ読み替えない。
    fixture_run(tmp_path / "moved", "0000", schema=1)
    assert records.run_references(runs_dir / "0001")["sources"] == [str(old[0].parent / "0000")]


def test_lock_file_is_kept_in_output_directory(tmp_path):
    run = fixture_run(tmp_path, warning=True)
    root, output = run.parent.parent.parent, run.parent.parent
    complete(run, warning=True)
    second = fixture_run(tmp_path, "0002", warning=True)
    records.carry_reviews_auto(second)
    overview = records.record_review(second, "overview", None, "確認者", "一覧を確認")
    records.revoke_review(second, overview["batch"], "確認者", "見直す")
    records.check_build(run, final=True)
    assert records.export_build(run)["status"] == "exported"
    # 資料のディレクトリにロックのファイルを残さない（バージョン管理で未追跡のファイルにならない）。
    assert [path.name for path in root.iterdir() if path.name.startswith(".")] == []
    assert (output / records.LOCK_NAME).is_file()
    # runs だけを含んでいた（slidekit が作った）実行保存先には、生成と同じく .gitignore を置く。
    assert (output / ".gitignore").read_text(encoding="utf-8") == runs.GITIGNORE
    # 実行保存先がなければ作る。利用者のファイルがある実行保存先には .gitignore を置かない。
    fresh, custom = tmp_path / "fresh" / "out", tmp_path / "custom"
    custom.mkdir()
    (custom / "memo.txt").write_text("メモ", encoding="utf-8")
    for place in (fresh, custom):
        with records._output_lock(place):
            pass
    assert sorted(path.name for path in fresh.iterdir()) == [".gitignore", records.LOCK_NAME]
    assert sorted(path.name for path in custom.iterdir()) == [records.LOCK_NAME, "memo.txt"]


def test_run_outside_runs_directory_leaves_no_lock(tmp_path):
    """runs/ の外へ複製した実行は、ロックを取らずに検査し、記録の操作は拒否する（関係のない場所にロックのファイルを作らない）。"""
    original = fixture_run(tmp_path, warning=True)
    complete(original)
    home = tmp_path / "home"
    first, second = home / "Downloads" / "0001", home / "Downloads" / "0002"
    shutil.copytree(original, first)
    shutil.copytree(fixture_run(tmp_path, "0002", warning=True), second)
    for final in (False, True):
        records.check_build(first, final=final)
    assert (first / records.CHECK_FILE).is_file() and (first / records.FINAL_CHECK_FILE).is_file()
    operations = {
        "record": lambda: records.record_review(second, "content", "all", "確認者", "内容を確認"),
        "accept": lambda: records.accept_warning(second, "d_overlap", "確認者", "図の意図に沿う重なりである"),
        "revoke": lambda: records.revoke_review(first, records._review(first)["records"][0]["id"], "確認者", "見直す"),
        "carry": lambda: records.carry_reviews(first, second),
        "auto": lambda: records.carry_reviews_auto(second),
    }
    for operation in operations.values():
        with pytest.raises(records.RecordError, match="実行保存先の runs/ にない実行には記録できない"):
            operation()
    assert records._review(second)["records"] == [] and records._review(first) == records._review(original)
    assert records.review_status(second)["status"] == "incomplete"
    # 入力が変わった後に示す再生成のコマンドは、生成記録の資料を指す（複製した場所から資料の場所を求めない）。
    (tmp_path / "project" / "slides.py").write_text("変更した元資料", encoding="utf-8")
    assert records.review_status(second)["next"] == [f"slidekit build --project {tmp_path / 'project' / 'deck.toml'}"]
    assert sorted(path.relative_to(home).as_posix() for path in home.iterdir()) == ["Downloads"]
    assert sorted(path.name for path in (home / "Downloads").iterdir()) == ["0001", "0002"]


# ロック


class WindowsLocks:
    """Windows の msvcrt の代わり。LockFile と同じく、ロックした範囲はほかのハンドルから読むことができない。

    ロックは先頭の 1 バイトだけを模す。閉じたときにロックが残っていたかを記録し、OS と同じく閉じたら解く。
    """

    LK_UNLCK, LK_NBLCK = 0, 2

    def __init__(self):
        self.owner = None
        self.held_at_close = []
        self.guard = threading.Lock()

    def locking(self, fd, mode, nbytes):
        # msvcrt.locking はファイル記述子の現在の位置から固定する。
        assert (os.lseek(fd, 0, os.SEEK_CUR), nbytes) == (0, 1)
        with self.guard:
            if mode == self.LK_NBLCK and self.owner is None:
                self.owner = fd
            elif mode == self.LK_UNLCK and self.owner == fd:
                self.owner = None
            else:
                raise PermissionError(13, "ロックが競合した")

    def install(self, monkeypatch):
        """records の _lock を、この msvcrt と、読み取りを制限したファイルで動かす。"""
        locks = self

        class Handle:
            def __init__(self, stream):
                self.stream = stream

            def read(self, *args):
                if locks.owner not in (None, self.stream.fileno()):
                    raise PermissionError(13, "ほかのハンドルがロックした範囲は読めない")
                return self.stream.read(*args)

            def close(self):
                with locks.guard:
                    held = locks.owner == self.stream.fileno()
                    locks.held_at_close.append(held)
                    if held:
                        locks.owner = None
                self.stream.close()

            def __getattr__(self, name):
                return getattr(self.stream, name)

        class LockedPath(type(Path())):
            def open(self, *args, **kwargs):
                return Handle(super().open(*args, **kwargs))

        monkeypatch.setitem(sys.modules, "msvcrt", self)
        monkeypatch.setattr(records, "os", SimpleNamespace(name="nt", fstat=os.fstat))
        monkeypatch.setattr(records, "Path", LockedPath)


@pytest.mark.parametrize("platform", ["native", "windows"])
def test_lock_waits_until_released(tmp_path, monkeypatch, platform):
    windows = WindowsLocks() if platform == "windows" else None
    if windows:
        windows.install(monkeypatch)
    path = tmp_path / records.LOCK_NAME
    with records._lock(path):
        # 待ち時間を過ぎると、待った秒数を示して失敗する。
        with pytest.raises(records.RecordError, match=r"別の処理が実行中である（0\.3 秒待った）"):
            with records._lock(path, wait_seconds=0.3):
                pass
    held, release, released = threading.Event(), threading.Event(), threading.Event()

    def hold():
        with records._lock(path, wait_seconds=0):
            held.set()
            release.wait(10)
            released.set()

    sleeps = []

    def sleep(seconds):
        sleeps.append(seconds)
        # 2 回目に取り直しを待つ間に、別の処理がロックを解放する。
        if len(sleeps) == 2:
            release.set()
        time.sleep(seconds)

    # records の中の待ちだけを集計する（ほかのスレッドの time.sleep は集計しない）。
    monkeypatch.setattr(records, "time", SimpleNamespace(monotonic=time.monotonic, sleep=sleep))
    thread = threading.Thread(target=hold)
    thread.start()
    try:
        assert held.wait(10)
        # 0 秒なら取り直さずに直ちに失敗する。
        with pytest.raises(records.RecordError, match="別の処理が実行中である: "):
            with records._lock(path, wait_seconds=0):
                pass
        assert sleeps == []
        with records._lock(path, wait_seconds=10):
            assert released.is_set()
        assert len(sleeps) >= 2 and all(0 < seconds <= 0.2 for seconds in sleeps)
    finally:
        release.set()
        thread.join()
    for wrong in (-1, float("nan"), True, "1"):
        with pytest.raises(ValueError, match="待ち時間"):
            with records._lock(path, wait_seconds=wrong):
                pass
    # ロックのファイルのディレクトリがなければ作る（まだ生成していない資料の実行保存先など）。
    with records._lock(tmp_path / "missing" / records.LOCK_NAME):
        pass
    assert (tmp_path / "missing" / records.LOCK_NAME).is_file()
    if windows:
        # ロックしたハンドルは、閉じる前にロックを解いた（閉じた後に OS が解くのを待たない）。
        assert len(windows.held_at_close) == 6 and not any(windows.held_at_close)
        assert path.read_bytes() == b"0"


def test_project_lock_prevents_concurrent_export(tmp_path, monkeypatch):
    run = fixture_run(tmp_path)
    complete(run)
    monkeypatch.setattr(records, "LOCK_WAIT_SECONDS", 0)
    # 資料単位のロックは実行保存先にある。
    with records._lock(run.parent.parent / records.LOCK_NAME):
        with pytest.raises(records.RecordError, match="実行中"):
            records.export_build(run)


# 形式 1（slidekit 0.1）の実行と記録


def legacy_copy(run, source, identifier="r_legacy00000000"):
    """形式 1 の写し（引き継ぎ元は写した時点の 1 つ前の実行）を確認記録に加える。"""
    review = json.loads((run / "review.json").read_text(encoding="utf-8"))
    review["records"].append(
        {
            "id": identifier,
            "kind": "content",
            "slide_id": "intro",
            "reviewer": "確認者",
            "note": "内容を確認",
            "source": {"run": str(source), "record_id": "r_0000000000000000", "record_hash": "0" * 64},
        }
    )
    save(run / "review.json", review)


def test_legacy_run_is_rejected_with_guidance(tmp_path):
    legacy = fixture_run(tmp_path, warning=True, schema=1)
    current = fixture_run(tmp_path, "0002", warning=True)
    complete(current, warning=True)
    operations = {
        "record": lambda: records.record_review(legacy, "content", "all", "確認者", "内容を確認"),
        "accept": lambda: records.accept_warning(legacy, "d_overlap", "確認者", "意図した重なり"),
        "revoke": lambda: records.revoke_review(legacy, "r_0000000000000000", "確認者", "見直す"),
        "carry_from": lambda: records.carry_reviews(legacy, current),
        "carry_to": lambda: records.carry_reviews(current, legacy),
        "check": lambda: records.check_build(legacy),
        "final": lambda: records.check_build(legacy, final=True),
        "export": lambda: records.export_build(legacy),
        "status": lambda: records.review_status(legacy),
        "carry_auto": lambda: records.carry_reviews_auto(legacy),
        "accept_code": lambda: records.accept_warnings(
            legacy, reviewer="確認者", reason="意図した重なり", code="element_overlap"
        ),
    }
    for name, operation in operations.items():
        with pytest.raises(
            records.LegacyRecordError,
            match=r"古い形式（schema_version 1）の実行である。現在の slidekit で再生成してから確認する",
        ) as caught:
            operation()
        assert caught.value.code == "legacy_run", name
    assert sorted(path.name for path in legacy.iterdir()) == [
        "deck.json",
        "deck.pdf",
        "deck.pptx",
        "diagnostics.json",
        "manifest.json",
        "notes.md",
        "renders",
        "review.json",
    ]
    assert [record["kind"] for record in records._review(current)["records"]] == [
        "visual",
        "overview",
        "content",
        "accept",
    ]


def test_run_references_reads_legacy_and_new_runs(tmp_path):
    # 形式 2：引き継ぎを重ねても、写しは人が記録した元の記録の実行を指す。
    first = fixture_run(tmp_path / "current")
    complete(first)
    second = fixture_run(tmp_path / "current", "0002")
    records.carry_reviews(first, second)
    third = fixture_run(tmp_path / "current", "0003")
    records.carry_reviews(second, third)
    assert records.run_references(third) == {"status": "built", "exportable": True, "sources": [str(first)]}
    assert records.run_references(first) == {"status": "built", "exportable": True, "sources": []}
    # 形式 1：写しは 1 つ前の実行を指す。
    old = [fixture_run(tmp_path / "legacy", f"{number:04d}", schema=1) for number in (1, 2, 3)]
    legacy_copy(old[1], old[0])
    legacy_copy(old[2], old[1])
    assert records.run_references(old[2]) == {"status": "built", "exportable": False, "sources": [str(old[1])]}
    # 確認記録や生成記録のない実行。
    bare = tmp_path / "bare"
    bare.mkdir()
    assert records.run_references(bare) == {"status": None, "exportable": False, "sources": []}


def test_export_after_legacy_adoption_records_previous_run(tmp_path):
    legacy = fixture_run(tmp_path, schema=1)
    root = legacy.parent.parent.parent
    for name in ("presentation.pptx", "presentation.pdf", "speaker_notes.md"):
        (root / name).write_text(f"旧版の {name}", encoding="utf-8")
    save(
        root / "deliverables.json",
        {
            "schema_version": 1,
            "run": str(legacy),
            "run_id": "0001",
            "project": str(root / "deck.toml"),
            "exported_at": "2026-10-01T00:00:00+00:00",
            "artifacts": {str(root / "presentation.pptx"): "0" * 64},
        },
    )
    before = (root / "deliverables.json").read_bytes()
    run = fixture_run(tmp_path, "0002")
    complete(run)
    result = records.export_build(run)
    history = Path(result["history"])
    assert records.read_record(history / "transaction.json")["previous_run"] == "out/runs/0001"
    # 旧形式の採用記録と完成物は、内容を変えずに退避する。
    assert (history / "previous-record").read_bytes() == before
    assert (history / "previous-pptx").read_text(encoding="utf-8") == "旧版の presentation.pptx"
    protected = runs.protected_runs(project_for(run), run.parent, run.parent.parent)
    assert protected[legacy.resolve()] == ["旧版として退避した実行"]


def test_legacy_pending_transaction_is_recovered(tmp_path, monkeypatch):
    first, second, deliverables, pending = interrupt_second_export(tmp_path, monkeypatch)
    # slidekit 0.1 が中断した反映の記録（構造は同じで、形式の番号だけが 1）。
    transaction = records.read_record(pending)
    save(pending, {**transaction, "schema_version": 1})
    assert records.export_build(second)["status"] == "exported"
    assert not pending.exists()
    recovered = records.read_record(Path(transaction["history"]) / "transaction.json", allow_legacy=True)
    assert recovered["status"] == "recovered"
    assert Path(deliverables["pdf"]).read_bytes() == (second / "deck.pdf").read_bytes()


def test_legacy_readers_accept_only_version_one(tmp_path):
    values = {}
    for name, version in (("legacy", 1), ("current", 2), ("future", 3), ("boolean", True), ("missing", None)):
        values[name] = {"run": "out/runs/0001"} | ({} if version is None else {"schema_version": version})
        save(tmp_path / f"{name}.json", values[name])
    assert records._read_legacy_export_record(tmp_path / "legacy.json") == values["legacy"]
    for name in ("current", "future", "boolean", "missing"):
        with pytest.raises(records.RecordError, match="形式 1 の記録ではない"):
            records._read_legacy_export_record(tmp_path / f"{name}.json")
    legacy = fixture_run(tmp_path / "old", schema=1)
    legacy_copy(legacy, legacy.parent / "0000")
    assert records._legacy_run_references(legacy) == {"status": "built", "sources": [str(legacy.parent / "0000")]}
    with pytest.raises(records.RecordError, match="形式 1 の確認記録ではない"):
        records._legacy_run_references(fixture_run(tmp_path / "new"))
    # 公開の読み取りは allow_legacy のときだけ形式 1 を読む。
    assert records.read_record(tmp_path / "legacy.json", allow_legacy=True) == values["legacy"]
    with pytest.raises(records.LegacyRecordError, match="古い形式"):
        records.read_record(tmp_path / "legacy.json")
    assert records.read_record(tmp_path / "current.json") == values["current"]
    for name in ("future", "boolean", "missing"):
        with pytest.raises(records.RecordError, match="対応していない記録形式"):
            records.read_record(tmp_path / f"{name}.json", allow_legacy=True)
