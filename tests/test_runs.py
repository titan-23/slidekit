"""実行ディレクトリの割り当て、生成中の印、実行の一覧と整理を確かめる。"""

import json
import os
from pathlib import Path
import socket
import subprocess
import sys
import time
from types import SimpleNamespace

import pytest

from slidekit import records, runs
from test_records import complete, fixture_run, project_for

HAS_PROC = Path(f"/proc/{os.getpid()}/stat").exists()
# どの環境でも使われていないプロセス番号（Linux の上限は 4194304）。
DEAD_PID = 999999999
SELECTION = {"spec": "2", "slide_ids": ["method"], "numbers": [2], "total_slides": 3}
LEGACY_REASON = "旧形式の生成中の印です。生成が動いていないことを確かめてから .building を消してください"


def save(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")


def make_project(tmp_path, output_dir=None):
    """整理と一覧が使う属性だけを持つ資料。deck.toml は出力先を読む試験のために置く。"""
    root = tmp_path / "project"
    root.mkdir(parents=True, exist_ok=True)
    config = root / "deck.toml"
    setting = f'output_dir = "{output_dir}"\n' if output_dir else ""
    config.write_text("schema_version = 1\n" + setting, encoding="utf-8")
    return SimpleNamespace(
        path=config,
        root=root,
        output_dir=(root / (output_dir or "out")).resolve(),
        config={"deliverables": {"record": str(root / "deliverables.json")}},
    )


def make_run(project, name, status="built", *, schema=2, mode="full", selection=None, sources=()):
    """生成記録と確認記録だけを持つ実行。schema=1 は mode・selection・exportable のない 0.1 の形式（確認記録も形式 1）。"""
    run = project.output_dir / "runs" / name
    run.mkdir(parents=True)
    manifest = {
        "schema_version": schema,
        "run_id": name,
        "status": status,
        "created_at": f"2026-10-01T00:00:{int(name):02d}+00:00",
    }
    if schema >= 2:
        manifest.update(
            mode=mode, selection=selection, exportable=status == "built" and mode == "full" and selection is None
        )
    save(run / "manifest.json", manifest)
    copies = [
        {
            "id": f"r_{index:016x}",
            "kind": "content",
            "source": {"run": str(source), "record_id": "r_0000000000000000", "record_hash": "0"},
        }
        for index, source in enumerate(sources)
    ]
    save(run / "review.json", {"schema_version": records.SCHEMA_VERSION if schema >= 2 else schema, "records": copies})
    return run


def rewrite_marker(run, **changes):
    path = run / ".building"
    marker = {**json.loads(path.read_text(encoding="utf-8")), **changes}
    save(path, marker)
    return marker


def run_ids(result):
    return {item["run_id"] for item in result["candidates"]}


def test_clean_filters_combine(tmp_path):
    project = make_project(tmp_path)
    make_run(project, "0001")
    make_run(project, "0002", "failed")
    make_run(project, "0003", "stopped")
    make_run(project, "0004", mode="draft")
    make_run(project, "0005", selection=SELECTION)
    interrupted = runs.new_run(project, command="build", options={"draft": False, "slides": None})
    rewrite_marker(interrupted, pid=DEAD_PID)
    assert interrupted.name == "0006"
    everything = {"0001", "0002", "0003", "0004", "0005", "0006"}
    cases = [
        ({"failed": True}, {"0002", "0003", "0006"}),
        ({"drafts": True}, {"0004", "0005"}),
        ({"failed": True, "drafts": True}, {"0002", "0003", "0004", "0005", "0006"}),
        ({"keep": 2}, {"0001", "0002", "0003", "0004"}),
        ({"keep": 0}, everything),
        ({"keep": 6}, set()),
        ({"before": "0004"}, {"0001", "0002", "0003"}),
        ({"before": "4"}, {"0001", "0002", "0003"}),
        ({"failed": True, "keep": 1}, {"0002", "0003"}),
        ({"drafts": True, "before": "0005"}, {"0004"}),
        ({"before": "0006", "keep": 2}, {"0001", "0002", "0003"}),
    ]
    for conditions, expected in cases:
        result = runs.clean_runs(project, **conditions)
        assert result["status"] == "preview" and result["removed"] == []
        assert run_ids(result) == expected, conditions
        assert result["candidate_paths"] == [item["run"] for item in result["candidates"]]
        assert result["bytes"] == sum(item["bytes"] for item in result["candidates"])
        assert all(item["bytes"] > 0 and item["reason"] for item in result["candidates"])
    assert {path.name for path in (project.output_dir / "runs").iterdir()} == everything
    result = runs.clean_runs(project, failed=True, drafts=True, keep=1)
    assert {item["run_id"]: (item["status"], item["mode"], item["reason"]) for item in result["candidates"]} == {
        "0002": ("failed", "full", "失敗した生成、新しい 1 件より古い"),
        "0003": ("stopped", "full", "変換の前に止めた生成、新しい 1 件より古い"),
        "0004": ("built", "draft", "下書き、新しい 1 件より古い"),
        "0005": ("built", "full", "部分生成、新しい 1 件より古い"),
    }
    result = runs.clean_runs(project, failed=True, before="0007")
    interrupted_item = next(item for item in result["candidates"] if item["run_id"] == "0006")
    assert (interrupted_item["status"], interrupted_item["mode"], interrupted_item["reason"]) == (
        None,
        None,
        "中断した生成、0007 より前",
    )
    assert result["conditions"] == {"before": "0007", "keep": None, "failed": True, "drafts": False}


def test_clean_requires_a_condition(tmp_path):
    project = make_project(tmp_path)
    old = make_run(project, "0001")
    with pytest.raises(ValueError, match="--before、--keep、--failed、--drafts のいずれかを指定してください"):
        runs.clean_runs(project, apply=True)
    for wrong in (
        {"keep": -1},
        {"keep": True},
        {"keep": 1.5},
        {"before": ""},
        {"before": ".."},
        {"before": "../0002"},
        {"before": 2},
    ):
        with pytest.raises(ValueError):
            runs.clean_runs(project, apply=True, **wrong)
    assert old.exists()


def test_clean_keeps_protected_runs(tmp_path):
    # 0.1 の test_clean_protects_adoption_history_carry_sources_and_running_builds と同じ状況を、記録のファイルから作る。
    project = make_project(tmp_path)
    first, second, third = (make_run(project, name) for name in ("0001", "0002", "0003"))
    fourth = make_run(project, "0004", sources=[third])
    fifth = make_run(project, "0005")
    allocated = project.output_dir / "runs" / "0006"
    allocated.mkdir()
    running = runs.new_run(project, command="build", options={})
    exporting = make_run(project, "0008")
    rolled_back = make_run(project, "0009")
    assert running.name == "0007"
    version = records.SCHEMA_VERSION
    exports = project.output_dir / "exports"
    save(project.root / "deliverables.json", {"schema_version": version, "run": str(second)})
    save(
        exports / "20261001T000000-aaaaaaaa" / "transaction.json",
        {"schema_version": version, "status": "committed", "run": str(first), "previous_run": None},
    )
    save(
        exports / "20261002T000000-bbbbbbbb" / "transaction.json",
        {"schema_version": version, "status": "committed", "run": str(second), "previous_run": str(first)},
    )
    save(
        exports / "20261003T000000-cccccccc" / "transaction.json",
        {"schema_version": version, "status": "rolled_back", "run": str(rolled_back), "previous_run": str(second)},
    )
    save(
        exports / "pending.json",
        {"schema_version": version, "status": "prepared", "run": str(exporting), "previous_run": str(second)},
    )
    protected = runs.protected_runs(project, project.output_dir / "runs", project.output_dir)
    assert protected[first] == ["完成物として反映した実行", "旧版として退避した実行"]
    assert protected[second] == ["現在採用している実行", "完成物として反映した実行", "復旧に必要な旧版の実行"]
    assert protected[third] == ["確認記録の引き継ぎ元"]
    assert protected[allocated] == ["生成記録も生成中の印もない（割り当ての直後か、手で作った実行）"]
    assert protected[running] == [f"実行中（PID {os.getpid()}、段階 construct）"]
    assert protected[exporting] == ["完成物の反映が未完了の実行"]
    assert not {fourth, fifth, rolled_back} & set(protected)
    preview = runs.clean_runs(project, before="0010")
    assert preview["candidate_paths"] == [str(fourth), str(fifth), str(rolled_back)]
    assert [item["run"] for item in preview["protected"]] == [
        str(path) for path in (first, second, third, allocated, running, exporting)
    ]
    assert fourth.exists() and fifth.exists() and rolled_back.exists()
    applied = runs.clean_runs(project, before="0010", apply=True)
    assert applied["status"] == "cleaned"
    assert applied["removed"] == [str(fourth), str(fifth), str(rolled_back)]
    assert all(path.exists() for path in (first, second, third, allocated, running, exporting))
    assert not any(path.exists() for path in (fourth, fifth, rolled_back))


def test_clean_uses_output_lock_shared_with_records(tmp_path, monkeypatch):
    # 整理は確認記録の操作と同じ実行保存先のロックを取り、資料のディレクトリにロックのファイルを残さない。
    project = make_project(tmp_path)
    make_run(project, "0001", status="failed")
    monkeypatch.setattr(records, "LOCK_WAIT_SECONDS", 0)
    with records._output_lock(project.output_dir):
        with pytest.raises(records.RecordError, match="実行中"):
            runs.clean_runs(project, failed=True, apply=True)
    assert (project.output_dir / "runs" / "0001").is_dir()
    assert runs.clean_runs(project, failed=True, apply=True)["removed"] == [str(project.output_dir / "runs" / "0001")]
    assert not (project.root / records.LOCK_NAME).exists()
    assert (project.output_dir / records.LOCK_NAME).is_file()


def test_clean_preview_writes_nothing(tmp_path, monkeypatch):
    # 候補の表示（apply なし）はロックを取らず、実行保存先にロックのファイルも .gitignore も置かない（以前は、
    # runs・exports だけを含む実行保存先に 2 つのファイルを作っていた）。別の処理がロックを持っていても待たない。
    project = make_project(tmp_path)
    make_run(project, "0001", status="failed")
    (project.output_dir / ".gitignore").unlink(missing_ok=True)
    before = sorted(path.relative_to(tmp_path) for path in tmp_path.rglob("*"))
    preview = runs.clean_runs(project, failed=True)
    assert preview["candidate_paths"] == [str(project.output_dir / "runs" / "0001")]
    assert sorted(path.relative_to(tmp_path) for path in tmp_path.rglob("*")) == before
    monkeypatch.setattr(records, "LOCK_WAIT_SECONDS", 0)
    with records._output_lock(project.output_dir):
        assert runs.clean_runs(project, failed=True)["candidate_paths"] == preview["candidate_paths"]
        with pytest.raises(records.RecordError, match="実行中"):
            runs.clean_runs(project, failed=True, apply=True)


def test_clean_protects_adoption_history_carry_sources_and_running_builds(tmp_path):
    # 確認記録の操作（records）が書いた採用記録・反映の記録・写しを、整理が読んで保護する。 0.1 の records.clean_runs の呼び方で確かめる。
    first = fixture_run(tmp_path, "0001")
    complete(first)
    records.export_build(first)
    second = fixture_run(tmp_path, "0002")
    complete(second)
    records.export_build(second)
    third = fixture_run(tmp_path, "0003")
    complete(third)
    fourth = fixture_run(tmp_path, "0004")
    records.carry_reviews(third, fourth)
    fifth = fixture_run(tmp_path, "0005")
    running = first.parent / "0006"
    running.mkdir()
    project = project_for(first)
    preview = records.clean_runs(project, "0007")
    assert set(preview["candidates"]) == {str(fourth), str(fifth)}
    assert fourth.exists() and fifth.exists()
    applied = records.clean_runs(project, "0007", apply=True)
    assert set(applied["removed"]) == {str(fourth), str(fifth)}
    assert all(path.exists() for path in (first, second, third, running))
    with pytest.raises(records.RecordError, match="--before"):
        records.clean_runs(project, "../0001")


def test_clean_keeps_legacy_protected_runs(tmp_path):
    # slidekit 0.1 の形式の採用記録・反映の記録・写しも読み、従来どおり保護する。
    project = make_project(tmp_path)
    adopted, committed, previous, exporting, recovery, source = (
        make_run(project, f"{number:04d}", schema=1) for number in range(1, 7)
    )
    # 形式 1 の写しは写した時点の 1 つ前の実行を指す。
    copy = make_run(project, "0007", schema=1, sources=[source])
    plain = make_run(project, "0008", schema=1)
    exports = project.output_dir / "exports"
    save(project.root / "deliverables.json", {"schema_version": 1, "run": str(adopted)})
    save(
        exports / "20261001T000000-aaaaaaaa" / "transaction.json",
        {"schema_version": 1, "status": "committed", "run": str(committed), "previous_run": str(previous)},
    )
    save(
        exports / "pending.json",
        {"schema_version": 1, "status": "prepared", "run": str(exporting), "previous_run": str(recovery)},
    )
    protected = runs.protected_runs(project, project.output_dir / "runs", project.output_dir)
    assert {path.name: reasons for path, reasons in protected.items()} == {
        "0001": ["現在採用している実行"],
        "0002": ["完成物として反映した実行"],
        "0003": ["旧版として退避した実行"],
        "0004": ["完成物の反映が未完了の実行"],
        "0005": ["復旧に必要な旧版の実行"],
        "0006": ["確認記録の引き継ぎ元"],
    }
    applied = runs.clean_runs(project, before="0009", apply=True)
    assert applied["removed"] == [str(copy), str(plain)]
    assert all(path.exists() for path in (adopted, committed, previous, exporting, recovery, source))


def test_unreadable_or_unfinished_records_are_protected(tmp_path):
    project = make_project(tmp_path)
    broken = make_run(project, "0001")
    (broken / "manifest.json").write_text("{", encoding="utf-8")
    odd = make_run(project, "0002", ["built"])
    unfinished = make_run(project, "0003", "running")
    result = runs.clean_runs(project, keep=0)
    assert result["candidates"] == []
    reasons = {item["run"]: item["reasons"] for item in result["protected"]}
    assert len(reasons[str(broken)]) == 1 and reasons[str(broken)][0].startswith("生成記録を読めない（")
    assert reasons[str(odd)] == ["生成が終わっていない（status: ['built']）"]
    assert reasons[str(unfinished)] == ["生成が終わっていない（status: running）"]
    # 確認記録や採用記録を読むことができなければ、保護すべき実行が分からないので整理を止める。
    review = make_run(project, "0004") / "review.json"
    review.write_text("{", encoding="utf-8")
    with pytest.raises(records.RecordError):
        runs.clean_runs(project, keep=0, apply=True)
    save(review, {"schema_version": records.SCHEMA_VERSION, "records": []})
    (project.root / "deliverables.json").write_text("{", encoding="utf-8")
    with pytest.raises(records.RecordError):
        runs.clean_runs(project, keep=0, apply=True)
    assert all(path.exists() for path in (broken, odd, unfinished, review.parent))


def test_relative_record_paths_are_read_from_project_directory(tmp_path):
    project = make_project(tmp_path)
    adopted = make_run(project, "0001")
    source = make_run(project, "0002")
    make_run(project, "0003", sources=["out/runs/0002"])
    save(project.root / "deliverables.json", {"schema_version": records.SCHEMA_VERSION, "run": "out/runs/0001"})
    protected = runs.protected_runs(project, project.output_dir / "runs", project.output_dir)
    assert protected[adopted] == ["現在採用している実行"]
    assert protected[source] == ["確認記録の引き継ぎ元"]


def test_new_run_writes_gitignore_only_for_slidekit_output(tmp_path):
    fresh = make_project(tmp_path / "fresh")
    runs.new_run(fresh, command="build", options={})
    assert (fresh.output_dir / ".gitignore").read_text(encoding="utf-8") == (
        "# slidekit が作る実行記録と退避した完成物を追跡しない\n*\n"
    )
    existing = make_project(tmp_path / "existing")
    (existing.output_dir / "runs" / "0001").mkdir(parents=True)
    (existing.output_dir / "exports").mkdir()
    assert runs.new_run(existing, command="build", options={}).name == "0002"
    assert (existing.output_dir / ".gitignore").read_text(encoding="utf-8") == runs.GITIGNORE
    shared = make_project(tmp_path / "shared")
    shared.output_dir.mkdir()
    (shared.output_dir / "memo.txt").write_text("利用者のメモ", encoding="utf-8")
    runs.new_run(shared, command="build", options={})
    assert not (shared.output_dir / ".gitignore").exists()
    custom = make_project(tmp_path / "custom")
    custom.output_dir.mkdir()
    (custom.output_dir / ".gitignore").write_text("runs/\n", encoding="utf-8")
    runs.new_run(custom, command="build", options={})
    assert (custom.output_dir / ".gitignore").read_text(encoding="utf-8") == "runs/\n"


def test_new_run_allocates_next_number_and_updates_marker(tmp_path):
    project = make_project(tmp_path)
    (project.output_dir / "runs" / "0003").mkdir(parents=True)
    (project.output_dir / "runs" / "notes").mkdir()
    run = runs.new_run(project, command="build", options={"draft": True, "slides": "2-3"})
    assert run == project.output_dir / "runs" / "0004"
    marker = json.loads((run / ".building").read_text(encoding="utf-8"))
    keys = ("schema_version", "pid", "host", "project", "command", "options", "stage")
    assert {key: marker[key] for key in keys} == {
        "schema_version": 1,
        "pid": os.getpid(),
        "host": socket.gethostname(),
        "project": str(project.path),
        "command": "build",
        "options": {"draft": True, "slides": "2-3"},
        "stage": "construct",
    }
    assert marker["started_at"] == marker["updated_at"]
    assert {"boot_id", "process_start"} <= set(marker)
    if HAS_PROC:
        assert marker["process_start"] == runs._process_stat(os.getpid())[1]
    runs.update_marker(run, stage="layout")
    updated = json.loads((run / ".building").read_text(encoding="utf-8"))
    assert updated == {**marker, "stage": "layout", "updated_at": updated["updated_at"]}
    assert updated["updated_at"] >= marker["started_at"]
    assert [path.name for path in run.iterdir()] == [".building"]
    runs.remove_marker(run)
    runs.remove_marker(run)
    assert runs.marker_state(run)["state"] == "none"
    with pytest.raises(FileNotFoundError, match="生成中の印がありません"):
        runs.update_marker(run, stage="pptx")
    # JSON にできない値は、実行ディレクトリを作る前に誤りにする（空の実行ディレクトリを残さない）。
    with pytest.raises(TypeError):
        runs.new_run(project, command="build", options={"output": object()})
    assert sorted(path.name for path in (project.output_dir / "runs").iterdir()) == ["0003", "0004", "notes"]


def test_marker_of_dead_process_is_interrupted(tmp_path):
    project = make_project(tmp_path)
    run = runs.new_run(project, command="build", options={})
    runs.update_marker(run, stage="pptx")
    rewrite_marker(run, pid=DEAD_PID)
    state = runs.marker_state(run)
    assert (state["state"], state["pid"], state["stage"]) == ("interrupted", DEAD_PID, "pptx")
    assert str(DEAD_PID) in state["reason"]
    # 中断した実行は保護せず、--failed の候補にする。
    assert runs.clean_runs(project, failed=True)["candidates"][0]["reason"] == "中断した生成"


def test_marker_of_current_process_is_running(tmp_path):
    project = make_project(tmp_path)
    run = runs.new_run(project, command="build", options={})
    marker = json.loads((run / ".building").read_text(encoding="utf-8"))
    state = runs.marker_state(run)
    assert state == {
        "state": "running",
        "pid": os.getpid(),
        "host": socket.gethostname(),
        "started_at": marker["started_at"],
        "stage": "construct",
        "reason": state["reason"],
    }
    assert runs.clean_runs(project, keep=0)["candidates"] == []


@pytest.mark.skipif(not HAS_PROC, reason="/proc でプロセスの開始時刻を読めない環境")
def test_reused_pid_with_other_start_time_is_interrupted(tmp_path):
    project = make_project(tmp_path)
    run = runs.new_run(project, command="build", options={})
    marker = rewrite_marker(run)
    # 同じ番号のプロセスはあるが、開始時刻が違う（番号が再利用された）。
    rewrite_marker(run, process_start=marker["process_start"] + 1)
    state = runs.marker_state(run)
    assert state["state"] == "interrupted" and "再利用" in state["reason"]


def test_old_text_marker_is_unknown_and_protected(tmp_path):
    project = make_project(tmp_path)
    run = project.output_dir / "runs" / "0001"
    run.mkdir(parents=True)
    # 0.1 の印は資料のパスだけの文字列だった。
    (run / ".building").write_text(str(project.path), encoding="utf-8")
    state = runs.marker_state(run)
    assert (state["state"], state["reason"]) == ("unknown", LEGACY_REASON)
    result = runs.clean_runs(project, before="0002", failed=True, apply=True)
    assert result["candidates"] == [] and run.exists()
    assert result["protected"] == [{"run": str(run), "reasons": [f"実行中か判定できない: {LEGACY_REASON}"]}]


def test_other_host_is_unknown(tmp_path):
    project = make_project(tmp_path)
    run = runs.new_run(project, command="build", options={})
    rewrite_marker(run, host="another-host", pid=DEAD_PID)
    state = runs.marker_state(run)
    assert (state["state"], state["host"]) == ("unknown", "another-host")
    assert "another-host" in state["reason"]
    assert runs.clean_runs(project, failed=True)["candidates"] == []


def test_marker_from_other_system_is_not_judged_by_local_processes(tmp_path):
    project = make_project(tmp_path)
    run = runs.new_run(project, command="build", options={})
    original = json.loads((run / ".building").read_text(encoding="utf-8"))

    def state(remove=(), **changes):
        marker = {**original, **changes}
        for key in remove:
            marker.pop(key)
        save(run / ".building", marker)
        return runs.marker_state(run)["state"]

    # WSL と Windows の Python は、端末名が同じでもプロセス番号の空間が違う。
    assert state(pid=DEAD_PID, platform="win32") == "unknown"
    if original["pid_namespace"]:
        # WSL の別のディストリビューションなどは、起動の識別子が同じでも PID 名前空間が違う。
        assert state(pid=DEAD_PID, pid_namespace="pid:[1]") == "unknown"
    if original["boot_id"]:
        # 再起動の後には、生成のプロセスは残っていない。
        assert state(boot_id="00000000-0000-0000-0000-000000000000") == "interrupted"
    # OS と PID 名前空間の項目がない印（仕様の形）も、プロセスの有無で判定する。
    assert state(remove=("platform", "pid_namespace"), pid=DEAD_PID) == "interrupted"
    assert state(remove=("platform", "pid_namespace")) == "running"
    for pid in (0, -1, True, "1", 2**40, None):
        assert state(pid=pid) == "unknown"
    assert state(schema_version=2) == "unknown"


def test_windows_python_keeps_markers_unknown(tmp_path, monkeypatch):
    project = make_project(tmp_path)
    run = runs.new_run(project, command="build", options={})
    rewrite_marker(run, pid=DEAD_PID)
    monkeypatch.setattr(runs, "_WINDOWS", True)
    state = runs.marker_state(run)
    assert state["state"] == "unknown" and "Windows" in state["reason"]
    assert runs.clean_runs(project, failed=True)["candidates"] == []
    # Windows の Python はプロセスの開始時刻を記録しない（この環境で試験できない ctypes の処理を持たない）。
    second = runs.new_run(project, command="build", options={})
    assert json.loads((second / ".building").read_text(encoding="utf-8"))["process_start"] is None


@pytest.mark.skipif(not HAS_PROC, reason="/proc でプロセスの状態を読めない環境")
def test_exited_but_unreaped_process_is_interrupted(tmp_path):
    project = make_project(tmp_path)
    run = runs.new_run(project, command="build", options={})
    child = subprocess.Popen([sys.executable, "-c", ""], stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL)
    try:
        deadline = time.monotonic() + 30
        while runs._process_stat(child.pid)[0] != "Z":
            assert time.monotonic() < deadline, "子プロセスが終了しない"
            time.sleep(0.01)
        rewrite_marker(run, pid=child.pid, process_start=runs._process_stat(child.pid)[1])
        # 番号はまだあり os.kill は成功するが、プロセスは終了している。
        assert runs.marker_state(run)["state"] == "interrupted"
    finally:
        child.wait()


def test_stopped_run_can_be_cleaned(tmp_path):
    project = make_project(tmp_path)
    stopped = make_run(project, "0001", "stopped")
    built = make_run(project, "0002")
    # 0.1 は built・failed 以外の状態を実行中として保護していたので、変換の前に止めた実行を消すことができなかった。
    assert "stopped" in runs.FINISHED_STATUSES
    assert stopped not in runs.protected_runs(project, project.output_dir / "runs", project.output_dir)
    result = runs.clean_runs(project, failed=True, apply=True)
    assert result["removed"] == [str(stopped)]
    assert result["candidates"][0]["reason"] == "変換の前に止めた生成"
    assert not stopped.exists() and built.exists()


def test_clean_apply_reads_each_run_once(tmp_path, monkeypatch):
    project = make_project(tmp_path)
    for number in range(1, 6):
        make_run(project, f"{number:04d}")
    reads = []
    read_manifest = runs._read_manifest

    def counted(run):
        reads.append(run.name)
        return read_manifest(run)

    monkeypatch.setattr(runs, "_read_manifest", counted)
    result = runs.clean_runs(project, before="0006", apply=True)
    assert len(result["removed"]) == 5
    # 削除のたびに全実行の生成記録を読み直さない。
    assert sorted(reads) == ["0001", "0002", "0003", "0004", "0005"]


def test_list_runs_reads_both_record_formats(tmp_path):
    project = make_project(tmp_path)
    make_run(project, "0001", schema=1)
    make_run(project, "0002", mode="draft")
    full = make_run(project, "0003")
    (full / "renders").mkdir()
    (full / "renders" / "deck.pdf").write_bytes(b"%PDF" * 250)
    # 主の PDF のハードリンクは大きさの合計に 1 回だけ含める。
    os.link(full / "renders" / "deck.pdf", full / "deck.pdf")
    running = runs.new_run(project, command="build", options={"draft": True, "slides": "2"})
    (project.output_dir / "runs" / "README.txt").write_text("実行ではないファイル", encoding="utf-8")
    listed = runs.list_runs(project)
    assert [item["run_id"] for item in listed] == ["0001", "0002", "0003", "0004"]
    by_id = {item["run_id"]: item for item in listed}
    # 形式 1 の記録には mode と exportable がない。全ページの通常の生成として示し、完成物にはしない。
    assert (by_id["0001"]["status"], by_id["0001"]["mode"], by_id["0001"]["exportable"]) == ("built", "full", False)
    assert by_id["0001"]["created_at"] == "2026-10-01T00:00:01+00:00"
    assert (by_id["0002"]["mode"], by_id["0002"]["exportable"]) == ("draft", False)
    assert (by_id["0003"]["exportable"], by_id["0003"]["state"]) == (True, "none")
    files = [path for path in full.rglob("*") if path.is_file() and path != full / "deck.pdf"]
    assert by_id["0003"]["bytes"] == sum(path.stat().st_size for path in files)
    marker = json.loads((running / ".building").read_text(encoding="utf-8"))
    assert by_id["0004"] == {
        "run": str(running),
        "run_id": "0004",
        "status": None,
        "mode": None,
        "selection": None,
        "exportable": False,
        "state": "running",
        "bytes": (running / ".building").stat().st_size,
        "created_at": marker["started_at"],
    }
    # deck.toml のパスでも、そのディレクトリでも同じ一覧になる。
    assert runs.list_runs(project.path) == runs.list_runs(project.root) == listed


def test_latest_run_accepts_project_or_deck_path(tmp_path):
    project = make_project(tmp_path, output_dir="build/out")
    assert runs.latest_run(project) is None
    built = make_run(project, "0001")
    make_run(project, "0002", mode="draft")
    make_run(project, "0003", selection=SELECTION)
    make_run(project, "0004", "failed")
    legacy = make_run(project, "0005", schema=1)
    # 生成中の実行は最新にしない。
    runs.new_run(project, command="build", options={})
    for value in (project, project.path, project.root, str(project.path)):
        assert runs.latest_run(value) == built
    assert runs.latest_run(project, exportable_only=False) == legacy
    child = tmp_path / "child"
    child.mkdir()
    (child / "deck.toml").write_text('schema_version = 1\nextends = "../project/deck.toml"\n', encoding="utf-8")
    # 出力先は extends の親から引き継がない。
    assert runs.latest_run(child) is None
