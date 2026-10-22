"""ファイルのハッシュ値、正規化した JSON のハッシュ値、原子的な書き出しを確かめる。"""

from contextlib import contextmanager
from datetime import date
import hashlib
import json
import os

import pytest

from slidekit import files
from slidekit.files import atomic_output, canonical_json, file_hash, json_digest, stage_copy, write_json


@contextmanager
def umask(value):
    previous = os.umask(value)
    try:
        yield
    finally:
        os.umask(previous)


def test_json_digest_matches_previous_formulas():
    value = {"slide": {"id": "intro", "title": "題名", "area": [1.5, 2, 3]}, "rules": {"numbers": True}, "none": None}
    # 生成記録と確認記録のハッシュ値に使っていた式。固定値は files.py にまとめる前のコードで求めた。
    previous = hashlib.sha256(
        json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode()
    ).hexdigest()
    assert json_digest(value) == previous == "e528ed6ccba5ac6143cfa58701dc6d8010717e430085a9f5ed3f3dbcbe62f24e"
    # 診断のハッシュ値と同じく、非有限の数と JSON にない値は文字列にする。
    assert (
        canonical_json({"x": float("nan"), "day": date(2026, 1, 2), "pair": (1, 2)})
        == '{"day":"2026-01-02","pair":[1,2],"x":"nan"}'
    )


def test_file_hash_matches_sha256_of_contents(tmp_path):
    data = b"slidekit" * 400_000
    path = tmp_path / "large.bin"
    path.write_bytes(data)
    assert file_hash(path) == file_hash(str(path)) == hashlib.sha256(data).hexdigest()


def test_write_json_is_synced_replaced_and_follows_umask(tmp_path, monkeypatch):
    synced = []
    fsync = os.fsync
    monkeypatch.setattr(files.os, "fsync", lambda descriptor: (synced.append(descriptor), fsync(descriptor)))
    path = tmp_path / "run" / "manifest.json"
    value = {"題名": "例", "values": [1, 2.5], "none": None}
    with umask(0o027):
        write_json(path, value)
        write_json(path, {**value, "values": [3]})
    assert (
        path.read_text(encoding="utf-8")
        == json.dumps({**value, "values": [3]}, ensure_ascii=False, indent=2, sort_keys=True) + "\n"
    )
    assert len(synced) == 2
    assert [item.name for item in path.parent.iterdir()] == ["manifest.json"]
    if os.name != "nt":
        assert path.stat().st_mode & 0o777 == 0o640


def test_atomic_output_keeps_previous_file_when_writing_fails(tmp_path):
    path = tmp_path / "deck.pptx"
    path.write_bytes(b"previous")
    with pytest.raises(RuntimeError, match="途中"):
        with atomic_output(path) as stream:
            stream.write(b"partial")
            raise RuntimeError("途中で失敗")
    assert path.read_bytes() == b"previous"
    assert [item.name for item in tmp_path.iterdir()] == ["deck.pptx"]


def test_stage_copy_keeps_destination_until_caller_replaces_it(tmp_path):
    source = tmp_path / "run" / "deck.pptx"
    source.parent.mkdir()
    source.write_bytes(b"new version")
    destination = tmp_path / "project" / "presentation.pptx"
    destination.parent.mkdir()
    destination.write_bytes(b"old version")
    with umask(0o027):
        staged = stage_copy(source, destination)
    assert staged.parent == destination.parent and staged != destination
    assert staged.read_bytes() == b"new version"
    assert destination.read_bytes() == b"old version"
    if os.name != "nt":
        assert staged.stat().st_mode & 0o777 == 0o640
    os.replace(staged, destination)
    assert destination.read_bytes() == b"new version"
    with pytest.raises(FileNotFoundError):
        stage_copy(tmp_path / "missing.pptx", destination)
    assert [item.name for item in destination.parent.iterdir()] == ["presentation.pptx"]
