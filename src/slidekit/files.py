"""ファイルのハッシュ値、正規化した JSON のハッシュ値、原子的な書き出しをまとめる。

slidekit のほかのモジュールを読み込まない。一時ファイルは書き出し先と同じディレクトリに作り、
権限は通常のファイルの作成と同じく umask にしたがう（`tempfile.mkstemp` の 600 にしない）。
"""

from __future__ import annotations

from contextlib import contextmanager
import errno
import hashlib
import json
import math
import os
from pathlib import Path
import secrets
import shutil
from typing import IO, Any, Iterator

_OPEN_FLAGS = (
    os.O_RDWR
    | os.O_CREAT
    | os.O_EXCL
    | getattr(os, "O_NOFOLLOW", 0)
    | getattr(os, "O_BINARY", 0)
    | getattr(os, "O_NOINHERIT", 0)
)


def file_hash(path: str | Path) -> str:
    """ファイルの SHA-256 を、全体を一度に読み込まずに求める。"""
    with Path(path).open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def json_safe(value: Any) -> Any:
    """非有限の数を文字列にし、辞書のキーを文字列にする。不正な座標も診断の根拠として保存できるようにする。"""
    if isinstance(value, float) and not math.isfinite(value):
        return str(value)
    if isinstance(value, dict):
        return {str(key): json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [json_safe(item) for item in value]
    return value


def canonical_json(value: Any) -> str:
    """キーを並べ、空白を除いた JSON。JSON にない値は文字列にする。ハッシュ値と同一性の判定に使う。"""
    return json.dumps(
        json_safe(value), ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str, allow_nan=False
    )


def json_digest(value: Any) -> str:
    """正規化した JSON の SHA-256。"""
    return hashlib.sha256(canonical_json(value).encode("utf-8")).hexdigest()


def _temporary(destination: Path) -> tuple[int, Path]:
    """書き出し先と同じディレクトリに、新しい一時ファイルを作る。"""
    for _ in range(100):
        path = destination.with_name(f".{destination.name}.{secrets.token_hex(4)}.tmp")
        try:
            return os.open(path, _OPEN_FLAGS, 0o666), path
        except FileExistsError:
            continue
    raise FileExistsError(errno.EEXIST, "一時ファイルを作れない", str(destination.parent))


@contextmanager
def atomic_output(path: str | Path, *, text: bool = False) -> Iterator[IO]:
    """一時ファイルへ書かせ、終わったら fsync して path を置き換える。

    途中で例外が起きたときは一時ファイルを消し、path は変えない。text が真なら UTF-8 の文字列で書く。
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary = _temporary(path)
    try:
        with os.fdopen(descriptor, "w" if text else "wb", **({"encoding": "utf-8"} if text else {})) as stream:
            yield stream
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def write_json(path: str | Path, value: Any) -> None:
    """JSON を原子的に書く（キーを並べ、字下げ 2、末尾に改行）。"""
    with atomic_output(path, text=True) as target:
        json.dump(value, target, ensure_ascii=False, indent=2, sort_keys=True)
        target.write("\n")


def stage_copy(source: str | Path, destination: str | Path) -> Path:
    """source を destination と同じディレクトリの一時ファイルへ複製し、fsync してその一時ファイルを返す。

    destination の置き換え（`os.replace`）は呼び出し側が行う。
    """
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary = _temporary(destination)
    try:
        with os.fdopen(descriptor, "wb") as target, Path(source).open("rb") as stream:
            shutil.copyfileobj(stream, target)
            target.flush()
            os.fsync(target.fileno())
    except BaseException:
        temporary.unlink(missing_ok=True)
        raise
    return temporary
