"""入力の誤りを表す例外と、例外の発生場所・失敗の診断。

このモジュールは slidekit のほかのモジュールを読み込まない（model・config などがこのモジュールを読み込むため）。
診断を作る `validators.make_diagnostic` は、関数の中で読み込む。
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Iterable

# slidekit の部品の場所。資料コードのフレームと区別する。
LIBRARY_ROOT = Path(__file__).resolve().parent


def _restore(kind: type, args: tuple) -> BaseException:
    """pickle から戻す。派生クラスの __init__ の引数の形によらず、args と属性をそのまま戻す。"""
    return kind.__new__(kind, *args)


class InputError(ValueError):
    """資料の入力の誤り。ページと要素の識別子と付加情報を持つことができる。

    `code` は診断の識別子として使う（既定 `input_error`）。`ValueError` の派生なので、
    `ValueError` を捕まえる処理でも扱うことができる。子プロセスから戻すとき（pickle）も、属性を保つ。
    """

    code = "input_error"

    def __init__(
        self,
        message: str,
        *,
        code: str | None = None,
        slide_id: str | None = None,
        element_id: str | None = None,
        **details: Any,
    ) -> None:
        super().__init__(message)
        self.code = type(self).code if code is None else code
        if not isinstance(self.code, str) or not self.code:
            raise ValueError(f"InputError の code は空でない文字列にする: {self.code!r}")
        self.slide_id = slide_id
        self.element_id = element_id
        self.details = details

    def __reduce__(self):
        # 派生クラスの __init__ を呼ばずに戻す（引数の形が違う派生でも、文言と属性を変えない）。
        return _restore, (type(self), self.args), dict(self.__dict__)


class ArgumentNameError(InputError, TypeError):
    """受け付けない引数名。`TypeError` としても捕まえることができる。

    `code` は `unknown_argument`（未知の名前）か `unsupported_argument`（その種類が受け付けない名前）。
    """

    code = "unknown_argument"
    CODES = frozenset({"unknown_argument", "unsupported_argument"})

    def __init__(self, message: str, *, code: str | None = None, **kwargs: Any) -> None:
        if code is not None and code not in self.CODES:
            raise ValueError(f"引数名の誤りの code は unknown_argument か unsupported_argument にする: {code}")
        super().__init__(message, code=code, **kwargs)


def _resolved(value: Any) -> Path | None:
    try:
        return Path(value).resolve()
    except (OSError, RuntimeError, TypeError, ValueError):
        return None


def _frame_path(frame: Any) -> Path | None:
    name = frame.filename
    # "<string>" などはファイルではない（カレントディレクトリが資料の場所でも、資料コードとみなさない）。
    if not name or name.startswith("<") and name.endswith(">"):
        return None
    return _resolved(name)


def _within(path: Path, root: Path) -> bool:
    return path == root or root in path.parents


def _short_path(path: Path | None, *, in_project: bool, base: Path | None, library: Path) -> str | None:
    if path is None:
        return None
    if in_project and base is not None:
        if _within(path, base):
            return path.relative_to(base).as_posix()
        try:
            return Path(os.path.relpath(path, base)).as_posix()
        except ValueError:
            return None
    if _within(path, library):
        return (Path(library.name) / path.relative_to(library)).as_posix()
    return None


def exception_location(
    error: BaseException, *, roots: Iterable[Path], library_root: Path | None = None, base: Path | None = None
) -> dict:
    """例外の場所。roots（資料の参照範囲）の中にあり、library_root（既定は slidekit）の外にある最後のフレームを探す。

    戻り値：{"path", "relative_path", "line", "function", "code", "in_project", "raised_at": {"path", "line", "function"}}。
    資料コードのフレームがなければ最後のフレーム（例外が起きた所）を示し、in_project を偽にする。
    構文の誤り（SyntaxError）で、誤りのあるファイルが資料の参照範囲の中にあれば、それを読み込んだ行でなく誤りの行を示す
    （function は None。並行して書く章のモジュールの誤りを、書いた人が見つけやすくする）。
    relative_path は表示用の短いパスで、資料コードなら base（既定は roots の最初）から、slidekit の中なら
    `slidekit/…` の形にする。どちらでもなければ None。
    """
    # traceback は読み込みに時間がかかるので、使うときだけ読み込む（errors はほとんどのモジュールが読み込む）。
    import traceback

    frames = traceback.extract_tb(error.__traceback__) if error.__traceback__ is not None else []
    library = _resolved(library_root if library_root is not None else LIBRARY_ROOT) or LIBRARY_ROOT
    resolved_roots = [root for root in (_resolved(value) for value in roots) if root is not None]
    base_path = _resolved(base) if base is not None else (resolved_roots[0] if resolved_roots else None)
    if isinstance(error, SyntaxError) and error.filename and isinstance(error.lineno, int) and frames:
        path = (
            _resolved(error.filename) if not (error.filename.startswith("<") and error.filename.endswith(">")) else None
        )
        if path is not None and not _within(path, library) and any(_within(path, root) for root in resolved_roots):
            raised = frames[-1]
            raised_path = _frame_path(raised)
            return {
                "path": str(path),
                "relative_path": _short_path(path, in_project=True, base=base_path, library=library),
                "line": error.lineno,
                "function": None,
                "code": (error.text or "").strip() or None,
                "in_project": True,
                "raised_at": {
                    "path": str(raised_path) if raised_path is not None else raised.filename,
                    "line": raised.lineno,
                    "function": raised.name,
                },
            }
    chosen, in_project = None, False
    for frame in reversed(frames):
        path = _frame_path(frame)
        if path is not None and not _within(path, library) and any(_within(path, root) for root in resolved_roots):
            chosen, in_project = frame, True
            break
    if chosen is None:
        if not frames:
            return {
                "path": None,
                "relative_path": None,
                "line": None,
                "function": None,
                "code": None,
                "in_project": False,
                "raised_at": None,
            }
        chosen = frames[-1]
    path = _frame_path(chosen)
    raised = frames[-1]
    raised_path = _frame_path(raised)
    return {
        "path": str(path) if path is not None else chosen.filename,
        "relative_path": _short_path(path, in_project=in_project, base=base_path, library=library),
        "line": chosen.lineno,
        "function": chosen.name,
        "code": chosen.line or None,
        "in_project": in_project,
        "raised_at": {
            "path": str(raised_path) if raised_path is not None else raised.filename,
            "line": raised.lineno,
            "function": raised.name,
        },
    }


def _plain(value: Any, depth: int = 0) -> Any:
    """付加情報の値を、JSON が扱う型に直す（失敗の診断の保存そのものが失敗しないようにする）。"""
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if depth > 20:
        return str(value)
    if isinstance(value, dict):
        return {str(key): _plain(item, depth + 1) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_plain(item, depth + 1) for item in value]
    if isinstance(value, (set, frozenset)):
        return sorted((_plain(item, depth + 1) for item in value), key=str)
    return str(value)


def _message(error: BaseException) -> str:
    try:
        return str(error)
    except Exception:
        return repr(error)


def failure_diagnostic(
    error: BaseException, *, stage: str, project: Any = None, traceback_path: Path | None = None
) -> dict:
    """例外を、場所とトレースバックのファイルを持つ診断にする。

    code は例外の `code`（空でない文字列のとき。なければ build_failed）、scope は stage、slide_id・element_id は
    例外の属性から取る。observations は {"exception", "message", "location", "traceback", ...例外の details}。
    traceback は実行ディレクトリからの相対パスで、実行ディレクトリの直下のファイルならファイル名になる。
    """
    from .validators import make_diagnostic

    code = getattr(error, "code", "build_failed")
    if not isinstance(code, str) or not code:
        code = "build_failed"
    roots: tuple = ()
    base = None
    if project is not None:
        try:
            roots = tuple(project.read_roots)
        except Exception:
            roots = ()
        base = getattr(project, "root", None)
    trace = None
    if traceback_path is not None:
        path = Path(traceback_path)
        trace = path.name if path.is_absolute() else path.as_posix()
    details = getattr(error, "details", None)
    message = _message(error)
    observations = {
        **(_plain(details) if isinstance(details, dict) else {}),
        "exception": type(error).__name__,
        "message": message,
        "location": exception_location(error, roots=roots, base=base),
        "traceback": trace,
    }
    slide_id, element_id = getattr(error, "slide_id", None), getattr(error, "element_id", None)
    return make_diagnostic(
        code,
        message or type(error).__name__,
        scope=stage,
        observations=observations,
        slide_id=slide_id if isinstance(slide_id, str) else None,
        element_id=element_id if isinstance(element_id, str) else None,
    )
