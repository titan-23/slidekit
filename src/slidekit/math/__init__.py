"""編集可能な数式の出力と、数式フォントの MATH 表による寸法計算。

公開する名前は使うときに読み込む（`python -m slidekit.math.commands` が、このパッケージの読み込みで
`commands` を先に読み込まないようにするため）。`MathMeasurer` は `metrics` の計測器である。
"""

__all__ = ["append_math", "math_element", "MathMeasurer", "parse", "MathSyntaxError"]


def __getattr__(name):
    if name in ("append_math", "math_element"):
        from . import office_math

        return getattr(office_math, name)
    if name in ("parse", "MathSyntaxError"):
        from . import syntax

        return getattr(syntax, name)
    if name == "MathMeasurer":
        from .metrics import MathMeasurer

        return MathMeasurer
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
