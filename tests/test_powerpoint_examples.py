"""SLIDEKIT_TEST_POWERPOINT=1 を指定したときだけ、同梱の使用例を実際の PowerPoint で生成する。

使用例のディレクトリを、実行保存先（src/out/）と完成物を除いて一時フォルダーへ写し、`pipeline.build()` で生成する。
確かめること：生成が完了して終了コード 0 になり、エラーと、資料のコードが許容を宣言していない警告がないこと。
binary_search では、数式を含むすべての要素（スライド直下の図形）に互換画像が付くこと。
フィクスチャ windows_fonts_environment で Windows の既定フォントを使う。直列に実行する（`-p no:xdist`）。
"""

import json
import os
from pathlib import Path
import shutil
import tomllib

import pytest

pytestmark = [
    pytest.mark.powerpoint,
    pytest.mark.skipif(
        os.environ.get("SLIDEKIT_TEST_POWERPOINT") != "1", reason="実機 PowerPoint を使う任意の検査である。"
    ),
]

EXAMPLES = Path(__file__).resolve().parents[1] / "examples"


def copy_example(name: str, destination: Path) -> Path:
    """使用例を、実行保存先と完成物を除いて写し、写した deck.toml を返す。

    使用例は推奨の構成で、編集元は src/ に、完成物は直下にある。除くパスは src/deck.toml の設定から求める。
    """
    source = EXAMPLES / name
    deck = source / "src" / "deck.toml"
    with deck.open("rb") as stream:
        config = tomllib.load(stream)
    excluded = {(deck.parent / config.get("output_dir", "out")).resolve()}
    excluded |= {(deck.parent / value).resolve() for value in config.get("deliverables", {}).values()}
    shutil.copytree(
        source,
        destination,
        ignore=lambda directory, names: [
            name for name in names if (Path(directory) / name).resolve() in excluded or name == "__pycache__"
        ],
    )
    return destination / "src" / "deck.toml"


@pytest.mark.parametrize("name", ("proposal", "binary_search"))
def test_real_examples_build_without_diagnostics(name, tmp_path, windows_fonts_environment):
    from slidekit import pipeline

    deck = copy_example(name, tmp_path / name)
    result = pipeline.build(deck, carry_from="none")
    assert result["status"] == "built" and result["exit_code"] == 0, result["issues"]
    unresolved = [
        issue
        for issue in result["issues"]
        if issue["severity"] == "error" or (issue["severity"] == "warning" and "declared" not in issue)
    ]
    assert unresolved == []
    run = Path(result["build"])
    record = json.loads((run / "manifest.json").read_text(encoding="utf-8"))
    assert record["status"] == "built" and record["exportable"] is True
    if name != "binary_search":
        return
    scene = json.loads((run / "deck.json").read_text(encoding="utf-8"))
    math_elements = {
        (slide["id"], element["id"])
        for slide in scene["slides"]
        for element in slide["elements"]
        if any(fragment["kind"] == "math" for fragment in element.get("content", ()))
    }
    crops = {(crop["slide_id"], crop["element_id"]) for crop in record["operations"]["math_fallback"]["crops"]}
    assert math_elements and crops == math_elements
