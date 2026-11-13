# slidekit

Python で内容と配置を書き、編集可能な PowerPoint 資料を生成するライブラリ。
図形・表・数式・画像を扱い、PPTX、PDF、確認用画像をローカルで出力する。
AI は生成コードの作成と修正を担当し、ライブラリ自身は AI を呼び出さない。

## 導入

Python 3.12 以上と Windows 版 PowerPoint が要る。WSL で動作を確かめている。検証した版は [検証記録](validation.md) にある。
既定の書体は Yu Gothic UI、Arial、Cambria Math である。寸法の計算に使うフォントファイルは、Windows のフォントのフォルダー（システムと利用者単位）から自動で探す。

以下は `pyproject.toml` のあるディレクトリで実行する、WSL 向けの手順である。

```sh
uv venv
source .venv/bin/activate
uv pip install -e '.[svg,bibliography]'
```

追加依存の `svg` は SVG の素材、`bibliography` は文献ファイルに使う。数式の計測は必須の依存で行うので、追加依存は要らない。
検証した版にそろえるときは、ロックファイルから導入する。

```sh
uv pip install -r requirements-lock.txt
uv pip install --no-deps -e .
```

別のリポジトリの資料から使うときは、仮想環境を有効にするか、仮想環境の `bin/slidekit` を絶対パスで呼ぶ（例：`/path/to/slidekit/.venv/bin/slidekit build --project src/deck.toml`）。
相対パスの `../slidekit/.venv/bin/python -m slidekit` で起動すると、Python が `sys.prefix` の警告を毎回 2 行出す。
要約の「次の操作」のコマンドは、起動した方法（`slidekit`、絶対パス、`python -m slidekit`）に合わせて書かれるので、そのまま実行できる。

## 使用例を生成する

| 使用例 | 内容 | 主な機能 |
|---|---|---|
| [proposal](examples/proposal/) | 資料作成の進め方（16:9、4 枚） | 二列の図、数式付きのまとまり、表、数式の互換画像 |
| [binary_search](examples/binary_search/) | 二分探索の考え方（4:3、10 枚） | 雛形、表紙と章扉の飾り、分岐図のグループと接続、表、内容から決める寸法、独立した数式、準備済みの図、警告の許容、文献一覧、章ごとの作成、互換画像 |

初回と実行環境を変えたときは、`slidekit doctor --project examples/proposal/src/deck.toml` で PowerPoint とフォントを確かめる。

```sh
slidekit layout --project examples/proposal/src/deck.toml
slidekit build --project examples/proposal/src/deck.toml
slidekit review status --project examples/proposal/src/deck.toml
```

`layout` は表示ソフトを使わずに配置と記述規則を数秒で検査する。`build` の結果は `examples/proposal/src/out/runs/` に実行ごとに保存され、要約に次の操作が出る。
`build` は PowerPoint で変換するので、使用例でも 10 秒前後かかる（proposal 9.6 秒、binary_search 12.8 秒。PowerPoint の起動と終了を含む。[所要時間](docs/usage.md#生成の流れと所要時間)）。

## 資料を書く

資料ごとに `deck.toml` と `slides.py` を用意し、生成関数から `Deck` を返す。資料は任意のディレクトリに置くことができる。
直下に完成物と README、`src/` に編集元と作業用のファイルを置く構成を勧める（[ディレクトリの構成](docs/usage.md#ディレクトリの構成)）。

```python
from slidekit import Deck, Math, Rect, Text
from slidekit.markup import parse


def build(context):
    deck = Deck(context)
    slide = deck.slide(id="intro", title="処理の概要", notes="対象と手順を説明します。")
    slide.text(
        id="premise",
        text=[Text("対象の件数を "), Math("n"), Text(" とする")],
        area=Rect(48, 100, 864, None),
    )
    slide.text(id="cost", text=parse(r"全体の計算量は $O(n \log n)$ である"), area=Rect(48, 160, 864, None))
    return deck
```

座標と文字サイズはポイント単位である。`Rect(48, 100, 864, None)` の高さは内容から決まる。`slide.stack()` を使うと、要素を上から順に間隔を空けて置くことができる。
文字は枠の幅で自動的に折り返す（要素ごとに `word_wrap=False` で無効にできる）。文字の自動縮小は行わず、収まらない箇所を診断する。
数式は PowerPoint で編集でき、PowerPoint 以外のソフト向けの互換画像も PowerPoint から作成できる。
互換画像を付けると変換が 1 回増え、PPTX は数式を含む要素 1 個あたり数 KB〜数十 KB 大きくなる（binary_search で 90 KB から 607 KB。[コスト](docs/usage.md#雛形と数式の互換画像)）。

## 確認と完成物

`build` は前の実行から、変わっていないページの確認を自動で引き継ぐ。`review status` で残りの確認を調べ、画像と本文を開いて確かめてから `review record` で記録する。
`export` は最終検査を含み、確認を終えた実行を指定の保存先へ反映する。機械検査の成功を目で見た確認の代わりにしない。

## 文書

- [利用ガイド](docs/usage.md)：設定、作成 API、生成と確認の手順、実行環境、0.2 への移行、開発
- [設定と引数の一覧](docs/reference.md)：書式の引数、要素を作る関数、`deck.toml`、テーマ、記述規則、実行環境
- [診断の一覧](docs/diagnostics.md)：機械検査の診断の判定と直し方
- [数式の書き方](docs/math.md)：記法と対応する命令
- [資料を作成する手順](workflows/authoring.md)：AI または人が資料を作るときの作業規則
- [検証記録](validation.md)・[未解決の問題](docs/issues.md)
- [開発と資料作成の改善案](docs/improvements.md)

PDF の読み取りと画像化には pypdfium2（PDFium）と pypdf を使う。依存関係とそのライセンスは [開発と依存関係](docs/usage.md#開発と依存関係) にある。

## ライセンスと商標

slidekit 本体は [MIT ライセンス](LICENSE) で公開する。

Microsoft、PowerPoint、Windows は Microsoft Corporation の商標または登録商標である。
