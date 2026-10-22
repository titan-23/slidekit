# 資料作成フローの提案

資料を作る担当者へ、内容の整理から配布までの進め方を説明する、16:9・4 枚の資料を作る使用例である。
二列の図、数式付きのまとまり、表、数式の互換画像を使う。元資料と外部素材は使わない。

直下には完成物と README だけを置き、編集元と作業用のファイルは `src/` に置く（[ディレクトリの構成](../../docs/usage.md#ディレクトリの構成)）。

| ファイル | 内容 |
|---|---|
| `proposal.pptx`・`proposal.pdf`・`notes.md` | 確認を終えて反映した完成物（`slidekit export` の出力） |
| `src/deck.toml`・`src/theme.toml`・`src/rules.toml` | 資料の設定、テーマ、記述規則 |
| `src/brief.md` | 依頼内容（内容確認の証拠に入る） |
| `src/slides.py` | 生成コード |
| `src/deliverables.json` | 採用記録 |

生成と確認は、`pyproject.toml` のあるディレクトリで次のように行う。

```sh
slidekit layout --project examples/proposal/src/deck.toml
slidekit build --project examples/proposal/src/deck.toml
slidekit review status --project examples/proposal/src/deck.toml
```

手順の全体は [作成手順](../../workflows/authoring.md)、コマンドと記録の詳細は [利用ガイド](../../docs/usage.md) にある。
