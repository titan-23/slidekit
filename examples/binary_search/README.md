# 二分探索の考え方

二分探索の考え方と比較回数を説明する、4:3・10 枚・説明時間の合計 6 分の資料を作る使用例である。
題材は一般的なアルゴリズムで、文献は公開されている書誌情報だけを使う。
数式を含むすべての要素に互換画像を付ける（`[math] fallback = "required"`）。

## ページと機能

| ページ | 識別子 | 示す機能 |
|---|---|---|
| 1 表紙 | `cover` | 雛形（`template.pptx`）、表紙、題名の下の線、折り返しの無効化（副題の `word_wrap=False`） |
| 2 章扉 | `idea` | ページ種別の飾り（`theme.toml` の `[[page_kinds.section.decorations]]` のひし形 2 個と線） |
| 3 比較による分岐 | `branching` | `parse()` で書いた数式を含む文、高さを内容から決める領域（`Rect(x, y, w, None)`）、まとまり `decision`（ひし形・箱・`anchor(..., gap=6)` の矢印） |
| 4 探す範囲の推移 | `trace` | 見出しも数式の表 `values`（`header_style`・`band_fills`）、計算から作る状態図のまとまり `states`（識別子は `key_id()`）、半開区間の数式 `[0,\,8)`（数字の間のカンマは数の一部になるので、`\,` を挟んで後ろを空ける） |
| 5 繰り返しの間の条件 | `invariant` | `slide.content.columns()`・`distribute()`・`align()`・`slide.measure()` による配置、辺の上の位置を指定した矢印（`anchor("top", at=0, gap=4)`） |
| 6 章扉 | `cost` | 2 ページ目と同じ飾り |
| 7 比較回数の漸化式 | `recurrence` | `slide.stack()` で積む文と式、`equation()` の左揃えと中央揃え（役割 `equation` の上書き）、`\left\lfloor … \right\rfloor`、`\text{回}`、出典の文献キー、辺ごとの余白（`padding_pt` の辞書） |
| 8 比較回数の増え方 | `growth` | SVG の図（`vector=True`、代替テキスト、出典 `figure_script`）、使う場所を限った役割 `chart_label`（ページの `allow_roles`）、凡例のまとまり、※ の注、図に重ねたラベルと警告の許容（`slide.allow_warning()`） |
| 9 線形探索との比較 | `comparison` | 枠線付きで数式を含む表 `methods`（`column_styles`・`row_styles`・`math_role`）、強調（`Emphasis`）、リンク（`Text(..., link=...)`）、`slide.measure()` で寸法を決めた箱 |
| 10 参考文献 | `references` | `.bib` から作る文献一覧（`slide.references()`）、文献のページを数式の候補の検査から外す設定（`math_candidates.exclude_page_kinds`） |

まとまり `decision`・`states` と凡例 `legend` は数式を含まないので、PowerPoint でもグループとして選んで動かすことができる。
数式を含むまとまりはグループにならない（子を個別の図形として書き出す）ので、3・4 ページ目の条件と更新の数式は、まとまりの外に置いている。

`slides.py` はページを章ごとの関数で作り、`context.selected()` で選んだページを含む章だけを作る（`CHAPTERS`）。
`--slides` でページを選んだ検査と生成では、ほかの章の書きかけの誤りに止められずに、自分の章を確かめることができる。

## ファイル

直下には完成物と README だけを置き、編集元と作業用のファイルは `src/` に置く（[ディレクトリの構成](../../docs/usage.md#ディレクトリの構成)）。

| ファイル | 内容 |
|---|---|
| `binary_search.pptx`・`binary_search.pdf`・`speaker_notes.md` | 確認を終えて反映した完成物（`slidekit export` の出力） |
| `src/deck.toml`・`src/theme.toml`・`src/rules.toml` | 資料の設定、テーマ、記述規則 |
| `src/brief.md` | 依頼内容（内容確認の証拠に入る） |
| `src/slides.py` | 生成コード |
| `src/references.bib` | 書誌情報 2 件 |
| `src/make_template.py`・`src/template.pptx` | 雛形を作るスクリプトと、その出力 |
| `src/make_figure.py`・`src/figures/comparisons.svg` | 図を作るスクリプトと、その出力 |
| `src/deliverables.json` | 採用記録 |

作成者・最終更新者・雛形の作成者は、すべて `slidekit example` である。

## 雛形と図を作り直す

雛形と図は同梱している。作り直すときは、`pyproject.toml` のあるディレクトリで次を実行する。
同じ版の python-pptx・fontTools と同じ書体からは、同じ内容のファイルになる（日時などの変わる情報を入れない）。
図の目盛りの数字は DejaVu Sans（`/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf`）の字形の輪郭で描く。

```sh
python examples/binary_search/src/make_template.py
python examples/binary_search/src/make_figure.py
```

雛形は python-pptx に同梱の既定の雛形を基にし、テーマの名前・配色・フォントだけを変えて、見本のページを 1 枚入れている。
slidekit は雛形のテーマ・マスター・レイアウトを引き継ぎ、雛形のページは外す。
図は線形探索の `n` と二分探索の `⌊log2 n⌋ + 1` を式から計算して描き、図の文字は目盛りの数字だけにしている。
目盛りの数字は SVG の文字ではなく輪郭にしているので、表示する環境の書体に左右されず、PDF の検査でもスライドの文字と混ざらない。
軸名と凡例はスライドの文字と線で置き、図に重ねるラベルの位置は `make_figure.plot_point()` で求める。

## 生成と確認

各コマンドは要約を出し、最後に次の操作を示す。機械で読むときは `--json` を付ける。

```sh
# 配置と記述規則だけを検査する（PowerPoint を使わず、ファイルを書かない）
slidekit layout --project examples/binary_search/src/deck.toml
# 自分の章だけを検査する（選んだページを含む章だけを作る）
slidekit layout --project examples/binary_search/src/deck.toml --slides trace
# 下書き（互換画像を省く）で表示を確かめ、直し終えたら通常の生成を行う
slidekit build --draft --project examples/binary_search/src/deck.toml
slidekit build --project examples/binary_search/src/deck.toml
# 残っている確認と、記録のコマンドの形を示す
slidekit review status --project examples/binary_search/src/deck.toml
```

8 ページ目の重なりの警告 2 件は、資料のコードで許容を宣言したもの（`slide.allow_warning()`）である。
診断には残るが、宣言の理由と場所が付き、確認記録なしで最終検査を通る（`review status` と `check --final` の要約に一覧が出る）。
`layout` と `build` の要約では「警告 2 件、うち宣言で許容 2 件」と示し、警告の行の末尾に「（宣言で許容）」を付ける。
ほかのエラー・警告はない。

画像を開いて確かめてから、確かめた範囲だけを `slidekit review record` で記録し、`review status` が完了になったら `slidekit export` で直下の完成物へ反映する。
数式に互換画像を付けるので、標準の表示に加えて、互換画像を使う表示（`--variant fallback`）の確認も要る。
`build` は前の実行から確認を自動で引き継ぐ。引き継ぎが要らないときは `--carry-from none` を付ける。
手順の全体は [作成手順](../../workflows/authoring.md)、コマンドと記録の詳細は [利用ガイド](../../docs/usage.md) にある。
