# 利用ガイド

slidekit で資料を書き、生成し、確認して完成物へ反映するまでの手順と考え方をまとめる。
設定と引数の全項目は [設定と引数の一覧](reference.md)、診断は [診断の一覧](diagnostics.md)、数式の記法は [数式の書き方](math.md) にある。
コマンド例は `pyproject.toml` のあるディレクトリで、仮想環境を有効にして実行する。

## 資料の設定

資料は任意のディレクトリに置くことができる。`deck.toml` と、生成関数 `build(context)` を定義した `slides.py` を用意する。最小の例は次のとおりである。

```toml
schema_version = 1
size = "16:9"

[metadata]
title = "資料の題名"
author = "作成者名"
last_modified_by = "編集者名"
```

`title`・`author`・`last_modified_by` は必須である。匿名の場合は名前を空文字列で明示する。
相対パスはその設定を書いた `deck.toml` の場所を基準にする。主な設定は次のとおりで、全項目は [資料の設定](reference.md#資料の設定decktoml) にある。

| 設定 | 用途 |
|---|---|
| `entrypoint` | 生成関数（既定 `slides.py:build`） |
| `theme`・`rules` | テーマと記述規則の TOML。同梱の `default`・`general` に重ねる |
| `inputs.read_roots` | 作成 API が読む入力の範囲。既定は資料のディレクトリ |
| `inputs.documents`・`inputs.bibliography`・`inputs.asset_dirs` | 元資料、文献ファイル、素材のディレクトリ |

| `math.fallback` | 数式の互換画像を付けるか |
| `deliverables` | 完成物（PPTX、PDF、発表者ノート、採用記録）の保存先 |
| `template` | 雛形の PPTX |

生成関数から読み込んだ補助の Python ファイルは、参照範囲の中にあれば入力として自動で追跡する。
`inputs.python` は、生成関数から読み込まれないファイルを追跡したいときだけ書く。
作成の目的・対象者・必須事項は任意の `brief.md` に書くことができる。資料のディレクトリ、継承元のディレクトリの順に探し、一覧画像の確認の証拠に入れる。

設定の誤り（未知の設定名、不正な値、参照できない役割や色）は、すべてを集めて 1 つのエラーにし、ファイル・設定の位置・候補を 1 件 1 行で示す。
設定ファイルの読み込みに失敗したとき（構文の誤り、`extends` で継承元をたどる際の誤り）は、その問題を示し、読み込んだ設定だけでは決まらない参照の検査を行わない（誤検出した問題で、読み込みに失敗した原因を隠さないため）。
Python のコード自体は隔離しない。信頼できる編集元だけを実行する。

### ディレクトリの構成

資料のディレクトリの直下には完成物（PPTX、PDF、発表者ノート）と README だけを置き、編集元と作業用のファイルは `src/` にまとめることを勧める。
ディレクトリを開いたときに完成物がすぐ見つかり、生成のたびに増えるファイルと混ざらない。同梱の使用例もこの構成である。

```text
seminar/
├── README.md             資料の説明と作り直し方
├── seminar.pptx          完成物
├── seminar.pdf
├── speaker_notes.md
└── src/
    ├── deck.toml         完成物の保存先は ../ にする
    ├── slides.py         生成関数（補助の Python ファイルも同じ場所に置く）
    ├── theme.toml
    ├── rules.toml
    ├── brief.md
    ├── figures/          素材と、素材を作るスクリプト
    ├── deliverables.json 採用記録
    └── out/              実行ディレクトリ（build が .gitignore を置く）
```

`deck.toml` では、完成物を一つ上のディレクトリに保存する。

```toml
[deliverables]
pptx = "../seminar.pptx"
pdf = "../seminar.pdf"
notes = "../speaker_notes.md"
record = "deliverables.json"
```

コマンドには `--project seminar/src/deck.toml` のように `src/` の `deck.toml` を渡す。
採用記録のパスは `deck.toml` のあるディレクトリからの相対パスなので、完成物は `../seminar.pptx` のように記録される。

### 設定の継承

`extends` で別の `deck.toml` を継承できる。同じ資料の写しを別の実行保存先で生成するときなどに使う。

```toml
schema_version = 1
extends = "../original/deck.toml"
output_dir = "out"
```

- 表は再帰的に合成し、値と配列は継承する側の値で置き換える。
- `output_dir` と `deliverables` は引き継がない（実行保存先と完成物を資料どうしで共有しないため）。
- 継承元に書いたパスは継承元の場所を基準にする。継承元の `deck.toml` のディレクトリは参照範囲に加わる。
- 継承は 8 段までで、循環はエラーになる。

## 基本 API

`Deck(context).slide(id=..., kind=..., title=..., notes=...)` でページを作り、生成関数は `Deck` を返す。
`kind` は `cover`（表紙）、`section`（章扉）、`content`（既定）、`references`（文献）、`blank` から選ぶ。
`slide.content` はページ種別の本文領域で、`columns()` と `rows()` で分けることができる。

| 関数 | 用途 |
|---|---|
| `text` | 文章。文字列か `Text`・`Emphasis`・`Math` の断片の列 |
| `box` | 図形（`rectangle`・`rounded_rectangle`・`ellipse`・`diamond`）。文章を置くことができる。書式 `tail_side` で吹き出しにする（[吹き出し](#吹き出し)） |
| `equation` | 編集可能な Office Math の数式。数式だけなら独立した数式になる |
| `line`・`arrow` | 座標か要素の `anchor()` を結ぶ線と矢印 |
| `image` | PNG・JPEG・SVG・PDF の画像。PDF には 1 始まりの `page` を指定する |
| `table` | 図形で組む表（[表](#表)） |
| `group` | まとまり。条件を満たせば PowerPoint のグループになる（[まとまりと識別子](#まとまりと識別子)） |
| `references` | 文献一覧（[文献一覧](#文献一覧)） |
| `stack`・`measure` | 寸法と配置を決める補助（[寸法と配置を決める](#寸法と配置を決める)） |
| `allow_warning` | 意図した警告の宣言（[警告の許容](#警告の許容)） |

座標と文字サイズはポイント単位である。各関数の引数は [要素を作る関数](reference.md#要素を作る関数) にある。

ページと要素には、編集後も意味が変わらない識別子（`id`）を付ける。識別子は空でない文字列で、`/` と制御文字を含めず、`_theme` で始めない。
ページの中と表の中で重複できない。繰り返し作る要素は番号ではなく内容のキーから `key_id()` で識別子を作る（並べ替えても同じ要素を指す）。

```python
from slidekit import key_id

for key, label in steps.items():
    slide.box(id=key_id("step", key), area=..., text=label)   # 例: step_prepare
```

文字の用途は `role` で指定する。`slide_title` はページの題名、`heading` は本文の小見出し、`body` は本文、`note` は注記である。
小さい文字などの例外にはテーマで最小値の低い役割を定義し、`restricted = true` で使う場所を限る。
使うページは `deck.slide(..., allow_roles=["chart_label"])`、要素は `allow_roles=["chart_label"]` で宣言する。
まとまりや表に付けた宣言は、その子にも及ぶ。テーマの `allowed_slide_ids` でページを列挙することもできる。
宣言のない場所で使うと、ページと要素の識別子と設定の場所を示すエラーになる。要素から最小文字サイズを下げることはできない。

入力の誤りは `slidekit.errors.InputError`（`ValueError` の派生）になり、`slide_id`・`element_id`・`code`（診断の識別子）を持つ。
受け付けない引数名は `ArgumentNameError`（`TypeError` の派生でもある）で、近い名前を候補として示す。
`layout`・`doctor`・`build` の構築では、続行できる誤り（書式の名前と値、役割、断片の書式とリンク、出典）を例外にせずに集め、まとめて診断として示す。
集めるかどうかは `context.collect_errors`（既定は偽）で決まり、これらのコマンドは生成関数を呼ぶ前に真にする。
生成関数の外で Python から直接 `Deck` を使うとき（`slidekit.config.load_project()` で読んだ資料など）は偽のままなので、最初の誤りで `InputError` になる。

## 書式と優先順位

書式は次の順に重ね、後ろほど優先する。

1. ライブラリの既定
2. テーマの `[defaults]`
3. テーマの `[defaults.<種類>]`（表の本文のセルは `box` → `table_cell`、見出しのセルは `box` → `table_cell` → `table_header`、矢印は `line` → `arrow` の順）
4. 役割
5. 要素の `style=` の辞書
6. 要素のキーワード引数

`min_size_pt` は常に役割の値になる。色はテーマの `[colors]` の名前か 6 桁の 16 進数で書く。

```toml
# theme.toml：図形の既定の塗りと枠線を替える
[defaults.box]
fill = "background"
line_color = "primary"
```

```python
callout = {"fill": "surface", "line_color": "accent", "padding_pt": {"left": 12}}
slide.box(id="note", area=Rect(60, 400, 400, None), text="補足", style=callout, bold=True)
```

値の `None` は「指定しない」で、下の層の値を使う。ただし、文字・図形・数式・表のセルの `fill`・`line_color` の `None` は `"none"`（塗らない・枠線なし）と同じである。
`padding_pt` は数値か、`left`・`top`・`right`・`bottom` の一部を持つ辞書で書く。辞書は指定した辺だけを上書きし、ほかの辺は下の層の値を引き継ぐ。
要素の種類が受け付けない書式名、未知の名前（`z_index` など）はエラーになる。種類ごとに受け付ける書式は [書式の引数](reference.md#書式の引数) にある。

## 文字の配置と折り返し

文字は枠の幅で自動的に折り返す（書式 `word_wrap`、既定 `true`）。文字サイズは自動で縮小せず、収まらない箇所を診断する（`text_overflow` など）。
1 行で収めたい要素は `word_wrap=False` にし、資料全体で折り返さないなら、テーマの `[defaults]` に `word_wrap = false` を書く。
`equation()` は同梱のテーマで折り返さない（`[defaults.equation] word_wrap = false`）。PowerPoint が枠より長い独立した数式を枠の幅で改行し、計測と合わなくなるためである。資料のテーマで `word_wrap = true` にすれば折り返す。

改行 `\n` は意味の区切りで行を分けたい箇所だけに書く。行数が変わる要素の高さは、`Rect(x, y, w, None)` か `slide.stack()` で内容から決める。
明示改行の前後に、行頭や行末に置かない文字（行頭の「、」など）があると警告する（`line_break`）。

折り返しの位置は PowerPoint の配置に合わせて計算する（PowerPoint 16.0 の build 20430 で確かめた）。

- 禁則は追い出しだけで処理し、追い込みとぶら下げは行わない。行頭に置かない文字には小書きの仮名と長音記号を含め、同じ一覧を PPTX の `p:kinsoku` に書く。
- 全角空白は半角空白と同じに扱い、前で分けず、行末では幅に含めない。
- 枠の内側より長い欧文の語は、PowerPoint と同じく語の途中で改行し、エラー `text_overflow` で知らせる。
- 文中の数式は、行の残りに入らないときも式の途中で改行せず、式ごと次の行へ送る。枠より長い式は右へはみ出す（`text_overflow`）。
- 12 pt 以上の文字は欧文の書体の kern 表で字間を詰める。断片の境界でも、書体・大きさ・太さが同じで書体の区切りをまたがなければ詰める。
- 折り返しの位置が幅の限界に近い要素は、参考の診断 `text_wrap_near_limit` で知らせる。行数は生成後の PDF の検査で確かめる。
- 図形の文字領域は形で決まる。ひし形は外形の中央の幅・高さとも半分、楕円は約 0.71 倍、角丸の長方形は角の半径の約 0.29 倍だけ内側で、その中で折り返す。吹き出しの文字領域は本体の形で決まり、しっぽの寸法は計測に影響しない。
- 自動で折り返した段落の最後の行が短い要素（`text_short_line`）、和文の文節の途中で折り返した要素（`text_wrap_mid_phrase`）、欧文の語のハイフンの後で折り返した要素（`text_wrap_hyphen`）は、参考の診断で知らせる。文節の境は BudouX の日本語のモデルで判定する。行の文字列は `slidekit layout --wraps` で見る。
- 書式 `phrase_wrap=True`（テーマの `[defaults]` や役割にも書くことができる）は、和文の文字の間を文節の境だけで折り返す。内側の幅より長い文節の中は今までどおり折り返す。PowerPoint は和文の文字の間ならどこでも折り返すので、場面を作るときに、計測した折り返しの位置へ明示改行を入れて書き出す（要素の枠の幅で折り返しを固定する。PPTX を編集して幅を変えると、明示改行の位置は変わらない）。
- 数式の描画が送り幅より左右へ張り出す分（入れ子の分数の線など）も、インクの範囲として、はみ出しと余白への入り込みの判定に使う。文字の送り（行の幅と揃え）は送り幅で決める。

書体は文字ごとに選ぶ。和文の文字は `font_japanese`、それ以外は `font_latin` で書き、欧文の書体にない記号は和文の書体で書く。
幅の曖昧な記号は、PowerPoint が既定で描く書体と同じ書体で計測して書き出す。
○ × □ ● ± “ ” ‘ ’ ° § ♪ ① などの PowerPoint が既定で和文の書体で描く記号は和文の書体（和文の書体に字形がなければ欧文の書体）、→ … — – などは欧文の書体である。
どちらの書体にも字形がない文字はエラー（`font_glyph_missing`）になる。

リンク（断片の `link`）は文字の色で表示する（PowerPoint の `hlinkClr` の指定。Office 2019 以降）。PowerPoint ではリンクに常に下線が付き、`underline=False` でも消えない。

書き出しは雛形の既定に左右されないように、次を明示する。

- 文字枠：折り返すときは `wrap="square"`、折り返さないときは `wrap="none"`。文字の自動の縮小と調整はしない（`a:noAutofit`）。
- 段落：`eaLnBrk="1"`・`latinLnBrk="0"`・`hangingPunct="0"`。通常の文字に `kern="1200"`（12 pt 以上で詰める）。
- 書体：和文と欧文の書体の区切りごとに書く。欧文の書体にない文字と幅の曖昧な記号の区切りは、`a:latin` と `a:ea` の両方を計測した書体にする。
- 明示改行の `a:br` と段落の末尾（`a:endParaRPr`）には、その位置の断片の書式を書く（書かないと PowerPoint は空の行を雛形の既定の大きさで置き、段落の最後のリンクを PDF のリンクにしない）。

### 行の配置

行の位置も PowerPoint の配置に合わせて計算する。計測の方式を変えたときは `measure.MEASURE_VERSION` を上げる（今は 9）。版は生成記録と診断の証拠に入る。
以下で s は文字サイズ、sp は行間である。行間 `line_spacing` は文字サイズの 1.2 倍を 1 とする倍率で、`line_spacing_pt` は 1.2 × s で割った行間として扱う。

- 行の書体の組は欧文の書体と、行の文字を計測した書体である。和文の文字を含む行には和文の書体も加え、空の行と半角空白だけの行は欧文の書体だけにする。
  書体ごとに A_f = 1.2 × usWinAscent ÷ (usWinAscent + usWinDescent)、D_f = 1.2 − A_f を求め（OS/2 表の値）、組の上の割合を A = 1.2 × A* ÷ (A* + D*)、下の割合を D = 1.2 − A とする（A*・D* は組の A_f・D_f の最大）。
- 行の上側は、行間がちょうど 1 なら A × s、1 より大きければ sp × min(A, 0.9035) × s、1 未満なら max(A − 1.2 × (1 − sp), sp × min(A, 0.9035)) × s である。行送りは前の行の下側（1.2 × sp × s − 上側）と次の行の上側の和で、同じ組の行どうしは 1.2 × sp × s になる。
- 文字のまとまりの下端は最後の基線から D × s で、行間によらない。
- 数式は、数式の寸法から決めた上下で行を広げる。数式の計測の上側を L とする。行間がちょうど 1 の行の上側は、インクの上 + 0.147 em 以上にする。
- 分数（横書きの `/` を除く）・`\binom`・伸びる括弧・大型演算子・下に置くものを含む背の高い式（インクの上下の和が文字サイズを超える式）は、a = インクの上 + 0.147 em（行の組の行間 1.0 の上側を下限にする）、d = インクの下、h = a + d とする。
  インクの上下の和が 1.3 em 以上の式と、下に置くもの（`\underset`・`\limits` の下）を含む式は、行間で増える分 e = (sp − 1) × h が 1.17 × d 以下ならすべてを行の上に、超えれば 4 分の 3 を上に置き、残りを下に置く。
  行間が 1 より大きいときは、上側を全高 sp × h の 4 分の 3 以下に抑え、減らした分を下に回す。分子側の入れ子などで上に偏る式も、行全体の高さは保つ。
  それより低い式は、上側を 1.2 × sp × インクの上、下側を d にする。どちらも上側は sp × L より小さくしない。文字の段落の最後の行の下側は d で、行間を掛けない。
  行間 1 未満の背の高い式は確かめていないので、上側を max(sp × L, 1.2 × sp × インクの上)、下側を行間 1.0 の下側 × sp にする。
- 文中の背の高い根号は行間 1 以上で次のように置く。当てはまるのは、根号を含み、分数などを含まず、根号の外に上へ出る部分（上付きの添字、上に置くもの、アクセント、上線）がなく、
  インクの上下の和が文字サイズを超える式である（根号の中の添字などは問わない）。行の上側は、行間がちょうど 1 なら T1 = max(L, 1.22 × インクの上)、
  1 より大きければ max(sp × L, 1.17 × sp × インクの上) にする。次の行への行送りに使う下側は sp × (T1 + d) − 上側、最後の行の下側は d である。
  独立した数式の根号と、行間 1 未満の行には当てない（ほかの数式と同じに置く）。
- 上の規則に当たらない数式の行間 1.0 の上側 T = max(L, インクの上 + 0.147 em) が文字の上側 A × s より高いときは、
  H = T + max(数式の計測の下側, D × s) を行の全高とする。行間が 1 より大きい文字の段落では、sp × H の 4 分の 3 を基線の上、残りを下に置く。
  独立した数式だけの行と、これより低い式では、数式の計測の上下に行間を掛ける。
- 行間が 1 より大きい文字の段落では、分数などの専用の規則に当たらない数式と文中の背の高い根号の下 d が組の下 D × s より深いとき、文字から求めた上側に 0.5 × sp × (d − D × s) を加える。行の上側はその値と数式から求めた上側の大きい方にし、最後の行の下側は d にする（行間 1.2 で確かめた）。
- 独立した数式だけの行は「数式だけの行」として置き、文中の数式だけの段落は欧文の書体の文字の行として置く。
- 行のインクの下端は、基線より上だけの字形（× や “ など）の行では負になる（生成記録の `ink_bottom_pt` も同じ）。
- PDF の読み取りでは、枠の外に描かれた数式の字形を、最も近い枠の要素に割り当てる。

確かめた範囲は、PowerPoint 16.0（build 20430）の校正用の資料の 18〜28 pt、行間 1.0〜2.0 である。行間 1 未満は 0.9 だけを確かめた。
上付きに横分数・斜線・上下の分数がある行は、12・20・32 pt、行間 1.0・1.15・1.2・1.5、欧文・和文・複数行と縦の揃えを変えて確かめた。
分子側が入れ子の分数の行は、12・20・32 pt、行間 1.0〜2.0、文中・独立した数式と縦の揃え・複数行を確かめた。
行間で増える分の分け方の境（e ÷ d が 1.17 の近く）は確かめていない。
分母側が高い入れ子分数では、行間 2.0 で基線が合わないケースがある（[未解決の問題](issues.md)）。

## 寸法と配置を決める

`Rect(x, y, w, h)` の幅か高さを `None` にすると、内容から決める。

- `Rect(x, y, w, None)`：幅 `w` で折り返した文字の高さにする。
- `Rect(x, y, None, h)`：折り返さない文字の幅にする。画像は縦横比から決める。
- `Rect(x, y, None, None)`：文字・図形・数式は折り返さない内容の幅と高さにする。画像では幅か高さの一方を数値で指定する（両方 `None` は `invalid_geometry` のエラー）。

文字のない要素（空の文字、文字のない図形）は幅を内容から決めることができないので、幅を数値で指定する（`invalid_geometry`）。高さは `None` でよい。

内容から決めた寸法は要素を作った後の `element.area` で数値になる。`None` のままの領域では `right`・`bottom`・`center`・`anchor()` などを求めることができない。

`slide.stack(area, gap=12)` は、要素を領域の上から順に間隔を空けて置く。まとまりの中でも `group.stack()` を使うことができる。

```python
stack = slide.stack(slide.content, gap=12)
stack.text(id="lead", text="二分探索は、探す範囲を半分ずつ狭める")
stack.equation(id="bound", expression=r"\lfloor \log_2 n \rfloor + 1", place="center", width=300)
stack.space(24)
area = stack.take(120)          # 高さ 120 pt の領域を受け取り、自分で使う
```

- `text`・`box`・`equation`・`image`・`table` は元の関数と同じ引数をとり、`area` の代わりに `height`・`width`・`indent`・`place`・`gap` を受け取る。
  幅の既定は領域の幅から `indent` を除いた幅、高さの既定は内容から決める。`place` は幅が領域より狭いときの横位置（`left`・`center`・`right`）で、文字の揃え `align` とは別である。`gap` はその要素の後の間隔である。
- `take(height)` は領域を返して進み、`space(amount)` は位置を下げる。`cursor`（次の上端）、`remaining`（残りの領域）、`elements`（置いた要素）を読むことができる。
- 例外になるのは幅が領域を超えたときだけである。領域の高さが数値なら、高さを超えた要素や `take` の領域は配置の検査が `layout_bounds_overflow` で報告する。
  その診断は、同じ積み上げの領域に置いた要素の全体で足りない量（`stack_overflow_pt`）と各要素の高さ（`stack_elements`）も示す。どの要素をどれだけ詰めればよいかを、1 回の検査で判断できる。
- `slide.stack(area, gap=12, avoid="footer")` は、ページ番号の領域を避けて積む。`avoid` には `Rect` か `"footer"`、またはその列を書く。幅を指定しない要素が避ける領域に重なるときは、領域の左に収まるように幅を縮めて作り直す（左に領域の幅の半分以上が残らなければ、領域の下へ送る）。幅を指定した要素は動かさない。
  ページ番号の領域は `slide.footer_area`（番号を置かないページは `None`）で読むことができる。

`slide.measure(text, kind="text", width=None, role=None, shape="rectangle", style=None, **書式)` は、要素を作るときと同じ書式の解決で内容を計測し、
`measure.TextLayout` を返す（要素は作らない）。`layout.height` は `Rect(x, y, w, None)` で決まる高さ、`layout.width` は枠の幅（`width` を省くと `Rect(x, y, None, h)` で決まる幅）である。
`kind` は `text`・`box`・`equation`、`width` は枠の幅（`None` なら折り返さない）である。使う場所を限った役割も宣言なしで計測できる。
`layout.lines` の各行は、文字列（`text`。数式は `$…$`）、幅、行末の自動の折り返しの種類（`wrap`）を持つ。`layout.natural_width` は折り返さないときの最も広い行の幅で、1 行に収めるのに要る内側の幅である。
`element.measure()` は作った要素を今の書式と枠で計測する（配置の検査と同じ値）。

同じ計測はコマンドでも行うことができる。資料を作り直さずに、文が 1 行に収まるかや、どこで折り返すかを確かめるのに使う。

```sh
slidekit measure --project deck.toml --width content "比較のたびに探す範囲を半分にする" "短い文"
slidekit measure --project deck.toml --width 618 --role note --markup '書誌 $O(\log n)$ の行'
```

要約は文字列ごとに行数・枠の高さ・各行の文字列と幅を示し、折り返したときは 1 行に収める枠の幅を添える。`--width content` はページ種別（`--page-kind`、既定 `content`）の本文領域の幅である。
`context.measure_text()`・`context.measure_math()` は互換用の低水準の関数である。`slide.measure()` を勧める。

領域の計算には次を使う。

- `Rect` の `inset()`・`columns()`・`rows()`・`contain()`・`anchor()`・`translated()` と、`right`・`bottom`・`center`。
- `align(rects, area, horizontal=..., vertical=...)`：各領域を辺か中央へそろえる。`distribute(rects, area, axis="horizontal", gap=None)`：寸法を変えずに等間隔に並べる。どちらも `from slidekit import align, distribute` で読み込む。
- 要素の `move(dx, dy)` と `area` への代入。まとまりを動かすと子も動く。

表の `column_widths="auto"`・`row_heights="auto"` はセルごとに解決した書式で内容から列幅と行高を決める（[表](#表)）。

## 表

`slide.table()` はセルを図形として並べた編集可能な表を作り、`Table` を返す。

```python
table = slide.table(
    id="methods",
    rows=[["線形探索", Math("n")], ["二分探索", Math(r"\lfloor \log_2 n \rfloor + 1")]],
    row_keys=["linear", "binary"],
    column_keys=["method", "count"],
    headers=["方法", "比較回数"],
    area=Rect(60, 120, None, None),
    column_widths="auto",
    row_heights="auto",
    header_style={"bold": True},
    band_fills=["background", "surface"],
    column_styles={"count": {"align": "center"}},
)
table.cell("binary", "count")           # セルの要素。見出しの行キーは "_header"
```

- `rows`・`row_keys`・`column_keys` は必須で、キーは識別子の規則に従う。`headers` を渡すと見出し行を作る。
  `header` は互換用の真偽値で、`False` なら `headers` があっても見出し行を作らず、`True` なら `headers` を必須にする。省くと `headers` の有無で決まる。
  `table_format` は `"shapes"`（セルを図形として並べる。既定）だけを受け付ける。
- `column_widths` は `None`（等分）、`"auto"`（内容から）、数値（全列）、数値の列のいずれかである。`row_heights` も同じで、数値の列は見出し行を含む。`header_height` は見出し行だけの高さで、`row_heights` より優先する。
  領域の幅・高さが `None` なら、`None` の列幅・行高を `"auto"` として扱う。列幅には折り返しの誤差の帯を含め、列幅が決まれば行高は折り返しを反映する。
- セルの役割は、`cell_roles["行キー/列キー"]`、見出し行なら `header_role`（既定 `heading`）、内容が数式だけなら `math_role`、`column_roles[列キー]`、`role`、ページの既定の役割の順に選ぶ。
- セルの書式は役割の上に、見出しのセルでは表全体 < `column_styles` < `header_style` < `cell_styles`、本文のセルでは表全体 < `band_fills` < `column_styles` < `row_styles` < `cell_styles` の順に重ねる。
  表全体の書式は `style=` とキーワード引数で書く。`cell_styles`・`cell_roles` のキーは `"行キー/列キー"`（見出しは `"_header/列キー"`）である。存在しないキーはエラーになる。
- セルの識別子は「表の識別子/行キー/列キー」である。戻り値では `area` が実際の表の範囲で、`row_keys`・`column_keys`・`column_widths`・`row_heights` を読むことができる。表に要素は追加できない。
- 表は数式を含まなければ PowerPoint のグループになる（`pptx_group`）。`sources`・`description`・`allow_roles` は表のまとまりに付く。

## 線と接続

`line()`・`arrow()` の端点は座標 `(x, y)` か要素の `anchor(side, gap=0, at=0.5)` で指定する。
`side` は `left`・`right`・`top`・`bottom`・`center`、`gap` は辺から外側への距離（pt）、`at` は辺の上の位置（上下の辺は左端 0〜右端 1、左右の辺は上端 0〜下端 1）である。
回転・反転した要素では辺を要素自身の向きで決め、画像は表示される範囲を基準にする。端点は、要素を動かすと追従する。

`anchor()` で文字枠・図形・表示範囲が配置領域と一致する画像の辺に結んだ端は、PowerPoint の接続になる（`center` には接続点がないので接続しない）。
PowerPoint で接続先を動かすと、両端を接続した線は両端とも辺の中央の接続点へ移り、`gap` と `at` の位置はなくなる。
片方だけを接続した線は、線全体が接続先と同じだけ平行に動き、`gap` と `at` の位置は保たれる（PowerPoint 16.0 の build 20430 で、図形の位置の変更で確かめた。画面でのドラッグは確かめていない）。
互換画像を付ける資料でも、数式を含む図形に接続する。

接続した線と接続先、端が図形の辺に接する線（図形への入り込みは線幅か 2 pt の大きいほうまで）は、重なりの警告の対象にしない。
吹き出しの接続点は本体の辺の中央である。しっぽのある辺には線を結ばない（[吹き出し](#吹き出し)）。

## 吹き出し

`slide.box()` の書式 `tail_side` で本体の辺に小さな三角形を付けると、吹き出しになる。三角形を「しっぽ」、長方形の部分を「本体」、しっぽのとがった頂点を「先端」、本体の辺の上の 2 頂点を「付け根」と呼ぶ。
しっぽを付けることができるのは `shape` が `rectangle`・`rounded_rectangle` の図形だけである。`slide.text()`・楕円・ひし形には付かない。

```python
# 枠線だけの補足。左の辺の上寄りにしっぽを付け、左の本文を指す
slide.box(id="note", area=Rect(420, 150, 240, 70), text="ここで $x_i$ と比べる",
          fill="none", line_color="text", line_width_pt=1, tail_side="left", tail_at=0.2)
# 角丸で塗りのある吹き出し。下の辺の中央から下の要素を指す
slide.box(id="hint", area=Rect(100, 80, 200, 50), shape="rounded_rectangle", text="要点だけを残す",
          corner_radius_pt=6, fill="surface", tail_side="bottom")
```

- `tail_at` は辺の上の付け根の中心の位置（上下の辺は左端 0〜右端 1、左右の辺は上端 0〜下端 1。既定 0.5）、`tail_width_pt` は付け根の幅（既定 10）、`tail_length_pt` は辺から先端までの長さ（既定 8）である。
  付け根は角丸の弧にかからない位置へ寄せるので、`tail_at=0` は角に最も近い位置になる。付け根の幅が辺の直線部分（辺の長さから角丸の半径の 2 倍を除いた長さ）を超えると `invalid_style` になり、しっぽを付けない。
- `area` は本体の長方形のままで、しっぽは `area` の外へ出る。文字領域、内容から決める寸法、`anchor()`、積み上げは本体で決まる。上下の辺にしっぽを付けて積むときは、間隔を `tail_length_pt` 以上にする。
- 役割に `tail_side` などを書くと、役割だけで吹き出しを定義できる。要素の側で `tail_side="right"`・`tail_at=0.8` などを上書きし、`tail_side="none"` でしっぽを取り消す（`tail_side=None` は指定なしで、役割の値が残る）。
  寸法と位置はテーマの `[defaults]`・`[defaults.box]` にも書くことができるが、`tail_side` は書くことができない（すべての図形と表のセルに届くため）。
  吹き出し向けの役割では `corner_radius_pt` も指定する。指定がないと角丸の半径が短い辺の約 1/6 になり、小さい吹き出しでは付け根が収まらないことがある。

```toml
# theme.toml
[roles.callout]
inherits = "note"
fill = "none"
line_color = "text"
line_width_pt = 1
padding_pt = 6
tail_side = "left"
tail_at = 0.2
```

`tail_toward` に同じページの別の要素の `anchor()` か座標 `(x, y)` を渡すと、しっぽの辺と付け根の位置をその点から決める。

```python
body = slide.text(id="body", area=Rect(40, 120, 340, 200), text="...")
slide.box(id="why", area=Rect(410, 130, 220, None), text="補足", line_color="text",
          tail_toward=body.anchor("right", at=0.2))
```

- 辺は、点が本体の外へ出ている横と縦の距離の大きい方（等しいときは左右の辺）にする。付け根の中心は点を辺へ垂直に下ろした位置で、長さは `tail_length_pt` のまま（先端は点まで伸ばさない）。
  付け根を角丸の弧の外へ寄せたために点が付け根の正面にないとき（点が本体の斜め外側にあるとき）は、先端を点の方へ傾ける。先端は付け根の中心と点を結ぶ直線の上に来て、傾きは辺の法線から 45° まで（先端を辺に沿って `tail_length_pt` までずらす）である。
- `style=` かキーワード引数の `tail_side` を併用すると、辺をその値に固定する。役割の `tail_side` は `tail_toward` が上書きする。`tail_at` は併用できない。配置を少し変えると辺が切り替わる位置関係では、`tail_side` を併用する。
- 点が本体の内側にあるときと、固定した辺の外側にないときは `invalid_geometry` になり、しっぽを付けない。まとまりの子の座標は、まとまりの左上が原点である。

配置の検査と書き出しは次のように扱う。

- ページと本文領域のはみ出しは、しっぽまでの範囲で判定する。積み上げの領域（`layout_bounds_overflow`）は本体で判定する。
- 重なりは本体としっぽを分けて判定する。しっぽ（3 頂点の外接長方形）がほかの要素に重なると `element_overlap` になり、観測値の `part` が `tail` になる。先端が相手の辺に接するだけなら警告しない。
  しっぽの先端を図や表の内側へ入れると警告になる。指す相手を `tail_toward` で渡すと、その要素とのしっぽの重なりは判定しない（本体が重なれば警告する）。
- 接続点は本体の 4 辺の中点で、しっぽの辺や位置では変わらない。しっぽのある辺に線を結ぶと線の端がしっぽの開いた部分に来るので、その辺には線を結ばない。
- しっぽのある図形は回転・反転できない（`unsupported_argument`）。向きは `tail_side` で決める。
- 本体としっぽは 1 つの多角形（PowerPoint の「フリーフォーム」）として書き出し、枠線の角は丸める。PowerPoint で大きさを変えると、しっぽと角丸も比率で伸縮する。調整ハンドルはない。
- PowerPoint 16.0 の build 20527 で、文字の位置・本体の角・枠線が同じ本体のしっぽのない図形と同じに描かれ、付け根に継ぎ目が出ないことを確かめた。
  図形の位置と大きさを変えると、接続した線は本体の辺の中央の接続点に追従し、文字領域と接続点は比率で伸縮する（画面でのドラッグは確かめていない）。

## まとまりと識別子

`slide.group(id=..., area=...)` はまとまりを作り、戻り値の同じ名前の関数で子を作る。子の座標はまとまりの左上を原点とし、子の識別子は「まとまりの識別子/子の識別子」になる。
まとまりを接頭辞として使うと、繰り返す図の識別子を整理できる。まとまりを動かすと子も動く。まとまりは回転・反転できない。

```python
decision = slide.group(id="decision", area=Rect(80, 120, 600, 300))
ask = decision.box(id="ask", area=Rect(180, 0, 280, 120), text="中央と比べる", shape="diamond")
left = decision.box(id="left", area=Rect(0, 200, 200, 60), text="左半分を探す")
decision.arrow(id="to_left", start=ask.anchor("left"), end=left.anchor("top", gap=6))
```

`pptx_group` が `None`（既定）か `True` のまとまりは、次の 3 つの条件をすべて満たすと PowerPoint のグループとして書き出す。

1. 描く子孫がある。
2. 子孫がまとまりの直後に続けて作られている（ほかの要素を間に挟まない。描画順を変えないため）。
3. 子孫に数式がない。

条件を満たさないまとまりと `pptx_group=False` のまとまりは、子をスライドの図形として書き出す。表も同じ条件でグループになる。文献一覧は既定でグループにしない。

## 文章の断片と記法

文章は文字列か、`Text`・`Emphasis`・`Math` の断片の列で書く。

| 断片 | 意味 | 引数 |
|---|---|---|
| `Text(文字列)` | 通常の文字 | `bold`・`color`・`size_pt`・`underline`・`link`・`superscript` |
| `Emphasis(文字列)` | 強調（既定は太字と強調色） | `Text` と同じ |
| `Math(式)` | 数式 | `bold`・`color`・`size_pt`・`emphasis`（真なら強調色と太字） |

断片の書式は、指定した項目だけが要素の書式より優先する。`link` は `https://`・`http://`・`mailto:` で始まる外部のリンク先で、slidekit はリンク先へ接続しない。
`superscript=True` は上付き（脚注の番号など）で、PowerPoint の文字の基準線のずらし（`baseline="30000"`）を書く。PowerPoint は上付きの文字を 3 分の 2 の大きさで、文字サイズの 30 % 上に描き、計測も同じにする（PowerPoint 16.0.20527.20150 の PDF で確かめた）。数式の指数（`^{1}`）と違い、直前の式の一部とは読まれない。
文章に改行 `\n` 以外の制御文字を含めるとエラーになる。

`slidekit.markup.parse()` は、`$…$` を数式、`**…**` を強調として断片の列にする。要素を作る関数は文字列の記法を自動では解釈しないので、使うときは明示して呼ぶ。

```python
from slidekit.markup import parse

slide.text(id="premise", text=parse(r"件数を $n$ とし、**最大 $k$ 件**まで扱う"), area=Rect(48, 100, 864, None))
```

強調の中の数式は強調色になる。`\$`・`\*`・`\\` はその文字になり、1 つの `*` は通常の文字である。閉じていない記号は位置と前後の文字を示すエラー（`invalid_markup`）になる。
`parse(text, math=False)`・`emphasis=False` で、その記号を通常の文字として扱うことができる。

## ページの題名と飾り

ページの題名はページ種別の `title_role`・`title_area` で置く。`deck.slide()` の引数でページごとに替えることができる。

- `title`：文字列・断片・断片の列。数式を含む題名も書くことができる。`None` は題名なし（`notes=None` は空のノート）。
- `title_role`・`title_area`・`title_style`：このページの題名の役割・領域・書式。ページ種別の値に重ねる。

章扉や表紙の飾り（線・矢印・図形）は、テーマのページ種別に書くと、その種別の全ページに置かれる。

```toml
[[page_kinds.section.decorations]]
kind = "box"
id = "mark"
area = [0.10, 0.30, 0.02, 0.04]
shape = "diamond"
fill = "accent"

[[page_kinds.section.decorations]]
kind = "line"
id = "rule"
start = [0.10, 0.56]
end = [0.90, 0.56]
line_color = "primary"
```

位置はページに対する比（`area`・`start`・`end`）か pt（`area_pt`・`start_pt`・`end_pt`）で書く。飾りには文字を置かない。
書式は種類の既定に重ね、役割の書式は重ねない。描く順は背景の後・題名の前である。`deck.slide(..., decorations=False)` で、そのページだけ飾りを置かない。
ページ種別の `background = "none"` は背景の図形を作らない（雛形のマスターの図形を見せるときに使う）。項目は [飾りの項目](reference.md#飾りの項目) にある。

## 文献一覧

`inputs.bibliography` に `.bib` ファイルを書くと、要素の出典に文献キーを付けることができる（`sources=[{"bib_key": "knuth1998"}]`、または `sources="knuth1998"`）。
読み込みには追加依存 `bibliography`（bibtexparser と pylatexenc）が要る。項目の LaTeX 表記（アクセント、`--` などの記号、書式の命令）は文字に直し、対応していない命令は知らせずに消すのではなくエラーにする。

`slide.references(id=..., keys=None, area=None, ...)` は文献一覧を作り、項目をまとめたまとまりを返す。

- `keys` を省くと、それまでに作ったページの出典の文献キーを、初めて現れた順に載せる（`deck.cited_keys()`）。空・重複・未定義のキーはエラーになる。
- `area` の既定は本文領域で、高さは `None` でよい。項目は幅で高さを計測し、`gap_pt`（既定 8）の間隔で上から積む。下端を越えても縮めず、配置の検査に任せる。
- 項目の識別子は文献キー（`/` は `_` に置き換える）で、出典にその文献キーが付く。
- `formatter` は `bibliography.Reference` を受け取って断片の列を返す関数で、既定の `bibliography.format_reference` は著者、題名（太字）、掲載情報、URL の行を作る。
- `text`・`sources`・`height`・`width`・`indent`・`place`・`gap` は使うことができない。書式は各項目に適用される。既定で PowerPoint のグループにしない（項目を個別に編集しやすくするため）。

文献は `deck.bibliography`（または `context.bibliography`）で、キーから項目を引くことができる。

## 画像

`slide.image(id=..., path=..., area=...)` は PNG・JPEG・SVG・PDF を置く。入力は参照範囲の中のファイルに限る。`area` の幅か高さを `None` にすると、もう一方を縦横比から決める。
`fit` は `contain`（既定。縦横比を保って中央に収める）か `cover` である。PDF には 1 始まりの `page` を指定する。

PDF・SVG は配置寸法 × `raster_scale`（1 pt あたりの画素数、既定 2）の PNG にする。SVG の素材は `vector=True` にすると SVG でも埋め込み、PowerPoint で拡大しても劣化しない。
PDF の素材は PNG だけにする。PDF に `vector=True` を指定すると `invalid_image_option` の `InputError` になる。拡大しても劣化させたくない PDF の図は、
SVG に変換して使う（例：Poppler の `pdftocairo -svg -f 1 -l 1 図.pdf 図.svg`、Inkscape）。図を作るスクリプトがあれば、SVG を直接書き出すとよい。
どちらも資料の既定を `deck.toml` の `[images]` で、個別の値を画像ごとの引数で指定できる。指定できるのは PDF・SVG の素材だけで、PDF の素材には `vector=True` を指定できない（資料の既定の `vector` は SVG の素材だけに当てはめる）。
外部の参照や実行される内容を持つ SVG と、文書型宣言を持つ SVG（文字コードによらない）は拒否し、XML として解析できない SVG はその旨を示して止める。
画像ファイルがないときは `asset_missing` の `InputError` になる。元の素材は変更しない。

## 数式

数式は文章の断片 `Math(r"...")` か `slide.equation(expression=r"...")` で指定する。記法は LaTeX に似た slidekit の記法で、`$` で囲まない。
受け付けない記法はエラーになる。記法、対応する命令、独立した数式と文中の数式の置き方、寸法の計算は [数式の書き方](math.md) にある。

`equation()` は、役割を指定しなければテーマの役割 `equation`（同梱のテーマでは `body` を継承する）を使う。式の文字サイズや揃えは、資料のテーマで `[roles.equation]` を上書きして決める。
文章中の変数は自動では数式にならない。数式にすべき文字は記述規則の `math_candidates` で見つけることができる（[記述規則](reference.md#記述規則)）。

## 雛形と数式の互換画像

既存の PPTX のテーマ、マスター、レイアウトを引き継ぐには、雛形を指定する。

```toml
[template]
path = "template.pptx"
layout_index = 6
show_master_shapes = false
```

雛形の元のページは外し、文書情報は `deck.toml` の値で上書きする。`layout_index` は雛形の中の空白レイアウトの 0 始まりの番号で、プレースホルダーは日付・フッター・ページ番号だけを許す。
既存のページの内容を取り込む機能ではない。マスターの図形を見せるには、`show_master_shapes = true` と、ページ種別の `background = "none"` を使う。

`math.fallback` は、数式を表示できないソフト向けの互換画像を付けるかを決める。

| 値 | 動作 |
|---|---|
| `required` | 数式がある資料に付ける |
| `off`（既定） | 付けない |

互換画像は PowerPoint で変換した PDF から数式の図形を切り出した画像で、PPTX の `mc:AlternateContent` の `mc:Fallback` に入る。PowerPoint は編集可能な数式を表示する。
互換画像を付けると、次のコストがかかる。PowerPoint だけで使い、ほかのソフトで開かない資料には要らない。

- 変換が 2 回になる（標準表示 1 回と、数式のページだけの互換表示 1 回）。互換画像を付けない場合と比べ、使用例の生成は binary_search（互換画像 55 個）で 10.8 秒から 12.8 秒に、proposal（4 個）で 8.8 秒から 9.6 秒になった（計測した条件は [生成の流れと所要時間](#生成の流れと所要時間) と同じ）。
- PPTX の容量が数式の数と面積に比例して増える。互換画像は数式を含む要素ごとに 1 枚で、binary_search の PPTX は 90 KB から 607 KB に増えた（画像 1 枚は 1.3〜47 KB、要素 1 個あたりの増加は平均 9.4 KB）。倍率は `math.fallback_scale`（1 pt あたりの画素数、既定 4）で決まる。
- 数式のあるページは互換表示の確認も要る。

修正の途中は `build --draft` で互換画像を省くことができる。
切り出す範囲は、数式の図形の矩形に、枠の外へはみ出した数式の描画（伸びた括弧の下端など）と枠線を加えた範囲である。矩形から 16 pt を超えるはみ出しは切り詰め、警告 `fallback_crop_clipped` で知らせる。
同じページのほかの要素の描画は範囲に含めない。外接矩形がその要素の範囲とほぼ一致する塗り（矢印の全体、箱の塗りなど）と、線・矢印の線分の近くにあり数式の矩形の外へ出る塗り（別に描いた矢印の先端など）を、その要素の描画とみなす。
それでも切り詰めたときは、超えた部分に範囲が掛かるほかの要素の識別子を警告の観測値 `nearby_elements` に示す。超えた描画がその要素のものなら、数式は欠けない。

標準表示の確認は最終 PPTX と同じスライドの XML を変換した画像で行う。互換画像を付ける資料の生成は、互換画像だけを仮の画像にした最終形の PPTX を PowerPoint で 1 回変換し、
その PDF を標準表示の確認に使い、同じ PDF から互換画像を切り出して最終 PPTX に差し替える。
最終 PPTX と最終形の PPTX のスライドの XML と関係が一致することを生成の中で確かめ、生成記録（`operations.final_form`）に残す。一致しなければエラー `final_form_mismatch` になる。
互換表示の確認は最終 PPTX から数式を互換画像に置き換えた確認用の複製を変換して行う。

## 生成と確認の手順

### 作成の流れ

1. `slidekit layout`：編集のたびに使う。表示ソフトを使わず、配置と記述規則の誤りを数秒で示す。
2. `slidekit build --draft`（必要なら `--slides`）：見た目を確かめる。互換画像を省き、配置のエラーがあれば変換しない。確認記録と完成物の対象外である。
3. `slidekit build`：確認と完成物のための通常の生成。前の実行から確認を自動で引き継ぐ。
4. `slidekit review status`：残りの確認を調べ、画像と本文を確かめてから `review record` で記録する。
5. `slidekit export`：最終検査に通った実行を完成物へ反映する。

`slidekit doctor` は、初回と、実行環境（フォント、PowerPoint、`environment.toml`）を変えたときに使う。
`build` は同じ事前確認を内部で行うので、毎回 `doctor` を前に実行する必要はない。

### コマンドの一覧

すべてのコマンドが `--json`（要約の代わりに結果の辞書を JSON で出す）、`--quiet`（進捗を出さない）、`--verbose`（許容した警告と生成の参考の一覧も要約に出す）を受け付ける。`slidekit --version` は版を示す。

| コマンド | 引数 | 用途 |
|---|---|---|
| `slidekit layout` | `--project P`（必須）、`--slides SPEC`、`--wraps` | 配置と記述規則だけを検査する。表示ソフトを使わず、ファイルを書かない。`--wraps` は自動で折り返した要素の行の文字列と幅を示す |
| `slidekit measure` | `--project P`（必須）、文字列（複数可）、`--width PT\|content`、`--kind text\|box\|equation`、`--role`、`--size PT`、`--page-kind`、`--markup` | 文字列を資料の書式で計測し、折り返した行・枠の寸法・1 行に収める枠の幅を示す。要素を作らず、ファイルを書かない |
| `slidekit doctor` | `--project P`（必須）、`--wait 秒` | 配置の検査に加え、出力先・実行環境・表示ソフトを事前に確かめる |
| `slidekit build` | `--project P`（必須）、`--slides SPEC`、`--draft`、`--stop-on-error`・`--no-stop-on-error`、`--jobs N`、`--wait 秒`、`--keep-intermediate`、`--carry-from auto\|none\|latest\|RUN` | 新しい実行を生成する |
| `slidekit check` | `--build RUN`（既定 `latest`）、`--project P`、`--final` | 生成記録と成果物を検査する。`--final` は最終検査 |
| `slidekit export` | `--build RUN`（必須）、`--project P` | 最終検査に通った完成物を固定パスへ反映する |
| `slidekit clean` | `--project P`（必須）、`--before RUN`、`--keep N`、`--failed`、`--drafts`、`--apply`（条件を 1 つ以上） | 古い実行の削除候補を示し、`--apply` で削除する |
| `slidekit review status` | `--build RUN`（既定 `latest`）、`--project P` | 残りの確認・未対応の警告と、次の操作を示す |
| `slidekit review record` | `--build RUN`、`--project P`、`--kind K[,K]` か `--pending`、`--slides SPEC`、`--variant standard\|fallback\|all`、`--status-id ID`、`--reviewer`、`--note`、`--note-visual`・`--note-overview`・`--note-content` | 画像や本文を開いて確かめた確認を記録する。`--pending` は `review status` が示した未確認をまとめて記録する |
| `slidekit review sheet` | `--build RUN`（既定 `latest`）、`--project P`、`--slides SPEC`（既定 `pending`）、`--variant standard\|fallback\|all`、`--compare`・`--no-compare` | ページの画像を、前の実行の画像と互換表示の画像と横に並べた比較画像を作る。前の実行と差のある領域を枠で示す。確認の記録はしない |
| `slidekit review accept` | `--build RUN`、`--project P`、`--diagnostic ID[,ID]` か `--code CODE`（`--slides SPEC`・`--element ID` で絞る）、`--reviewer`、`--reason` | 警告を理由付きで許容する |
| `slidekit review revoke` | `--build RUN`、`--project P`、`--record ID[,ID]`、`--reviewer`、`--reason` | 確認か許容の記録を取り消す |
| `slidekit review carry` | `--from auto\|RUN`、`--to RUN\|latest`、`--project P` | 前の実行の確認と許容を手で引き継ぐ |
| `slidekit renderer status` | `--project P`（任意） | 表示ソフトの排他・残存記録・使用中の処理と、PowerPoint の処理（slidekit が起動したものか）を示す |
| `slidekit renderer clear` | `--project P`（任意）、`--force`、`--stop-owned` | 安全を確かめてから残存記録を解除する。`--stop-owned` は slidekit が起動した PowerPoint だけを終了してから解除する |

- `P` は `deck.toml` かそのディレクトリである。
- `RUN` は実行ディレクトリのパス、番号（`12`・`0012`）、`latest`（完成物にできる最新の実行）である。番号と `latest` は `--project`（省くとカレントディレクトリの `deck.toml`）の実行保存先から探す。
  `review record`・`accept`・`revoke` は `latest` を受け付けない（画像を見た後に別の生成があると、見ていない実行に記録してしまうため）。
- `SPEC` はページ番号（`3`）、範囲（`5-7`）、識別子（`intro`）、`id:識別子` のカンマ区切りである。空白で区切って並べてもよい（`--slides 3 5-7` は `--slides 3,5-7` と同じ）。番号は資料の元のページ番号で、数字だけの識別子は `id:2024` と書く。
  確認の記録では `all` と `pending` も使うことができる。`all`・`pending` と同じ名前の識別子も `id:` を付けて書く。カンマを含む識別子は指定できない。
- `renderer` の `--project` を省くと、共通の実行環境だけを読む。

### コマンドの出力

既定の出力は人が読む要約で、標準出力に 1 回だけ出す。機械が読む用途には `--json` で結果の辞書を出す。

```text
配置の検査: 合格（エラー 0 件、警告 0 件）
資料: examples/proposal/src/deck.toml（4 ページ、互換画像あり）
所要時間: … 秒（配置の検査 …、構築 …、その他 …）
次の操作: slidekit build --draft --project examples/proposal/src/deck.toml
結果: 合格（エラー 0、終了コード 0）
```

- 進捗（`[  12.3 秒] 構築` の形の行）、資料のコードの `print`、トレースバックは標準エラーへ出す。`--quiet` は進捗だけを止める。
- 要約は、状態、診断の件数と主な診断、所要時間の内訳（段階をまとめた時間の長いものから 4 つまでと、残りの「その他」）、次の操作を示す。`build` は実行ディレクトリと確認の引き継ぎの結果も示す。
- どのコマンドも、最後の行を「結果: 合格（エラー 0、終了コード 0）」の形の合否の行にする。残りのエラー・未完了の確認・未対応の警告の件数と終了コードを示すので、末尾の 1 行だけで結果が分かる。
- 宣言か記録で許容した警告と、`build` の参考の診断は、既定では件数だけを示す（「宣言で許容した警告 3 件（一覧は --verbose）」）。一覧は `--verbose` で出す。新しく出た警告が許容済みの警告に埋もれない。
- 「次の操作」のコマンドは、起動した方法に合わせて書く。`python -m slidekit` で起動したときはその Python の絶対パスと `-m slidekit`、PATH にない `slidekit` を絶対パスで起動したときはそのパスにする。
- 終了コードは、成功が `0`、品質上の未解決事項（エラーの診断、未確認、未対応の警告）が `1`、入力・環境・変換の失敗と中断が `2` である。
- `build` の結果と生成記録の `status` は、`built`（生成した。エラーの診断があってもよい）、`stopped`（`--stop-on-error` で変換の前に止めた）、`failed`（例外で止まった）である。品質は `counts`（重大度ごとの件数）と終了コードで示す。
- 説明のための例外（確認記録の操作の拒否、入力と設定の誤り）は、トレースバックを出さずに文言だけを示す。設定の問題は 1 件 1 行で出す。それ以外の例外はトレースバックを標準エラーへ出す（`--quiet` のときも）。
- 引数の誤り（`--slides`・`--jobs`・`--wait`・`--keep`・`--carry-from` の書式、条件のない `clean`）は、使い方を示して終了コード `2` になり、処理を始めない。
- `--slides` の範囲外の番号と存在しない識別子は、構築の後に判明するので、診断 `invalid_selection`（終了コード `2`）で示す。入力の誤りなのでトレースバックは出さず、次の操作は `--slides` の指定を直すことである。

### 配置だけを検査する

`slidekit layout --project deck.toml` は、構築・素材の準備（一時フォルダー）・配置と記述規則の検査だけを行う。表示ソフトを使わず、資料のディレクトリと実行保存先に何も書かない。
別の変換の実行中でも、何人でも同時に動かすことができる。`--slides` を付けると、検査と報告を選んだページに限る。
`--json` の結果の `layout` には `layout.json` と同じ形の予測（要素ごとの行と寸法）が入る。各行の `wrap` は行末の自動の折り返しの種類（`space`・`math`・`punctuation`・`phrase`・`mid_phrase`・`hyphen`・`split`・`other`。明示改行で終わる行と最後の行は `null`）である。
`--wraps` を付けると、要約に自動で折り返した要素の行の文字列と幅を出し、文節の途中・ハイフンの後・語の途中の折り返しに印を付ける。画像を開く前に、折り返しの位置を文字で確かめることができる。

```text
自動で折り返した要素:
  p.3 lead（内側の幅 600.0 pt）
    「通知が届いたた」 590.0 pt ← 文節の途中
    「め、次の処理へ進む」 210.0 pt
```

折り返しの位置について、次の参考の診断を出す（どれも許容は要らない）。

- `text_short_line`：自動で折り返した段落の最後の行が、2 文字以下か内側の幅の 2 割未満である。
- `text_wrap_mid_phrase`：和文の文節の途中で折り返している。文節の境は BudouX の日本語のモデルで判定する。
- `text_wrap_hyphen`：欧文の語のハイフンの後で折り返している（`time-` / `fusible` のような折り返し）。

### 一部のページと下書きの生成

- `build --slides SPEC`：選んだページだけを書き出す。番号は資料の元のページ番号である。
- `build --draft`：互換画像を省く。
- `--stop-on-error`：配置・記述規則・入力の検査でエラーがあれば、変換の前に止める（`status` は `stopped`）。下書きでは既定で有効で、`--no-stop-on-error` で無効にできる。

部分生成と下書きの実行は、確認の記録・許容・引き継ぎ・最終検査・`export` の対象外である。完成させるときは、全ページを通常の `build` で生成する。

`--slides` の指定は生成関数に `context.selection`（指定の文字列。指定がなければ `None`）として渡る。
`context.selected(*識別子)` は、識別子のどれかが選ばれていれば真を返す（選択がなければ常に真、番号や範囲を含む指定でも常に真）。
資料のコードは選ばれたページを含む章だけを作ってよい。章だけを作った検査と生成では、ページを番号でなく識別子で示す（[作成の手順](../workflows/authoring.md#複数人で並行して書く)）。

### 生成の流れと所要時間

`build` は次の段階で進み、進捗を標準エラーに 1 行ずつ出す。段階ごとの所要時間は生成記録の `timings` に残り、要約は内訳を示す。

1. 構築、素材の準備、配置の検査（`deck.json`・`notes.md`・`layout.json`・暫定の `diagnostics.json` を書く）
2. PPTX の書き出し、実行環境の記録
3. 表示ソフトの事前確認（変換待ちを含む）
4. 互換画像を付ける資料では、最終形の PPTX の作成、PowerPoint で変換（標準表示）、互換画像の作成、互換表示の複製と変換（数式のページだけ）
5. 互換画像のない資料では、PowerPoint で変換（標準表示）
6. 最後の変換の後の PowerPoint の終了（要約では「表示ソフトの終了」にまとめる）
7. PPTX の検査、PDF の検査と画像化（変換の順）、結果の集約、生成記録の保存

`timings` と `--json` の結果の段階の名前と、進捗に出る名前は次のとおりである。`layout`・`doctor` の結果の `timings` にも、使った段階が入る。

| 段階 | 進捗の名前 | 使うコマンド |
|---|---|---|
| `construct` | 構築 | `layout`・`doctor`・`build` |
| `assets` | 素材の準備 | `layout`・`doctor`・`build` |
| `layout` | 配置の検査 | `layout`・`doctor`・`build` |
| `pptx` | PPTX の書き出し | `build` |
| `environment` | 実行環境の記録 | `doctor`・`build` |
| `doctor` | 表示ソフトの事前確認（変換待ちを含む） | `doctor`・`build` |
| `final_form` | 最終形の PPTX の作成（互換画像を付ける資料） | `build` |
| `render:powerpoint:<表示>` | 「PowerPoint で変換（標準表示）」の形 | `build` |
| `math_fallback` | 互換画像の作成（互換画像を付ける資料） | `build` |
| `fallback_clone` | 互換表示の複製（互換画像を付ける資料） | `build` |
| `close:powerpoint` | PowerPoint の終了 | `build` |
| `pptx_check` | PPTX の検査 | `build` |
| `pdf:powerpoint:<表示>` | 「PDF の検査と画像化（標準表示）」の形 | `build` |
| `collect` | 結果の集約 | `build` |
| `record` | 生成記録の保存 | `build` |

`<表示>` は `standard`（標準表示）・`fallback`（互換表示）である。
要約の内訳は、`render:…` を「変換」、`pdf:…` を「PDF の検査と画像化」、`close:…` を「表示ソフトの終了」、`doctor` を「事前確認」、`math_fallback` と `final_form` を「互換画像」にまとめる。

PowerPoint の変換は利用者ごとに 1 つずつで、生成は事前確認から最後の変換まで排他を持つ。PowerPoint を使わない処理（ページ画像と互換画像の切り出し）は、
`--jobs`、環境変数 `SLIDEKIT_JOBS`、実行環境の `jobs`、CPU の数と 8 の小さいほうの順に決めた数のプロセスで並列に行う。生成記録の `operations.parallel` に `jobs` が残る。
所要時間の内訳の合計は全体を超えない。

使用例の生成の所要時間を次に示す（2026 年 10 月 20 日、WSL 2 と PowerPoint 16.0。生成の前に PowerPoint は起動していない。生成記録の `elapsed_seconds` の 3 回の中央値）。
`--slides 3` は 3 ページ目だけの生成で、どちらの使用例でも数式を含むので互換画像を作る。

| 使用例 | 通常 | `--draft` | `--slides 3` |
|---|---|---|---|
| proposal（16:9、4 ページ、互換画像 4 個） | 9.6 秒 | 8.8 秒 | 8.5 秒 |
| binary_search（4:3、10 ページ、互換画像 55 個） | 12.8 秒 | 11.2 秒 | 8.6 秒 |

binary_search の通常の生成の内訳は、変換 4.9 秒（標準表示 3.9、互換表示 1.0）、表示ソフトの終了 2.8 秒、事前確認 1.3 秒、構築 1.2 秒、互換画像 0.6 秒、PDF の検査と画像化 0.5 秒である。
ページ数によらない時間が約 5 秒ある（PowerPoint の起動を含む事前確認 約 1.3 秒、最後の変換の後の PowerPoint の終了 2.3〜2.8 秒、Windows への問い合わせなど 約 1 秒）。そのため、1 ページだけの生成でも 8〜9 秒かかる。
ページが多い資料では変換が大半を占める。変換の速さは環境によって変わり、同じ資料の変換の時間が 2 倍近く違ったこともある。
1 回の生成で、PowerShell を 2 回（Windows への問い合わせと変換の作業プロセス）起動する。PowerPoint は、起動していなければ 1 回だけ起動し、最後の変換の後に終了する。

`doctor` の結果と生成記録の `environment.font_directories` には、PowerPoint の事前確認が通ったときだけ自動で探したフォルダーを入れ、それ以外は設定の値か `null` を入れる。
フォルダーは確認の証拠に入れない（使ったフォントの実体のハッシュ値は入れる）。

### 実行ディレクトリと記録

`build` は実行保存先（既定 `out/`）の `runs/NNNN/` に新しい実行ディレクトリを作る。実行保存先を新しく作ったときは、中身を追跡しない `.gitignore` を置く。

| ファイル | 内容 |
|---|---|
| `deck.pptx` | 最終 PPTX |
| `deck.pdf` | 標準表示の PDF（`renders/` の PDF へのハードリンク。できなければ複製） |
| `deck.json` | 構築した場面データ |
| `notes.md` | 発表者ノート（「## 番号. 題名」とノートを並べる） |
| `layout.json` | 計測の予測（要素ごとの行と寸法） |
| `renders/powerpoint/<standard\|fallback>/` | 変換した PDF、ページ画像 `slide-NNN.png`、一覧画像 `contact-NNN.png`（標準表示）、予測と実際の行の照合 `layout.json` |
| `assets/` | 配置用に変換した素材 |
| `intermediate/` | 中間 PPTX。確認用の複製と最終形の PPTX は `--keep-intermediate` のときだけ残す |
| `manifest.json` | 生成記録：入力・設定・環境・成果物のハッシュ値、段階と所要時間、確認の証拠 |
| `diagnostics.json` | 機械検査の診断 |
| `review.json` | 表示・内容の確認、警告の許容と取り消し、引き継ぎ |
| `check.json`・`final_check.json` | 直近の整合検査と最終検査の結果 |
| `traceback.txt` | 例外で止まったときのトレースバック |
| `.building` | 生成中の印（生成が終わると消える） |

### 失敗したとき

生成関数の例外と構築の誤りは、失敗の診断（識別子は例外の `code`。なければ `build_failed`）になり、資料コードの中の場所（ファイルと行）を示す。
章のモジュールの構文の誤りは、読み込んだ行でなく誤りのあるファイルと行を示す。全文は `traceback.txt` と標準エラーにある。
生成の途中の中断（Ctrl+C）も実行ディレクトリに記録を書き、診断 `build_interrupted` を加える。
表示ソフトの使用中・時間切れ・残存記録で失敗したときは、要約が次の操作（変換の終了を待つ、`renderer status` で確かめる、`renderer clear` で解除するなど）を示す（[実行環境](#実行環境)）。

### 確認と完成物の保存

`build` は前の実行のうち証拠が一致する確認と許容を自動で引き継ぐ。残りの確認は `review status` で調べる。
一覧画像・全ページの画像・互換表示の画像を開き、本文とノートを依頼内容に照合してから記録する。

```sh
slidekit build --project examples/proposal/src/deck.toml
slidekit review status --project examples/proposal/src/deck.toml
slidekit review record --project examples/proposal/src/deck.toml --build 1 --kind visual --variant all --slides all --reviewer editor --note "全ページの文字・図表・数式を確認"
slidekit review record --project examples/proposal/src/deck.toml --build 1 --kind overview --reviewer editor --note "ページ順・全体の配置・文書情報・依頼内容との対応を確認"
slidekit review record --project examples/proposal/src/deck.toml --build 1 --kind content --slides all --reviewer editor --note "依頼内容と本文・ノートを照合"
slidekit export --project examples/proposal/src/deck.toml --build latest
```

- 確認の種類は、`visual`（ページの表示）、`overview`（一覧画像。ページ順・全体の配置・文書情報）、`content`（本文・ノート・出典の内容）である。`--kind visual,content` のようにまとめて指定できる。
- `--slides` は `visual` と `content` で必須である。`all` は、その確認が要る全ページを選ぶ（互換表示では数式のあるページ）。明示したページにその確認が要らなければエラーになる。
- `--variant` は `standard`（既定）・`fallback`（互換表示）・`all` である。
- 未確認のページだけを記録するときは、`review status --json` が示す照合用の値を使い `--slides pending --status-id …` と書く。
- 画像と本文をすべて見た後は、`review record --build N --pending --status-id … --reviewer … --note …` で、`review status` が示した未確認の確認（表示・一覧画像・内容のすべての組）を 1 回で記録できる。照合用の値は `review status` の結果の `status_id`（要約の「次の操作」の先頭のコマンドにも出る）で、その後に未確認の対象が変わっていれば記録しない。種類ごとに確認内容を分けるときは `--note-visual`・`--note-overview`・`--note-content` を使う（`--note` より優先する）。
- 記録の結果は、残っている確認と未対応の警告の件数（`remaining`）を示す。残りがなければ次の操作は `export` になる。
- `review sheet --build N` は、表示の確認が未完了のページ（`--slides` で替える）の画像を、前の実行の画像・互換表示の画像と横に並べた比較画像を、実行ディレクトリの `sheets/` に作る。前の実行と差のある領域を今の画像に赤い枠で示す（ページ番号の領域は除いて比べる）。比較画像は確認を助ける表示で、見たことを確認の記録としない。ページ画像を開いて確かめてから記録する。
- `review status` が示す記録と許容のコマンドは、確認者・確認内容・理由を `<確認者>`・`<確認した内容>`・`<理由>` の置き場所で示す。画像を見て書き換えてから実行する。置き場所のままの値と、全体を `< >` で囲んだ値では、記録も許容もしない。
- 1 回のコマンドで作った記録には同じ `batch`（`b_…`）が付く。
- 確認の途中で入力を直しても、その実行に記録できる（要約が知らせる）。その実行は反映できないが、再生成した後に変わっていないページの記録が引き継がれる。

`review status` の状態は、`complete`（終了コード 0）、`incomplete`（未確認か未対応の警告がある。1）、`blocked`（エラーの診断がある 1、成果物・入力の不一致や完成物にできない実行 2）である。

確認の引き継ぎは、証拠（ページの部品、内容、表示環境、画像）が一致する確認対象に限る。変更ごとの扱いは次のとおりである。

| 変更 | 内容の確認 | 表示の確認 | 警告の許容 |
|---|---|---|---|
| ページに表示する内容（文字・数式・題名・画像・図形と線の接続）の変更 | そのページだけやり直す | そのページだけやり直す | 関係する警告をやり直す |
| ノート・出典だけの変更（ページに表示しない内容） | そのページだけやり直す | 引き継ぐ | 関係する警告をやり直す |
| 配置・書式だけの変更 | 引き継ぐ | そのページだけやり直す | 関係する警告をやり直す |
| ページ番号だけの変更（桁数が同じ） | 引き継ぐ | 引き継ぐ（番号の領域を除いて画像を比べる） | 引き継ぐ（ページ番号の要素の警告を除く） |
| ページ番号の桁数・書式・枠の位置と大きさが変わる変更（9 から 10、`page_number_format`、`footer_area_pt` など） | 引き継ぐ | 引き継ぐ（番号の図形を除いた証拠が一致し、変更前と後の番号の領域を除いた画像が一致するとき） | そのページの PDF の警告と、ページ番号の要素の警告をやり直す |
| 参照する元資料・文献項目の変更 | 参照するページだけやり直す | 引き継ぐ | — |
| 文書情報・依頼内容・ページ順の変更 | 引き継ぐ（一覧画像の確認で扱う） | 引き継ぐ | — |
| 記述規則の変更 | 引き継ぐ | 引き継ぐ | その規則を使う警告だけやり直す |
| フォント・PowerPoint・画像化の変更 | 引き継ぐ | 全ページやり直す | 計測に使う警告をやり直す |

表示の確認の証拠はページの PPTX の部品（ノートを含まない）と表示環境である。部品のページ番号の数字は 0 に置き換えて比べるので、桁数が同じ番号の変化では証拠が変わらず、桁数が変わると変わる。
生成記録には、ページ番号の図形を除いた証拠（`visuals` の `body_hash`）も記録する。証拠が違っても、これが一致するページは、変更前と後の番号の領域を除いて画像を比べ、差がなければ「ページ番号だけ変わったページ」として表示の確認を引き継ぐ（写しの `render_check` に `footer_only` と元の記録の証拠を残す）。
ただし、ページか雛形に表示されるフィールド（`a:fld`。雛形のページ番号や日付）があるページは、資料の中の位置（日付のフィールドなら生成した日も）を証拠に入れるので、位置が変わると表示の確認をやり直す。
内容の確認の証拠は題名・文字・数式・ノート・出典・図の構造と、参照する元資料・文献項目である（座標・書式・ページ番号は入れない）。
一覧画像の確認は毎回行う（引き継がない）。表示の確認を引き継ぐときは、元の記録の画像と新しい画像を 1 画素のずれを許して比べ、大きな差があれば引き継がない。
引き継いだ記録（写し）は、人が記録した元の記録を指す。どちらかを取り消すと、同じ元の記録の写しはすべて無効になる。

- `build --carry-from` が受け付ける引き継ぎ元は、`auto`（既定。同じ実行保存先の前の実行から、新しい順に探す）、`none`（引き継がない）、`latest`（生成の前の、完成物にできる最新の実行）、実行の番号かパスである。
  それ以外の書式は生成の前に引数の誤りになる。完成物にできる実行がないときと、存在しない番号を指定したときは、生成は行い、引き継ぎの失敗として示す。
  引き継ぎに失敗しても生成は失敗にせず、終了コードも変わらない。
- 要約の「確認の引き継ぎ」の行は、件数、引き継ぎ元の実行、ページ番号だけ変わったページの数と比べた実行（数式のページの標準表示と互換表示の 2 つの確認も 1 ページとして集計する）、候補にしなかった実行（古い形式の実行など）を示す。
  続く行に、引き継がなかった確認の理由（表示は「ページの表示の変更」「表示環境の変更」、内容は「本文の変更」「元資料の変更（元資料の識別子）」「文献項目の変更」、どちらも「前の確認の記録がない」）と、確認の記録がある最も新しい実行との表示環境の違い（PowerPoint・パッケージ・Python・フォント・計測の版）を示す。
  `review status` の未確認のページにも、理由ごとのページを添える（例：「内容 3,5-7（4 ページ。本文の変更 p.3,5、元資料の変更 p.6-7（proofs））」）。
  `--json` の結果の `carry` は `status`・`mode`・`count`・`sources`・`renumbered`・`renumbered_from`・`skipped`・`skipped_runs`・`remaining`・`pending_reasons`・`compared_run`・`environment_changes`・`notices` を持ち、`review.json` の `carry_attempts` にも残る。
- `build` の要約は、前の完了した実行（下書きと部分生成を含む）とページ画像を比べた結果も示す（例：「前の実行 0133 との比較: 変更なし 34、ページ番号だけ変わった 1、変わった 1（p.14）」）。`--json` の結果では `comparison`。
- `review carry --from auto --to latest` で手で引き継ぐことができる。2 回実行しても重複して写さない。
- 資料のディレクトリを実行保存先ごと移しても（何度移しても）、同じ実行保存先の実行どうしは同じ資料として扱い、確認を引き継ぐことができる。反映した実行と退避した旧版も、整理で保護され続ける（古い形式の反映の履歴は絶対パスなので、移した後は保護の理由にならない）。
  `runs/` の外へ複製した実行には記録・許容・取り消し・引き継ぎができない（検査と `review status` はできる）。

`check --final` と `export` の最終検査は、成果物と入力が生成記録と一致し、エラーの診断がなく、必要な確認（全ページの内容と標準表示、数式のページの互換表示、一覧画像）がそろい、
すべての警告が許容か宣言で対応済みであることを確かめる。画像を実際に見たか、主張が正しいかを自動で証明するものではない。
`export` は最終検査を含み、完成物を `deliverables` の保存先へ反映し、旧版を `out/exports/<識別子>/` に退避する。
採用記録（`deliverables.json`）と反映の履歴（`transaction.json`）のパスは資料のディレクトリからの相対パスで、反映した完成物の権限は umask に従う。

確認記録の操作・`export`・`clean --apply` は、実行保存先のロックのファイル（`.slidekit-export.lock`）の中で行う。ロックを 30 秒待っても取得できなければ、別の処理の終了を待って再実行する。
古い形式（slidekit 0.1）の実行には確認の記録・許容・取り消し・引き継ぎ・検査・反映を行わない（`legacy_run`）。再生成してから確認する。

### 警告の許容

エラーは許容できない。編集元か設定を直して再生成する。警告は原因を確かめて直し、意図した状態だけを理由付きで許容する。方法は 2 つある。

- 資料コードで宣言する：`slide.allow_warning("element_overlap", element=label, related=figure, reason="図の上にラベルを重ねる")`。
  2 つの要素の組の警告（`element_overlap` など）では `element` が必須で、`related` は組の相手である。それ以外の警告で `element` を省くと、そのページの同じ種類の警告すべてが対象になる。
  宣言した警告は確認記録なしで最終検査を通る。対応する警告がない宣言は、参考の診断 `allow_warning_unused` になる。宣言しても、そのページの表示確認は要る。
  宣言で許容した警告は、`review status` と `check --final` の要約に「宣言で許容した警告」として件数を示し、`--verbose` を付けるとページ・種類・要素・理由・宣言の場所とともに並べる（`--json` の結果では `declared_warnings`）。
  `layout`・`doctor`・`build` の要約では、警告の件数に「うち宣言で許容 N 件」を添える。`--verbose` を付けると、その警告の行を並べて末尾に「（宣言で許容）」を付ける（`--json` の診断では `declared`）。
- 生成の後に記録する：`slidekit review accept --build 1 --code element_overlap --slides 3-5 --reviewer editor --reason "意図した重なりである"`、または `--diagnostic` に診断の識別子を書く。

許容は警告の根拠（`evidence_hash`）に結び付き、根拠が変わると失効する。`allow_warning` は、使う場所を限った役割の宣言 `allow_roles` とは別の機能である。

```sh
slidekit review revoke --build 1 --record r_0123456789abcdef --reviewer editor --reason "確認を見直す"
```

`review revoke` は記録の識別子（`r_…`）か、1 回のコマンドで作った記録すべて（`b_…`）を取り消す。

### 古い実行の整理

`slidekit clean --project deck.toml` は、条件に当たる古い実行の削除候補を示し、`--apply` のときだけ削除する。候補の表示はロックを取らず、何も書かない。

- 条件（1 つ以上）：`--before RUN`（実行の識別子。`0005` ならそれより前の実行）、`--keep N`（条件に当たる実行のうち新しい N 件を残す）、`--failed`（失敗・停止・中断した実行）、`--drafts`（下書きと部分生成）。
- 保護する実行：採用中の実行、完成物として反映した実行と退避した旧版、反映が未完了の実行、確認記録の引き継ぎ元、実行中の実行と実行中か判定できない実行、生成記録も生成中の印もない実行、別資料の実行と所属を確認できない実行。古い形式の実行の採用記録・反映の履歴・引き継ぎ元も保護する。
- 資料を実行保存先ごと移した場合も、生成記録から実行保存先に対する資料の相対位置が同じと確認できれば整理の対象にする。生成中の印（`.building`）にも実行保存先を記録するので、生成記録ができる前に中断した実行も同じ相対位置で判定する。
- 別の資料の実行は「別の資料の実行（資料: …）」、資料か実行保存先の記録がない実行（古い形式の印など）は「現在の資料の実行と確認できない（記録に資料か実行保存先がない）」と示して残す。
- `--json` の `protected` の各項目は、`run`・`run_id`・`reasons`（文言）・`reason_codes`（`adopted`・`exported`・`previous_export`・`export_pending`・`export_recovery`・`running`・`running_unknown`・`unreadable`・`no_record`・`unfinished`・`other_project`・`unknown_project`・`carry_source`）・`details` を持つ。`details` の別資料の項目は、記録された資料の場所（`recorded_project`）と、実行保存先から見た資料の位置（`recorded_place`・`current_place`）を持つ。
- 要約で、生成記録のない実行（中断した生成）は「生成記録なし」と示す。

## 機械検査

`layout`・`doctor`・`build` は、配置と記述規則を検査して診断を返す。`build` はさらに PPTX と、変換した PDF を検査する。診断の種類・判定・直し方は [診断の一覧](diagnostics.md) にある。

- 配置：計測で、文字のはみ出し（`text_overflow`。図形の文字領域と余白が基準）、余白への食い込み、ページと本文領域の外、要素の重なり、最小文字サイズを調べる。
- 記述規則：禁止文字、数式の候補、明示改行の禁則、発表者ノートの有無と文末。
- PPTX：場面の通りに書き出したか（文字、書体の区切り、数式、互換画像、接続、まとまり、パッケージの参照）。
- PDF：文字が欠けずに描かれたか、PowerPoint が書体を置き換えていないか、要素のインクが枠に収まるか、行数が予測と合うか。
  PDF の文字と描画を、予測した要素の範囲で要素に割り当てて比べる（`renders/powerpoint/<表示>/layout.json`）。割り当てられない文字は `pdf_text_unassigned`、
  描かれていない文字と数式は `pdf_text_missing` になる。インクの比べ方と許容幅は、計測の照合データで確かめた誤差に合わせてある。

重なりの警告 `element_overlap` は次の場合には出ない。

- 先に作った図形の内側に、後から置いた要素。図形に文章があるときは、その文章と重ならないもの。塗りも文章もない図形の中のもの
- まとまりとその子孫、線どうし
- `anchor()` で結んだ線と、その接続先。端が図形の辺に接する線
- 吹き出しのしっぽと、`tail_toward` で指した要素。しっぽを横切らない線。吹き出し全体が収まる枠（本体は通常どおり判定する）
- `semantic="background"`・`semantic="decoration"` を付けた要素

同じまとまりの中の要素どうしも検査する。枠は文字より先に作る。それ以外の意図した重なりは、[警告の許容](#警告の許容) で理由を残す。

計算した寸法は PowerPoint の表示と完全には一致しない。最終 PPTX の表示を確認する。

## 実行環境

共通の設定は、Linux と WSL では `~/.config/slidekit/environment.toml`、Windows では `%APPDATA%/slidekit/environment.toml` に置く。
`SLIDEKIT_ENVIRONMENT` で場所を替えることができ、資料だけの上書きは `deck.toml` の `environment_override` で指定する。項目は [実行環境](reference.md#実行環境environmenttoml) にある。

### 既定の書体の探し方

寸法の計算には PowerPoint と同じ実際のフォントファイルを使う。テーマが使う書体のファイルは次のように決める。

1. 実行環境の `fonts.<書体>.path` があれば、そのファイルを使う。
2. なければ、既定の書体を探すフォルダーから探す。フォルダーは、実行環境に `font_directories` があるかどうかで、次のどちらか一方に決まる。
   - `font_directories` を書いたとき：そのフォルダーだけを探し、自動では探さない（`[]` なら探さない）。存在しないフォルダーはエラーになる。
   - 書かないとき：自動で探したフォルダーを使う。WSL では PowerShell で Windows のフォントのフォルダー（システムと利用者単位）を求める（PowerShell を使用できなければ `/mnt/c/Windows/Fonts`）。
     Windows の Python では `%WINDIR%\Fonts` と `%LOCALAPPDATA%\Microsoft\Windows\Fonts`。それ以外の環境では探さない。

同じ名前のファイルは前のフォルダーを優先する。自動で探すのは候補の書体（Yu Gothic UI、Yu Gothic、Arial、Cambria Math、Calibri、Segoe UI、Meiryo、BIZ UDGothic）だけで、
候補にない書体は `fonts.<書体>.path` で指定する。`doctor` の結果の `font_directories` に、探したフォルダーが入る。
変換の前に、計測に使ったフォントと Windows に登録された書体の実体を照合し、違えば `windows_font_mismatch` になる。

### 変換の排他と待ち時間

PowerPoint の変換は Windows の利用者ごとに 1 つずつ行う。別の処理が変換中なら、生成は排他を待ち、待つ間は標準エラーに
「変換待ち: 別の処理が PowerPoint を使用中（実行 …、PID …、経過 N 秒）」を出す（PID は相手の Python の PID。分からなければ Windows の作業プロセスの PID と明記する）。
待つ上限は `--wait` か実行環境の `wait_seconds`（既定 600 秒）で、超えると `renderer_busy` で失敗する。生成は事前確認から最後の変換まで排他を持つ。

利用者の PowerPoint が起動していれば、その PowerPoint で変換し、利用者の資料を閉じず、PowerPoint も終了しない。
slidekit が起動した PowerPoint は、最後の変換の後に終了を待ってから排他を放す。
生成記録の `operations.sessions.powerpoint` には、`powerpoint_started`（接続の前後で PowerPoint の処理が増えたときだけ真）と、
`powershell_launches`（そのセッションが起動した PowerShell の数。作業プロセスと、セッションの中で初めて行った Windows への問い合わせ）が残る。
フォントのフォルダーを自動で探す生成では、探索がセッションの前に Windows へ問い合わせるので `powershell_launches` は 1 になり、`font_directories` を設定した生成では 2 になる（生成全体の起動はどちらも 2 回）。

### 時間切れと中断

1 回の変換が `timeout`（既定 90 秒）を超えると `renderer_timeout` で失敗する。時間切れと中断（端末の Ctrl+C を含む）の後は、切り離した作業プロセスが変換を終え、
開いた複製を閉じ、残存記録・使用中の記録・作業フォルダーを消し、自分で起動した PowerPoint を終了する。Python が先に終わっても後片付けは続き、終わるまで `renderer status` は使用中を示す。
準備が完了する前の中断では、作業プロセスもその場で終わる。利用者の PowerPoint は強制終了しない。

`build` は、生成の間に受けた SIGTERM と SIGHUP（端末を閉じたときなど）も Ctrl+C と同じ中断として扱う。差し替えるのは既定の処理のままの信号だけで、2 回目以降の信号は後片付けのために無視する。
作業プロセスが後片付けの前に終わった場合（作業プロセスが強制終了された場合など）や、起動した PowerPoint を終了できなかった場合は、PowerPoint と記録が残る。
`renderer status` で確かめ、`renderer clear --stop-owned` で解除する（次の節）。

### 残存記録の確認と解除

後片付けができずに残った記録（残存記録 `pending.json`）がある間は、新しい変換を始めない（`renderer_pending`）。

```sh
slidekit renderer status
slidekit renderer clear
slidekit renderer clear --stop-owned
```

- `renderer status` は、排他（`idle`・`busy`・`pending`）、残存記録、その作業プロセスが動いているか、検査用の複製を排他で開くことができないか、PowerPoint が起動しているかを示す。
  `abandoned` は、排他を持っていた処理が放さずに終わっていたことを示す（この確認で取って放すので、後には残らない。slidekit が起動した PowerPoint が残っていることがある）。
- `renderer status` は、PowerPoint の処理ごとに、PID・ウィンドウの題名・開始時刻と、slidekit が起動した PowerPoint の記録と一致するかを示す。
  作業プロセスは PowerPoint を起動すると、その PID と開始時刻を残存記録と使用中の記録の `powerpoint` に書く。PID と開始時刻の両方が一致した処理だけを、slidekit が起動したものとみなす。
  `--json` では `powerpoint_processes`（`pid`・`window_title`・`has_window`・`started_at`・`owned`・`record`・`stoppable`）と、`powerpoint_records`（記録と、その処理が動いているか `running`）に入る。
- `renderer clear` は、排他を取得したうえで、作業プロセスが動いておらず、検査用の複製を排他で開くことに成功したときに、残存記録を `cleared/` へ移し、作業フォルダーを消す。24 時間より古く、どの記録も指していない作業フォルダーも消す。
- PowerPoint は読み取り専用で開いた複製のファイルを開いたままにしないので、複製を排他で開くことができても、PowerPoint が複製を開いていないとは断定できない。
  PowerPoint が起動しているときは、PowerPoint の終了を待つか、PowerPoint が検査用の複製を開いていないことを確かめてから `renderer clear --force` を実行する。
  残存記録の変換に使った PowerPoint が slidekit の起動したものと記録されていて、もう動いていなければ、ほかの PowerPoint が起動していても `--force` は要らない。
- `renderer clear --stop-owned` は、作業プロセスが動いていないことを確かめた後、slidekit が起動したと記録で確かめることができ、ウィンドウのない PowerPoint だけを強制終了してから解除する。
  記録と一致してもウィンドウがある PowerPoint は、利用者が資料を開いていることがあるので、終了せずに断る。ほかの PowerPoint は終了しない。終了した PowerPoint は結果の `stopped_powerpoint` に入る。
  残存記録がなくても、前の処理が終了させずに残した PowerPoint（使用中の記録と一致するもの）は終了する。
- 記録がない場合（古い形式の記録か、PowerPoint を起動する前に止まった記録）と、既に起動していた PowerPoint を変換に使った場合は、`--stop-owned` でも終了せず、理由を示して断る。
  `renderer status` で PowerPoint の一覧を確かめ、PowerPoint の終了を待つか、複製を開いていないことを確かめてから `--force` を使う。
- 前の処理が終了させずに残した PowerPoint だけが動いていて、ウィンドウがなければ、次の生成はそれを引き継いで使い、最後に資料がなければ終了する。

### Python のスクリプトから呼ぶとき

スクリプトから `slidekit.cli.main()` や `slidekit.pipeline.build()` を呼ぶときは、`if __name__ == "__main__":` の中で呼ぶ。
互換画像とページ画像の作業プロセスは `__main__` を読み込み直す（起動方法は slidekit が選び、Python の版によらず、Windows 以外では forkserver、Windows では spawn を使う）。守らないと 2 つ目の生成が始まり、PowerPoint の排他を待ち合って `--wait` の上限まで止まる。
Windows の Python では、仮想環境の実行ファイルは `.venv\Scripts\` に置かれる。

## 0.2 への移行

slidekit 0.2 では計測・書き出し・記録の形式・コマンドを改めた。0.1 で作った資料は次の変化を確かめてから使う。
既存の資料は、次に完成物へ反映するときに全ページの確認をやり直す（古い形式の確認記録は引き継ぐことができない）。

| 変更点 | 影響 | 対処 |
|---|---|---|
| 文字を枠の幅で自動的に折り返す（既定 `word_wrap = true`） | 枠の幅が足りない要素が複数行になる。`equation()` は折り返さない | 1 行で収める要素は `word_wrap=False` にする。資料のテーマの `[defaults]` に `word_wrap = false` を書けば 0.1 と同じ |
| `context.measure_text()` の値（和欧文の境界の 0.25 em をなくした、12 pt 以上の kern、行間 1.0 の高さを書体の値で決める、欧文だけの行は欧文の書体だけで計測する、行末の空白を幅に含めない、どちらの書体にもない文字と制御文字はエラー） | 計測値から枠を決める資料コードの枠の幅と高さが変わる。行末の空白を計測して幅を加えていた枠は、空白の分だけ狭くなる | 表示を確かめる。寸法は `Rect(x, y, w, None)`・`slide.measure()` で内容から決める |
| `context.measure_math()` の値（数式の書体の MATH 表で計算する。`\text` を文字の書体で計測する、文中の数式の末尾の斜体補正、添字の位置、関数名の後と数字の間のカンマの空き、関数の適用・横書きの分数の前後の空きなど） | 数式の幅と高さが変わる。描画の見込みの倍率を掛けて枠を決めていた資料コードでは、倍率が要らなくなる | 倍率を外し、計測値で枠を決める |
| 数式だけの段落を、段落の揃えのとおりに置く（0.1 は揃えによらず中央に置いた）。明示改行で分けた数式だけの行も同じ。中央揃えの式は枠の中心にそろえる | 中央に置かれる前提で枠をずらしていた式が、左や右に寄る。中央揃えの式が少し右へ動く | 枠の左端を式の左端に合わせる。式の後に空白を置いて揃えていたなら空白を消してよい |
| 普通の括弧 `(`・`[` は伸びない | 背の高い中身の括弧が小さく表示される | 伸ばす括弧は `\left(`・`\right)` で書く |
| 字形：大文字のギリシャ文字は立体、`*` は ∗、`\epsilon`・`\phi` は ϵ・ϕ、`\mid` の前後に空き、`\bmod` の前後に空白、`:` は句読点 | 数式の見た目が変わる | 0.1 の字形にするなら `\varGamma`・`\varepsilon`・`\varphi` を使い、`\mid` の代わりに縦線の文字を書く |
| 大型演算子の引数は、次の二項演算・関係・句読点の前まで | `\sum`・`\int` の後の空きが変わる | 表示を確かめる |
| 未知の引数名・書式名、`z_index`、`parent_id`、`equation(text=)`、まとまりと `style=` の回転・反転はエラー | 0.1 で知らせずに無視された指定がエラーになる | 候補に示す名前に直す。数式は `expression` で書く |
| 役割とテーマの `[defaults.<種類>]` の `fill`・`line_color` などが `box()` に適用される（0.1 は `box()` の既定値が優先した） | 図形の塗りと枠線が変わることがある | 表示を確かめ、要素かテーマで指定する |
| 同梱のテーマで角丸の半径の既定を 6 pt にした（`[defaults.box]` の `corner_radius_pt`。0.1 は短い辺の約 1/6） | 角丸の長方形の角が小さくなり、文字領域が少し広がる | 表示を確かめる。大きくするなら、資料のテーマの `[defaults.box]` か要素の `corner_radius_pt` で指定する |
| 表：存在しないキー（`cell_roles` など）はエラー、戻り値は `Table` で `area` は実際の範囲、`auto` の行高は折り返しを反映する、まとまりの書式からセル用の書式を除く | 表の寸法と、戻り値を使う資料コードの座標が変わる | 戻り値の `area` を使う |
| 線・矢印を `anchor()` で結んだ端は PowerPoint の接続になる。箱の `dashed` は枠線を破線にする。条件を満たすまとまりと表は PowerPoint のグループになる | PowerPoint で動かしたときの動きと、図形の選び方が変わる | 接続とグループにしたくなければ座標で結ぶか `pptx_group=False` にする |
| 互換画像は、枠の外へはみ出した数式の描画と枠線を含めて切り出す。境界の薄い線がなくなった | 互換表示が標準表示に近づく | 互換表示を確かめ直す |
| 記号の書体：○ × ± “ ” などは和文の書体、→ … — などは欧文の書体で書き出す（PowerPoint の既定と同じ） | 0.1 と同じ見た目になる | なし |
| 実行環境の `temporary_dir`・`cache_dir`・`fonts.*.windows_path`・`windows_bold_path`、`deck.toml` の `size = {width, height}`、`notes_style` の `desu_masu`・`です・ます` は受け付けない | 設定の読み込みがエラーになる | 削除する。ノートの文体は `notes_style = "polite"`。寸法は `[幅, 高さ]` |
| 既定の書体を Windows のフォントのフォルダーから自動で探す | 実行環境に書体のパスを書かなくてよい | 候補にない書体だけを `fonts.<書体>.path` で指定する |
| Python 3.12 以上。数式の計測に matplotlib を使わない（追加依存 `math` は空） | Python 3.11 では導入できない | 3.12 以上の仮想環境を作り直す |
| 移行用の `slidekit.migration.add_legacy_elements()` と `Deck.bibliography_keys()` を削除した | 呼ぶ資料コードが動かない | 作成 API で要素を作る。文献は `deck.bibliography`・`deck.cited_keys()` を使う |
| コマンドの既定の出力は要約になった | JSON を読んでいた処理が動かない | `--json` を付ける |
| `build` の `status` は、エラーの診断があっても `built` | `status` で品質を判定していた処理が判定を間違える | 終了コードか `counts` を使う |
| `build` は前の実行から確認を自動で引き継ぐ | 手で `review carry` しなくてよい | 引き継がないなら `--carry-from none` |
| `review record` の `--slides` は `visual`・`content` で必須、記録・許容・取り消しに `latest` は使うことができない | 0.1 の手順のコマンドが拒否される | `--slides all` などを明示し、実行の番号を書く |
| 記録の形式を 2 にした。古い形式の実行には確認を記録できない | 0.1 の実行は確認・反映できない | 再生成して確認する。完成物と退避した旧版はそのまま扱うことができる |
| 採用記録と反映の履歴のパスは相対パス。ロックのファイルは実行保存先（`out/`）に置く | 資料のディレクトリを移しても確認を引き継ぐことができる | 資料のディレクトリに残った古い `.slidekit-export.lock` は使わないので、手で消してよい |
| `inputs.python` は任意 | 入口と読み込んだ補助ファイルは自動で追跡する | 書いたままでもよい |
| PDF の読み取りと画像化を PyMuPDF から pypdfium2（PDFium）・pypdf に替え、PDF の素材の `vector=True` を受け付けない | 画像に `vector=True` を指定した PDF の図は `invalid_image_option` で止まる。資料の既定 `[images] vector = true` は SVG の素材だけに当てはめ、PDF の図は PNG だけにする。前の実行の表示確認は引き継がない（画像化の版が変わる） | PDF の図を SVG に変換して置くか、`vector` を外して PNG にする。表示を確かめ直す |

## 開発と依存関係

| 追加依存 | 内容 |
|---|---|
| `svg` | SVG の素材を PNG にする（CairoSVG） |
| `bibliography` | 文献ファイルを読み、LaTeX 表記を直す（bibtexparser、pylatexenc） |
| `test` | 試験（pytest、pytest-xdist、reportlab） |
| `dev` | 整形（Black） |
| `math` | 空。旧版の導入コマンド（`.[math]`）が動くように名前だけを残している |

```sh
uv pip install -e '.[svg,bibliography,test,dev]'
python -m pytest tests -n auto
SLIDEKIT_TEST_POWERPOINT=1 python -m pytest tests -m powerpoint -p no:xdist
black --check src tests examples tools
```

- 通常の試験は PowerPoint を使わず、並列に実行できる。利用者の実行環境の設定と Windows のフォントを読まない（`tests/conftest.py`）。使用例を構築する試験（目印 `examples`）は、Windows の既定の書体がある環境でだけ実行する。
- PowerPoint の試験（目印 `powerpoint`）は直列に実行する。`SLIDEKIT_TEST_POWERPOINT=1` で `-n` を付けると、開始前にエラーになる。利用者の別の変換と重ならないときに実行する。
- 整形は Black（行の長さ 120）で行い、`black src tests examples tools` で直す。
- `requirements-lock.txt` は、全追加依存を検証した組み合わせである。そろえるときは `uv pip install -r requirements-lock.txt` の後に `uv pip install --no-deps -e .` を実行する。
  作り直すときは、ファイルの先頭の注記にある `uv pip compile` のコマンドを使う。作り直すと先頭の注記が消えるので、書き戻す（`tests/test_packaging.py` が注記と依存の固定を確かめる）。
- PDF の読み取りと画像化は pypdfium2（PDFium）・pypdf・fontTools で行う（`slidekit.pdf`）。試験の PDF は reportlab で作る。
- 和文の文節の境（`text_wrap_mid_phrase` と書式 `phrase_wrap`）は BudouX（Apache-2.0、依存なしの Python のパッケージ）の日本語のモデルで判定する。読み込むことができない環境では、文節の判定だけを省く。
- 必須の依存のライセンスは MIT・BSD・Apache-2.0 などである。追加依存 `svg` の CairoSVG は LGPL-3.0 以降で提供される。PowerPoint 本体とフォントは同梱しない。

### 照合データの作り直し

計測を PowerPoint の表示と照合する試験は、PowerPoint で変換した PDF から計測した照合データを使う。照合データと手元のフォントのファイルが違う環境では省略する。
PowerPoint・フォント・計測の版を変えたとき（`measure.MEASURE_VERSION` か数式の計測の版を上げたとき）に作り直し、差を確かめてから入れ替える。PowerPoint を使う作成は一度に 1 つだけ行う。

- 文字の照合データ `tests/data/layout/powerpoint_layout.json`（`MEASURE_VERSION` 9、PowerPoint 16.0.20527.20196、校正用の資料の 24 ページ・390 要素）：校正用の資料 `tests/data/layout/deck/` を継承する `deck.toml` を作業用の場所に置いて `slidekit build` し、
  `python tools/make_layout_fixture.py --run 実行ディレクトリ --output tests/data/layout/powerpoint_layout.json` で作る。試験は `tests/test_layout_fixture.py`（環境変数 `SLIDEKIT_LAYOUT_FIXTURE` で照合データを替えることができる）。
- 数式の照合データ `tests/data/powerpoint_math.json`（道具の 166 式を文中と独立で置いた 332 件）：`python tests/tools/powerpoint_math_data.py --output … --work …` で作り、`python -m slidekit.math.metrics` で計測と比べる（[照合データと誤差](math.md#照合データと誤差)）。

開発用の共通の入口 `tools/calibrate.py` は、設定ファイル（TOML）から照合用の資料を作り、`slidekit build` で変換して計測し、今の計測と既存の照合データとの差を示す。
設定の `deck` は、`cases`（式・文字サイズ・行間・文中か独立か・揃え・前後の文字の組み合わせ）、`layout`（校正用の資料）、`math`（数式の道具の式の一覧。`expressions` で絞ることができる）のどれかにする。

```toml
schema_version = 1
deck = "cases"

[defaults]
line_spacing = [1.0, 1.2]

[[cases]]
name = "script"
expressions = ["m^{1/4}", 'x^{\frac{a}{b}}']
size_pt = [12, 20, 32]
tags = ["text_baseline"]
```

```sh
python tools/calibrate.py run 設定.toml --work 作業ディレクトリ          # 資料を作り、変換して比べる
python tools/calibrate.py run --work 作業ディレクトリ --measure-only     # 変換せず、保存した PDF を計測し直す
python tools/calibrate.py adopt --work 作業ディレクトリ [--only 項目 …]  # 計測値を tests/data の照合データに採用する
```

作業ディレクトリは slidekit の外の新しいディレクトリにする。設定の写し、生成条件 `conditions.json`、実行 `out/runs/`（PDF と生成記録）、計測の結果 `results/実行番号/` が残る。
`run` の終了コードは、今の計測が許容差に入れば 0、違いがあれば 1、照合できなければ 2 である。
採用は `layout` と `math` の資料だけで行う。`cases` の要素は、校正用の資料か数式の道具の式の一覧に加えてから作り直す。`--only` は、計測の環境が照合データと同じときに既存の項目だけを置き換える。

`tools/layout_overlay.py` は、実行の標準表示の PDF の画像に、要素の予測（`layout.json`。`--remeasure` では今の計測）と PDF から読んだ基線・行の範囲・インクを重ね、差を pt と em で示す。
`python tools/layout_overlay.py --run 実行ディレクトリ --element ページ/要素 --output 出力先` のように使い、`--run` と `--element` は複数指定できる。配置を直すときの補助であり、表示確認の記録には使わない。

照合ケースごとに厳しく比べる項目と許容差は `tests/data/layout/tolerances.toml` に書く。規則はタグか要素で選び、基線（`baseline`）・行の幅と左端（`width`）・インクの左右（`ink_x`）・インクの上下（`ink_y`）の許容差を `tolerance_pt` か `tolerance_em` で決め、根拠を `reason` に書く。

照合データと校正用の資料は公開してよい内容の試験の対象になる。ローカルのパスや非公開の語を入れない（`tools/make_layout_fixture.py` はパスに見える文字列があれば書かずに止まる）。

### 公開してよい内容の試験

`tests/test_examples.py` は、使用例・文書・`src`・`tests`（試験用のデータを含む）・`tools` と、使用例の完成物の文書情報と文字に、利用者のホームなどのローカルのパスがないことを確かめる。
環境変数 `SLIDEKIT_PRIVATE_TERMS` が 1 行 1 語の UTF-8 のファイルを指すときは、同じ範囲にその語がないことも確かめる。語の一覧はリポジトリに置かない。
