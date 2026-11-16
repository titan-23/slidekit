---
name: slidekit
description: slidekit（Python で内容と配置を書き、編集可能な PowerPoint 資料を生成・検査するライブラリ）で、スライドを新しく作る、直す、確認して完成物へ反映するときに使う。deck.toml・slides.py・slidekit コマンド（layout・build・review・export）を扱う作業、日本語の技術資料、数式・図・表のある資料の作成が対象。
---

# slidekit で資料を作る

資料を Python のコード（`slides.py`）で書き、`slidekit` コマンドで PPTX・PDF・確認用の画像を生成する。
文字の折り返しとはみ出しは PowerPoint の配置に合わせて計算し、最後は Windows の PowerPoint で変換した PDF で確かめる。
この文書は作業に要る要点だけをまとめる。細部は末尾の文書を、必要な節だけ開いて読む。

## 前提

- Python 3.12 以上と Windows 版 PowerPoint が要る（WSL で使う）。コマンドは slidekit の仮想環境で実行する（`source .venv/bin/activate`）。
- 初回と、フォント・PowerPoint・実行環境の設定を変えたときは `slidekit doctor --project <deck.toml>` で確かめる。
- PowerPoint を使うコマンド（`build`・`doctor`）は 1 つずつ順に実行する。利用者の PowerPoint や資料を閉じない。

## 作業の流れ

| 段階 | コマンド | 内容 |
|---|---|---|
| 1 | `slidekit layout --project P [--wraps]` | 編集のたびに実行する。表示ソフトを使わず、配置と記述規則の誤りを数秒で示す。ファイルを書かない。`--wraps` は自動で折り返した行の文字列を示す |
| 2 | `slidekit build --draft --project P --slides SPEC` | 見た目を確かめる下書き。互換画像を省き、エラーがあれば変換の前に止まる。確認と完成物の対象外 |
| 3 | `slidekit build --project P` | 全ページの通常の生成。前の実行から、変わっていないページの確認を自動で引き継ぐ |
| 4 | `slidekit review status --project P` | 残りの確認と、記録に使うコマンドを示す。`review sheet` で未確認のページの比較画像を作る |
| 5 | 画像を開いて確かめ、`slidekit review record ...` | 実際に見た範囲だけを記録する。すべて見たら `--pending` で 1 回で記録できる |
| 6 | `slidekit export --project P --build latest` | 最終検査に通った実行を完成物（PPTX・PDF・ノート）へ反映する |

- `P` は `deck.toml` かそのディレクトリである。`SPEC` はページ番号（`3`）、範囲（`5-7`）、識別子（`intro`）のカンマ区切り。
- 終了コードは、成功が 0、品質上の未解決事項（エラーの診断、未確認、未対応の警告）が 1、入力・環境・変換の失敗が 2 である。どのコマンドも、要約の最後の行が「結果: 合格（エラー 0、終了コード 0）」の形の合否の行になる。合否は末尾の 1 行で読む。
- 許容した警告と生成の参考は件数だけが出る。一覧が要るときは `--verbose` を付ける。
- 機械が読むときは `--json` を付ける。要約で足りるときは付けない（出力が長くなる）。
- 「次の操作」のコマンドは起動した方法（`slidekit`、絶対パス、`python -m slidekit`）に合わせて出るので、そのまま使う。
- 生成した PPTX・JSON・`review.json` を手で直さない。直すのは `slides.py` と設定だけである。

### 確認と記録

- `build` は `<出力先>/runs/NNNN/` を作る。画像は `renders/powerpoint/standard/` の `contact-NNN.png`（一覧画像）と `slide-NNN.png`、互換表示は `renders/powerpoint/fallback/slide-NNN.png` にある。
- 一覧画像でページ順と全体の配置を見てから、`review status` が示す未確認のページの画像だけを開く。文字の欠け・はみ出し・重なり・数式・線の接続・表を確かめる。
- `build` の要約の「前の実行 NNNN との比較」は、変わっていないページ・ページ番号だけ変わったページ・変わったページを示す。画像を開くのは変わったページと新しいページだけでよい。
- `slidekit review sheet --project P --build N` は、未確認のページを前の実行の画像・互換表示の画像と横に並べ、差のある領域を赤い枠で示した比較画像（`runs/NNNN/sheets/`）を作る。1 枚で標準表示と互換表示を見ることができる。比較画像を見ただけで記録しない。
- 引き継がなかった確認の理由（本文の変更・元資料の変更・表示環境の変更など）は、`build` と `review status` の要約に出る。
- 記録のコマンドの例を次に示す（`--build` には実行の番号を書く。`latest` は受け付けない）。

```sh
slidekit review record --project P --build 4 --kind visual --variant all --slides all --reviewer "確認者" --note "全ページの文字・図表・数式を確認"
slidekit review record --project P --build 4 --kind overview --reviewer "確認者" --note "ページ順・全体の配置・文書情報を確認"
slidekit review record --project P --build 4 --kind content --slides all --reviewer "確認者" --note "依頼内容と本文・ノートを照合"
```

- 未確認をすべて見た後は、`review status` の「次の操作」の先頭にある `slidekit review record --build N --pending --status-id s_… --reviewer "確認者" --note "…"` で、表示・一覧画像・内容をまとめて記録できる（種類ごとの注記は `--note-visual`・`--note-overview`・`--note-content`）。記録の結果に残りの件数が出る。
- `review status` が示すコマンドの `<確認者>`・`<確認した内容>` は、見た後に書き換える。機械検査の成功を、目で見た確認として記録しない。画像を開くことができないときは記録せず、利用者に伝える。
- `export` が最終検査で止まったときは、足りない確認のページと記録のコマンドが要約に出る。
- 意図した警告は資料のコードで `slide.allow_warning(...)` を宣言するか、`slidekit review accept --project P --build N --code CODE --slides SPEC --reviewer R --reason "理由"` で許容する。エラーは許容できない。

### トークンを節約する

- まずこの文書だけで書き始め、詳しい文書は見出しを grep して必要な節だけを読む。
- 直すたびに `build` せず、`layout` の文字の診断で直す。画像を見るのは `build` の後だけにする。
- 折り返しの位置は `layout --wraps --slides <ページ>` の行の文字列で確かめる（文節の途中・ハイフンの後の折り返しに印が付く）。画像を開かずに判断できる。
- 文が幅に収まるかは `slidekit measure --project P --width 618 "文"` か `slide.measure()` で、資料を作り直さずに確かめる。
- 見た目を確かめるときは `build --draft --slides <直したページ>` にする。
- 確認では一覧画像の後に未確認のページだけを開く。`deck.json`・`layout.json` を丸ごと表示しない。

## 新しい資料

資料のディレクトリの直下に完成物と README、`src/` に編集元を置く。コマンドには `src/deck.toml` を渡す。

```text
seminar/
├── README.md
├── seminar.pptx・seminar.pdf・speaker_notes.md   export が書く完成物
└── src/
    ├── deck.toml・slides.py・theme.toml・rules.toml
    ├── brief.md        目的・対象者・必須事項（任意）
    ├── figures/        素材
    └── out/            実行ディレクトリ（build が作る）
```

```toml
# src/deck.toml
schema_version = 1
size = "16:9"            # 960×540 pt。"4:3" は 720×540 pt
theme = "theme.toml"     # 同梱の default に重ねる（任意）
rules = "rules.toml"     # 同梱の general に重ねる（任意）

[metadata]
title = "資料の題名"
author = "作成者名"            # 必須。匿名なら空文字列を明示する
last_modified_by = "作成者名"  # 必須

[deliverables]
pptx = "../seminar.pptx"
pdf = "../seminar.pdf"
notes = "../speaker_notes.md"
record = "deliverables.json"
```

- テーマ（`theme.toml`）: `[colors]`（`background`・`text`・`primary`・`accent`・`muted`・`line`・`surface`）、`[fonts]`、`[margins]`、`[defaults]`・`[defaults.<種類>]`、`[roles.<名前>]`、`[page_kinds.<種別>]`。
- 記述規則（`rules.toml`）: `require_notes = true`、`notes_style = "polite"`（ノートをです・ます調にする）、`forbidden_characters`、`[math_candidates]`（数式にすべき文字を警告する）。
- 雛形の PPTX は `[template] path = "template.pptx"` で指定する（テーマとマスターを引き継ぐ。既存のページの内容は取り込まない）。
- 数式の互換画像（PowerPoint 以外のソフト向け）は `[math] fallback = "required"` で付ける。既定の `auto` では LibreOffice を必須にしない資料に付けない。

## slides.py の書き方

座標と文字サイズは pt である。`Rect(x, y, w, h)` の幅か高さを `None` にすると内容から決まる。

```python
from slidekit import Deck, Math, Rect, Text, key_id
from slidekit.markup import parse


def build(context):
    deck = Deck(context)
    deck.slide(id="cover", kind="cover", title="二分探索の紹介", notes="二分探索の考え方を紹介します。")

    slide = deck.slide(id="idea", title="範囲を半分にする", notes="比較のたびに探す範囲を半分にします。")
    stack = slide.stack(slide.content, gap=16)  # 本文領域の上から順に積む
    stack.text(id="lead", text=parse(r"昇順の配列で、中央の値 $a_m$ と $x$ を比べる"))
    stack.equation(id="count", expression=r"\lfloor \log_2 n \rfloor + 1", place="center", width=360)
    left, right = stack.remaining.columns(2, gap=48)
    ask = slide.box(id="ask", area=Rect(left.x, left.y, left.w, None), text="中央と比べる",
                    shape="rounded_rectangle", align="center")
    done = slide.box(id="done", area=Rect(right.x, right.y, right.w, None),
                     text=[Text("範囲が空なら終了 "), Math("l = r")], align="center")
    slide.arrow(id="next", start=ask.anchor("right"), end=done.anchor("left"))  # PowerPoint の接続になる
    return deck
```

### 要素を作る関数

| 関数 | 必須の引数 | 主な任意の引数 |
|---|---|---|
| `deck.slide` | `id` | `kind`（`cover`・`section`・`content`・`references`・`blank`）、`title`、`notes`、`seconds`、`allow_roles`、`decorations` |
| `slide.text` | `id`、`text`、`area` | `role`、書式、`sources` |
| `slide.box` | `id`、`area` | `text`、`shape`（`rectangle`・`rounded_rectangle`・`ellipse`・`diamond`）、`role`、書式、`tail_toward`（吹き出しのしっぽを向ける点。別の要素の `anchor()` か `(x, y)`） |
| `slide.equation` | `id`、`area`、`expression` | `align`、書式。数式だけなら独立した数式になる |
| `slide.line`・`slide.arrow` | `id`、`start`、`end` | `line_color`・`line_width_pt`・`dashed`。端は `(x, y)` か `要素.anchor(side, gap=0, at=0.5)` |
| `slide.image` | `id`、`path`、`area` | `fit`（`contain`・`cover`）、`page`（PDF は必須。1 始まり）、`vector`（SVG だけ。PDF は PNG にするので、劣化させたくない図は SVG にする） |
| `slide.table` | `id`、`rows`、`row_keys`、`column_keys`、`area` | `headers`、`column_widths`・`row_heights`（`"auto"` で内容から）、`column_styles`・`cell_styles`、`band_fills` |
| `slide.group` | `id` | `area`。戻り値の同じ名前の関数で子を作る（座標はまとまりの左上が原点、識別子は `まとまり/子`） |
| `slide.references` | `id` | `keys`（既定は出典の文献キー）。`inputs.bibliography` の `.bib` が要る |
| `slide.stack` | `area` | `gap`、`avoid`（`"footer"` か `Rect` の列。ページ番号の領域などを避けて幅を縮める）。`text`・`box`・`equation`・`image`・`table` は `height`・`width`・`indent`・`place`・`gap` を受け取る。`take(h)`・`space(h)`・`remaining`・`cursor`。1 行を加えても下の要素の y を計算し直さなくてよい |
| `slide.measure` | `text` | `kind`・`width`・`role`。要素を作らずに寸法（`height`・`width`）と行（`lines` の `text`・`width`）を求める |
| `slide.allow_warning` | `code`、`reason` | `element`、`related`（重なりの警告では `element` が必須） |

- 領域の計算: `slide.content`（本文領域）、`Rect` の `inset()`・`columns(n か重みの列, gap=)`・`rows()`・`translated()`・`right`・`bottom`・`center`、`align()`・`distribute()`。作った要素の `area` は数値になり、`move(dx, dy)` で動かすことができる。
- 識別子: 編集後も意味が変わらない名前を付ける。ページの中と表の中で重複させない。繰り返す要素は番号でなく `key_id("step", key)` で作る。
- 描画順は作成順である。枠や背景を先に、文字を後に作る。

### 文字と書式

- 文章は文字列か、`Text`・`Emphasis`・`Math` の断片の列で書く。`parse(r"...")` は `$…$` を数式、`**…**` を強調にする（明示して呼ぶ）。脚注の番号などの上付きは `Text("1", superscript=True)` で書く（数式の `^{1}` は直前の式の指数になる）。
- 和文の段落は `phrase_wrap=True`（書式。テーマの `[defaults]` にも書ける）にすると文節の境で折り返し、手で `\n` を入れなくてよい。
- 文字は枠の幅で自動的に折り返す。自動では縮小しない（はみ出すと `text_overflow`）。1 行で収めたい要素は `word_wrap=False` にする。改行 `\n` は意味の区切りだけに入れる。
- 役割（`role`）: `slide_title`・`heading`（24 pt）・`body`（20 pt、既定）・`note`（14 pt）など。各役割に最小文字サイズがあり、下回るとエラーになる。小さい文字は最小値の低い役割をテーマで定義し、`restricted = true` と `allow_roles` で使う場所を限る。
- 書式の引数: `size_pt`・`color`・`bold`・`align`（`left`・`center`・`right`・`justify`）・`valign`（`top`・`middle`・`bottom`）・`line_spacing`・`padding_pt`・`fill`・`line_color`・`line_width_pt`・`dashed`・`corner_radius_pt`（角丸の箱だけ。同梱テーマの既定は 6）。色はテーマの色の名前か 6 桁の 16 進数で書く。
- 吹き出し: `slide.box` の書式 `tail_side`（`left`・`right`・`top`・`bottom`。`none` で役割のしっぽを取り消す）で、本体の辺にしっぽを付ける。`tail_at`（辺の上の位置 0〜1、既定 0.5）・`tail_width_pt`（既定 10）・`tail_length_pt`（既定 8）。`rectangle`・`rounded_rectangle` だけに付き、`slide.text` には付かない。`area` は本体で、しっぽは外へ出る。しっぽが相手に重なると `element_overlap` になるので、本体と相手の間を `tail_length_pt` 以上あけるか、`tail_toward` で相手を指す。しっぽのある辺には線を結ばない。役割に書けば役割だけで吹き出しにできる。
- 重ねる順は「テーマの `[defaults]` < `[defaults.<種類>]` < 役割 < `style=` の辞書 < キーワード引数」で、後ろほど優先する。

### 数式

- 記法は LaTeX に似ている。`$` で囲まず、式の中に改行を書かない。日本語は `\text{…}` の中に書く。`\text{…}` の中の `-` はハイフンとして書き出す（`\mathrm{delete}\text{-}\mathrm{min}`）。
- 普通の括弧は伸びない。伸ばすときは `\left(`・`\right)` で囲む。
- `/` は横書きの分数、`\frac{a}{b}` は上下の分数。`log`・`max` などは `\` がなくても立体の関数名になる。
- `equation()` は既定で折り返さない。文中の数式は途中で改行せず、式ごと次の行へ送る。
- 数式を含む要素は回転・反転できない。数式を含むまとまりは PowerPoint のグループにならない。

### 表の例

```python
slide.table(
    id="methods",
    headers=["方法", "最悪の比較回数"],
    rows=[["線形探索", Math("n")], ["二分探索", Math(r"\lfloor \log_2 n \rfloor + 1")]],
    row_keys=["linear", "binary"],
    column_keys=["method", "count"],
    area=Rect(slide.content.x, slide.content.y, None, None),
    column_widths="auto",
    row_heights="auto",
    column_styles={"count": {"align": "center"}},
)
```

セルの識別子は `表/行キー/列キー`（見出しの行キーは `_header`）で、戻り値の `cell(行キー, 列キー)` で取り出す。

## よく出る診断

| 診断 | 直し方 |
|---|---|
| `text_overflow`（エラー） | 領域を広げるか文章を短くする。高さは `Rect(x, y, w, None)` か `stack` で内容から決める。要約の「1 行に収めるには内側の幅 N pt が要る」が直す量の目安 |
| `page_overflow`（エラー）・`content_area_overflow`（警告） | `area` を直す。本文は `slide.content` の中に置く |
| `layout_bounds_overflow`（エラー） | 高さを決めた `stack` に収まらない。領域を広げるか要素を減らす |
| `font_too_small`（エラー） | 文字サイズを上げるか、最小値の低い役割を使う |
| `element_overlap`（警告） | 位置を直す。意図した重なりは `anchor()` で接続するか `allow_warning` で宣言する |
| `text_padding_overlap`（警告） | 余白を小さくするか領域を広げる |
| `line_break`（警告） | 明示改行の位置を直す（行頭に句読点・閉じ括弧が来ている） |
| `math_not_marked`（警告） | 本文の変数や数字を `Math` で書く。書誌など役割ごとに外すときはテーマの `[roles.<名前>] skip_checks = ["math_not_marked"]` |
| `notes_missing`（エラー） | ノートを書く |
| `pdf_line_count_mismatch`（警告） | 計測と PowerPoint で行数が違う。枠を広げるか改行を明示し、表示を確かめる |
| `text_short_line`・`text_wrap_mid_phrase`・`text_wrap_hyphen`（参考） | 短い最後の行・文節の途中・ハイフンの後の折り返し。改行を明示するか `phrase_wrap=True` にするか、幅か文言を変える |
| `unknown_argument`・`invalid_style`（エラー） | 示された候補の名前に直す |

診断は `layout`・`build` の要約に 1 件 1 行で出る。全件の判定と直し方は [docs/diagnostics.md](docs/diagnostics.md) にある。

## 詳しい文書

- [docs/usage.md](docs/usage.md): 設定、作成 API、文字の配置、表・線・まとまり、数式の互換画像、生成と確認の手順、実行環境
- [docs/reference.md](docs/reference.md): 書式の引数、要素を作る関数、`deck.toml`、テーマ、記述規則、`environment.toml` の全項目
- [docs/math.md](docs/math.md): 数式の記法と対応する命令の一覧
- [workflows/authoring.md](workflows/authoring.md): 資料を作るときの作業規則（修正は原則 3 回まで、確認と報告の仕方）
- 使用例: [examples/proposal](examples/proposal/)（4 枚、二列の図・数式・表）、[examples/binary_search](examples/binary_search/)（10 枚、雛形・分岐図・文献一覧など）
