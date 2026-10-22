# 設定と引数の一覧

書式の引数、要素を作る関数、`deck.toml`、テーマ、記述規則、実行環境（`environment.toml`）の項目をまとめる。
手順と考え方は [利用ガイド](usage.md)、診断は [診断の一覧](diagnostics.md)、数式の記法は [数式の書き方](math.md) にある。
表の項目は実装の一覧と照合している（`tests/test_documents.py`）。未知の名前と不正な値はエラーになり、近い名前を候補として示す。

## 書式の引数

要素を作る関数のキーワード引数か `style=` の辞書で指定する。テーマの `[defaults]`・`[defaults.<種類>]`・役割にも同じ名前で書くことができる。
「対象の要素」の `table_cell` は表の本文のセル、`table_header` は見出しのセルである。

| 引数 | 対象の要素 | 意味 | 値・単位 | 指定がないとき |
|---|---|---|---|---|
| `size_pt` | `text`・`box`・`equation`・`table_cell`・`table_header` | 文字サイズ | 正の数（pt） | 役割の値 |
| `min_size_pt` | `text`・`box`・`equation`・`table_cell`・`table_header` | 最小文字サイズ。要素からは変更できず、役割と違う値はエラー（`minimum_override`） | 正の数（pt） | 役割の値 |
| `color` | `text`・`box`・`equation`・`table_cell`・`table_header` | 文字の色 | 色の名前か 6 桁の 16 進数 | `text` |
| `emphasis_color` | `text`・`box`・`equation`・`table_cell`・`table_header` | `Emphasis` と `Math(..., emphasis=True)` の色 | 色 | `accent` |
| `bold` | `text`・`box`・`equation`・`table_cell`・`table_header` | 太字 | 真偽値 | `false` |
| `align` | `text`・`box`・`equation`・`table_cell`・`table_header` | 横の揃え | `left`・`center`・`right`・`justify` | `left` |
| `valign` | `text`・`box`・`equation`・`table_cell`・`table_header` | 縦の揃え | `top`・`middle`・`bottom` | `top` |
| `line_spacing` | `text`・`box`・`equation`・`table_cell`・`table_header` | 行間（文字サイズの 1.2 倍を 1 とする倍率） | 正の数 | `1.15` |
| `line_spacing_pt` | `text`・`box`・`equation`・`table_cell`・`table_header` | 固定の行送り。指定すると `line_spacing` より優先する | 正の数（pt） | なし |
| `padding_pt` | `text`・`box`・`equation`・`table_cell`・`table_header` | 枠と文字の間の余白 | 0 以上の数（pt）か、`left`・`top`・`right`・`bottom` の一部を持つ辞書 | `0`（同梱テーマの `box` と表のセルは `6`） |
| `font_japanese` | `text`・`box`・`equation`・`table_cell`・`table_header` | 和文の書体 | 書体名 | テーマの `fonts.japanese` |
| `font_latin` | `text`・`box`・`equation`・`table_cell`・`table_header` | 欧文の書体 | 書体名 | テーマの `fonts.latin` |
| `font_math` | `text`・`box`・`equation`・`table_cell`・`table_header` | 数式の書体（MATH 表が要る） | 書体名 | テーマの `fonts.math` |
| `word_wrap` | `text`・`box`・`equation`・`table_cell`・`table_header` | 枠の幅で自動的に折り返すか | 真偽値 | `true`（同梱テーマの `equation` は `false`） |
| `fill` | `text`・`box`・`equation`・`table_cell`・`table_header` | 塗りの色。`"none"` と `None` は塗らない | 色か `"none"` | 塗らない（同梱テーマの `box` は `surface`、本文のセルは `background`、見出しは `surface`） |
| `line_color` | `text`・`box`・`equation`・`table_cell`・`table_header`・`line`・`arrow` | 枠線・線の色。文字枠では `"none"` と `None` は線なし | 色か `"none"` | 文字枠は線なし（同梱テーマの `box` と表のセルは `line`）、線と矢印は `color` |
| `line_width_pt` | `text`・`box`・`equation`・`table_cell`・`table_header`・`line`・`arrow` | 枠線・線の太さ | 0 以上の数（pt） | `1` |
| `dashed` | `text`・`box`・`equation`・`table_cell`・`table_header`・`line`・`arrow` | 破線にする。文字枠は枠線の色があるときだけ適用される | 真偽値 | `false` |
| `corner_radius_pt` | `box` | 角丸の半径。`shape="rounded_rectangle"` だけに指定できる | 0 以上の数（pt） | ライブラリは図形の既定（短い辺の約 1/6）、同梱テーマは `6` |

重ねる順は「ライブラリの既定 < テーマの `[defaults]` < `[defaults.<種類>]` < 役割 < 要素の `style=` < キーワード引数」で、後ろほど優先する。
`[defaults.<種類>]` は、表の本文のセルでは `box` → `table_cell`、見出しのセルでは `box` → `table_cell` → `table_header`、矢印では `line` → `arrow` の順に重ねる。
`min_size_pt` は常に役割の値になる。値の `None` は「指定しない」で、下の層の値を使う。ただし文字枠を持つ要素の `fill`・`line_color` の `None` は `"none"` と同じである。
`padding_pt` の辞書は指定した辺だけを上書きし、指定しない辺は下の層の値を引き継ぐ。

画像とまとまりには書式の引数がない。表の書式の引数（`style=` とキーワード引数）は、すべてのセルの書式（表全体の層）になる（[表](usage.md#表)）。
次の名前は受け付けない：`z_index`（描画順は作成順）、`italic`、`shadow`、`autofit`・`shrink_to_fit`（文字は自動で縮小しない）、
`underline`・`link`（断片の引数で指定する）、`parent_id`（まとまりの戻り値から子を作る）、`font`・`font_family`（`font_japanese` などを使う）。

## 要素の属性

書式でない引数である。まとまり（表を含む）は `description` と `allow_roles` だけを受け付ける。

| 属性 | 意味 | 値 |
|---|---|---|
| `description` | 代替テキスト（PowerPoint の説明） | 文字列 |
| `semantic` | 要素の意味。`background`・`decoration` を付けた要素は重なりの検査から外す | `background`・`decoration` |
| `rotation` | 時計回りの回転 | 角度（度）。数式を含む要素には指定できない |
| `flip_h` | 左右の反転 | 真偽値。数式を含む要素には指定できない |
| `flip_v` | 上下の反転 | 真偽値。数式を含む要素には指定できない |
| `allow_roles` | この要素（まとまりなら子孫も）で使う、使う場所を限った役割の宣言 | 役割名の列 |

すべての要素は `id`（必須）、`role`、`sources`（出典。元資料の `document_id` か文献の `bib_key` を持つ辞書、またはその識別子の文字列、その列）、`style=`（書式の辞書）を受け付ける。

## 要素を作る関数

座標と寸法は pt である。`area` は `Rect(x, y, w, h)` か 4 つの数の列で、幅か高さの `None` は内容から決める（[寸法と配置を決める](usage.md#寸法と配置を決める)）。

| 関数 | 必須の引数 | 任意の引数 | 戻り値 |
|---|---|---|---|
| `deck.slide` | `id` | `kind`（`cover`・`section`・`content`・`references`・`blank`、既定 `content`）、`title`（文字列・断片・断片の列。`None` は題名なし）、`notes`（`None` は空）、`seconds`（説明時間）、`title_role`、`title_area`、`title_style`、`decorations`（偽でページ種別の飾りを置かない）、`allow_roles` | ページ |
| `slide.text` | `id`、`text`、`area` | `role`、書式、属性、`sources` | 要素 |
| `slide.box` | `id`、`area` | `text`、`shape`（`rectangle`・`rounded_rectangle`・`ellipse`・`diamond`、既定 `rectangle`）、`role`、書式、属性、`sources` | 要素 |
| `slide.equation` | `id`、`area`、`expression` | `role`（既定はテーマの役割 `equation`）、書式、属性、`sources` | 要素 |
| `slide.line`・`slide.arrow` | `id`、`start`、`end`（`(x, y)` か要素の `anchor()`） | 線の書式、属性、`sources` | 要素 |
| `slide.image` | `id`、`path`、`area` | `fit`（`contain`・`cover`、既定 `contain`）、`page`（PDF のページ。1 始まり。PDF では必須）、`raster_scale`、`vector`、属性、`sources` | 要素 |
| `slide.table` | `id`、`rows`、`row_keys`、`column_keys`、`area` | [表](usage.md#表)の引数、書式（すべてのセル。表全体の層） | `Table` |
| `slide.group` | `id` | `area`（既定は大きさ 0）、`role`、`pptx_group`（`None`・`True`・`False`）、`description`、`allow_roles`、`sources` | まとまり |
| `slide.references` | `id` | `keys`（既定は出典の文献キー）、`area`（既定は本文領域）、`formatter`、`gap_pt`（既定 8）、`role`、`pptx_group`（既定 `False`）、書式 | まとまり |
| `slide.stack` | `area` | `gap`（既定 12） | 積み上げの補助 |
| `slide.measure` | `text` | `kind`（`text`・`box`・`equation`）、`width`、`role`、`shape`、`style`、書式 | `TextLayout` |
| `slide.allow_warning` | `code`、`reason` | `element`、`related` | なし |

まとまりの戻り値は `text`・`box`・`equation`・`image`・`line`・`arrow`・`table`・`group`・`references`・`stack`・`measure` を持ち、子をまとまりの左上を原点とする座標で作る。
子の識別子は「まとまりの識別子/子の識別子」になる。積み上げの補助（`stack`）の `text`・`box`・`equation`・`image`・`table` は、`area` の代わりに `height`・`width`・`indent`・`place`（`left`・`center`・`right`）・`gap` を受け取る。
`image` の `raster_scale` は PDF・SVG を PNG にするときの配置寸法 1 pt あたりの画素数（1 以上 16 以下）で、PDF・SVG の素材だけに指定できる。
`vector` は SVG の素材を SVG でも埋め込むかで、PDF・SVG の素材だけに指定できる。PDF の素材には真を指定できない（PNG だけにする）。

公開の名前：`slidekit` の `Deck`・`Rect`・`Text`・`Emphasis`・`Math`・`Table`（表の戻り値の型）・`align`・`distribute`・`key_id`、`slidekit.markup` の `parse`。
生成関数が受け取る `context`（`Project`）は、`selection`・`selected()`・`bibliography`・`fonts`・`measure_text()`・`measure_math()` を持つ。
`Deck` は `slide()`・`cited_keys()`（出典の文献キーを初めて現れた順に返す）・`bibliography`（文献）を持つ。
要素は `id`・`area`（代入できる）・`anchor(side, gap=0, at=0.5)`・`measure()`・`move(dx, dy)` を持ち、`Table` はさらに `cell(行キー, 列キー)`・`row_keys`・`column_keys`・`column_widths`・`row_heights` を持つ。
ページは `content`（本文領域）・`default_role`・`title` を持つ。

## 資料の設定（deck.toml）

相対パスはその設定を書いた `deck.toml` の場所を基準にする。`inputs.documents` は `id` と `path` を持つ表の配列である。

| 設定 | 値 | 既定値 | 内容 |
|---|---|---|---|
| `schema_version` | `1` | `1` | 設定の形式 |
| `extends` | `deck.toml` かそのディレクトリのパス | なし | 継承元の設定。`output_dir`・`deliverables` は引き継がない |
| `entrypoint` | `ファイル:関数` | `"slides.py:build"` | 生成関数 |
| `size` | `"16:9"`・`"4:3"`・`[幅, 高さ]`（pt） | `"16:9"` | ページ寸法。`16:9` は 960×540 pt、`4:3` は 720×540 pt |
| `theme` | 同梱の設定名かファイルのパス | `"default"` | テーマ。同梱の `default` に重ねる |
| `rules` | 同梱の設定名かファイルのパス | `"general"` | 記述規則。同梱の `general` に重ねる |
| `output_dir` | パス | `"out"` | 実行保存先 |
| `environment_override` | パス | なし | この資料だけの実行環境の設定。共通の設定に重ねる |
| `metadata.title` | 文字列 | 必須 | 題名 |
| `metadata.author` | 文字列 | 必須 | 作成者。匿名なら空文字列を明示する |
| `metadata.last_modified_by` | 文字列 | 必須 | 最終更新者 |
| `metadata.subject` | 文字列 | なし | 件名 |
| `metadata.keywords` | 文字列 | なし | キーワード |
| `metadata.comments` | 文字列 | なし | コメント |
| `metadata.category` | 文字列 | なし | 分類 |
| `metadata.content_status` | 文字列 | なし | 状態 |
| `metadata.identifier` | 文字列 | なし | 識別子 |
| `metadata.language` | 文字列 | なし | 言語 |
| `metadata.version` | 文字列 | なし | 版 |
| `inputs.read_roots` | パスの配列 | `["."]` | 作成 API が読む入力の範囲。継承したときは継承元の `deck.toml` のディレクトリも加わる |
| `inputs.python` | パスの配列 | `[]` | 任意。生成関数から読み込まれない Python ファイルを追跡するときに書く。入口と、構築の間に読み込まれた範囲内の Python ファイルは自動で追跡する |
| `inputs.documents` | `id` と `path` を持つ表の配列 | `[]` | 元資料。出典の `document_id` で参照する |
| `inputs.bibliography` | `.bib` のパスの配列 | `[]` | 文献ファイル。追加依存 `bibliography` が要る |
| `inputs.asset_dirs` | パスの配列 | `[]` | 素材のディレクトリ |
| `render.primary` | `powerpoint`・`libreoffice` | `"powerpoint"` | 主表示ソフト。下書きと `deck.pdf` に使う |
| `render.required` | 表示ソフトの配列 | `["powerpoint"]` | 表示を確かめる表示ソフト |
| `math.fallback` | `auto`・`required`・`off` | `"auto"` | 数式の互換画像を付けるか（[雛形と数式の互換画像](usage.md#雛形と数式の互換画像)） |
| `math.fallback_renderer` | `"powerpoint"` | `"powerpoint"` | 互換画像を作る表示ソフト。`powerpoint` だけを受け付ける（旧版の設定のために残す） |
| `math.fallback_scale` | 1 より大きく 8 以下の数 | `4` | 互換画像の倍率（1 pt あたりの画素数） |
| `math.functions` | 英字だけの名前の配列 | `[]` | 資料に固有の関数名（立体の関数名にする） |
| `images.raster_scale` | 1 以上 16 以下の数 | `2` | PDF・SVG を PNG にするときの 1 pt あたりの画素数 |
| `images.vector` | 真偽値 | `false` | SVG の素材を SVG でも埋め込む（拡大しても劣化しない）。PDF の素材には当てはめない |
| `deliverables.pptx` | パス | `"presentation.pptx"` | 完成物の PPTX |
| `deliverables.pdf` | パス | `"presentation.pdf"` | 完成物の PDF |
| `deliverables.notes` | パス | `"speaker_notes.md"` | 完成物の発表者ノート |
| `deliverables.record` | パス | `"deliverables.json"` | 採用記録 |
| `template.path` | PPTX のパス | なし | 雛形 |
| `template.layout_index` | 0 以上の整数 | `6` | 雛形の空白レイアウトの番号（0 始まり） |
| `template.show_master_shapes` | 真偽値 | `false` | マスターの図形を表示する |

完成物の保存先は互いに重ならず、実行保存先の外にする。資料のディレクトリとその親は実行保存先にできない。

## テーマ

`theme` に指定した TOML は同梱のテーマ `default` に重ねる。表は再帰的に合成し、値と配列は資料の値で置き換える。

### テーマの表

| 項目 | 意味 | 値 | 同梱の既定値 |
|---|---|---|---|
| `schema_version` | 設定の形式 | `1` | `1` |
| `fonts.japanese` | 和文の書体 | 書体名 | `"Yu Gothic UI"` |
| `fonts.latin` | 欧文の書体 | 書体名 | `"Arial"` |
| `fonts.math` | 数式の書体 | MATH 表のある書体名 | `"Cambria Math"` |
| `colors.<名前>` | 色の名前。書式の色には名前か 16 進数を書くことができる。名前に `none` は使用できない | 6 桁の 16 進数 | `background`・`text`・`primary`・`accent`・`muted`・`line`・`surface` |
| `margins.left` | 本文領域の左の余白 | 0 以上の数（pt） | `36` |
| `margins.right` | 本文領域の右の余白 | 0 以上の数（pt） | `36` |
| `margins.top` | 本文領域の上の余白（`blank` のページは `left` と同じ） | 0 以上の数（pt） | `72` |
| `margins.bottom` | 本文領域の下の余白 | 0 以上の数（pt） | `36` |
| `defaults.<書式>` | すべての要素の書式の既定値 | [書式の引数](#書式の引数) | `color`・`align`・`valign`・`line_spacing`・`padding_pt`・`line_width_pt` |
| `defaults.<種類>` | 種類ごとの既定値の表（`text`・`box`・`equation`・`table_cell`・`table_header`・`line`・`arrow`） | 書式の表 | `box`・`table_cell`・`table_header`・`equation` |
| `roles.<名前>` | 文字の役割 | [役割の項目](#役割の項目) | `slide_title`・`cover_title`・`section_title`・`heading`・`body`・`note`・`references`・`equation` |
| `page_kinds.<種別>` | ページ種別の設定（`cover`・`section`・`content`・`references`・`blank`） | [ページ種別の項目](#ページ種別の項目) | 各種別 |

### 役割の項目

役割には下の項目と [書式の引数](#書式の引数) のすべて（`fill`・`line_color`・`corner_radius_pt` などを含む）を書くことができる。役割の書式のうち、その役割を使う要素の種類が受け付ける書式だけが適用される。

| 項目 | 意味 | 値 |
|---|---|---|
| `size_pt` | 基本の文字サイズ。`inherits` で継承しない役割では必須 | 正の数（pt） |
| `min_size_pt` | 最小文字サイズ。`inherits` で継承しない役割では必須 | 正の数（pt） |
| `inherits` | 継承元の役割。継承元の項目に重ねる | 役割名 |
| `allowed_content` | 内容の種類を限る | `math`（数式だけ）か `text`（文字だけ） |
| `restricted` | 真なら、`allow_roles` で宣言した場所だけで使用できる | 真偽値 |
| `allowed_slide_ids` | 宣言がなくても使用を認めるページの識別子 | 識別子の配列 |

同梱の役割は `slide_title`（32 pt、最小 28）、`cover_title`（36 pt、最小 28）、`section_title`（48 pt、最小 36）、`heading`（24 pt、最小 20）、
`body`（20 pt、最小 18）、`note`（14 pt、最小 12）、`references`（`body` を継承、18 pt、最小 16）、`equation`（`body` を継承。役割を指定しない `equation()` が使う）である。

### ページ種別の項目

| 項目 | 意味 | 値 |
|---|---|---|
| `title_role` | 題名の役割 | 役割名 |
| `title_area` | 題名の領域 | ページに対する比 `[x, y, 幅, 高さ]` |
| `title_area_pt` | 題名の領域（`title_area` より優先） | `[x, y, 幅, 高さ]`（pt） |
| `title_style` | 題名の書式 | 文字の書式の表 |
| `content_area` | 本文領域（`slide.content`）。省略すると `margins` から決める | ページに対する比 |
| `content_area_pt` | 本文領域（`content_area` より優先） | pt |
| `default_role` | 要素の既定の役割 | 役割名（省略すると `body`） |
| `page_number` | ページ番号を置く | 真偽値 |
| `footer_area_pt` | ページ番号の領域（省略するとページの右下の幅 36 pt・高さ 20 pt） | pt |
| `footer_style` | ページ番号の書式（役割は `note`、右揃え） | 文字の書式の表 |
| `background` | 背景の塗り（省略すると色 `background`）。`"none"` は背景の図形を作らない（雛形のマスターの図形を見せるときに使う） | 色か `"none"` |
| `decorations` | ページ種別の飾り（`[[page_kinds.<種別>.decorations]]`） | [飾りの項目](#飾りの項目)の表の配列 |

### 飾りの項目

飾りは線・矢印か図形で、文字を置かない。描く順は背景の後、題名の前である。
書式は種類の既定（`[defaults.<種類>]`）に重ね、役割の書式は重ねない。線・矢印には `line_color`・`line_width_pt`・`dashed`、図形にはさらに `fill` と `corner_radius_pt` を書くことができる。
飾りは場面で `_theme/decoration/<識別子>` の要素になり、`semantic` は `decoration` である。

| 項目 | 意味 | 値 |
|---|---|---|
| `kind` | 飾りの種類 | `line`・`arrow`・`box` |
| `id` | 飾りの識別子（ページ種別の中で一意） | `/` と制御文字を含まない文字列 |
| `area` | 図形の領域 | ページに対する比 `[x, y, 幅, 高さ]` |
| `area_pt` | 図形の領域（`area` とどちらか一方） | pt |
| `shape` | 図形の形 | `rectangle`・`rounded_rectangle`・`ellipse`・`diamond` |
| `start` | 線の始点 | ページに対する比 `[x, y]` |
| `end` | 線の終点 | ページに対する比 `[x, y]` |
| `start_pt` | 線の始点（`start`・`end` の組とどちらか一方） | pt |
| `end_pt` | 線の終点 | pt |

## 記述規則

`rules` に指定した TOML は同梱の記述規則 `general` に重ねる。同梱の規則は句点やノートの文体を制限しない。制限は資料ごとの規則で加える。

| 設定 | 既定値 | 内容 |
|---|---|---|
| `schema_version` | `1` | 設定の形式 |
| `forbidden_characters` | `[]` | 本文に使わない文字。1 項目に 1 文字を書く |
| `require_notes` | `false` | 真なら、全ページに発表者ノートを必須にする |
| `notes_style` | `"unrestricted"` | `"polite"` なら、です・ます調でないノートの文末を警告する |
| `math_candidates.enabled` | `false` | 真なら、数式として書くべき文字を警告する（`math_not_marked`） |
| `math_candidates.numbers` | `true` | 数字も候補にする |
| `math_candidates.variables` | `[]` | 候補に加える名前（2 文字以上の変数名など） |
| `math_candidates.operations` | `[]` | 候補に加える操作名。扱いは `variables` と同じ |
| `math_candidates.exclude_patterns` | `[]` | 一致した部分を候補から外す正規表現 |
| `math_candidates.exclude_roles` | `[]` | 検査しない文字の役割 |
| `math_candidates.exclude_page_kinds` | `[]` | 検査しないページ種別 |

`math_candidates` は数式とリンクを除いた文字の中の、単独の英字（全角を含む）と英字に数字・添字が続くもの（`n`、`n0`、`x_1`、`n²`）、
単独のギリシャ文字と `\alpha` のような数式の命令、`numbers` が真のときの数字、`variables`・`operations` に書いた名前を候補にする。
URL・DOI・arXiv の番号の中は候補にしない。題名は検査し、ページ番号と飾りは検査しない。年などは `exclude_patterns = ['\d{4}\s*年']` のように外す。

## 実行環境（environment.toml）

共通の設定の場所は、Linux と WSL では `~/.config/slidekit/environment.toml`、Windows では `%APPDATA%/slidekit/environment.toml` で、環境変数 `SLIDEKIT_ENVIRONMENT` で変更できる。
資料の `environment_override` の設定は共通の設定に重ねる。相対パスはその設定ファイルの場所を基準にする。

| 設定 | 値 | 既定値 | 内容 |
|---|---|---|---|
| `schema_version` | `1` | `1` | 設定の形式 |
| `jobs` | 1 以上の整数 | CPU の数と 8 の小さいほう | PowerPoint を使わない処理（ページ画像と互換画像の切り出し）の同時プロセス数。`--jobs` と環境変数 `SLIDEKIT_JOBS` が優先する |
| `font_directories` | フォルダーのパスの配列 | 自動で探す | 既定の書体を探すフォルダー。書くとこのフォルダーだけを探し、自動では探さない。`[]` なら探さない（[既定の書体の探し方](usage.md#既定の書体の探し方)） |
| `fonts.<書体>.path` | フォントファイルのパス | 自動で探した結果 | 寸法の計算に使うフォントファイル。書体名ごとの表に書く |
| `fonts.<書体>.index` | 0 以上の整数 | `0` | TTC の中の書体の番号 |
| `fonts.<書体>.bold_path` | フォントファイルのパス | 自動で探した結果 | 太字のフォントファイル。太字の文字を使う書体では必須 |
| `fonts.<書体>.bold_index` | 0 以上の整数 | `0` | 太字の書体の番号 |
| `renderers.powerpoint.powershell` | 実行ファイル | `"powershell.exe"` | PowerShell |
| `renderers.powerpoint.timeout` | 正の秒数 | `90` | 1 回の変換の制限時間 |
| `renderers.powerpoint.wait_seconds` | 0 以上の秒数 | `600` | 別の変換が終わるのを待つ上限。`--wait` が優先する。`0` は待たない |
| `renderers.libreoffice.executable` | 実行ファイル | `"libreoffice"` | LibreOffice |
| `renderers.libreoffice.timeout` | 正の秒数 | `90` | 1 回の変換の制限時間 |

既定の書体の候補は Yu Gothic UI、Yu Gothic、Arial、Cambria Math、Calibri、Segoe UI、Meiryo、BIZ UDGothic である。テーマが使う書体のうち、`fonts` に指定がなく候補にあるものだけを探す。
候補にない書体は自動では探さないので、`fonts.<書体>.path` を指定する。

```toml
font_directories = ["/mnt/c/Windows/Fonts"]

[renderers.powerpoint]
timeout = 120
wait_seconds = 300

[fonts."Noto Sans JP"]
path = "fonts/NotoSansJP-Regular.ttf"
bold_path = "fonts/NotoSansJP-Bold.ttf"
```
