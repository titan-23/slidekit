# slidekit の未解決の問題

2026 年 10 月 11 日の調査でまとめた 61 項目は、0.2 の改修で解消した。
その後に記録した数式の寸法・上付きの横分数を含む行のずれと、実行整理の資料判定には、2026 年 10 月 23〜24 日に対応した。
2026 年 11 月 13 日には、資料作成で使って気づいた改善点（折り返しの検出、確認の記録と引き継ぎ、要約、残存記録の解除など）と、[開発と資料作成の改善案](improvements.md) に対応した。

各項目には、確認した版、最小の入力、実行コマンド、期待する結果、実際の結果、対応する回帰試験を書く。
解消した項目は削除し、ここには再現できる未解決事項だけを残す。許容幅の中に残る計測差は、[数式の書き方](math.md#残る差) にまとめる。

## 分母側が高い入れ子分数で、広い行間の基線が合わない（低）

- 確認した版：slidekit 0.2（`MEASURE_VERSION` 8、数式の計測 `opentype-4`）、PowerPoint 16.0.20430.20182（2026 年 10 月 24 日）。
- 最小の入力：独立した数式の `\frac{a}{\frac{b}{c}}` を、20 pt・行間 2.0・上揃えの `equation()` に置く。
- 実行コマンド：`python tools/calibrate.py run CONFIG --work 作業ディレクトリ`（CONFIG は `deck = "cases"` で上の式・大きさ・行間を指定する）。
- 期待する結果：計測の基線が、PDF の外側の分数線から求めた基線と照合の許容差の中で合う。
- 実際の結果：PDF の基線は枠の上端から 88.57 pt、計測は 73.30 pt で、計測が約 15.27 pt 上に出る。数式自体の寸法は合い、行間で増える高さを上下へ分ける規則が違う。枠切れや字形の割り当て違いではない。
- 回帰試験：なし（1 ケースだけに合わせた条件は加えていない）。この式の 20 pt・行間 1.5 までは照合データ（`tests/data/layout/powerpoint_layout.json`）で差が小さいことを確かめた。
- 回避：広い行間を使う場合は、生成後の PDF と表示で位置を確認する。
- 対応案：分母側が高い別の式と行間を `tools/calibrate.py` で照合し、行間の増分をすべて上に配る条件を確かめる。

## PowerPoint 16.0.20527 で、数式の × の前後が空かない（中）

- 確認した版：slidekit 0.2（`MEASURE_VERSION` 9、数式の計測 `opentype-4`）、PowerPoint 16.0.20527.20150（2026 年 11 月 13 日）。照合データ `tests/data/powerpoint_math.json` は 16.0.20430.20182 で作った。
- 最小の入力：`a\times b` を 20 pt の文中の数式と独立した数式で置く（数式の照合データの道具の式の一覧にある）。
- 実行コマンド：`printf 'schema_version = 1\ndeck = "math"\n' > math.toml && python tools/calibrate.py run math.toml --work 作業ディレクトリ`
- 期待する結果：照合データと同じく、× の前後に二項演算の空き（4/18 em ずつ）が入り、計測と合う。
- 実際の結果：332 件のうち `a\times b` の 2 件だけが照合データより 8.88 pt（0.444 em）狭い。PDF の字形の位置では、𝑎 の後と 𝑏 の前の空きが二項演算の空きより小さい。同じ資料の `a\cdot b` は照合データと計測に合う。
  × は `EAST_ASIAN_FONT_CHARACTERS`（PowerPoint が和文の書体で描く記号）に入っている。PowerPoint の更新で数式の中の × の扱いが変わった見込みだが、原因は確かめていない。
- 回帰試験：なし。`test_math_metrics_match_powerpoint` は保存した照合データと計測を比べるので、PowerPoint の更新では失敗しない。
- 回避：× を使う式を枠いっぱいに置くときは、生成後の PDF の検査（`pdf_element_text_overflow`・`pdf_line_count_mismatch`）と表示で確かめる。計測は実際より広いので、はみ出しを見落とす向きの誤差ではない。
- 対応案：`tools/calibrate.py` の `cases` で `\times`・`\div`・`\pm` を、数式の文字の言語（`lang`）を替えて照合し、空きが消える条件を確かめる。原因が分かったら照合データを作り直す。

## 変わったページだけを変換し直すことができない（中）

- 確認した版：slidekit 0.2（2026 年 11 月 13 日）。
- 最小の入力：33 ページの資料の 1 ページの 1 行だけを直す。
- 実行コマンド：`slidekit build --project src/deck.toml`
- 期待する結果：前の実行から変わっていないページの変換結果を使い回し、変わったページだけを変換する。
- 実際の結果：全ページを標準表示と互換表示の両方で変換する（約 25 秒）。下書きの確認には `build --draft --slides 14` で指定したページだけを変換できるが、その実行は完成物にできない。
- 回帰試験：なし（機能がない）。
- 直していない理由：完成物の PDF と表示確認の画像は、最終 PPTX と同じスライドの XML を変換したものにする必要がある。ページごとの PDF をつなぐと、リンク・しおり・書体の埋め込みが 1 回の変換と同じになることを確かめる必要がある。
- 対応案：表示の証拠（`evidence_hash`）が前の実行と同じページの PDF のページを使い回し、変わったページだけの PPTX を変換してつなぐ。つないだ PDF と 1 回で変換した PDF の違いを PowerPoint で照合してから採用する。

## 元資料の変更で内容の確認がやり直しになったとき、変わった節を示さない（低）

- 確認した版：slidekit 0.2（2026 年 11 月 13 日）。
- 最小の入力：`inputs.documents` の `proofs` を参照する 10 ページの資料で、`proofs` の 1 節だけを直して生成する。
- 実行コマンド：`slidekit build --project src/deck.toml` の後に `slidekit review status --project src/deck.toml`
- 期待する結果：内容の確認が未確認に戻った理由と、元資料の変わった節の見出しを示す。
- 実際の結果：理由は「元資料の変更 p.3-12（proofs）」のように示すが、変わった節は示さない。生成記録は元資料のハッシュ値だけを持つ。
- 回帰試験：`tests/test_records.py::test_pending_reasons_split_page_and_document_changes`（理由の区別まで）。
- 対応案：元資料が Markdown のときは、節ごとのハッシュ値を生成記録に入れて比べる。

## Windows の Python から直接実行すると、Ctrl+C が変換の処理にも届く（低）

- 確認した版：slidekit 0.2（2026 年 11 月 13 日、コードの確認だけ）。
- 最小の入力：Windows の Python で `slidekit build` を実行し、PowerPoint の変換中に Ctrl+C を押す。
- 実行コマンド：`python -m slidekit build --project src/deck.toml`
- 期待する結果：WSL と同じく、変換の処理（PowerShell）は現在の変換を終えてから PowerPoint を閉じ、残存記録を消す。
- 実際の結果：変換の処理を別のセッションで起動する指定（`start_new_session`）は POSIX でだけ働くので、Ctrl+C が変換の処理にも届く見込みである。WSL では起こらない。
- 回帰試験：なし（Windows の Python では試験していない）。
- 対応案：Windows では `CREATE_NEW_PROCESS_GROUP` で起動する。
