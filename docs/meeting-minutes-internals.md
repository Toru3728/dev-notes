# 議事録自動化の仕組み

体裁を変えたい・チェック項目を足したいときだけ読む。
普段の使い方は [議事録作成の自動化](meeting-minutes-automation.md) を参照。

## 構成

```
.claude/commands/my-gijiroku.md          入口。何を受け取って何をするか
docs/meeting-minutes-automation.md       使う人向けの手順書
docs/meeting-minutes-internals.md        このファイル
templates/
├── meeting-minutes-template.docx        標準体裁の空テンプレート（default.json から生成）
├── build-meeting-minutes-template.js    プロファイル → .docx
├── check-minutes.py                     配布前チェック
└── profiles/
    ├── default.json                     標準の体裁
    ├── careplan-4.json                  ケアプラン第4表（サービス担当者会議の要点）
    └── sample-office.json               写真から起こしたプロファイルの見本（架空）
```

実在の事業所のプロファイルは `company/secretary/minutes/formats/<事業所名>.json`。
**dev-notes は公開リポジトリなので、事業所名を含むものは置かない。**

## 書式プロファイル

事業所ごとの様式を JSON 1枚で表す。`templates/profiles/default.json` が標準、
`sample-office.json` が写真から起こした例。

| キー | 意味 |
|---|---|
| `page.orientation` | `portrait` / `landscape` |
| `page.marginMm` | 余白（mm）。四辺共通 |
| `font.name` / `font.sizePt` | 本文フォントと文字サイズ |
| `headingColor` | 見出しの色（16進6桁）。様式が黒なら `000000` |
| `title` | 1行目の表題。様式の呼び方をそのまま |
| `headerText` | ページ上部の見出し。空文字ならヘッダなし |
| `infoLabelWidth` | 会議情報表のラベル列幅（DXA）。省略時は本文幅の26% |
| `infoFields[]` | 会議情報の表。`label` と `placeholder` |
| `sections[].heading` | 節の見出し。**様式の項目名をそのまま** |
| `sections[].type` | `text`（自由記述）／ `table`（表） |
| `sections[].role` | `decisions` / `todo` / `needs-check`。improvements がどの節に効くかの目印 |
| `sections[].columns[]` | 表の列。`width`（DXA 固定）か `weight`（比率） |
| `sections[].rows` | 空の記入行の数 |
| `sections[].note` | 表の下に出す注記 |
| `improvements` | 下記3つを足すかどうか |

`improvements` は様式に無くても足すもの。`false` にすれば足さない。

| キー | 効果 |
|---|---|
| `todoOwnerDeadline` | `role: "todo"` の表に「担当」「期限」列を足す（同名の列があれば足さない） |
| `needsCheckSection` | `role: "needs-check"` の節が無ければ末尾に足す |
| `preDistributionChecklist` | 末尾に配布前チェック欄を足す |

### 列幅の指定

`weight` で書いておけば、余白や用紙向きが変わっても合計が本文幅にそろう。
`width` で DXA を直接指定すると様式の列幅を再現できる。どちらでも
**合計は必ず本文幅に一致させる**（ずれると Word 以外で開いたときに崩れる）。
端数は一番広い列に寄せて自動でそろえる。

DXA は 1440 = 1インチ。A4 縦の幅は 11906、余白 20mm（1134）なら本文幅は 9638。

## テンプレートの生成

```bash
npm install docx    # 初回だけ
node templates/build-meeting-minutes-template.js
  # → templates/profiles/default.json から templates/meeting-minutes-template.docx

node templates/build-meeting-minutes-template.js <プロファイル.json>
  # → templates/built/<プロファイル名>.docx

node templates/build-meeting-minutes-template.js <プロファイル.json> -o <出力.docx>
```

`templates/built/` と `node_modules/` は `.gitignore` 済み。

`.docx` はバイナリで GitHub 上の差分が読めないため、体裁を変えるときは
**JSON かスクリプトを直して再生成する**。Word で直接編集して上書きしない。

## 配布前チェック

```bash
python3 templates/check-minutes.py <議事録.docx>
python3 templates/check-minutes.py <議事録.docx> --preview <出力先ディレクトリ>
```

追加の依存は無し（Python 標準ライブラリだけ）。`--preview` は LibreOffice があれば
目視用の JPEG を出す。無ければその旨を出して続行する。

終了コード: `0` 配布可 / `1` 直すところがある / `2` ファイルが読めない

### 検査の内容

| 検査 | 重み | 何を見るか |
|---|---|---|
| `［要確認］` の残り | 配布不可 | 本文と表に「要確認」を含む記述 |
| プレースホルダの残り | 配布不可 | `〔…〕` の形 |
| ToDo の担当・期限 | 配布不可 | 「担当」「期限」列の空欄 |
| 配布前チェック欄の残り | 配布不可 | 配布時に削除する前提の欄 |
| 担当・期限が「未定」 | 警告 | `未定` `TBD` `未記入` `要記入` |
| 仮名化されていない氏名 | 警告 | `◯◯様` `◯◯さん` `◯◯氏`。`A様` 形と役割語（利用者・ご家族・職員 等）は除外 |
| 列幅の合計 | 警告 | 本文幅とのずれ |

配布前チェック欄より後ろの段落は他の検査から除外している。
あの欄は「［要確認］が残っていないか」という点検文を含むので、
混ぜると自分自身を誤検知する。

### 法定様式のとき（第4表など）

ケアプラン第4表のような法定様式では `improvements` を全て `false` にする。
**様式に列や節を足すと実地指導で問題になる。**
未確認の印は本文中の `［AI提案・要確認］` タグで表し、
ケアマネジャーが確認してタグを削除するまで提出しない運用にする。

チェッカーは `［AI提案・要確認］`（AIが推測した箇所）と
`［要確認］`（聞き取れなかった箇所）を分けて報告する。直す人が違うため
（前者はケアマネジャーが判断、後者は音声に戻って確認）。
「タグを削除」を含む注記行は、タグの使い方を説明している文なので数に入れない。

### 検査を足すとき

`check-minutes.py` に `check_*(paragraphs, tables, report)` の形の関数を足して、
`main()` から呼ぶ。`report.blocker()` が配布不可、`report.warn()` が警告、
`report.note()` が参考情報。出力の体裁は `Report.render()` がまとめて面倒を見る。
