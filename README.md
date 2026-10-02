# 内訳アシスト（Excel アドイン）

保温工事の見積書「内訳」シートに、品目・施工箇所・仕様・サイズを選んで明細を登録する Excel アドインです。

- 公開先：https://lograck-jp.github.io/uchiwake-assist/
- 作業ウィンドウ：`taskpane.html`（`taskpane.js` / `core.js` / `taskpane.css`）
- Excel 登録用：`manifest.xml`
- 設定マスタ（選択肢・仕様候補・単価履歴）は利用者のパソコンの Excel ファイルから読み込みます。このリポジトリには単価などのデータは入っていません。
- 導入手順は「内訳アシスト_導入手順.html」を参照してください。

## ファイル
| ファイル | 役割 |
|---|---|
| taskpane.html / .css / .js | 作業ウィンドウの画面と Excel への書き込み |
| core.js | 内訳シートの解析、登録位置の計算（Excel に依存しない部分） |
| commands.html | リボンボタン用（中身は空） |
| manifest.xml | Excel にアドインを登録するための設定ファイル |
| assets/ | アイコン |
| lib/xlsx.full.min.js | 設定マスタ（.xlsx）を読むためのライブラリ（SheetJS、Apache-2.0） |

## 更新するとき
ファイルを差し替えたら、`taskpane.html` の `?v=1.0.1` の数字を上げてからアップロードします（Excel 側の古い読み込みを避けるため）。


## 更新履歴
- 1.0.1（2026/10/02）サイズ選択中にカーソルが単価欄へ移らないよう修正。Tab は数量→次のサイズの数量へ。単価未登録の欄は点線表示。
- 1.0.0（2026/09/28）初版
