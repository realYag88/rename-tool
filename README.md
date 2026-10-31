# リネームツール

フォルダ内や選択したファイルを、連番か文字列の置換でまとめてリネームするWindows用ツールです。

- 説明ページ: https://caprame.com/tools/rename-tool/
- ダウンロード: [Releases](../../releases/latest) の `RenameTool.zip`

## 配布の仕組み(自分用メモ)

main に push すると、GitHub Actions がツール一式を `RenameTool.zip` にまとめて Releases に出します。
サイトのダウンロードボタン(`releases/latest/download/RenameTool.zip`)は常に最新版を指します。
README.md だけの変更では Releases は作られません。
