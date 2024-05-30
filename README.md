# json2excel

JSONファイルをExcelファイルに変換する

## Description

- JSONをEXCELに変換するAPI
- 変換の定義はYAMLで定義します
- ハイパーリンクなどの処理はdecoratorとして変換のYAMLとは別で定義します
- JSON 2 EXCEL までの処理や、既存のExcelファイルへハイパーリンクの設定のみ行う処理も実行できます
- APIはFastAPIを使ってます

## Getting Started

### local

1. cd docker
2. docker compose up -d
3. http://localhost:1239/docs/

<!-- ### PaaS

- renderの無料枠でdeployしています
- 🌐 https://json2excel.onrender.com/docs/ -->
