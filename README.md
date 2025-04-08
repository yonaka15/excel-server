# Excel Server

Excel ServerはModel Context Protocol（MCP）を使用して、xlwings-rpc APIにアクセスし、Excelワークブックを操作するためのサーバーアプリケーションです。

## 機能

- HTTP経由でxlwings-rpc APIにアクセス
- MCPプロトコルを使用したツールとして各Excel操作機能を提供
- 環境変数による設定

## セットアップ

1. 必要なパッケージをインストール:

```bash
npm install
```

2. TypeScriptをコンパイル:

```bash
npm run build
```

3. サーバーを起動:

```bash
npm start
```

## 環境変数

| 環境変数 | 説明 | デフォルト値 |
|----------|------|-------------|
| XLWINGS_HOST | xlwings-rpcサーバーのホスト | 0.0.0.0 |
| XLWINGS_PORT | xlwings-rpcサーバーのポート | 8000 |

## 使用方法

サーバーはStdioでMCPリクエストを受け付けます。

### 利用可能なツール

#### アプリケーション操作

- `app.list` - すべての実行中のExcelアプリケーションを取得
- `app.get` - 指定されたPIDまたはアクティブなExcelアプリケーションを取得
- `app.create` - 新しいExcelアプリケーションを作成
- `app.quit` - Excelアプリケーションを終了

#### ワークブック操作

- `book.list` - すべての開いているワークブックを取得
- `book.get` - 指定されたワークブックを取得
- `book.open` - ワークブックを開く
- `book.create` - 新しいワークブックを作成
- `book.close` - ワークブックを閉じる

#### シート操作

- `sheet.list` - ワークブック内のすべてのシートを取得
- `sheet.get` - 特定のシートを取得

#### レンジ操作

- `range.get_value` - セル範囲の値を取得
- `range.set_value` - セル範囲に値を設定
- `range.get_formula` - セル範囲の数式を取得
- `range.set_formula` - セル範囲に数式を設定

## ライセンス

MIT
