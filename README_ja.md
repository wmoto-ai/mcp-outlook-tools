# MCP Outlook Tools

Microsoft Outlookとの連携を可能にするModel Context Protocol (MCP)サーバー実装です。AIアシスタントがカレンダー管理、メール操作、検索機能を実行できるようになります。

## 機能

- 📅 **カレンダー管理**
  - 指定期間内のカレンダー項目を取得
  - 詳細情報付きで新しい予定を追加
  - カテゴリーと予定状態のサポート

- 📧 **メール操作**
  - To/CC宛先を指定してメール送信
  - 送信前の確認表示
  - 本文フォーマットの完全サポート

- 🔍 **メール検索**
  - 日付とキーワードでメール検索
  - メールアドレスからユーザー情報を抽出
  - 日本語テキストのエンコーディングサポート

## 必要条件

- Windows OS (pywin32のため必須)
- Microsoft Outlookがインストール・設定済み
- Python 3.10以上
- MCP対応のAIアシスタント（例：Claude Desktop）

## インストール

1. リポジトリをクローン：
```bash
git clone https://github.com/yourusername/mcp-outlook-tools.git
cd mcp-outlook-tools
```

2. uvを使用して依存関係をインストール：
```bash
uv pip install -e .
```

またはpipを使用：
```bash
pip install -e .
```

## 設定

### Claude Desktop向け

Claude Desktopの設定ファイルに以下を追加：

**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "outlook-tools": {
      "command": "uv",
      "args": [
        "--directory",
        "C:/path/to/mcp-outlook-tools",
        "run",
        "--with-editable",
        ".",
        "-m",
        "outlook_tools.server"
      ],
      "cwd": "C:/path/to/mcp-outlook-tools"
    }
  }
}
```

## 使い方

設定が完了すると、AIアシスタントで以下のツールが利用可能になります：

### `add_appointment`
```
Outlookカレンダーに新しい予定を追加
パラメータ：
- subject: 予定のタイトル
- start_time: 開始日時 (YYYY-MM-DD HH:MM)
- end_time: 終了日時 (YYYY-MM-DD HH:MM)
- location: 会議場所（任意）
- description: 詳細説明（任意）
- categories: カンマ区切りのカテゴリー（任意）
- busy_status: 0=空き時間、1=仮予定、2=予定あり、3=外出中（デフォルト：1）
```

### `get_calendar`
```
指定期間のカレンダー項目を取得
パラメータ：
- start_date: 開始日 (YYYY-MM-DD)
- end_date: 終了日 (YYYY-MM-DD)
```

### `send_email`
```
Outlook経由でメール送信
パラメータ：
- to: 宛先メールアドレス（セミコロン区切り）
- cc: CC宛先（セミコロン区切り）
- subject: メール件名
- body: メール本文
```

## プロジェクト構成

```
mcp-outlook-tools/
├── src/
│   └── outlook_tools/
│       ├── __init__.py
│       ├── server.py           # MCPサーバー実装
│       ├── calendar_service.py # カレンダー操作
│       └── search_service.py   # メール検索操作
├── test/                       # テストファイル
├── examples/                   # サンプルスクリプト
├── pyproject.toml             # プロジェクト設定
├── README.md                  # 英語版README
└── README_ja.md               # 日本語版README（このファイル）
```

## 開発

### テストの実行
```bash
pytest test/
```

### 型チェック
```bash
pyright src/
```

### リンティング
```bash
ruff check src/
```

## セキュリティに関する注意

- このツールはローカルのOutlookインストールへのアクセスが必要です
- メールは送信前に確認のため表示されます
- コード内に認証情報は保存されません
- すべての操作は既存のOutlook認証を使用したWindows COMインターフェース経由で行われます

## 制限事項

- Windowsのみ対応（pywin32依存のため）
- Outlookがインストール・設定済みである必要があります
- タイムゾーン処理はJST（+9時間）を想定

## 貢献

プルリクエストは歓迎します！お気軽にご投稿ください。

## ライセンス

このプロジェクトはMITライセンスの下で公開されています。詳細はLICENSEファイルをご覧ください。

## 謝辞

- [FastMCP](https://github.com/modelcontextprotocol/fastmcp)フレームワークを使用
- Outlook COMインターフェースにはpywin32を使用