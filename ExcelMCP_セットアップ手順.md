# ExcelMCP セットアップ手順

CursorでExcelMCP（excel-mcp-server）を使用できるようにするためのセットアップ手順です。

## 概要

ExcelMCPは、Microsoft Excelがインストールされていない環境でも、Excelファイルの作成・読み取り・編集ができるModel Context Protocol（MCP）サーバーです。

**参考**: [excel-mcp-server GitHub](https://github.com/haris-musa/excel-mcp-server)

---

## 前提条件

- **Cursor** がインストールされていること
- **Python 3.8以上** がインストールされていること（`uv` の実行に必要）
- **インターネット接続** があること（`uv` と `excel-mcp-server` のインストールに必要）

---

## セットアップ手順

### ステップ1: `uv` / `uvx` のインストール

`excel-mcp-server` は `uvx` コマンド経由で実行されます。まず `uv`（`uvx` を含む）をインストールします。

#### Windows の場合

**方法A: 公式インストーラー（推奨）**

PowerShellを開いて、以下のコマンドを実行します：

```powershell
irm https://astral.sh/uv/install.ps1 | iex
```

インストール後、**PowerShellを一度閉じて開き直す**（PATH環境変数の反映のため）。

**方法B: winget を使用する場合**

```powershell
winget install --id Astral.Uv -e
```

インストール後、**ターミナルを再起動**してください。

#### macOS / Linux の場合

ターミナルを開いて、以下のコマンドを実行します：

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

インストール後、**ターミナルを再起動**するか、以下のコマンドでPATHを反映します：

```bash
source $HOME/.cargo/env
```

#### インストール確認

インストールが完了したら、以下のコマンドで確認します：

**Windows (PowerShell)**:
```powershell
uv --version
uvx --version
```

**macOS / Linux**:
```bash
uv --version
uvx --version
```

正常にインストールされていれば、バージョン情報が表示されます（例: `uv 0.9.26`）。

---

### ステップ2: CursorのMCP設定ファイルの作成

プロジェクトのルートディレクトリに `.cursor` フォルダを作成し、`mcp.json` ファイルを配置します。

#### 設定ファイルの場所

```
プロジェクトルート/
  └── .cursor/
      └── mcp.json
```

#### `mcp.json` の内容

以下の内容で `.cursor/mcp.json` を作成します：

```json
{
  "mcpServers": {
    "excel": {
      "command": "uvx",
      "args": ["excel-mcp-server", "stdio"],
      "disabled": false
    }
  }
}
```

#### 設定ファイルの作成方法

**Windows (PowerShell)**:
```powershell
# プロジェクトルートで実行
mkdir .cursor
@"
{
  "mcpServers": {
    "excel": {
      "command": "uvx",
      "args": ["excel-mcp-server", "stdio"],
      "disabled": false
    }
  }
}
"@ | Out-File -FilePath .cursor\mcp.json -Encoding utf8
```

**macOS / Linux**:
```bash
# プロジェクトルートで実行
mkdir -p .cursor
cat > .cursor/mcp.json << 'EOF'
{
  "mcpServers": {
    "excel": {
      "command": "uvx",
      "args": ["excel-mcp-server", "stdio"],
      "disabled": false
    }
  }
}
EOF
```

または、エディタで直接作成しても構いません。

---

### ステップ3: ExcelMCPの動作確認（オプション）

設定が正しく動作するか確認するため、ターミナルで以下のコマンドを実行します：

**Windows (PowerShell)**:
```powershell
uvx excel-mcp-server --help
```

**macOS / Linux**:
```bash
uvx excel-mcp-server --help
```

初回実行時は、`excel-mcp-server` とその依存パッケージが自動的にダウンロード・インストールされます。数秒〜数十秒かかる場合があります。

正常に動作していれば、以下のようなヘルプメッセージが表示されます：

```
Usage: excel-mcp-server [OPTIONS] COMMAND [ARGS]...

Excel MCP Server

Options:
  --install-completion          Install completion for the current shell.
  --show-completion             Show completion for the current shell, to copy
                                it or customize the installation.
  --help                        Show this message and exit.

Commands:
  sse               Start Excel MCP Server in SSE mode
  streamable-http   Start Excel MCP Server in streamable HTTP mode
  stdio             Start Excel MCP Server in stdio mode
```

---

### ステップ4: Cursorの再起動

設定ファイルを作成・編集した後は、**Cursorを再起動**してください。

再起動後、CursorがMCPサーバーを自動的に検出し、`excel` サーバーが利用可能になります。

---

## 動作確認

### Cursor内での確認方法

1. Cursorを再起動後、MCPサーバーの状態を確認します
2. AIアシスタントに「Excelファイルを読み取って」などと依頼し、ExcelMCPが正常に動作するか確認します

### 利用可能な機能

ExcelMCPでは、以下のような操作が可能です：

- Excelファイルの読み取り
- セルへの値の書き込み
- セルの書式設定（フォント、色、罫線など）
- シートの作成・コピー
- テーブルの作成
- スクリーンショットの取得（Windowsのみ）

---

## トラブルシューティング

### `uvx` コマンドが見つからない

**症状**: `uvx: command not found` または `uvx は、内部コマンドまたは外部コマンド、操作可能なプログラムまたはバッチ ファイルとして認識されていません。`

**解決方法**:
1. ターミナル（PowerShell / ターミナル）を**完全に再起動**してください
2. インストールが正しく完了しているか確認：
   - Windows: `$env:USERPROFILE\.local\bin\uv.exe --version`
   - macOS/Linux: `~/.cargo/bin/uv --version`
3. PATH環境変数に `uv` のインストール先が含まれているか確認してください

### ExcelMCPサーバーが起動しない

**症状**: CursorでExcelMCPが利用できない、エラーメッセージが表示される

**解決方法**:
1. `.cursor/mcp.json` のJSON構文が正しいか確認（カンマ、引用符など）
2. `uvx excel-mcp-server stdio` をターミナルで直接実行し、エラーメッセージを確認
3. インターネット接続を確認（初回実行時にパッケージをダウンロードします）
4. Cursorを再起動

### パッケージのダウンロードが遅い / 失敗する

**症状**: `uvx excel-mcp-server` の実行に時間がかかる、またはエラーになる

**解決方法**:
1. インターネット接続を確認
2. ファイアウォールやプロキシの設定を確認
3. 再試行（初回実行時は依存パッケージのダウンロードに時間がかかります）

---

## 補足情報

### グローバル設定との違い

この手順では、**プロジェクト固有の設定**（`.cursor/mcp.json`）を使用しています。

**グローバル設定**を使用する場合の設定ファイルの場所：

- **Windows**: `%APPDATA%\Code\User\globalStorage\tencent-cloud.coding-copilot\settings\Craft_mcp_settings.json`
- **macOS**: `~/Library/Application Support/Code/User/globalStorage/tencent-cloud.coding-copilot/settings/Craft_mcp_settings.json`
- **Linux**: `~/.config/Code/User/globalStorage/tencent-cloud.coding-copilot/settings/Craft_mcp_settings.json`

プロジェクト固有の設定の方が、プロジェクトごとに異なるMCPサーバーを管理しやすいため推奨です。

### Python環境について

`uv` はPythonの実行環境を自動的に管理するため、事前にPythonをインストールする必要はありませんが、Python 3.8以上がシステムにインストールされている必要があります。

### 参考リンク

- [excel-mcp-server GitHub](https://github.com/haris-musa/excel-mcp-server)
- [uv 公式ドキュメント](https://github.com/astral-sh/uv)
- [Cursor MCP ドキュメント](https://docs.cursor.com/advanced/model-context-protocol)

---

## まとめ

1. ✅ `uv` / `uvx` をインストール
2. ✅ `.cursor/mcp.json` を作成
3. ✅ Cursorを再起動
4. ✅ ExcelMCPが利用可能になったことを確認

これで、CursorからExcelファイルの操作が可能になります！
