# 清掃スケジュール Excel変換ツール

清掃スケジュールの画像を Claude AI で解析し、Excel ファイルに変換するデスクトップアプリケーションです。

## 📋 機能

- 📁 画像ファイル選択（JPG, PNG対応）
- 👁️ 画像プレビュー表示
- 🤖 Claude API による高精度な表認識
- 📊 Excel ファイル自動生成（レイアウト再現）
- 📈 進行状況表示
- ⚙️ APIキー設定機能

## 🚀 クイックスタート

### Windows の場合

#### 方法1: 実行ファイル版（推奨）

1. `CleaningSchedule.exe` をダブルクリック
2. 設定ボタンから Claude API キーを入力
3. 画像を選択して変換開始

#### 方法2: Pythonスクリプト版

1. Python 3.9 以上をインストール
2. コマンドプロンプトで以下を実行:

```cmd
cd "C:\path\to\清掃スケジュール対応"
pip install -r requirements.txt
python main.py
```

### macOS の場合

1. ターミナルを開く
2. 以下を実行:

```bash
cd "/Users/hyodo/InProc/清掃スケジュール対応"
pip3 install -r requirements.txt
python3 main.py
```

## 🔑 APIキーの取得

1. https://console.anthropic.com/ にアクセス
2. アカウント作成（無料）
3. APIキーを作成
4. アプリの「設定」ボタンから入力

## 📦 必要な環境

- Python 3.9 以上
- インターネット接続
- Claude API キー

## 🛠️ インストール方法

### 1. 依存ライブラリのインストール

```bash
pip install -r requirements.txt
```

### 2. APIキーの設定

アプリを起動後、「設定」ボタンから Claude API キーを入力してください。

または、`config.json` を手動で作成:

```json
{
  "claude_api_key": "sk-ant-api03-..."
}
```

## 💻 使い方

### 基本的な使い方

1. **アプリを起動**
   ```bash
   python main.py
   ```

2. **APIキーを設定**（初回のみ）
   - 「設定」ボタンをクリック
   - Claude API キーを入力
   - 「保存」をクリック

3. **画像を選択**
   - 「📁 画像ファイルを選択」ボタンをクリック
   - 清掃スケジュール画像を選択

4. **変換開始**
   - 「🚀 Excel変換開始」ボタンをクリック
   - 処理完了を待つ（通常10-30秒）

5. **完了**
   - デスクトップに Excel ファイルが保存されます
   - 「フォルダを開く」で確認

## 📁 ファイル構成

```
清掃スケジュール対応/
├── main.py                  # メインアプリケーション
├── requirements.txt         # 依存ライブラリ
├── config.json              # 設定ファイル（自動生成）
├── README.md                # このファイル
├── BUILD_WINDOWS.md         # Windows用ビルド手順
├── schedule_image/          # サンプル画像
│   └── 清掃スケジュール.jpg
└── .gitignore               # Git除外設定
```

## 🔧 Windows用実行ファイルのビルド

詳細は [BUILD_WINDOWS.md](BUILD_WINDOWS.md) を参照してください。

簡易手順:

```cmd
pip install pyinstaller
pyinstaller --onefile --windowed --name CleaningSchedule main.py
```

生成物: `dist/CleaningSchedule.exe`

## ⚠️ トラブルシューティング

### エラー: "APIキーが無効です"

- APIキーが正しいか確認
- https://console.anthropic.com/ で APIキーの有効性を確認
- クレジットカードが登録されているか確認

### エラー: "モジュールが見つかりません"

```bash
pip install -r requirements.txt
```

を再実行してください。

### 画像が読み込めない

- ファイル形式を確認（JPG, PNG のみ対応）
- ファイルが破損していないか確認

### Excel生成に失敗する

- Claude API の解析結果を確認
- 画像の品質を確認（鮮明な画像を使用）

## 💰 コスト

- Claude API: 画像1枚あたり 約5-15円
- 月間100枚処理: 約500-1500円

## 📝 ライセンス

このプロジェクトは個人利用・商用利用ともに自由に使用できます。

## 🔄 バージョン履歴

### v1.0.0 (2025-10-20)
- 初回リリース
- 基本機能実装
  - 画像選択
  - Claude API 解析
  - Excel 生成
  - プレビュー表示
  - 進行状況表示

## 📧 サポート

問題が発生した場合は、以下を確認してください:

1. Python バージョン (3.9以上)
2. 依存ライブラリのインストール状況
3. API キーの有効性
4. インターネット接続

## 🙏 使用ライブラリ

- CustomTkinter: モダンな GUI フレームワーク
- Anthropic: Claude API クライアント
- OpenPyXL: Excel ファイル操作
- Pillow: 画像処理

