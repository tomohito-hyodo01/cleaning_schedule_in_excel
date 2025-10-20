# Windows版 クイックスタートガイド

## 🚀 初めてのユーザー向け簡単ガイド

### 必要なもの

1. Windows 10 または 11
2. インターネット接続
3. Claude API キー（無料で取得可能）

---

## ⚡ 5分で始める

### ステップ1: Python をインストール（初回のみ）

1. https://www.python.org/downloads/ にアクセス
2. 「Download Python 3.12.x」をクリック
3. ダウンロードしたファイルを実行
4. **重要**: 「Add Python to PATH」にチェック ✅
5. 「Install Now」をクリック

### ステップ2: アプリを準備

1. プロジェクトフォルダを開く
2. アドレスバーに `cmd` と入力して Enter（コマンドプロンプトが開く）
3. 以下をコピー＆ペーストして実行:

```cmd
pip install -r requirements.txt
```

### ステップ3: アプリを起動

```cmd
python main.py
```

### ステップ4: APIキーを設定（初回のみ）

1. アプリが起動したら「⚙️ 設定」ボタンをクリック
2. Claude API キーを入力
3. 「保存」をクリック

#### APIキーの取得方法:

1. https://console.anthropic.com/ にアクセス
2. アカウントを作成（無料）
3. 「API Keys」→「Create Key」
4. 生成されたキーをコピー

### ステップ5: 画像を変換

1. 「📁 画像ファイルを選択」をクリック
2. 清掃スケジュール画像を選択
3. 「🚀 Excel変換開始」をクリック
4. 完了！デスクトップにExcelファイルが保存されます

---

## 🎁 実行ファイル版（簡単！）

Python のインストールが面倒な場合は、実行ファイル版を作成できます。

### ビルド方法

```cmd
pip install pyinstaller
pyinstaller --onefile --windowed --name CleaningSchedule main.py
```

### 使用方法

`dist\CleaningSchedule.exe` をダブルクリックするだけ！

---

## 📖 詳しい説明

- 完全なドキュメント: [README.md](README.md)
- ビルド手順: [BUILD_WINDOWS.md](BUILD_WINDOWS.md)

---

## ❓ よくある質問

### Q: Python がインストールされているか確認したい

```cmd
python --version
```

`Python 3.x.x` と表示されればOK

### Q: エラーが出る

1. Python がインストールされているか確認
2. `pip install -r requirements.txt` を再実行
3. APIキーが正しいか確認

### Q: 画像が選択できない

JPG または PNG 形式の画像のみ対応しています

### Q: 料金は？

Claude API の利用料金のみ（画像1枚あたり約5-15円）

---

## 📞 サポートが必要な場合

README.md の「トラブルシューティング」セクションを確認してください。

