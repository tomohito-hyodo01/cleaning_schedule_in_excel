# Windows用実行ファイルのビルド手順

このドキュメントでは、Pythonスクリプトから Windows 用の実行ファイル (.exe) を作成する手順を説明します。

## 🎯 概要

PyInstaller を使用して、`main.py` を単一の実行ファイル `CleaningSchedule.exe` に変換します。

## 📋 前提条件

- Windows 10/11
- Python 3.9 以上がインストール済み
- インターネット接続

## 🚀 ビルド手順

### ステップ1: 依存ライブラリのインストール

コマンドプロンプトまたは PowerShell を開き、プロジェクトフォルダに移動:

```cmd
cd "C:\path\to\清掃スケジュール対応"
```

必要なライブラリをインストール:

```cmd
pip install -r requirements.txt
pip install pyinstaller
```

### ステップ2: PyInstallerでビルド

以下のコマンドを実行:

```cmd
pyinstaller --onefile --windowed --name CleaningSchedule --icon=icon.ico main.py
```

#### オプションの説明:

- `--onefile`: 単一の実行ファイルを生成
- `--windowed`: コンソールウィンドウを表示しない（GUIアプリ用）
- `--name CleaningSchedule`: 生成ファイル名を指定
- `--icon=icon.ico`: アイコンファイルを指定（オプション、なくても可）

### ステップ3: ビルド完了

ビルドが完了すると、以下のフォルダが作成されます:

```
清掃スケジュール対応/
├── build/              # ビルド中間ファイル（削除可能）
├── dist/               # 実行ファイルの出力先
│   └── CleaningSchedule.exe  ← これが実行ファイル
└── CleaningSchedule.spec  # ビルド設定ファイル
```

### ステップ4: 実行ファイルのテスト

`dist/CleaningSchedule.exe` をダブルクリックして動作確認:

1. アプリが起動するか
2. 設定ボタンでAPIキーが設定できるか
3. 画像選択ができるか
4. Excel変換が正常に動作するか

## 📦 配布方法

### 配布ファイルの準備

`dist/CleaningSchedule.exe` を配布します。

#### 推奨: ZIP ファイルで配布

```cmd
# PowerShellで実行
Compress-Archive -Path "dist\CleaningSchedule.exe" -DestinationPath "CleaningSchedule_Windows.zip"
```

または、手動で:

1. `dist` フォルダを開く
2. `CleaningSchedule.exe` を右クリック
3. 「送る」→「圧縮 (zip) フォルダー」

### 配布先のユーザーへの案内

**使い方:**

1. ZIP ファイルを解凍
2. `CleaningSchedule.exe` をダブルクリック
3. 初回起動時に「設定」ボタンから Claude API キーを入力
4. 画像を選択して変換開始

**注意事項:**

- Python のインストールは不要
- インターネット接続が必要（Claude API 使用のため）
- Windows Defender で警告が出る場合があります（「詳細情報」→「実行」で起動可能）

## 🔧 高度な設定

### アイコンを追加する場合

1. `.ico` 形式のアイコンファイルを用意
2. プロジェクトフォルダに `icon.ico` として配置
3. ビルドコマンドに `--icon=icon.ico` を追加

### ファイルサイズを小さくする

以下のオプションを追加:

```cmd
pyinstaller --onefile --windowed --name CleaningSchedule ^
  --exclude-module tkinter ^
  --exclude-module matplotlib ^
  main.py
```

### デバッグ情報を表示する

開発中は `--windowed` を外してコンソールを表示:

```cmd
pyinstaller --onefile --name CleaningSchedule main.py
```

## 📊 ビルド結果

### ファイルサイズ

- 通常: 約 50-80 MB
- 最適化後: 約 40-60 MB

### 起動時間

- 初回起動: 2-3秒
- 2回目以降: 1-2秒

## ⚠️ トラブルシューティング

### エラー: "module not found"

必要なライブラリがインストールされていない可能性があります:

```cmd
pip install -r requirements.txt
pip install pyinstaller
```

### 実行ファイルが起動しない

1. `--windowed` を外してコンソール版でビルド
2. エラーメッセージを確認:

```cmd
pyinstaller --onefile main.py
dist\main.exe
```

### Windows Defender で警告が出る

PyInstaller で作成した実行ファイルは、ウイルススキャンで誤検知される場合があります:

1. 「詳細情報」をクリック
2. 「実行」ボタンをクリック

または、デジタル署名を追加（高度な方法）。

### 生成ファイルが大きすぎる

不要なモジュールを除外:

```cmd
pyinstaller --onefile --windowed ^
  --exclude-module pytest ^
  --exclude-module notebook ^
  main.py
```

## 🔄 再ビルド

コードを修正した後、再ビルドする場合:

```cmd
# クリーンビルド
rmdir /s /q build dist
del CleaningSchedule.spec

# 再ビルド
pyinstaller --onefile --windowed --name CleaningSchedule main.py
```

## 📝 spec ファイルのカスタマイズ

より高度な設定が必要な場合、`CleaningSchedule.spec` を編集:

```python
# CleaningSchedule.spec
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],  # 追加ファイルをここに指定
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CleaningSchedule',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # UPX圧縮を有効化
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUIモード
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico'  # アイコン指定
)
```

編集後、以下でビルド:

```cmd
pyinstaller CleaningSchedule.spec
```

## 🎁 配布パッケージの作成

プロフェッショナルな配布パッケージを作成する場合:

```
CleaningSchedule_v1.0_Windows/
├── CleaningSchedule.exe
├── README.txt
└── sample/
    └── 清掃スケジュール_サンプル.jpg
```

ZIPで圧縮して配布。

## ✅ チェックリスト

ビルド前:

- [ ] すべての機能が動作することを確認
- [ ] config.json が含まれていないか確認（APIキー漏洩防止）
- [ ] requirements.txt が最新か確認

ビルド後:

- [ ] .exe が起動するか
- [ ] すべての機能が動作するか
- [ ] ファイルサイズが妥当か（100MB以下推奨）
- [ ] 配布用ドキュメントを用意したか

## 📞 サポート

ビルドに問題がある場合:

1. Python バージョンを確認: `python --version`
2. PyInstaller バージョンを確認: `pyinstaller --version`
3. 依存ライブラリを再インストール: `pip install -r requirements.txt --force-reinstall`

## 🔗 参考リンク

- PyInstaller 公式ドキュメント: https://pyinstaller.org/
- CustomTkinter: https://github.com/TomSchimansky/CustomTkinter
- Anthropic API: https://docs.anthropic.com/

