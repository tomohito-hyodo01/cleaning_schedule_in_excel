# パッケージ化手順（PyInstaller）

このドキュメントでは、PyInstallerを使用してアプリケーションを単一の実行ファイルにパッケージ化する手順を説明します。

## 前提条件

- Python環境がセットアップされていること
- すべての依存関係がインストールされていること
- アプリケーションが正常に動作することを確認済み

## 基本的なパッケージ化

### Windows

#### 1. PyInstallerのインストール

```bash
pip install pyinstaller
```

#### 2. specファイルの生成

```bash
pyi-makespec --onefile --windowed --name="Excel画像変換" main.py
```

これにより `Excel画像変換.spec` ファイルが生成されます。

#### 3. specファイルの編集

生成された `.spec` ファイルを編集して、必要なデータファイルを含めます：

```python
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        # PaddleOCRのモデルファイル（自動ダウンロードされるため通常は不要）
        # ('C:/Users/<username>/.paddleocr', 'paddleocr'),
    ],
    hiddenimports=[
        'paddleocr',
        'paddle',
        'pytesseract',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.cell',
        'PySide6',
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
        'cv2',
        'numpy',
        'PIL',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'pandas',
        'tkinter',
    ],
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
    name='Excel画像変換',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUIアプリなのでコンソールを非表示
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # アイコンファイルのパスを指定可能
)
```

#### 4. ビルドの実行

```bash
pyinstaller "Excel画像変換.spec"
```

ビルドが完了すると、`dist/Excel画像変換.exe` が生成されます。

### macOS

macOSでは同様の手順ですが、いくつか違いがあります：

```bash
# specファイルの生成
pyi-makespec --onefile --windowed --name="Excel画像変換" main.py

# ビルド
pyinstaller "Excel画像変換.spec"
```

生成されるファイル：
- `dist/Excel画像変換.app` - macOSアプリケーションバンドル
- または `dist/Excel画像変換` - 単一実行ファイル

### Linux

Linuxでも同様の手順です：

```bash
pyi-makespec --onefile --name="excel_converter" main.py
pyinstaller excel_converter.spec
```

生成されるファイル：
- `dist/excel_converter` - 実行ファイル

## 高度な設定

### アイコンの追加

アイコンファイル（`.ico`）を用意し、specファイルで指定：

```python
exe = EXE(
    ...
    icon='icon.ico',  # Windowsの場合
    icon='icon.icns',  # macOSの場合
)
```

### ファイルサイズの削減

#### 不要なモジュールを除外

```python
a = Analysis(
    ...
    excludes=[
        'matplotlib',
        'scipy',
        'pandas',
        'tkinter',
        'IPython',
        'jupyter',
    ],
)
```

#### UPXで圧縮

UPXをインストール：
- Windows: [UPX releases](https://github.com/upx/upx/releases)をダウンロード
- macOS: `brew install upx`
- Linux: `sudo apt-get install upx`

specファイルで有効化：

```python
exe = EXE(
    ...
    upx=True,
    upx_exclude=[
        'vcruntime140.dll',  # 圧縮すると問題が起きることがある
        'python3.dll',
    ],
)
```

### ディレクトリモード（より高速）

`--onefile` の代わりにディレクトリモードを使用すると、起動が高速になります：

```bash
pyi-makespec --windowed --name="Excel画像変換" main.py
pyinstaller "Excel画像変換.spec"
```

この場合、`dist/Excel画像変換/` ディレクトリ全体を配布する必要があります。

## モデルファイルの同梱

### PaddleOCR

PaddleOCRのモデルは初回実行時に自動ダウンロードされます。事前に同梱する場合：

1. モデルをダウンロード（初回実行時に `~/.paddleocr/` に保存される）
2. specファイルに追加：

```python
import os
from pathlib import Path

paddle_model_dir = Path.home() / '.paddleocr'

a = Analysis(
    ...
    datas=[
        (str(paddle_model_dir), 'paddleocr_models'),
    ],
)
```

3. 実行時にモデルディレクトリを設定：

```python
# src/ocr.py内
import sys
if getattr(sys, 'frozen', False):
    # PyInstallerで実行されている場合
    model_dir = os.path.join(sys._MEIPASS, 'paddleocr_models')
    os.environ['PADDLEOCR_MODEL_DIR'] = model_dir
```

### Tesseract

Tesseractは外部プログラムなので、以下の方法があります：

1. **ユーザーに別途インストールを依頼**（推奨）
2. **Tesseractを同梱**:
   - Tesseractのバイナリとtessdataを同梱
   - specファイルに追加
   - 実行時にパスを設定

## ビルド後のテスト

### Windows

```cmd
# 実行ファイルをテスト
dist\Excel画像変換.exe

# CLIモードでテスト
dist\Excel画像変換.exe --cli schedule_image\清掃スケジュール.jpg -o test.xlsx
```

### macOS/Linux

```bash
# 実行権限を付与
chmod +x dist/Excel画像変換

# 実行
./dist/Excel画像変換
```

## 配布

### Windows

単一ファイル（`Excel画像変換.exe`）を配布するか、ZIPで圧縮して配布。

### macOS

`.app` バンドルをDMGファイルにパッケージ化して配布（オプション）：

```bash
# hdiutilを使用
hdiutil create -volname "Excel画像変換" -srcfolder dist/Excel画像変換.app -ov -format UDZO Excel画像変換.dmg
```

### インストーラーの作成

より専門的な配布を行う場合：

- **Windows**: Inno Setup、NSIS
- **macOS**: pkgbuild、productbuild
- **Linux**: .deb、.rpm、AppImage

## トラブルシューティング

### インポートエラー

```
ModuleNotFoundError: No module named 'xxx'
```

→ specファイルの `hiddenimports` に追加

### ファイルが見つからない

```
FileNotFoundError: [Errno 2] No such file or directory
```

→ specファイルの `datas` に追加

### 起動が遅い

- `--onefile` は起動時に一時ディレクトリに展開するため遅い
- ディレクトリモードを使用するか、一部のファイルを外部化

### サイズが大きすぎる

- `excludes` で不要なモジュールを除外
- UPXで圧縮
- 仮想環境を使用して最小限の依存関係のみインストール

### OCRエンジンが動作しない

- PaddleOCR: モデルファイルのパスを確認
- Tesseract: インストールとPATHを確認

## バッチビルドスクリプト

### Windows（build.bat）

```bat
@echo off
echo Excel画像変換ツールをビルドしています...

REM 依存関係を確認
pip install -r requirements.txt
pip install pyinstaller

REM 古いビルドを削除
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM ビルド実行
pyinstaller --onefile --windowed --name="Excel画像変換" main.py

echo ビルド完了！
echo 実行ファイル: dist\Excel画像変換.exe
pause
```

### macOS/Linux（build.sh）

```bash
#!/bin/bash
echo "Excel画像変換ツールをビルドしています..."

# 依存関係を確認
pip install -r requirements.txt
pip install pyinstaller

# 古いビルドを削除
rm -rf build dist

# ビルド実行
pyinstaller --onefile --windowed --name="Excel画像変換" main.py

echo "ビルド完了！"
echo "実行ファイル: dist/Excel画像変換"
```

## まとめ

1. `pip install pyinstaller`
2. `pyi-makespec --onefile --windowed main.py`
3. specファイルを編集（必要に応じて）
4. `pyinstaller <specfile>`
5. `dist/` ディレクトリの実行ファイルをテスト
6. 配布

パッケージ化に問題がある場合は、PyInstallerのログ（`build/` ディレクトリ）を確認してください。

