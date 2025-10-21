# クイックスタートガイド

このガイドでは、最速でアプリケーションを試す方法を説明します。

## 1. 環境セットアップ（5分）

### Python 3.8以上がインストールされているか確認

```bash
python --version
```

### 依存関係のインストール

```bash
# リポジトリをクローン（またはダウンロード）
cd cleaning_schedule_in_excel

# 依存関係をインストール
pip install -r requirements.txt
```

**注意**: 初回のPaddleOCR実行時に、モデルファイルが自動ダウンロードされます（数百MB）。

## 2. GUIで試す（推奨）

### 起動

```bash
python main.py
```

または

```bash
python main.py --gui
```

### 使い方

1. 画像をドラッグ＆ドロップ、または「画像を選択」ボタンでファイルを選択
2. 必要に応じて設定を調整
3. 「解析してExcel作成」ボタンをクリック
4. 完了したらExcelファイルが生成されます

## 3. CLIで試す

### サンプル画像を変換

#### Windows

```cmd
run_sample.bat
```

または

```cmd
python main.py --cli schedule_image\清掃スケジュール.jpg -o output.xlsx
```

#### macOS/Linux

```bash
chmod +x run_sample.sh
./run_sample.sh
```

または

```bash
python main.py --cli schedule_image/清掃スケジュール.jpg -o output.xlsx
```

### 自分の画像を変換

```bash
python main.py --cli your_image.png -o output.xlsx
```

## 4. 結果の確認

- **Excelファイル**: 指定した出力パスに生成されます
- **デバッグ画像**: `debug_output/` ディレクトリに保存されます
  - `*_binary.png`: 二値化された画像
  - `*_cells.png`: 検出されたセルの境界

## トラブルシューティング

### PaddleOCRのインストールに失敗する

```bash
# CPU版を明示的にインストール
pip install paddlepaddle -i https://mirror.baidu.com/pypi/simple
pip install paddleocr
```

### Tesseractを使いたい

1. Tesseractをインストール（[ダウンロード](https://github.com/UB-Mannheim/tesseract/wiki)）
2. 日本語データを含めてインストール
3. OCRエンジンを指定して実行

```bash
python main.py --cli input.png -o output.xlsx --engine tesseract
```

### メモリ不足

画像が大きすぎる場合は、事前に縮小してください：

```python
import cv2
img = cv2.imread("large.png")
resized = cv2.resize(img, None, fx=0.5, fy=0.5)
cv2.imwrite("resized.png", resized)
```

### OCRの精度が低い

1. 画像の解像度を確認（推奨: 150-300 DPI）
2. 設定でOCRエンジンを変更
3. `--log-level DEBUG` でログを確認

## 次のステップ

- [README.md](README.md) - 詳細なドキュメント
- [PACKAGING.md](PACKAGING.md) - パッケージ化手順
- `config_example.json` - 設定ファイルのサンプル

## サポート

問題が発生した場合は、以下を確認してください：

1. Python 3.8以上を使用しているか
2. すべての依存関係がインストールされているか
3. 入力画像が正しい形式（PNG/JPG）か
4. ログメッセージにエラーがないか

それでも解決しない場合は、Issueを作成してください。

