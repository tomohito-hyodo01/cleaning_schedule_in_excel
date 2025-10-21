# Excel画像変換ツール

Excelレイアウトのスクリーンショット画像（JPG/PNG）を解析し、元画像と見た目がほぼ同一のExcel（.xlsx）を生成するPythonアプリケーションです。

## 特徴

- **GUI & CLI対応**: PySide6による直感的なGUIと、自動化に便利なCLI両方に対応
- **高精度なOCR**: PaddleOCR（推奨）またはTesseractによる日本語認識
- **表構造抽出**: OpenCVによる罫線検出とセル分割
- **精密な再現**: 列幅、行高、罫線スタイル、セル結合を忠実に再現
- **デバッグ機能**: 検出結果の可視化と精度レポート

## 技術スタック

- **UI**: PySide6（GUI）、argparse（CLI）
- **画像処理**: OpenCV
- **OCR**: PaddleOCR（推奨）/ Tesseract
- **Excel出力**: openpyxl
- **パッケージ化**: PyInstaller

## インストール

### 必要要件

- Python 3.8以上
- Windows/macOS/Linux

### 依存関係のインストール

```bash
# リポジトリをクローン
git clone <repository-url>
cd cleaning_schedule_in_excel

# 仮想環境を作成（推奨）
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 依存関係をインストール
pip install -r requirements.txt
```

### OCRエンジンの設定

#### PaddleOCR（推奨）

PaddleOCRは自動的にモデルをダウンロードします。初回実行時に数百MBのダウンロードが発生します。

#### Tesseract

Tesseractを使用する場合は、事前にインストールが必要です。

**Windows:**
1. [Tesseract-OCR](https://github.com/UB-Mannheim/tesseract/wiki)をダウンロード
2. インストール時に「Additional language data」で日本語を選択
3. PATHに追加

**macOS:**
```bash
brew install tesseract tesseract-lang
```

**Linux:**
```bash
sudo apt-get install tesseract-ocr tesseract-ocr-jpn tesseract-ocr-jpn-vert
```

## 使い方

### GUIモード（デフォルト）

```bash
python main.py
```

または

```bash
python main.py --gui
```

GUIでは以下の操作が可能です：
- 画像をドラッグ＆ドロップ
- 設定パネルで各種パラメータを調整
- 進捗とログをリアルタイム表示
- 完了後にExcelファイルを直接開く

### CLIモード

```bash
# 基本的な使い方
python main.py --cli input.png -o output.xlsx

# 設定を指定
python main.py --cli input.png -o output.xlsx \
  --scale-col 1.2 \
  --scale-row 0.9 \
  --engine paddleocr \
  --log-level DEBUG

# 詳細なヘルプを表示
python main.py --cli --help
```

#### CLIオプション

- `input`: 入力画像ファイル（必須）
- `-o, --output`: 出力Excelファイル（デフォルト: output.xlsx）
- `--config`: 設定ファイル（JSON）のパス
- `--scale-col`: 列幅の倍率（デフォルト: 1.0）
- `--scale-row`: 行高の倍率（デフォルト: 1.0）
- `--min-line`: 罫線の最小長さ（ピクセル、デフォルト: 30）
- `--double-line-th`: 二重線の閾値（デフォルト: 2.0）
- `--engine`: OCRエンジン（paddleocr/tesseract、デフォルト: paddleocr）
- `--no-vertical`: 縦書き検出を無効化
- `--log-level`: ログレベル（DEBUG/INFO/WARNING/ERROR）
- `--no-debug-images`: デバッグ画像の保存を無効化
- `--debug-dir`: デバッグ画像の出力ディレクトリ

### サンプル実行

```bash
# サンプル画像を変換
python main.py --cli schedule_image/清掃スケジュール.jpg -o 清掃スケジュール_出力.xlsx
```

## 設定ファイル

JSON形式の設定ファイルで詳細な設定が可能です。

```json
{
  "preprocess": {
    "enable_rotation_correction": true,
    "binarization_method": "otsu"
  },
  "table_detection": {
    "min_line_length": 30,
    "thin_line_threshold": 2
  },
  "ocr": {
    "engine": "paddleocr",
    "confidence_threshold": 0.85
  },
  "excel": {
    "column_width_scale": 1.0,
    "row_height_scale": 1.0,
    "default_font": "MS Gothic"
  }
}
```

使用例：
```bash
python main.py --cli input.png -o output.xlsx --config config.json
```

## テスト

```bash
# すべてのテストを実行
python -m unittest discover tests

# 特定のテストを実行
python -m unittest tests.test_preprocess
```

## パッケージ化（PyInstaller）

### Windows

```bash
# 単一の実行ファイルを作成
pyinstaller --onefile --windowed --name="Excel画像変換" main.py

# 実行ファイルは dist/ ディレクトリに生成されます
```

### 詳細な手順

1. **specファイルの作成**

```bash
pyi-makespec --onefile --windowed --name="Excel画像変換" main.py
```

2. **specファイルの編集**

PaddleOCRのモデルファイルやTesseractのデータを含める場合は、specファイルを編集します：

```python
# Excel画像変換.spec

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        # PaddleOCRモデル（必要に応じて）
        ('venv/Lib/site-packages/paddleocr', 'paddleocr'),
    ],
    hiddenimports=[
        'paddleocr',
        'pytesseract',
        'openpyxl',
        'PySide6',
    ],
    ...
)
```

3. **ビルド実行**

```bash
pyinstaller "Excel画像変換.spec"
```

### サイズ削減のヒント

- `--exclude-module`: 不要なモジュールを除外
- `--strip`: デバッグシンボルを削除
- UPXで圧縮（`--upx-dir`）

## プロジェクト構造

```
.
├── src/                    # ソースコード
│   ├── __init__.py
│   ├── config.py          # 設定管理
│   ├── preprocess.py      # 画像前処理
│   ├── table_detect.py    # 表構造抽出
│   ├── ocr.py            # OCR処理
│   ├── excel_writer.py   # Excel生成
│   ├── cli.py            # CLIインターフェース
│   └── ui.py             # GUIインターフェース
├── tests/                 # テスト
│   ├── __init__.py
│   ├── test_preprocess.py
│   ├── test_table_detect.py
│   ├── test_ocr.py
│   └── test_excel_writer.py
├── examples/              # サンプル
├── schedule_image/        # 入力画像
├── debug_output/         # デバッグ出力（自動生成）
├── main.py               # メインエントリーポイント
├── requirements.txt      # 依存関係
└── README.md            # このファイル
```

## アルゴリズムの概要

### 1. 画像前処理

- **傾き補正**: HoughLinesPで直線を検出し、角度の中央値で回転
- **二値化**: Otsu法または適応的二値化
- **ノイズ除去**: モルフォロジー演算

### 2. 表構造抽出

- **罫線検出**: 形態学処理で横線・縦線を分離
- **グリッド生成**: 交点座標をソートしてグリッド化
- **セル結合判定**: 内部罫線の有無で判定

### 3. OCR

- **セル画像切り出し**: 各セルを個別に処理
- **テキスト認識**: PaddleOCRまたはTesseract
- **正規化**: 全角→半角、誤認識修正
- **再試行**: 低信頼度セルは画像強調して再実行

### 4. Excel生成

- **寸法変換**: ピクセル→Excel単位（列幅・行高）
- **罫線スタイル**: 線幅からthin/medium/thickを判定
- **セル結合**: rowspan/colspanを反映
- **書式設定**: フォント、配置、塗りつぶし

## 品質基準

- **罫線位置の再現**: 平均誤差 ≤ 3px
- **結合セルの再現率**: ≥ 98%
- **文字一致率**: ≥ 98%（正規化後）
- **罫線スタイル判定**: ≥ 95%

## 既知の制約

- **画像品質**: 低解像度や強い圧縮ノイズがある画像は精度が低下
- **複雑な結合**: 不規則な結合セルは誤検出の可能性
- **縦書き**: 縦書きテキストの精度はOCRエンジンに依存
- **色情報**: 背景色や文字色の完全な再現は困難
- **画像・グラフ**: セル内の画像やグラフは再現不可

## トラブルシューティング

### PaddleOCRのエラー

```
ImportError: cannot import name 'PaddleOCR'
```

→ PaddlePaddleとPaddleOCRを再インストール：
```bash
pip install --upgrade paddlepaddle paddleocr
```

### Tesseractが見つからない

```
TesseractNotFoundError
```

→ Tesseractをインストールし、PATHを設定

### メモリ不足

大きな画像を処理する際にメモリ不足になる場合は、画像を縮小してから処理：

```python
# 画像を縮小
import cv2
image = cv2.imread("large_image.png")
scale = 0.5
resized = cv2.resize(image, None, fx=scale, fy=scale)
cv2.imwrite("resized.png", resized)
```

## ライセンス

MIT License

## 作者

Created by AI Assistant

## 貢献

プルリクエストを歓迎します！バグ報告や機能要望はIssueでお願いします。

## 参考資料

- [OpenCV Documentation](https://docs.opencv.org/)
- [PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR)
- [openpyxl Documentation](https://openpyxl.readthedocs.io/)
- [PySide6 Documentation](https://doc.qt.io/qtforpython/)

