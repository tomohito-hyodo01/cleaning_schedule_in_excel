#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
サンプル画像を変換する（デバッグ版）

エラーを詳細に表示し、どの段階で問題が発生するか確認
"""

import sys
import os
from pathlib import Path
import traceback
import io

# Windowsのコンソールエンコーディング問題を回避
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# スクリプトのディレクトリをPythonパスに追加
script_dir = Path(__file__).parent.resolve()
if str(script_dir) not in sys.path:
    sys.path.insert(0, str(script_dir))

print("=" * 60)
print("Excel画像変換ツール - デバッグモード")
print("=" * 60)
print()

# schedule_imageディレクトリ内のJPG/PNG画像を探す
schedule_dir = script_dir / "schedule_image"
if not schedule_dir.exists():
    print(f"エラー: {schedule_dir} が見つかりません")
    sys.exit(1)

# 画像ファイルを検索
image_files = list(schedule_dir.glob("*.jpg")) + list(schedule_dir.glob("*.png"))

if not image_files:
    print(f"エラー: {schedule_dir} 内に画像ファイルが見つかりません")
    sys.exit(1)

# 最初の画像ファイルを使用
input_image = image_files[0]
output_excel = script_dir / f"{input_image.stem}_変換結果.xlsx"

print(f"入力画像: {input_image.name}")
print(f"出力Excel: {output_excel.name}")
print()

try:
    # 依存関係の確認
    print("[1/6] 依存関係を確認中...")
    import cv2
    import numpy as np
    import openpyxl
    print("  [OK] 必須モジュール確認完了")
    print()
    
    # 設定の作成
    print("[2/6] 設定を読み込み中...")
    from src.config import AppConfig
    config = AppConfig.default()
    config.debug.log_level = "INFO"
    config.debug.save_debug_images = True
    print("  [OK] 設定読み込み完了")
    print()
    
    # 画像の読み込み
    print("[3/6] 画像を読み込み中...")
    with open(input_image, 'rb') as f:
        file_bytes = np.frombuffer(f.read(), dtype=np.uint8)
    image = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)
    if image is None:
        raise ValueError("画像のデコードに失敗")
    print(f"  [OK] 画像読み込み完了: {image.shape[1]}x{image.shape[0]}")
    print()
    
    # 前処理
    print("[4/6] 画像を前処理中...")
    from src.preprocess import create_preprocessor
    preprocessor = create_preprocessor(config.preprocess)
    binary_image, preprocess_info = preprocessor.preprocess(image)
    print(f"  [OK] 前処理完了")
    print()
    
    # 表構造検出
    print("[5/6] 表構造を検出中...")
    from src.table_detect import create_detector
    detector = create_detector(config.table_detection)
    cells, border_styles, detect_info = detector.detect(binary_image, image)
    print(f"  [OK] 表構造検出完了")
    print(f"    - セル数: {len(cells)}")
    print(f"    - グリッド: {detect_info.get('grid_rows', 0)}行 × {detect_info.get('grid_cols', 0)}列")
    print()
    
    # OCR処理（エラーが発生しやすい部分）
    print("[6/6] OCR処理中...")
    print("  注意: この処理には時間がかかります（数分）")
    
    # セル数が多い場合は警告
    if len(cells) > 100:
        print(f"  警告: セル数が多いため、処理に時間がかかります（推定: {len(cells) * 2}秒）")
    
    from src.ocr import create_ocr_processor
    
    try:
        ocr_processor = create_ocr_processor(config.ocr)
        
        # プログレス表示付きでOCR実行
        cell_texts = {}
        total = len(cells)
        
        for i, cell in enumerate(cells, 1):
            if i % 10 == 0 or i == 1:
                print(f"    処理中... {i}/{total} ({i*100//total}%)")
            
            # セル画像を切り出し
            margin = 3
            x = max(0, cell.x + margin)
            y = max(0, cell.y + margin)
            x2 = min(image.shape[1], cell.x + cell.width - margin)
            y2 = min(image.shape[0], cell.y + cell.height - margin)
            
            if x2 <= x or y2 <= y:
                cell_texts[(cell.row, cell.col)] = ("", 0.0)
                continue
            
            cell_image = image[y:y2, x:x2]
            
            if cell_image.shape[0] < 5 or cell_image.shape[1] < 5:
                cell_texts[(cell.row, cell.col)] = ("", 0.0)
                continue
            
            # OCR実行（エラーが出てもスキップ）
            try:
                text, confidence = ocr_processor.engine.recognize(cell_image)
                if config.ocr.enable_normalization and text:
                    from src.ocr import TextNormalizer
                    text = TextNormalizer.normalize(text)
                cell_texts[(cell.row, cell.col)] = (text, confidence)
            except Exception as e:
                # エラーが出ても続行
                cell_texts[(cell.row, cell.col)] = ("", 0.0)
        
        print(f"  [OK] OCR処理完了: {len(cell_texts)}セル")
        
    except Exception as ocr_error:
        print(f"  警告: OCR処理でエラーが発生しましたが続行します")
        print(f"    エラー: {ocr_error}")
        # 空のテキストで埋める
        cell_texts = {(cell.row, cell.col): ("", 0.0) for cell in cells}
    
    print()
    
    # Excel生成
    print("[7/7] Excelファイルを生成中...")
    from src.excel_writer import create_writer
    writer = create_writer(config.excel)
    excel_info = writer.write(cells, cell_texts, border_styles, str(output_excel))
    print(f"  [OK] Excel生成完了")
    print()
    
    # 完了
    print("=" * 60)
    print("[OK] 変換完了！")
    print("=" * 60)
    print()
    print(f"出力ファイル: {output_excel}")
    print(f"サイズ: {excel_info.get('columns', 0)}列 × {excel_info.get('rows', 0)}行")
    print(f"結合セル: {excel_info.get('merged_cells', 0)}個")
    print()
    print(f"デバッグ画像: debug_output/")
    print()
    
    # Excelファイルを開くか確認
    try:
        response = input("Excelファイルを開きますか？ (y/n): ")
        if response.lower() == 'y':
            import subprocess
            import platform
            
            system = platform.system()
            if system == "Windows":
                os.startfile(str(output_excel))
            elif system == "Darwin":  # macOS
                subprocess.run(["open", str(output_excel)])
            else:  # Linux
                subprocess.run(["xdg-open", str(output_excel)])
    except KeyboardInterrupt:
        print("\n終了します")

except KeyboardInterrupt:
    print("\n\n処理がキャンセルされました")
    sys.exit(1)

except Exception as e:
    print()
    print("=" * 60)
    print("[NG] エラーが発生しました")
    print("=" * 60)
    print()
    print(f"エラーの種類: {type(e).__name__}")
    print(f"エラーメッセージ: {e}")
    print()
    print("詳細なトレースバック:")
    print("-" * 60)
    traceback.print_exc()
    print("-" * 60)
    print()
    print("対処方法:")
    print("  1. 画像サイズが大きすぎる場合は縮小してください")
    print("  2. メモリ不足の場合は他のアプリを閉じてください")
    print("  3. PaddleOCRのエラーの場合は --engine tesseract を試してください")
    print()
    sys.exit(1)

