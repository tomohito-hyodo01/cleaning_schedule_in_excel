#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
改善された設定でサンプル画像を変換

透かし文字対応、より良い二値化設定
"""

import sys
import os
from pathlib import Path
import io

# Windowsのコンソールエンコーディング問題を回避
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# スクリプトのディレクトリをPythonパスに追加
script_dir = Path(__file__).parent.resolve()
if str(script_dir) not in sys.path:
    sys.path.insert(0, str(script_dir))

print("=" * 70)
print("Excel画像変換ツール - 改善版設定")
print("=" * 70)
print()

# 画像ファイルを探す
schedule_dir = script_dir / "schedule_image"
image_files = list(schedule_dir.glob("*.jpg")) + list(schedule_dir.glob("*.png"))

if not image_files:
    print("エラー: 画像ファイルが見つかりません")
    sys.exit(1)

input_image = image_files[0]
output_excel = script_dir / f"{input_image.stem}_改善版.xlsx"

print(f"入力画像: {input_image.name}")
print(f"出力Excel: {output_excel.name}")
print(f"設定ファイル: config_for_sample.json")
print()

try:
    print("処理を開始します...")
    print("-" * 70)
    
    from src.cli import convert_image_to_excel
    from src.config import AppConfig
    
    # カスタム設定を読み込み
    config_path = script_dir / "config_for_sample.json"
    if config_path.exists():
        config = AppConfig.from_json(str(config_path))
        print("[OK] カスタム設定を読み込みました")
    else:
        config = AppConfig.default()
        print("[警告] デフォルト設定を使用します")
    
    print()
    
    # 変換実行
    info = convert_image_to_excel(
        str(input_image),
        str(output_excel),
        config
    )
    
    print()
    print("=" * 70)
    print("[OK] 変換完了！")
    print("=" * 70)
    print()
    print(f"出力ファイル: {output_excel}")
    
    if 'excel' in info:
        excel = info['excel']
        print(f"サイズ: {excel.get('columns', 0)}列 × {excel.get('rows', 0)}行")
        print(f"結合セル: {excel.get('merged_cells', 0)}個")
    
    if 'ocr_avg_confidence' in info:
        print(f"OCR平均信頼度: {info['ocr_avg_confidence']:.1%}")
    
    print()
    print("デバッグ画像を確認してください:")
    print("  - debug_output/清掃スケジュール_binary.png (二値化結果)")
    print("  - debug_output/清掃スケジュール_cells.png (セル検出結果)")
    print()
    
    # Excelファイルを開く
    try:
        response = input("Excelファイルを開きますか？ (y/n): ")
        if response.lower() == 'y':
            if sys.platform == 'win32':
                os.startfile(str(output_excel))
    except KeyboardInterrupt:
        print("\n終了します")
    
except KeyboardInterrupt:
    print("\n\n処理がキャンセルされました")
    sys.exit(1)
    
except Exception as e:
    print()
    print("=" * 70)
    print("[NG] エラーが発生しました")
    print("=" * 70)
    print()
    print(f"エラー: {e}")
    print()
    import traceback
    traceback.print_exc()
    sys.exit(1)

