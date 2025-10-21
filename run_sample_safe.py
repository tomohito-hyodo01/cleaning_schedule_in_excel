#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
サンプル画像を変換するPythonスクリプト

Windowsの文字エンコーディング問題を回避
"""

import sys
import os
from pathlib import Path

# スクリプトのディレクトリをPythonパスに追加
script_dir = Path(__file__).parent.resolve()
if str(script_dir) not in sys.path:
    sys.path.insert(0, str(script_dir))

print("=" * 60)
print("Excel画像変換ツール - サンプル実行")
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

# 依存関係の確認
print("依存関係を確認中...")
try:
    import cv2
    import openpyxl
    print("  [OK] 必須モジュール")
except ImportError as e:
    print(f"  [NG] 必須モジュールが不足: {e}")
    print()
    print("以下のコマンドを実行してください:")
    print("  pip install -r requirements.txt")
    sys.exit(1)

print()
print("変換を開始します...")
print("-" * 60)

# 変換を実行
from src.cli import convert_image_to_excel
from src.config import AppConfig

try:
    config = AppConfig.default()
    config.debug.log_level = "INFO"
    
    info = convert_image_to_excel(
        str(input_image),
        str(output_excel),
        config
    )
    
    print()
    print("=" * 60)
    print("変換完了！")
    print("=" * 60)
    print()
    print(f"出力ファイル: {output_excel}")
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

except Exception as e:
    print()
    print("=" * 60)
    print(f"エラー: 変換に失敗しました")
    print("=" * 60)
    print()
    print(f"詳細: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

