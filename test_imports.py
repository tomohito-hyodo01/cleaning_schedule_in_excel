"""
依存関係のインポートテスト

このスクリプトで依存関係が正しくインストールされているか確認できます
"""

import sys
import io

# Windowsのコンソールエンコーディング問題を回避
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

print("=" * 60)
print("依存関係チェック")
print("=" * 60)
print()

# 必須モジュール
required_modules = [
    ("numpy", "NumPy"),
    ("cv2", "OpenCV"),
    ("PIL", "Pillow"),
    ("openpyxl", "openpyxl"),
]

# オプショナルモジュール（GUIやOCR）
optional_modules = [
    ("paddleocr", "PaddleOCR"),
    ("paddle", "PaddlePaddle"),
    ("pytesseract", "pytesseract"),
    ("PySide6", "PySide6 (GUIに必要)"),
]

all_ok = True

print("【必須モジュール】")
for module_name, display_name in required_modules:
    try:
        __import__(module_name)
        print(f"  [OK] {display_name}")
    except ImportError as e:
        print(f"  [NG] {display_name} - 未インストール")
        all_ok = False

print()
print("【オプショナルモジュール】")
for module_name, display_name in optional_modules:
    try:
        __import__(module_name)
        print(f"  [OK] {display_name}")
    except ImportError:
        print(f"  [--] {display_name} - 未インストール (オプション)")

print()
print("=" * 60)

if all_ok:
    print("[OK] すべての必須モジュールがインストールされています")
    print()
    print("次のステップ:")
    print("  GUIモード: python main.py")
    print("  CLIモード: python main.py --cli schedule_image/清掃スケジュール.jpg -o output.xlsx")
else:
    print("[NG] いくつかの必須モジュールがインストールされていません")
    print()
    print("インストール方法:")
    print("  pip install -r requirements.txt")

print("=" * 60)

