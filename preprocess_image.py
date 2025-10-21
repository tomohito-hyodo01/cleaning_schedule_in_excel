#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
画像を前処理して、表構造検出を改善

透かし除去、コントラスト強調、ノイズ除去
"""

import cv2
import numpy as np
from pathlib import Path

print("=" * 70)
print("画像前処理ツール")
print("=" * 70)
print()

# 入力画像
input_path = Path("schedule_image/清掃スケジュール.jpg")
output_path = Path("schedule_image/清掃スケジュール_前処理済み.jpg")

# 画像を読み込み
with open(input_path, 'rb') as f:
    file_bytes = np.frombuffer(f.read(), dtype=np.uint8)
image = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)

print(f"元画像サイズ: {image.shape[1]}x{image.shape[0]}")
print()

# グレースケール化
gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

# 1. ガンマ補正で透かしを薄くする
print("[1/5] ガンマ補正...")
gamma = 0.6  # 暗くする
lookup = np.array([((i / 255.0) ** gamma) * 255 for i in range(256)]).astype(np.uint8)
gray_corrected = cv2.LUT(gray, lookup)

# 2. CLAHEでコントラスト強調
print("[2/5] コントラスト強調...")
clahe = cv2.createCLAHE(clipLimit=4.0, tileGridSize=(8, 8))
gray_clahe = clahe.apply(gray_corrected)

# 3. バイラテラルフィルタでノイズ除去（エッジ保持）
print("[3/5] ノイズ除去...")
gray_filtered = cv2.bilateralFilter(gray_clahe, 9, 75, 75)

# 4. シャープ化
print("[4/5] シャープ化...")
kernel_sharpen = np.array([[-1,-1,-1],
                           [-1, 9,-1],
                           [-1,-1,-1]])
gray_sharp = cv2.filter2D(gray_filtered, -1, kernel_sharpen)

# 5. アンシャープマスク
print("[5/5] 最終調整...")
gaussian = cv2.GaussianBlur(gray_sharp, (9, 9), 10.0)
unsharp = cv2.addWeighted(gray_sharp, 1.5, gaussian, -0.5, 0)

# カラーに戻す（3チャンネル）
result = cv2.cvtColor(unsharp, cv2.COLOR_GRAY2BGR)

# 保存
cv2.imwrite(str(output_path), result)

print()
print("=" * 70)
print("[OK] 前処理完了！")
print("=" * 70)
print()
print(f"出力ファイル: {output_path}")
print()
print("次のステップ:")
print("  1. 前処理済み画像を確認")
print(f"     → {output_path}")
print()
print("  2. 前処理済み画像で変換を実行")
print("     → python main.py --cli schedule_image/清掃スケジュール_前処理済み.jpg -o 結果.xlsx")
print()

# 比較画像を保存
comparison = np.hstack([
    cv2.cvtColor(gray, cv2.COLOR_GRAY2BGR),
    result
])
comparison_path = Path("debug_output/前処理_比較.jpg")
comparison_path.parent.mkdir(exist_ok=True)
cv2.imwrite(str(comparison_path), comparison)

print(f"比較画像も保存しました: {comparison_path}")
print("  （左: 元画像、右: 前処理後）")
print()

