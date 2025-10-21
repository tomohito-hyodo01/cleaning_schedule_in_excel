@echo off
chcp 65001 > nul
REM 最小限の依存関係をインストール（CLIモードのみ）

echo ========================================
echo 最小インストール（CLIモードのみ）
echo ========================================
echo.

echo 基本ライブラリをインストール中...
pip install numpy opencv-python Pillow openpyxl

echo.
echo OCRエンジンをインストール中...
echo （これには時間がかかります）
pip install paddlepaddle paddleocr

echo.
echo ========================================
echo インストール完了！
echo ========================================
echo.
echo テスト実行:
echo   python test_imports.py
echo.
echo サンプル実行:
echo   python main.py --cli schedule_image\清掃スケジュール.jpg -o output.xlsx
echo.
echo 注意: GUIモードを使用する場合は以下も必要です:
echo   pip install PySide6
echo.

pause

