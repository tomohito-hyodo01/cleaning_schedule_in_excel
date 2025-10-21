@echo off
chcp 65001 > nul
REM サンプル画像を変換するバッチファイル（Windows用）

echo ========================================
echo Excel画像変換ツール - サンプル実行
echo ========================================
echo.

REM 依存関係の確認
echo 依存関係を確認中...
python -c "import cv2, openpyxl" 2>nul
if errorlevel 1 (
    echo エラー: 必要なライブラリがインストールされていません
    echo 以下のコマンドを実行してください:
    echo   pip install -r requirements.txt
    pause
    exit /b 1
)

echo.
echo サンプル画像を変換しています...
echo 入力: schedule_image\清掃スケジュール.jpg
echo 出力: 清掃スケジュール_変換結果.xlsx
echo.

REM 変換実行（文字エンコーディング問題を回避するためPythonスクリプトを使用）
python run_sample_safe.py

if errorlevel 1 (
    echo.
    echo エラー: 変換に失敗しました
    pause
    exit /b 1
)

echo.
echo ========================================
echo 変換完了！
echo ========================================
echo.
echo 出力ファイル: 清掃スケジュール_変換結果.xlsx
echo デバッグ画像: debug_output\
echo.

REM Excelファイルを開くか確認
set /p OPEN="Excelファイルを開きますか？ (y/n): "
if /i "%OPEN%"=="y" (
    start 清掃スケジュール_変換結果.xlsx
)

pause
