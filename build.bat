@echo off
chcp 65001 > nul
echo ======================================
echo 清掃スケジュール Excel変換ツール
echo Windows実行ファイルビルドスクリプト
echo ======================================
echo.

echo [1/3] PyInstaller をインストール中...
pip install pyinstaller

if %errorlevel% neq 0 (
    echo.
    echo エラー: PyInstaller のインストールに失敗しました
    pause
    exit /b 1
)

echo.
echo [2/3] 実行ファイルをビルド中...
echo （数分かかる場合があります）
echo.

pyinstaller --onefile --windowed --name CleaningSchedule main.py

if %errorlevel% neq 0 (
    echo.
    echo エラー: ビルドに失敗しました
    pause
    exit /b 1
)

echo.
echo [3/3] ビルド完了！
echo.
echo 実行ファイルの場所:
echo   dist\CleaningSchedule.exe
echo.
echo このファイルをダブルクリックで起動できます
echo.

choice /C YN /M "エクスプローラーで開きますか？"
if %errorlevel% equ 1 (
    explorer dist
)

pause

