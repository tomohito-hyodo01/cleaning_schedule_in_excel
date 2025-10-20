@echo off
chcp 65001 > nul
echo 清掃スケジュール Excel変換ツールを起動中...
python main.py
if %errorlevel% neq 0 (
    echo.
    echo エラー: アプリの起動に失敗しました
    echo.
    echo セットアップを実行してください:
    echo   setup.bat
    echo.
    pause
)

