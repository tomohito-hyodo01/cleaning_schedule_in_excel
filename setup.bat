@echo off
chcp 65001 > nul
echo ======================================
echo 清掃スケジュール Excel変換ツール
echo セットアップスクリプト
echo ======================================
echo.

echo [1/2] 依存ライブラリをインストール中...
pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo.
    echo エラー: ライブラリのインストールに失敗しました
    echo.
    echo Python がインストールされているか確認してください:
    echo https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo.
echo [2/2] セットアップ完了！
echo.
echo アプリを起動するには以下を実行してください:
echo   python main.py
echo.
echo または、実行ファイル版をビルドするには:
echo   build.bat
echo.
pause

