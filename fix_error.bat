@echo off
chcp 65001 > nul
echo ======================================
echo エラー修正スクリプト
echo proxies エラーを修正します
echo ======================================
echo.

echo [1/3] 問題のあるライブラリをアンインストール中...
pip uninstall httpx anthropic -y

if %errorlevel% neq 0 (
    echo 警告: アンインストールに一部失敗しましたが続行します
)

echo.
echo [2/3] 修正版ライブラリをインストール中...
pip install httpx==0.27.2
pip install anthropic==0.18.1

if %errorlevel% neq 0 (
    echo.
    echo エラー: ライブラリのインストールに失敗しました
    pause
    exit /b 1
)

echo.
echo [3/3] インストール確認...
pip show httpx | findstr "Version"
pip show anthropic | findstr "Version"

echo.
echo ======================================
echo 修正完了！
echo ======================================
echo.
echo アプリを起動して再テストしてください:
echo   run.bat
echo.
pause

