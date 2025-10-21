@echo off
chcp 65001 > nul
REM 依存関係をインストールするバッチファイル

echo ========================================
echo Excel画像変換ツール - インストール
echo ========================================
echo.

REM Pythonのバージョン確認
echo Pythonバージョンを確認中...
python --version
if errorlevel 1 (
    echo エラー: Pythonがインストールされていません
    echo Python 3.8以上をインストールしてください
    pause
    exit /b 1
)

echo.
echo 依存関係をインストール中...
echo これには数分かかる場合があります...
echo.

REM 依存関係をインストール
pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo エラー: インストールに失敗しました
    pause
    exit /b 1
)

echo.
echo ========================================
echo インストール完了！
echo ========================================
echo.
echo 次のコマンドでアプリを起動できます:
echo   python main.py              (GUIモード)
echo   run_sample.bat              (サンプル実行)
echo.

pause

