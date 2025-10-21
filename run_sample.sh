#!/bin/bash
# サンプル画像を変換するシェルスクリプト（macOS/Linux用）

echo "========================================"
echo "Excel画像変換ツール - サンプル実行"
echo "========================================"
echo ""

# 依存関係の確認
echo "依存関係を確認中..."
python3 -c "import cv2, paddleocr, openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "エラー: 必要なライブラリがインストールされていません"
    echo "以下のコマンドを実行してください:"
    echo "  pip install -r requirements.txt"
    exit 1
fi

echo ""
echo "サンプル画像を変換しています..."
echo "入力: schedule_image/清掃スケジュール.jpg"
echo "出力: 清掃スケジュール_変換結果.xlsx"
echo ""

# 変換実行
python3 main.py --cli schedule_image/清掃スケジュール.jpg -o 清掃スケジュール_変換結果.xlsx --log-level INFO

if [ $? -ne 0 ]; then
    echo ""
    echo "エラー: 変換に失敗しました"
    exit 1
fi

echo ""
echo "========================================"
echo "変換完了！"
echo "========================================"
echo ""
echo "出力ファイル: 清掃スケジュール_変換結果.xlsx"
echo "デバッグ画像: debug_output/"
echo ""

# Excelファイルを開くか確認
read -p "Excelファイルを開きますか？ (y/n): " OPEN
if [ "$OPEN" = "y" ] || [ "$OPEN" = "Y" ]; then
    if [[ "$OSTYPE" == "darwin"* ]]; then
        # macOS
        open 清掃スケジュール_変換結果.xlsx
    else
        # Linux
        xdg-open 清掃スケジュール_変換結果.xlsx
    fi
fi

