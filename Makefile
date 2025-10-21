# Makefile for Excel Image Converter

.PHONY: help install test sample clean build

help:
	@echo "利用可能なコマンド:"
	@echo "  make install  - 依存関係をインストール"
	@echo "  make test     - テストを実行"
	@echo "  make sample   - サンプル画像を変換"
	@echo "  make clean    - 生成ファイルを削除"
	@echo "  make build    - 実行ファイルをビルド"

install:
	pip install -r requirements.txt

test:
	python -m unittest discover tests -v

sample:
	python main.py --cli schedule_image/清掃スケジュール.jpg -o 清掃スケジュール_変換結果.xlsx --log-level INFO

clean:
	rm -rf build dist *.spec
	rm -rf debug_output
	rm -rf __pycache__ src/__pycache__ tests/__pycache__
	rm -rf .pytest_cache
	rm -rf *.egg-info
	find . -name "*.pyc" -delete
	find . -name "*.pyo" -delete

build:
	pip install pyinstaller
	pyinstaller --onefile --windowed --name="Excel画像変換" main.py

