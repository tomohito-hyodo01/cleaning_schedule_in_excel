"""
セットアップスクリプト

pip install -e . でインストール可能にする
"""

from setuptools import setup, find_packages
from pathlib import Path

# README.mdを読み込む
readme_file = Path(__file__).parent / "README.md"
long_description = readme_file.read_text(encoding="utf-8") if readme_file.exists() else ""

setup(
    name="excel-image-converter",
    version="1.0.0",
    description="画像からExcelファイルを生成するツール",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="AI Assistant",
    author_email="",
    url="https://github.com/yourusername/excel-image-converter",
    packages=find_packages(),
    install_requires=[
        "numpy>=1.24.0",
        "opencv-python>=4.8.0",
        "Pillow>=10.0.0",
        "openpyxl>=3.1.0",
        "paddleocr>=2.7.0",
        "paddlepaddle>=2.5.0",
        "pytesseract>=0.3.10",
        "PySide6>=6.6.0",
        "tqdm>=4.66.0",
    ],
    extras_require={
        "dev": [
            "pyinstaller>=6.0.0",
            "pytest>=7.0.0",
        ],
    },
    entry_points={
        "console_scripts": [
            "img2xlsx=src.cli:main",
        ],
        "gui_scripts": [
            "excel-converter=src.ui:main",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.8",
)

