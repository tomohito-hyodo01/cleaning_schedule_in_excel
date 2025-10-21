"""
GUIインターフェース（PySide6）

ドラッグ＆ドロップ対応のGUIアプリケーション
"""

import sys
import os
from pathlib import Path
from typing import Optional
import logging

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QProgressBar, QTextEdit,
    QGroupBox, QFormLayout, QDoubleSpinBox, QSpinBox, QComboBox,
    QCheckBox, QMessageBox, QSplitter
)
from PySide6.QtCore import Qt, QThread, Signal, QUrl
from PySide6.QtGui import QPixmap, QDragEnterEvent, QDropEvent

from .config import AppConfig
from .cli import convert_image_to_excel


# ログハンドラ
class QTextEditLogger(logging.Handler):
    """QTextEditにログを出力するハンドラ"""
    
    def __init__(self, text_edit: QTextEdit):
        super().__init__()
        self.text_edit = text_edit
    
    def emit(self, record):
        msg = self.format(record)
        self.text_edit.append(msg)


# 変換ワーカースレッド
class ConversionWorker(QThread):
    """バックグラウンドで変換を実行するワーカー"""
    
    progress = Signal(int)  # 進捗（0-100）
    finished = Signal(dict)  # 完了時に情報を渡す
    error = Signal(str)  # エラー時にメッセージを渡す
    
    def __init__(self, input_path: str, output_path: str, config: AppConfig):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path
        self.config = config
    
    def run(self):
        """変換を実行"""
        try:
            # 変換を実行
            info = convert_image_to_excel(
                self.input_path,
                self.output_path,
                self.config
            )
            
            self.progress.emit(100)
            self.finished.emit(info)
            
        except Exception as e:
            self.error.emit(str(e))


# メインウィンドウ
class MainWindow(QMainWindow):
    """メインウィンドウ"""
    
    def __init__(self):
        super().__init__()
        
        self.input_image_path: Optional[str] = None
        self.output_excel_path: Optional[str] = None
        self.worker: Optional[ConversionWorker] = None
        
        self.init_ui()
        self.setup_logging()
    
    def init_ui(self):
        """UIを初期化"""
        self.setWindowTitle("Excel画像変換ツール")
        self.setGeometry(100, 100, 1200, 800)
        
        # 中央ウィジェット
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # メインレイアウト
        main_layout = QVBoxLayout(central_widget)
        
        # スプリッター（上下に分割）
        splitter = QSplitter(Qt.Vertical)
        main_layout.addWidget(splitter)
        
        # 上部：画像プレビューと設定
        top_widget = QWidget()
        top_layout = QHBoxLayout(top_widget)
        
        # 左：画像プレビュー
        preview_group = QGroupBox("画像プレビュー")
        preview_layout = QVBoxLayout()
        
        self.image_label = QLabel("画像をドロップしてください")
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setMinimumSize(400, 400)
        self.image_label.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                background-color: #f5f5f5;
            }
        """)
        self.image_label.setAcceptDrops(True)
        self.image_label.dragEnterEvent = self.drag_enter_event
        self.image_label.dropEvent = self.drop_event
        
        preview_layout.addWidget(self.image_label)
        
        # ファイル選択ボタン
        btn_layout = QHBoxLayout()
        self.btn_select_image = QPushButton("画像を選択")
        self.btn_select_image.clicked.connect(self.select_image)
        btn_layout.addWidget(self.btn_select_image)
        
        self.btn_select_output = QPushButton("出力先を選択")
        self.btn_select_output.clicked.connect(self.select_output)
        btn_layout.addWidget(self.btn_select_output)
        
        preview_layout.addLayout(btn_layout)
        
        self.label_output = QLabel("出力: 未設定")
        preview_layout.addWidget(self.label_output)
        
        preview_group.setLayout(preview_layout)
        top_layout.addWidget(preview_group, 2)
        
        # 右：設定パネル
        settings_group = QGroupBox("設定")
        settings_layout = QFormLayout()
        
        # OCRエンジン
        self.combo_ocr_engine = QComboBox()
        self.combo_ocr_engine.addItems(["paddleocr", "tesseract"])
        settings_layout.addRow("OCRエンジン:", self.combo_ocr_engine)
        
        # 列幅倍率
        self.spin_col_scale = QDoubleSpinBox()
        self.spin_col_scale.setRange(0.1, 5.0)
        self.spin_col_scale.setSingleStep(0.1)
        self.spin_col_scale.setValue(1.0)
        settings_layout.addRow("列幅倍率:", self.spin_col_scale)
        
        # 行高倍率
        self.spin_row_scale = QDoubleSpinBox()
        self.spin_row_scale.setRange(0.1, 5.0)
        self.spin_row_scale.setSingleStep(0.1)
        self.spin_row_scale.setValue(1.0)
        settings_layout.addRow("行高倍率:", self.spin_row_scale)
        
        # 最小線長
        self.spin_min_line = QSpinBox()
        self.spin_min_line.setRange(10, 200)
        self.spin_min_line.setValue(30)
        settings_layout.addRow("最小線長（px）:", self.spin_min_line)
        
        # 縦書き検出
        self.check_vertical = QCheckBox()
        self.check_vertical.setChecked(True)
        settings_layout.addRow("縦書き検出:", self.check_vertical)
        
        # デバッグ画像保存
        self.check_debug_images = QCheckBox()
        self.check_debug_images.setChecked(True)
        settings_layout.addRow("デバッグ画像保存:", self.check_debug_images)
        
        settings_group.setLayout(settings_layout)
        top_layout.addWidget(settings_group, 1)
        
        splitter.addWidget(top_widget)
        
        # 下部：変換ボタン、進捗、ログ
        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout(bottom_widget)
        
        # 変換ボタン
        self.btn_convert = QPushButton("解析してExcel作成")
        self.btn_convert.setEnabled(False)
        self.btn_convert.setMinimumHeight(50)
        self.btn_convert.clicked.connect(self.start_conversion)
        bottom_layout.addWidget(self.btn_convert)
        
        # 進捗バー
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        bottom_layout.addWidget(self.progress_bar)
        
        # ログビュー
        log_group = QGroupBox("ログ")
        log_layout = QVBoxLayout()
        
        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMaximumHeight(200)
        log_layout.addWidget(self.log_view)
        
        log_group.setLayout(log_layout)
        bottom_layout.addWidget(log_group)
        
        splitter.addWidget(bottom_widget)
        
        # スプリッターのサイズ比率
        splitter.setSizes([500, 300])
    
    def setup_logging(self):
        """ログを設定"""
        # ルートロガーを取得
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        
        # QTextEditハンドラを追加
        handler = QTextEditLogger(self.log_view)
        handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        root_logger.addHandler(handler)
    
    def drag_enter_event(self, event: QDragEnterEvent):
        """ドラッグエンター時の処理"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def drop_event(self, event: QDropEvent):
        """ドロップ時の処理"""
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.lower().endswith(('.png', '.jpg', '.jpeg')):
                self.load_image(file_path)
            else:
                QMessageBox.warning(self, "エラー", "PNG/JPG画像のみ対応しています")
    
    def select_image(self):
        """画像を選択"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "画像を選択",
            "",
            "画像ファイル (*.png *.jpg *.jpeg)"
        )
        
        if file_path:
            self.load_image(file_path)
    
    def load_image(self, file_path: str):
        """画像を読み込む"""
        self.input_image_path = file_path
        
        # プレビューを表示
        pixmap = QPixmap(file_path)
        scaled_pixmap = pixmap.scaled(
            self.image_label.size(),
            Qt.KeepAspectRatio,
            Qt.SmoothTransformation
        )
        self.image_label.setPixmap(scaled_pixmap)
        
        # デフォルトの出力パスを設定
        input_path = Path(file_path)
        default_output = input_path.parent / f"{input_path.stem}_変換結果.xlsx"
        self.output_excel_path = str(default_output)
        self.label_output.setText(f"出力: {self.output_excel_path}")
        
        # 変換ボタンを有効化
        self.btn_convert.setEnabled(True)
        
        self.log_view.append(f"画像を読み込みました: {file_path}")
    
    def select_output(self):
        """出力先を選択"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "出力先を選択",
            self.output_excel_path or "output.xlsx",
            "Excelファイル (*.xlsx)"
        )
        
        if file_path:
            self.output_excel_path = file_path
            self.label_output.setText(f"出力: {self.output_excel_path}")
    
    def get_config(self) -> AppConfig:
        """現在の設定を取得"""
        config = AppConfig.default()
        
        # OCRエンジン
        config.ocr.engine = self.combo_ocr_engine.currentText()
        
        # Excel設定
        config.excel.column_width_scale = self.spin_col_scale.value()
        config.excel.row_height_scale = self.spin_row_scale.value()
        
        # 表検出設定
        config.table_detection.min_line_length = self.spin_min_line.value()
        
        # OCR設定
        config.ocr.enable_vertical_text = self.check_vertical.isChecked()
        
        # デバッグ設定
        config.debug.save_debug_images = self.check_debug_images.isChecked()
        
        return config
    
    def start_conversion(self):
        """変換を開始"""
        if not self.input_image_path or not self.output_excel_path:
            QMessageBox.warning(self, "エラー", "入力画像と出力先を設定してください")
            return
        
        # UIを無効化
        self.btn_convert.setEnabled(False)
        self.btn_select_image.setEnabled(False)
        self.btn_select_output.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        self.log_view.clear()
        self.log_view.append("変換を開始します...")
        
        # ワーカースレッドを作成
        config = self.get_config()
        self.worker = ConversionWorker(
            self.input_image_path,
            self.output_excel_path,
            config
        )
        
        # シグナルを接続
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        
        # 変換を開始
        self.worker.start()
    
    def on_progress(self, value: int):
        """進捗更新"""
        self.progress_bar.setValue(value)
    
    def on_finished(self, info: dict):
        """変換完了"""
        self.log_view.append("\n変換が完了しました！")
        
        # 結果を表示
        if 'excel' in info:
            excel = info['excel']
            self.log_view.append(f"出力: {info['output']}")
            self.log_view.append(f"サイズ: {excel.get('columns', 0)}列 × {excel.get('rows', 0)}行")
            self.log_view.append(f"結合セル: {excel.get('merged_cells', 0)}個")
        
        # UIを有効化
        self.btn_convert.setEnabled(True)
        self.btn_select_image.setEnabled(True)
        self.btn_select_output.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        # 完了メッセージ
        reply = QMessageBox.question(
            self,
            "完了",
            f"Excelファイルを生成しました。\n\n{self.output_excel_path}\n\nファイルを開きますか？",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.open_excel_file()
    
    def on_error(self, error_msg: str):
        """エラー発生"""
        self.log_view.append(f"\nエラー: {error_msg}")
        
        # UIを有効化
        self.btn_convert.setEnabled(True)
        self.btn_select_image.setEnabled(True)
        self.btn_select_output.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        QMessageBox.critical(self, "エラー", f"変換に失敗しました:\n\n{error_msg}")
    
    def open_excel_file(self):
        """Excelファイルを開く"""
        if self.output_excel_path and Path(self.output_excel_path).exists():
            import subprocess
            import platform
            
            system = platform.system()
            try:
                if system == "Windows":
                    os.startfile(self.output_excel_path)
                elif system == "Darwin":  # macOS
                    subprocess.run(["open", self.output_excel_path])
                else:  # Linux
                    subprocess.run(["xdg-open", self.output_excel_path])
            except Exception as e:
                QMessageBox.warning(self, "警告", f"ファイルを開けませんでした: {e}")


def main():
    """GUIアプリケーションを起動"""
    app = QApplication(sys.argv)
    
    # アプリケーション情報
    app.setApplicationName("Excel画像変換ツール")
    app.setOrganizationName("ImageToExcel")
    
    # メインウィンドウを表示
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

