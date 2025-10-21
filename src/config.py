"""
設定管理モジュール

アプリケーション全体の設定を管理する
"""

from dataclasses import dataclass, field
from typing import Literal, Optional
import json


@dataclass
class PreprocessConfig:
    """画像前処理の設定"""
    
    # 自動傾き補正
    enable_rotation_correction: bool = True
    rotation_angle_threshold: float = 0.5  # 度
    
    # 二値化
    binarization_method: Literal["otsu", "adaptive", "simple"] = "otsu"
    adaptive_block_size: int = 15
    adaptive_c: int = 10
    
    # ノイズ除去
    denoise_kernel_size: int = 3
    morphology_iterations: int = 1
    
    # コントラスト調整
    enable_contrast_adjustment: bool = True
    clahe_clip_limit: float = 2.0
    clahe_tile_size: int = 8


@dataclass
class TableDetectionConfig:
    """表構造抽出の設定"""
    
    # 検出方法
    method: Literal["opencv", "pp_structure"] = "opencv"
    
    # 線の最小長さ（ピクセル）
    min_line_length: int = 30
    
    # 線幅の閾値
    thin_line_threshold: int = 2
    medium_line_threshold: int = 5
    double_line_threshold: float = 2.0  # 倍率
    
    # 罫線検出
    hough_threshold: int = 50
    hough_min_line_length: int = 100
    hough_max_line_gap: int = 10
    
    # 形態学処理
    horizontal_kernel_scale: int = 40
    vertical_kernel_scale: int = 40
    dilation_iterations: int = 2
    
    # セル結合判定
    merge_tolerance: int = 5  # ピクセル


@dataclass
class OCRConfig:
    """OCRの設定"""
    
    # OCRエンジン
    engine: Literal["paddleocr", "tesseract"] = "paddleocr"
    
    # PaddleOCR設定
    paddleocr_lang: str = "japan"
    paddleocr_use_angle_cls: bool = True
    paddleocr_use_gpu: bool = False
    
    # Tesseract設定
    tesseract_lang: str = "jpn+jpn_vert"
    tesseract_config: str = "--psm 6"
    
    # 縦書き検出
    enable_vertical_text: bool = True
    
    # 信頼度閾値
    confidence_threshold: float = 0.85
    
    # 再読取り
    enable_retry: bool = True
    retry_count: int = 2
    
    # 正規化ルール
    enable_normalization: bool = True


@dataclass
class ExcelConfig:
    """Excel生成の設定"""
    
    # 寸法変換係数
    column_width_scale: float = 1.00
    row_height_scale: float = 1.00
    
    # DPI想定（ピクセル→Excel単位変換用）
    assumed_dpi: int = 96
    
    # ピクセル→Excel単位変換係数
    # 列幅: Excelの単位は「文字数」（1文字 ≈ 7px @ 11pt）
    pixel_to_column_width: float = 1.0 / 7.0
    # 行高: Excelの単位はポイント（1pt = 1/72 inch）
    pixel_to_row_height: float = 72.0 / 96.0  # 0.75
    
    # フォント
    default_font: str = "MS Gothic"
    default_font_size: int = 11
    
    # 配置
    default_horizontal_alignment: str = "center"
    default_vertical_alignment: str = "center"
    wrap_text: bool = True
    
    # ヘッダ設定
    header_row_count: int = 1
    header_fill_color: str = "D9D9D9"  # 薄いグレー
    header_bold: bool = True


@dataclass
class DebugConfig:
    """デバッグ/検品の設定"""
    
    # デバッグ画像出力
    save_debug_images: bool = True
    debug_output_dir: str = "debug_output"
    
    # ログレベル
    log_level: Literal["DEBUG", "INFO", "WARNING", "ERROR"] = "INFO"
    
    # 精度レポート
    generate_accuracy_report: bool = True


@dataclass
class AppConfig:
    """アプリケーション全体の設定"""
    
    preprocess: PreprocessConfig = field(default_factory=PreprocessConfig)
    table_detection: TableDetectionConfig = field(default_factory=TableDetectionConfig)
    ocr: OCRConfig = field(default_factory=OCRConfig)
    excel: ExcelConfig = field(default_factory=ExcelConfig)
    debug: DebugConfig = field(default_factory=DebugConfig)
    
    @classmethod
    def from_json(cls, json_path: str) -> "AppConfig":
        """JSONファイルから設定を読み込む"""
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        return cls(
            preprocess=PreprocessConfig(**data.get("preprocess", {})),
            table_detection=TableDetectionConfig(**data.get("table_detection", {})),
            ocr=OCRConfig(**data.get("ocr", {})),
            excel=ExcelConfig(**data.get("excel", {})),
            debug=DebugConfig(**data.get("debug", {}))
        )
    
    def to_json(self, json_path: str) -> None:
        """設定をJSONファイルに保存"""
        data = {
            "preprocess": self.preprocess.__dict__,
            "table_detection": self.table_detection.__dict__,
            "ocr": self.ocr.__dict__,
            "excel": self.excel.__dict__,
            "debug": self.debug.__dict__
        }
        
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    
    @classmethod
    def default(cls) -> "AppConfig":
        """デフォルト設定を取得"""
        return cls()

