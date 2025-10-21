"""
CLIインターフェース

コマンドラインからの実行を提供する
"""

import argparse
import sys
import logging
from pathlib import Path
from typing import Optional
import cv2
import numpy as np

from .config import AppConfig
from .preprocess import create_preprocessor
from .table_detect import create_detector
from .ocr import create_ocr_processor
from .excel_writer import create_writer


def setup_logging(level: str = "INFO") -> None:
    """
    ログを設定する
    
    Args:
        level: ログレベル
    """
    numeric_level = getattr(logging, level.upper(), None)
    if not isinstance(numeric_level, int):
        numeric_level = logging.INFO
    
    logging.basicConfig(
        level=numeric_level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )


def convert_image_to_excel(
    input_path: str,
    output_path: str,
    config: Optional[AppConfig] = None
) -> dict:
    """
    画像をExcelに変換する
    
    Args:
        input_path: 入力画像パス
        output_path: 出力Excelパス
        config: 設定
        
    Returns:
        変換情報の辞書
    """
    logger = logging.getLogger(__name__)
    
    if config is None:
        config = AppConfig.default()
    
    info = {
        'input': input_path,
        'output': output_path
    }
    
    # 画像を読み込む（日本語パス対応）
    logger.info(f"画像を読み込み中: {input_path}")
    
    # OpenCVは日本語パスを扱えないため、バイトストリームから読み込む
    import numpy as np
    from pathlib import Path
    
    try:
        # ファイルをバイナリで読み込み
        with open(input_path, 'rb') as f:
            file_bytes = np.frombuffer(f.read(), dtype=np.uint8)
        
        # デコード
        image = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)
        
        if image is None:
            raise ValueError(f"画像をデコードできませんでした")
            
    except Exception as e:
        raise ValueError(f"画像を読み込めませんでした: {input_path} - {e}")
    
    info['input_size'] = (image.shape[1], image.shape[0])
    logger.info(f"画像サイズ: {info['input_size'][0]}x{info['input_size'][1]}")
    
    # 前処理
    logger.info("画像を前処理中...")
    preprocessor = create_preprocessor(config.preprocess)
    binary_image, preprocess_info = preprocessor.preprocess(image)
    info['preprocess'] = preprocess_info
    
    # デバッグ画像を保存
    if config.debug.save_debug_images:
        debug_dir = Path(config.debug.debug_output_dir)
        debug_dir.mkdir(exist_ok=True)
        
        debug_path = debug_dir / f"{Path(input_path).stem}_binary.png"
        cv2.imwrite(str(debug_path), binary_image)
        logger.info(f"デバッグ画像保存: {debug_path}")
    
    # 表構造を検出
    logger.info("表構造を検出中...")
    detector = create_detector(config.table_detection)
    cells, border_styles, detect_info = detector.detect(binary_image, image)
    info['detection'] = detect_info
    
    # デバッグ画像を保存（セルの境界を描画）
    if config.debug.save_debug_images:
        debug_image = image.copy()
        for cell in cells:
            cv2.rectangle(
                debug_image,
                (cell.x, cell.y),
                (cell.x + cell.width, cell.y + cell.height),
                (0, 255, 0),
                2
            )
        
        debug_path = debug_dir / f"{Path(input_path).stem}_cells.png"
        cv2.imwrite(str(debug_path), debug_image)
        logger.info(f"デバッグ画像保存: {debug_path}")
    
    # OCR処理
    logger.info("OCR処理中...")
    ocr_processor = create_ocr_processor(config.ocr)
    cell_texts = ocr_processor.process_cells(image, cells)
    
    # OCR結果の統計
    confidences = [conf for _, conf in cell_texts.values()]
    if confidences:
        info['ocr_avg_confidence'] = sum(confidences) / len(confidences)
        info['ocr_low_confidence_count'] = sum(1 for conf in confidences if conf < config.ocr.confidence_threshold)
    
    # Excelを生成
    logger.info("Excelファイルを生成中...")
    writer = create_writer(config.excel)
    excel_info = writer.write(cells, cell_texts, border_styles, output_path)
    info['excel'] = excel_info
    
    logger.info("変換完了！")
    
    # 精度レポートを表示
    if config.debug.generate_accuracy_report:
        print_accuracy_report(info)
    
    return info


def print_accuracy_report(info: dict) -> None:
    """
    精度レポートを表示する
    
    Args:
        info: 変換情報
    """
    print("\n" + "=" * 60)
    print("精度レポート")
    print("=" * 60)
    
    # 入力情報
    print(f"入力画像: {info['input']}")
    print(f"画像サイズ: {info['input_size'][0]}x{info['input_size'][1]}")
    
    # 前処理情報
    if 'preprocess' in info:
        preprocess = info['preprocess']
        if 'rotation_angle' in preprocess:
            print(f"傾き補正: {preprocess['rotation_angle']:.2f}度")
    
    # 検出情報
    if 'detection' in info:
        detection = info['detection']
        print(f"\n表構造:")
        print(f"  - グリッド: {detection.get('grid_rows', 0)}行 × {detection.get('grid_cols', 0)}列")
        print(f"  - セル総数: {detection.get('total_cells', 0)}")
        print(f"  - 結合セル: {detection.get('merged_cells', 0)}")
    
    # OCR情報
    if 'ocr_avg_confidence' in info:
        print(f"\nOCR:")
        print(f"  - 平均信頼度: {info['ocr_avg_confidence']:.2%}")
        print(f"  - 低信頼度セル: {info.get('ocr_low_confidence_count', 0)}")
    
    # Excel情報
    if 'excel' in info:
        excel = info['excel']
        print(f"\nExcel:")
        print(f"  - 出力: {info['output']}")
        print(f"  - サイズ: {excel.get('columns', 0)}列 × {excel.get('rows', 0)}行")
        print(f"  - 結合セル: {excel.get('merged_cells', 0)}")
    
    print("=" * 60 + "\n")


def main() -> int:
    """
    メイン関数
    
    Returns:
        終了コード
    """
    parser = argparse.ArgumentParser(
        description="画像からExcelファイルを生成する",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  %(prog)s input.png -o output.xlsx
  %(prog)s input.png -o output.xlsx --scale-col 1.2 --scale-row 0.9
  %(prog)s input.png -o output.xlsx --engine tesseract --log-level DEBUG
        """
    )
    
    # 必須引数
    parser.add_argument(
        "input",
        help="入力画像ファイル（PNG/JPG）"
    )
    
    # オプション引数
    parser.add_argument(
        "-o", "--output",
        default="output.xlsx",
        help="出力Excelファイル（デフォルト: output.xlsx）"
    )
    
    parser.add_argument(
        "--config",
        help="設定ファイル（JSON）"
    )
    
    # Excel設定
    parser.add_argument(
        "--scale-col",
        type=float,
        help="列幅の倍率（デフォルト: 1.0）"
    )
    
    parser.add_argument(
        "--scale-row",
        type=float,
        help="行高の倍率（デフォルト: 1.0）"
    )
    
    # 表検出設定
    parser.add_argument(
        "--min-line",
        type=int,
        help="罫線の最小長さ（ピクセル、デフォルト: 30）"
    )
    
    parser.add_argument(
        "--double-line-th",
        type=float,
        help="二重線の閾値（デフォルト: 2.0）"
    )
    
    # OCR設定
    parser.add_argument(
        "--engine",
        choices=["paddleocr", "tesseract"],
        help="OCRエンジン（デフォルト: paddleocr）"
    )
    
    parser.add_argument(
        "--no-vertical",
        action="store_true",
        help="縦書き検出を無効化"
    )
    
    # デバッグ設定
    parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default="INFO",
        help="ログレベル（デフォルト: INFO）"
    )
    
    parser.add_argument(
        "--no-debug-images",
        action="store_true",
        help="デバッグ画像の保存を無効化"
    )
    
    parser.add_argument(
        "--debug-dir",
        default="debug_output",
        help="デバッグ画像の出力ディレクトリ（デフォルト: debug_output）"
    )
    
    args = parser.parse_args()
    
    # ログを設定
    setup_logging(args.log_level)
    logger = logging.getLogger(__name__)
    
    # 設定を読み込む
    if args.config:
        try:
            config = AppConfig.from_json(args.config)
            logger.info(f"設定ファイルを読み込みました: {args.config}")
        except Exception as e:
            logger.error(f"設定ファイルの読み込みに失敗しました: {e}")
            return 1
    else:
        config = AppConfig.default()
    
    # コマンドライン引数で設定を上書き
    if args.scale_col is not None:
        config.excel.column_width_scale = args.scale_col
    
    if args.scale_row is not None:
        config.excel.row_height_scale = args.scale_row
    
    if args.min_line is not None:
        config.table_detection.min_line_length = args.min_line
    
    if args.double_line_th is not None:
        config.table_detection.double_line_threshold = args.double_line_th
    
    if args.engine is not None:
        config.ocr.engine = args.engine
    
    if args.no_vertical:
        config.ocr.enable_vertical_text = False
    
    if args.no_debug_images:
        config.debug.save_debug_images = False
    
    config.debug.debug_output_dir = args.debug_dir
    config.debug.log_level = args.log_level
    
    # 入力ファイルの存在確認
    if not Path(args.input).exists():
        logger.error(f"入力ファイルが見つかりません: {args.input}")
        return 1
    
    # 変換を実行
    try:
        convert_image_to_excel(args.input, args.output, config)
        return 0
    except Exception as e:
        logger.exception(f"変換に失敗しました: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())

