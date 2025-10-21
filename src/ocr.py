"""
OCRモジュール

PaddleOCRまたはTesseractを使ってテキスト認識を行う
"""

import cv2
import numpy as np
from typing import List, Tuple, Optional, Dict
import re
import logging

from .config import OCRConfig
from .table_detect import Cell

logger = logging.getLogger(__name__)


class OCREngine:
    """OCRエンジンの基底クラス"""
    
    def recognize(self, image: np.ndarray) -> Tuple[str, float]:
        """
        画像からテキストを認識
        
        Args:
            image: 入力画像
            
        Returns:
            認識されたテキストと信頼度
        """
        raise NotImplementedError


class PaddleOCREngine(OCREngine):
    """PaddleOCRエンジン"""
    
    def __init__(self, config: OCRConfig):
        """
        Args:
            config: OCR設定
        """
        self.config = config
        self._ocr = None
    
    def _initialize(self):
        """OCRエンジンを初期化（遅延初期化）"""
        if self._ocr is not None:
            return
        
        try:
            from paddleocr import PaddleOCR
            
            # PaddleOCRの初期化パラメータ（最新版対応、最小限のパラメータのみ）
            ocr_params = {
                'lang': self.config.paddleocr_lang,
                'use_angle_cls': self.config.paddleocr_use_angle_cls,
            }
            
            # 古いバージョンとの互換性のため、エラーが出るパラメータは自動的にスキップ
            try:
                self._ocr = PaddleOCR(**ocr_params)
            except (TypeError, ValueError) as e:
                error_msg = str(e)
                # エラーメッセージから問題のあるパラメータを特定して除外
                if 'use_angle_cls' in error_msg:
                    ocr_params.pop('use_angle_cls', None)
                    self._ocr = PaddleOCR(**ocr_params)
                else:
                    raise
            
            logger.info("PaddleOCRを初期化しました")
        except Exception as e:
            logger.error(f"PaddleOCRの初期化に失敗: {e}")
            raise
    
    def recognize(self, image: np.ndarray) -> Tuple[str, float]:
        """
        画像からテキストを認識
        
        Args:
            image: 入力画像（BGR or グレースケール）
            
        Returns:
            認識されたテキストと信頼度
        """
        self._initialize()
        
        try:
            # PaddleOCRで認識（最新版はpredict メソッドを使用）
            try:
                # 新しいAPI
                result = self._ocr.predict(image)
            except (AttributeError, TypeError):
                # 古いAPI
                try:
                    result = self._ocr.ocr(image, cls=self.config.paddleocr_use_angle_cls)
                except TypeError:
                    # clsパラメータがサポートされていない場合
                    result = self._ocr.ocr(image)
            
            if not result or not result[0]:
                return "", 0.0
            
            # 結果を結合
            texts = []
            confidences = []
            
            for line in result[0]:
                text = line[1][0]
                confidence = line[1][1]
                texts.append(text)
                confidences.append(confidence)
            
            # テキストを結合（改行で）
            full_text = "\n".join(texts) if texts else ""
            avg_confidence = sum(confidences) / len(confidences) if confidences else 0.0
            
            return full_text, avg_confidence
            
        except Exception as e:
            logger.warning(f"PaddleOCRでの認識に失敗: {e}")
            return "", 0.0


class TesseractOCREngine(OCREngine):
    """TesseractOCRエンジン"""
    
    def __init__(self, config: OCRConfig):
        """
        Args:
            config: OCR設定
        """
        self.config = config
    
    def recognize(self, image: np.ndarray) -> Tuple[str, float]:
        """
        画像からテキストを認識
        
        Args:
            image: 入力画像
            
        Returns:
            認識されたテキストと信頼度
        """
        try:
            import pytesseract
            
            # Tesseractで認識
            custom_config = f"-l {self.config.tesseract_lang} {self.config.tesseract_config}"
            text = pytesseract.image_to_string(image, config=custom_config)
            
            # 信頼度を取得
            data = pytesseract.image_to_data(image, config=custom_config, output_type=pytesseract.Output.DICT)
            confidences = [int(conf) for conf in data['conf'] if conf != '-1']
            avg_confidence = sum(confidences) / len(confidences) / 100.0 if confidences else 0.0
            
            return text.strip(), avg_confidence
            
        except Exception as e:
            logger.warning(f"Tesseractでの認識に失敗: {e}")
            return "", 0.0


class TextNormalizer:
    """テキスト正規化クラス"""
    
    @staticmethod
    def normalize(text: str) -> str:
        """
        テキストを正規化する
        
        Args:
            text: 入力テキスト
            
        Returns:
            正規化されたテキスト
        """
        if not text:
            return text
        
        # 全角英数字→半角
        text = TextNormalizer._zenkaku_to_hankaku(text)
        
        # よくある誤認識を修正
        text = TextNormalizer._fix_common_errors(text)
        
        # 前後の空白を削除
        text = text.strip()
        
        return text
    
    @staticmethod
    def _zenkaku_to_hankaku(text: str) -> str:
        """全角英数字を半角に変換"""
        # 全角数字→半角数字
        text = text.translate(str.maketrans(
            '０１２３４５６７８９',
            '0123456789'
        ))
        
        # 全角アルファベット→半角アルファベット
        text = text.translate(str.maketrans(
            'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ',
            'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
        ))
        
        return text
    
    @staticmethod
    def _fix_common_errors(text: str) -> str:
        """よくある誤認識を修正"""
        # I → 1（数字文脈）
        text = re.sub(r'(?<!\w)I(?=/)', '1', text)
        
        # O → 0（数字文脈）
        text = re.sub(r'(?<=\d)O(?=\d)', '0', text)
        text = re.sub(r'(?<!\w)O(?=\d)', '0', text)
        
        # l → 1（数字文脈）
        text = re.sub(r'(?<=\d)l(?=\d)', '1', text)
        
        # 1／Y → 1/Y（スラッシュ）
        text = text.replace('／', '/')
        
        return text


class CellOCRProcessor:
    """セル単位でOCRを実行するクラス"""
    
    def __init__(self, config: OCRConfig):
        """
        Args:
            config: OCR設定
        """
        self.config = config
        
        # OCRエンジンを作成
        if config.engine == "paddleocr":
            self.engine = PaddleOCREngine(config)
        else:
            self.engine = TesseractOCREngine(config)
        
        self.normalizer = TextNormalizer()
    
    def process_cells(self, image: np.ndarray, cells: List[Cell]) -> Dict[Tuple[int, int], Tuple[str, float]]:
        """
        各セルのテキストを認識する
        
        Args:
            image: 元画像（BGR）
            cells: セルのリスト
            
        Returns:
            {(row, col): (text, confidence)} の辞書
        """
        results = {}
        
        logger.info(f"OCR処理開始: {len(cells)}個のセル")
        
        for i, cell in enumerate(cells):
            if (i + 1) % 50 == 0:
                logger.info(f"OCR処理中: {i + 1}/{len(cells)}")
            
            # セル画像を切り出し
            cell_image = self._extract_cell_image(image, cell)
            
            # OCR実行
            text, confidence = self._recognize_cell(cell_image)
            
            # 低信頼度の場合は再試行
            if confidence < self.config.confidence_threshold and self.config.enable_retry:
                for retry in range(self.config.retry_count):
                    logger.debug(f"セル({cell.row}, {cell.col})を再試行 ({retry + 1}/{self.config.retry_count})")
                    
                    # 前処理を変えて再試行
                    enhanced_image = self._enhance_image(cell_image, retry)
                    text_retry, confidence_retry = self._recognize_cell(enhanced_image)
                    
                    if confidence_retry > confidence:
                        text, confidence = text_retry, confidence_retry
                    
                    if confidence >= self.config.confidence_threshold:
                        break
            
            results[(cell.row, cell.col)] = (text, confidence)
        
        logger.info(f"OCR処理完了")
        
        # 統計情報をログ出力
        confidences = [conf for _, conf in results.values()]
        if confidences:
            avg_conf = sum(confidences) / len(confidences)
            low_conf_count = sum(1 for conf in confidences if conf < self.config.confidence_threshold)
            logger.info(f"平均信頼度: {avg_conf:.2%}, 低信頼度セル: {low_conf_count}/{len(confidences)}")
        
        return results
    
    def _extract_cell_image(self, image: np.ndarray, cell: Cell) -> np.ndarray:
        """
        セル画像を切り出す
        
        Args:
            image: 元画像
            cell: セル
            
        Returns:
            切り出された画像
        """
        # マージンを少し削る（罫線を避ける）
        margin = 3
        x = max(0, cell.x + margin)
        y = max(0, cell.y + margin)
        x2 = min(image.shape[1], cell.x + cell.width - margin)
        y2 = min(image.shape[0], cell.y + cell.height - margin)
        
        if x2 <= x or y2 <= y:
            # 無効な領域
            return np.zeros((10, 10, 3), dtype=np.uint8)
        
        return image[y:y2, x:x2]
    
    def _recognize_cell(self, cell_image: np.ndarray) -> Tuple[str, float]:
        """
        セル画像を認識
        
        Args:
            cell_image: セル画像
            
        Returns:
            認識されたテキストと信頼度
        """
        # 画像が小さすぎる場合はスキップ
        if cell_image.shape[0] < 5 or cell_image.shape[1] < 5:
            return "", 0.0
        
        # OCR実行
        text, confidence = self.engine.recognize(cell_image)
        
        # 正規化
        if self.config.enable_normalization and text:
            text = self.normalizer.normalize(text)
        
        return text, confidence
    
    def _enhance_image(self, image: np.ndarray, retry_index: int) -> np.ndarray:
        """
        画像を強調する（再試行用）
        
        Args:
            image: 入力画像
            retry_index: 再試行回数
            
        Returns:
            強調された画像
        """
        # グレースケール化
        if len(image.shape) == 3:
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        else:
            gray = image.copy()
        
        if retry_index == 0:
            # 1回目: コントラスト強調
            clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
            enhanced = clahe.apply(gray)
        else:
            # 2回目以降: 二値化
            _, enhanced = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        return enhanced


def create_ocr_processor(config: Optional[OCRConfig] = None) -> CellOCRProcessor:
    """
    OCRプロセッサを作成する
    
    Args:
        config: 設定（Noneの場合はデフォルト）
        
    Returns:
        CellOCRProcessorインスタンス
    """
    if config is None:
        config = OCRConfig()
    
    return CellOCRProcessor(config)

