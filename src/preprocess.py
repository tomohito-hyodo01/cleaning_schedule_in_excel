"""
画像前処理モジュール

画像の傾き補正、二値化、ノイズ除去などを行う
"""

import cv2
import numpy as np
from typing import Tuple, Optional
import logging

from .config import PreprocessConfig

logger = logging.getLogger(__name__)


class ImagePreprocessor:
    """画像前処理クラス"""
    
    def __init__(self, config: PreprocessConfig):
        """
        Args:
            config: 前処理設定
        """
        self.config = config
    
    def preprocess(self, image: np.ndarray) -> Tuple[np.ndarray, dict]:
        """
        画像を前処理する
        
        Args:
            image: 入力画像（BGR or グレースケール）
            
        Returns:
            処理済み画像（二値画像）と処理情報の辞書
        """
        info = {}
        
        # グレースケール変換
        if len(image.shape) == 3:
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        else:
            gray = image.copy()
        
        logger.info(f"入力画像サイズ: {gray.shape}")
        
        # コントラスト調整
        if self.config.enable_contrast_adjustment:
            gray = self._adjust_contrast(gray)
            info['contrast_adjusted'] = True
        
        # 傾き補正
        if self.config.enable_rotation_correction:
            gray, angle = self._correct_rotation(gray)
            info['rotation_angle'] = angle
            logger.info(f"傾き補正: {angle:.2f}度")
        
        # 二値化
        binary = self._binarize(gray)
        info['binarization_method'] = self.config.binarization_method
        
        # ノイズ除去
        binary = self._denoise(binary)
        info['denoised'] = True
        
        return binary, info
    
    def _adjust_contrast(self, image: np.ndarray) -> np.ndarray:
        """
        コントラストを調整する（CLAHE使用）
        
        Args:
            image: グレースケール画像
            
        Returns:
            調整後の画像
        """
        clahe = cv2.createCLAHE(
            clipLimit=self.config.clahe_clip_limit,
            tileGridSize=(self.config.clahe_tile_size, self.config.clahe_tile_size)
        )
        return clahe.apply(image)
    
    def _correct_rotation(self, image: np.ndarray) -> Tuple[np.ndarray, float]:
        """
        傾きを自動補正する（HoughLinesPで直線を検出し、角度の中央値で回転）
        
        Args:
            image: グレースケール画像
            
        Returns:
            回転後の画像と回転角度（度）
        """
        # エッジ検出
        edges = cv2.Canny(image, 50, 150, apertureSize=3)
        
        # Hough変換で直線検出
        lines = cv2.HoughLinesP(
            edges,
            rho=1,
            theta=np.pi / 180,
            threshold=100,
            minLineLength=100,
            maxLineGap=10
        )
        
        if lines is None or len(lines) == 0:
            logger.warning("直線が検出されませんでした。傾き補正をスキップします。")
            return image, 0.0
        
        # 各直線の角度を計算
        angles = []
        for line in lines:
            x1, y1, x2, y2 = line[0]
            
            # 角度を計算（ラジアン→度）
            angle = np.arctan2(y2 - y1, x2 - x1) * 180 / np.pi
            
            # -45度〜+45度の範囲に正規化
            if angle < -45:
                angle += 90
            elif angle > 45:
                angle -= 90
            
            # 長い線のみを考慮
            length = np.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
            if length > 50:
                angles.append(angle)
        
        if not angles:
            logger.warning("有効な直線が見つかりませんでした。")
            return image, 0.0
        
        # 角度の中央値を計算
        median_angle = np.median(angles)
        
        # 閾値以下なら補正しない
        if abs(median_angle) < self.config.rotation_angle_threshold:
            return image, 0.0
        
        # 画像を回転
        h, w = image.shape
        center = (w // 2, h // 2)
        rotation_matrix = cv2.getRotationMatrix2D(center, median_angle, 1.0)
        
        # 回転後の画像サイズを計算
        cos = np.abs(rotation_matrix[0, 0])
        sin = np.abs(rotation_matrix[0, 1])
        new_w = int(h * sin + w * cos)
        new_h = int(h * cos + w * sin)
        
        # 平行移動成分を調整
        rotation_matrix[0, 2] += (new_w / 2) - center[0]
        rotation_matrix[1, 2] += (new_h / 2) - center[1]
        
        rotated = cv2.warpAffine(
            image,
            rotation_matrix,
            (new_w, new_h),
            flags=cv2.INTER_CUBIC,
            borderMode=cv2.BORDER_REPLICATE
        )
        
        return rotated, median_angle
    
    def _binarize(self, image: np.ndarray) -> np.ndarray:
        """
        画像を二値化する
        
        Args:
            image: グレースケール画像
            
        Returns:
            二値画像
        """
        if self.config.binarization_method == "otsu":
            # Otsu's法
            _, binary = cv2.threshold(
                image,
                0,
                255,
                cv2.THRESH_BINARY + cv2.THRESH_OTSU
            )
        elif self.config.binarization_method == "adaptive":
            # 適応的二値化
            binary = cv2.adaptiveThreshold(
                image,
                255,
                cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                cv2.THRESH_BINARY,
                self.config.adaptive_block_size,
                self.config.adaptive_c
            )
        else:  # simple
            # 単純な閾値（127）
            _, binary = cv2.threshold(image, 127, 255, cv2.THRESH_BINARY)
        
        return binary
    
    def _denoise(self, image: np.ndarray) -> np.ndarray:
        """
        ノイズを除去する
        
        Args:
            image: 二値画像
            
        Returns:
            ノイズ除去後の画像
        """
        # モルフォロジー演算（オープニング）でノイズ除去
        kernel = cv2.getStructuringElement(
            cv2.MORPH_RECT,
            (self.config.denoise_kernel_size, self.config.denoise_kernel_size)
        )
        
        # ノイズ除去（小さい黒点を除去）
        denoised = cv2.morphologyEx(
            image,
            cv2.MORPH_OPEN,
            kernel,
            iterations=self.config.morphology_iterations
        )
        
        # 小さい穴を埋める（クロージング）
        denoised = cv2.morphologyEx(
            denoised,
            cv2.MORPH_CLOSE,
            kernel,
            iterations=self.config.morphology_iterations
        )
        
        return denoised


def create_preprocessor(config: Optional[PreprocessConfig] = None) -> ImagePreprocessor:
    """
    前処理器を作成する
    
    Args:
        config: 設定（Noneの場合はデフォルト）
        
    Returns:
        ImagePreprocessorインスタンス
    """
    if config is None:
        config = PreprocessConfig()
    
    return ImagePreprocessor(config)

