"""
画像前処理のテスト
"""

import unittest
import numpy as np
import cv2

from src.config import PreprocessConfig
from src.preprocess import ImagePreprocessor


class TestImagePreprocessor(unittest.TestCase):
    """ImagePreprocessorのテスト"""
    
    def setUp(self):
        """セットアップ"""
        self.config = PreprocessConfig()
        self.preprocessor = ImagePreprocessor(self.config)
    
    def test_preprocess_grayscale_image(self):
        """グレースケール画像の前処理"""
        # 100x100のグレースケール画像を作成
        image = np.random.randint(0, 256, (100, 100), dtype=np.uint8)
        
        binary, info = self.preprocessor.preprocess(image)
        
        # 二値画像が返されることを確認
        self.assertEqual(binary.shape, image.shape)
        self.assertTrue(np.all((binary == 0) | (binary == 255)))
    
    def test_preprocess_color_image(self):
        """カラー画像の前処理"""
        # 100x100x3のカラー画像を作成
        image = np.random.randint(0, 256, (100, 100, 3), dtype=np.uint8)
        
        binary, info = self.preprocessor.preprocess(image)
        
        # 二値画像が返されることを確認
        self.assertEqual(len(binary.shape), 2)
        self.assertTrue(np.all((binary == 0) | (binary == 255)))
    
    def test_binarization_otsu(self):
        """Otsu法による二値化"""
        # 白と黒のグラデーション画像
        image = np.linspace(0, 255, 100*100).reshape(100, 100).astype(np.uint8)
        
        binary = self.preprocessor._binarize(image)
        
        # 二値画像であることを確認
        unique_values = np.unique(binary)
        self.assertTrue(len(unique_values) <= 2)
    
    def test_rotation_correction(self):
        """傾き補正"""
        # 白い背景に黒い長方形を描画
        image = np.ones((200, 200), dtype=np.uint8) * 255
        cv2.rectangle(image, (50, 50), (150, 150), 0, 2)
        
        # 傾き補正を実行（傾きがない場合は角度0が返される）
        rotated, angle = self.preprocessor._correct_rotation(image)
        
        # 角度が数値であることを確認
        self.assertIsInstance(angle, float)


class TestPreprocessConfig(unittest.TestCase):
    """PreprocessConfigのテスト"""
    
    def test_default_config(self):
        """デフォルト設定"""
        config = PreprocessConfig()
        
        self.assertTrue(config.enable_rotation_correction)
        self.assertEqual(config.binarization_method, "otsu")
        self.assertTrue(config.enable_contrast_adjustment)


if __name__ == "__main__":
    unittest.main()

