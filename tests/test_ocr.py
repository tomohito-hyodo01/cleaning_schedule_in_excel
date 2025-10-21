"""
OCRのテスト
"""

import unittest

from src.ocr import TextNormalizer


class TestTextNormalizer(unittest.TestCase):
    """TextNormalizerのテスト"""
    
    def test_zenkaku_to_hankaku_numbers(self):
        """全角数字→半角数字"""
        text = "０１２３４５６７８９"
        normalized = TextNormalizer.normalize(text)
        
        self.assertEqual(normalized, "0123456789")
    
    def test_zenkaku_to_hankaku_alphabet(self):
        """全角アルファベット→半角アルファベット"""
        text = "ＡＢＣａｂｃ"
        normalized = TextNormalizer.normalize(text)
        
        self.assertEqual(normalized, "ABCabc")
    
    def test_fix_common_errors(self):
        """よくある誤認識の修正"""
        # I → 1（数字文脈）
        text = "I/Y"
        normalized = TextNormalizer.normalize(text)
        self.assertEqual(normalized, "1/Y")
        
        # 全角スラッシュ→半角スラッシュ
        text = "1／2"
        normalized = TextNormalizer.normalize(text)
        self.assertEqual(normalized, "1/2")
    
    def test_strip_whitespace(self):
        """前後の空白を削除"""
        text = "  test  "
        normalized = TextNormalizer.normalize(text)
        
        self.assertEqual(normalized, "test")
    
    def test_empty_string(self):
        """空文字列"""
        text = ""
        normalized = TextNormalizer.normalize(text)
        
        self.assertEqual(normalized, "")


if __name__ == "__main__":
    unittest.main()

