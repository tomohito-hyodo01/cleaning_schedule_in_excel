"""
表構造抽出のテスト
"""

import unittest
import numpy as np
import cv2

from src.config import TableDetectionConfig
from src.table_detect import TableDetector, Cell, Line


class TestLine(unittest.TestCase):
    """Lineクラスのテスト"""
    
    def test_horizontal_line(self):
        """水平線の判定"""
        line = Line(10, 50, 100, 50)
        
        self.assertTrue(line.is_horizontal)
        self.assertFalse(line.is_vertical)
        self.assertEqual(line.position, 50)
        self.assertEqual(line.start, 10)
        self.assertEqual(line.end, 100)
        self.assertEqual(line.length, 90)
    
    def test_vertical_line(self):
        """垂直線の判定"""
        line = Line(50, 10, 50, 100)
        
        self.assertFalse(line.is_horizontal)
        self.assertTrue(line.is_vertical)
        self.assertEqual(line.position, 50)
        self.assertEqual(line.start, 10)
        self.assertEqual(line.end, 100)
        self.assertEqual(line.length, 90)


class TestCell(unittest.TestCase):
    """Cellクラスのテスト"""
    
    def test_cell_basic(self):
        """基本的なセル"""
        cell = Cell(row=0, col=0, x=10, y=20, width=30, height=40)
        
        self.assertEqual(cell.row, 0)
        self.assertEqual(cell.col, 0)
        self.assertEqual(cell.x, 10)
        self.assertEqual(cell.y, 20)
        self.assertEqual(cell.width, 30)
        self.assertEqual(cell.height, 40)
        self.assertEqual(cell.rowspan, 1)
        self.assertEqual(cell.colspan, 1)
    
    def test_merged_cell(self):
        """結合セル"""
        cell = Cell(
            row=0, col=0, x=10, y=20,
            width=60, height=80,
            rowspan=2, colspan=2
        )
        
        self.assertEqual(cell.rowspan, 2)
        self.assertEqual(cell.colspan, 2)


class TestTableDetector(unittest.TestCase):
    """TableDetectorのテスト"""
    
    def setUp(self):
        """セットアップ"""
        self.config = TableDetectionConfig()
        self.detector = TableDetector(self.config)
    
    def test_create_simple_grid(self):
        """シンプルなグリッド作成"""
        # 水平線と垂直線を定義
        horizontal_lines = [
            Line(0, 50, 200, 50),
            Line(0, 100, 200, 100)
        ]
        vertical_lines = [
            Line(50, 0, 50, 200),
            Line(100, 0, 100, 200)
        ]
        
        # グリッドを作成
        grid_rows, grid_cols = self.detector._create_grid(
            horizontal_lines,
            vertical_lines,
            (200, 200)
        )
        
        # グリッドが正しく生成されることを確認
        self.assertGreater(len(grid_rows), 0)
        self.assertGreater(len(grid_cols), 0)
    
    def test_create_cells_from_grid(self):
        """グリッドからセルを生成"""
        grid_rows = [0, 50, 100]
        grid_cols = [0, 50, 100]
        
        cells = self.detector._create_cells(grid_rows, grid_cols)
        
        # 2x2のセルが生成されることを確認
        self.assertEqual(len(cells), 4)
        
        # 最初のセル
        cell = cells[0]
        self.assertEqual(cell.row, 0)
        self.assertEqual(cell.col, 0)
        self.assertEqual(cell.x, 0)
        self.assertEqual(cell.y, 0)
        self.assertEqual(cell.width, 50)
        self.assertEqual(cell.height, 50)


if __name__ == "__main__":
    unittest.main()

