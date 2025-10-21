"""
表構造抽出モジュール

罫線検出、セル分割、結合セル判定などを行う
"""

import cv2
import numpy as np
from typing import List, Tuple, Dict, Optional
from dataclasses import dataclass
import logging

from .config import TableDetectionConfig

logger = logging.getLogger(__name__)


@dataclass
class Line:
    """直線を表すクラス"""
    x1: int
    y1: int
    x2: int
    y2: int
    thickness: int = 1
    
    @property
    def is_horizontal(self) -> bool:
        """水平線かどうか"""
        return abs(self.y2 - self.y1) < abs(self.x2 - self.x1)
    
    @property
    def is_vertical(self) -> bool:
        """垂直線かどうか"""
        return not self.is_horizontal
    
    @property
    def position(self) -> int:
        """線の位置（水平線ならy座標、垂直線ならx座標）"""
        if self.is_horizontal:
            return (self.y1 + self.y2) // 2
        else:
            return (self.x1 + self.x2) // 2
    
    @property
    def start(self) -> int:
        """線の開始位置"""
        if self.is_horizontal:
            return min(self.x1, self.x2)
        else:
            return min(self.y1, self.y2)
    
    @property
    def end(self) -> int:
        """線の終了位置"""
        if self.is_horizontal:
            return max(self.x1, self.x2)
        else:
            return max(self.y1, self.y2)
    
    @property
    def length(self) -> int:
        """線の長さ"""
        return self.end - self.start


@dataclass
class Cell:
    """セルを表すクラス"""
    row: int
    col: int
    x: int
    y: int
    width: int
    height: int
    rowspan: int = 1
    colspan: int = 1
    
    def __repr__(self) -> str:
        return f"Cell(row={self.row}, col={self.col}, x={self.x}, y={self.y}, w={self.width}, h={self.height}, rs={self.rowspan}, cs={self.colspan})"


@dataclass
class BorderStyle:
    """罫線スタイル"""
    top: str = "thin"  # "thin", "medium", "thick", "double", None
    bottom: str = "thin"
    left: str = "thin"
    right: str = "thin"


class TableDetector:
    """表構造検出クラス"""
    
    def __init__(self, config: TableDetectionConfig):
        """
        Args:
            config: 表検出設定
        """
        self.config = config
    
    def detect(self, binary_image: np.ndarray, original_image: Optional[np.ndarray] = None) -> Tuple[List[Cell], Dict[Tuple[int, int], BorderStyle], dict]:
        """
        表構造を検出する
        
        Args:
            binary_image: 二値画像
            original_image: 元画像（デバッグ用）
            
        Returns:
            セルのリスト、罫線スタイルの辞書、検出情報
        """
        info = {}
        
        if self.config.method == "opencv":
            return self._detect_opencv(binary_image, original_image, info)
        else:
            raise NotImplementedError("PP-Structureは未実装です")
    
    def _detect_opencv(self, binary_image: np.ndarray, original_image: Optional[np.ndarray], info: dict) -> Tuple[List[Cell], Dict[Tuple[int, int], BorderStyle], dict]:
        """
        OpenCVを使って表構造を検出
        
        Args:
            binary_image: 二値画像
            original_image: 元画像
            info: 検出情報を格納する辞書
            
        Returns:
            セルのリスト、罫線スタイルの辞書、検出情報
        """
        # 罫線を検出
        horizontal_lines, vertical_lines = self._detect_lines(binary_image)
        info['horizontal_lines'] = len(horizontal_lines)
        info['vertical_lines'] = len(vertical_lines)
        
        logger.info(f"罫線検出: 横{len(horizontal_lines)}本、縦{len(vertical_lines)}本")
        
        # 交点からグリッドを生成
        grid_rows, grid_cols = self._create_grid(horizontal_lines, vertical_lines, binary_image.shape)
        info['grid_rows'] = len(grid_rows)
        info['grid_cols'] = len(grid_cols)
        
        logger.info(f"グリッド生成: {len(grid_rows)}行 × {len(grid_cols)}列")
        
        # セルを生成
        cells = self._create_cells(grid_rows, grid_cols)
        
        # セル結合を判定
        cells, merged_map = self._detect_merged_cells(cells, binary_image)
        info['total_cells'] = len(cells)
        info['merged_cells'] = sum(1 for c in cells if c.rowspan > 1 or c.colspan > 1)
        
        logger.info(f"セル生成: 合計{len(cells)}個（結合{info['merged_cells']}個）")
        
        # 罫線スタイルを判定
        border_styles = self._detect_border_styles(cells, horizontal_lines, vertical_lines)
        
        return cells, border_styles, info
    
    def _detect_lines(self, binary_image: np.ndarray) -> Tuple[List[Line], List[Line]]:
        """
        罫線を検出する
        
        Args:
            binary_image: 二値画像
            
        Returns:
            水平線のリスト、垂直線のリスト
        """
        # 水平線を検出
        horizontal_kernel = cv2.getStructuringElement(
            cv2.MORPH_RECT,
            (self.config.horizontal_kernel_scale, 1)
        )
        horizontal_mask = cv2.morphologyEx(binary_image, cv2.MORPH_OPEN, horizontal_kernel)
        horizontal_mask = cv2.dilate(horizontal_mask, horizontal_kernel, iterations=self.config.dilation_iterations)
        
        # 垂直線を検出
        vertical_kernel = cv2.getStructuringElement(
            cv2.MORPH_RECT,
            (1, self.config.vertical_kernel_scale)
        )
        vertical_mask = cv2.morphologyEx(binary_image, cv2.MORPH_OPEN, vertical_kernel)
        vertical_mask = cv2.dilate(vertical_mask, vertical_kernel, iterations=self.config.dilation_iterations)
        
        # 輪郭から線を抽出
        horizontal_lines = self._extract_lines_from_mask(horizontal_mask, is_horizontal=True)
        vertical_lines = self._extract_lines_from_mask(vertical_mask, is_horizontal=False)
        
        return horizontal_lines, vertical_lines
    
    def _extract_lines_from_mask(self, mask: np.ndarray, is_horizontal: bool) -> List[Line]:
        """
        マスクから線を抽出する
        
        Args:
            mask: 罫線マスク
            is_horizontal: 水平線かどうか
            
        Returns:
            線のリスト
        """
        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        lines = []
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            
            # 最小長さフィルタ
            length = w if is_horizontal else h
            if length < self.config.min_line_length:
                continue
            
            # 線の太さ
            thickness = h if is_horizontal else w
            
            # 線を作成
            if is_horizontal:
                line = Line(x, y + h // 2, x + w, y + h // 2, thickness)
            else:
                line = Line(x + w // 2, y, x + w // 2, y + h, thickness)
            
            lines.append(line)
        
        # 位置でソート
        lines.sort(key=lambda l: l.position)
        
        return lines
    
    def _create_grid(self, horizontal_lines: List[Line], vertical_lines: List[Line], image_shape: Tuple[int, int]) -> Tuple[List[int], List[int]]:
        """
        グリッドを作成する
        
        Args:
            horizontal_lines: 水平線のリスト
            vertical_lines: 垂直線のリスト
            image_shape: 画像のサイズ (height, width)
            
        Returns:
            行座標のリスト、列座標のリスト
        """
        height, width = image_shape
        
        # Y座標（行）を抽出
        y_coords = [0]  # 上端
        for line in horizontal_lines:
            y = line.position
            # 重複を避ける
            if not any(abs(y - existing) < 5 for existing in y_coords):
                y_coords.append(y)
        y_coords.append(height)  # 下端
        y_coords.sort()
        
        # X座標（列）を抽出
        x_coords = [0]  # 左端
        for line in vertical_lines:
            x = line.position
            # 重複を避ける
            if not any(abs(x - existing) < 5 for existing in x_coords):
                x_coords.append(x)
        x_coords.append(width)  # 右端
        x_coords.sort()
        
        return y_coords, x_coords
    
    def _create_cells(self, grid_rows: List[int], grid_cols: List[int]) -> List[Cell]:
        """
        グリッドからセルを生成する
        
        Args:
            grid_rows: 行座標のリスト
            grid_cols: 列座標のリスト
            
        Returns:
            セルのリスト
        """
        cells = []
        
        for row_idx in range(len(grid_rows) - 1):
            for col_idx in range(len(grid_cols) - 1):
                y = grid_rows[row_idx]
                x = grid_cols[col_idx]
                height = grid_rows[row_idx + 1] - y
                width = grid_cols[col_idx + 1] - x
                
                cell = Cell(
                    row=row_idx,
                    col=col_idx,
                    x=x,
                    y=y,
                    width=width,
                    height=height
                )
                cells.append(cell)
        
        return cells
    
    def _detect_merged_cells(self, cells: List[Cell], binary_image: np.ndarray) -> Tuple[List[Cell], Dict[Tuple[int, int], Cell]]:
        """
        結合セルを検出する
        
        Args:
            cells: セルのリスト
            binary_image: 二値画像
            
        Returns:
            更新されたセルのリスト、結合マップ
        """
        # セルを行列に配置
        if not cells:
            return [], {}
        
        max_row = max(c.row for c in cells)
        max_col = max(c.col for c in cells)
        
        cell_map = {}
        for cell in cells:
            cell_map[(cell.row, cell.col)] = cell
        
        # 内部罫線の有無を確認してセル結合を判定
        visited = set()
        merged_cells = []
        
        for cell in cells:
            if (cell.row, cell.col) in visited:
                continue
            
            # 右方向への結合をチェック
            colspan = 1
            for c in range(cell.col + 1, max_col + 1):
                if (cell.row, c) not in cell_map:
                    break
                next_cell = cell_map[(cell.row, c)]
                
                # 境界に罫線があるかチェック
                boundary_x = cell.x + cell.width * colspan
                roi = binary_image[
                    cell.y + 5:cell.y + cell.height - 5,
                    boundary_x - 5:boundary_x + 5
                ]
                
                if roi.size == 0 or np.mean(roi == 0) < 0.3:  # 罫線が薄い
                    colspan += 1
                    visited.add((cell.row, c))
                else:
                    break
            
            # 下方向への結合をチェック
            rowspan = 1
            for r in range(cell.row + 1, max_row + 1):
                if (r, cell.col) not in cell_map:
                    break
                next_cell = cell_map[(r, cell.col)]
                
                # 境界に罫線があるかチェック
                boundary_y = cell.y + cell.height * rowspan
                roi = binary_image[
                    boundary_y - 5:boundary_y + 5,
                    cell.x + 5:cell.x + cell.width - 5
                ]
                
                if roi.size == 0 or np.mean(roi == 0) < 0.3:  # 罫線が薄い
                    rowspan += 1
                    visited.add((r, cell.col))
                else:
                    break
            
            # セルを更新
            if colspan > 1 or rowspan > 1:
                merged_cell = Cell(
                    row=cell.row,
                    col=cell.col,
                    x=cell.x,
                    y=cell.y,
                    width=cell.width * colspan,
                    height=cell.height * rowspan,
                    rowspan=rowspan,
                    colspan=colspan
                )
                merged_cells.append(merged_cell)
            else:
                merged_cells.append(cell)
            
            visited.add((cell.row, cell.col))
        
        # 新しいマップを作成
        new_map = {}
        for cell in merged_cells:
            new_map[(cell.row, cell.col)] = cell
        
        return merged_cells, new_map
    
    def _detect_border_styles(self, cells: List[Cell], horizontal_lines: List[Line], vertical_lines: List[Line]) -> Dict[Tuple[int, int], BorderStyle]:
        """
        罫線スタイルを判定する
        
        Args:
            cells: セルのリスト
            horizontal_lines: 水平線のリスト
            vertical_lines: 垂直線のリスト
            
        Returns:
            罫線スタイルの辞書
        """
        styles = {}
        
        for cell in cells:
            style = BorderStyle()
            
            # 上罫線
            for line in horizontal_lines:
                if abs(line.position - cell.y) < 5:
                    style.top = self._classify_line_style(line.thickness)
                    break
            
            # 下罫線
            for line in horizontal_lines:
                if abs(line.position - (cell.y + cell.height)) < 5:
                    style.bottom = self._classify_line_style(line.thickness)
                    break
            
            # 左罫線
            for line in vertical_lines:
                if abs(line.position - cell.x) < 5:
                    style.left = self._classify_line_style(line.thickness)
                    break
            
            # 右罫線
            for line in vertical_lines:
                if abs(line.position - (cell.x + cell.width)) < 5:
                    style.right = self._classify_line_style(line.thickness)
                    break
            
            styles[(cell.row, cell.col)] = style
        
        return styles
    
    def _classify_line_style(self, thickness: int) -> str:
        """
        線の太さからスタイルを分類
        
        Args:
            thickness: 線の太さ（ピクセル）
            
        Returns:
            スタイル名
        """
        if thickness <= self.config.thin_line_threshold:
            return "thin"
        elif thickness <= self.config.medium_line_threshold:
            return "medium"
        else:
            return "thick"


def create_detector(config: Optional[TableDetectionConfig] = None) -> TableDetector:
    """
    表検出器を作成する
    
    Args:
        config: 設定（Noneの場合はデフォルト）
        
    Returns:
        TableDetectorインスタンス
    """
    if config is None:
        config = TableDetectionConfig()
    
    return TableDetector(config)

