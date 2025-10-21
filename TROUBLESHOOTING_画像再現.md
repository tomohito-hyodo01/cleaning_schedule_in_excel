# 画像再現の問題と解決方法

## 現在の問題

清掃スケジュール画像の変換で、以下の問題が発生しています：

1. **二値化処理が不適切**
   - 透かし文字（"sample"）の影響で、表の左側が真っ黒になる
   - 罫線が途切れて正しく検出されない

2. **表構造の抽出失敗**
   - 罫線検出がうまくいかず、セルが正しく分割されない
   - 結果として、Excelファイルが元の画像を再現できない

## 解決方法

### 方法1: 改善版設定で実行（推奨）

```bash
python run_with_better_config.py
```

この設定では以下を改善：
- 適応的二値化（adaptive）を使用
- より大きなブロックサイズ（25）でローカル適応
- 罫線検出の閾値を調整
- カーネルサイズを大きく（60）

### 方法2: 画像を前処理

元画像の品質を改善してから変換：

```python
import cv2
import numpy as np

# 画像読み込み
img = cv2.imread('schedule_image/清掃スケジュール.jpg')

# 透かし除去（ガンマ補正）
gamma = 1.5
lookup = np.array([((i / 255.0) ** gamma) * 255 for i in range(256)]).astype(np.uint8)
img_corrected = cv2.LUT(img, lookup)

# 保存
cv2.imwrite('schedule_image/清掃スケジュール_前処理済み.jpg', img_corrected)
```

### 方法3: 設定パラメータの調整

`config_for_sample.json`を編集して、以下のパラメータを調整：

**二値化設定:**
```json
"binarization_method": "adaptive",  // "otsu" から変更
"adaptive_block_size": 25,          // 大きいほどグローバルに
"adaptive_c": 15                    // 閾値の調整値
```

**罫線検出:**
```json
"min_line_length": 50,              // より長い線のみ検出
"horizontal_kernel_scale": 60,      // より太いカーネル
"vertical_kernel_scale": 60
```

**コントラスト:**
```json
"clahe_clip_limit": 3.0,           // より強いコントラスト強調
"clahe_tile_size": 8
```

### 方法4: 手動で表部分のみ切り出し

画像編集ソフトで、表の部分のみを切り出してから変換：

1. Windowsの「ペイント」などで画像を開く
2. 表の部分のみを選択して切り取り
3. 新しい画像として保存
4. 変換を実行

## テスト手順

1. **設定を変更**
   ```bash
   # config_for_sample.json を編集
   ```

2. **実行**
   ```bash
   python run_with_better_config.py
   ```

3. **デバッグ画像を確認**
   ```
   debug_output/清掃スケジュール_binary.png  ← 二値化が適切か確認
   debug_output/清掃スケジュール_cells.png   ← セルが正しく検出されているか確認
   ```

4. **パラメータを調整**
   - 二値化が暗すぎる → `adaptive_c` を減らす
   - 罫線が検出されない → `min_line_length` を減らす
   - ノイズが多い → `denoise_kernel_size` を増やす

## 期待される結果

改善版設定では：
- 表構造が正しく検出される
- 罫線が適切に抽出される
- OCRで文字が読み取れる
- Excelファイルが元の画像に近い形で再現される

## さらなる改善

もし上記でも改善しない場合：

1. **画像の解像度を上げる**
   - より高解像度でスキャン/撮影

2. **照明・コントラストを調整**
   - 元画像の品質を改善

3. **OCRエンジンを変更**
   ```bash
   python main.py --cli input.jpg -o output.xlsx --engine tesseract
   ```

4. **手動補正**
   - 自動変換後、Excelで手動調整

