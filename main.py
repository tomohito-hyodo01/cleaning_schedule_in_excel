#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
清掃スケジュール Excel変換ツール
画像をClaude APIで解析してExcelファイルを生成します
"""

import os
import sys
import json
import base64
import threading
from pathlib import Path
from tkinter import filedialog, messagebox
import customtkinter as ctk
from PIL import Image, ImageTk
import anthropic
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# アプリケーション設定
APP_TITLE = "清掃スケジュール Excel変換ツール"
APP_VERSION = "1.0.0"
CONFIG_FILE = "config.json"
DEFAULT_SAVE_DIR = str(Path.home() / "Desktop")

# CustomTkinter テーマ設定
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class CleaningScheduleApp(ctk.CTk):
    """メインアプリケーションクラス"""
    
    def __init__(self):
        super().__init__()
        
        # ウィンドウ設定
        self.title(APP_TITLE)
        self.geometry("800x600")
        self.resizable(True, True)
        
        # 変数初期化
        self.image_path = None
        self.image_display = None
        self.api_key = None
        
        # 設定読み込み
        self.load_config()
        
        # UI構築
        self.create_widgets()
        
    def load_config(self):
        """設定ファイルを読み込む"""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.api_key = config.get('claude_api_key', '')
            except Exception as e:
                print(f"設定読み込みエラー: {e}")
    
    def save_config(self):
        """設定ファイルを保存"""
        try:
            config = {'claude_api_key': self.api_key or ''}
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"設定保存エラー: {e}")
    
    def create_widgets(self):
        """UI要素を作成"""
        
        # メインフレーム
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # タイトル
        title_label = ctk.CTkLabel(
            main_frame,
            text=APP_TITLE,
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=(0, 20))
        
        # 画像選択ボタン
        select_btn = ctk.CTkButton(
            main_frame,
            text="📁 画像ファイルを選択",
            command=self.select_image,
            font=ctk.CTkFont(size=16),
            height=40
        )
        select_btn.pack(pady=10)
        
        # 画像プレビューフレーム
        preview_frame = ctk.CTkFrame(main_frame)
        preview_frame.pack(fill="both", expand=True, pady=10)
        
        self.preview_label = ctk.CTkLabel(
            preview_frame,
            text="画像が選択されていません",
            font=ctk.CTkFont(size=14)
        )
        self.preview_label.pack(expand=True)
        
        # ファイル名表示
        self.filename_label = ctk.CTkLabel(
            main_frame,
            text="",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        self.filename_label.pack(pady=5)
        
        # 進行状況バー
        self.progress_bar = ctk.CTkProgressBar(main_frame)
        self.progress_bar.pack(fill="x", pady=10)
        self.progress_bar.set(0)
        
        # ステータスラベル
        self.status_label = ctk.CTkLabel(
            main_frame,
            text="待機中...",
            font=ctk.CTkFont(size=12)
        )
        self.status_label.pack(pady=5)
        
        # ボタンフレーム
        button_frame = ctk.CTkFrame(main_frame)
        button_frame.pack(pady=10)
        
        # 変換開始ボタン
        self.convert_btn = ctk.CTkButton(
            button_frame,
            text="🚀 Excel変換開始",
            command=self.start_conversion,
            font=ctk.CTkFont(size=16, weight="bold"),
            height=40,
            width=200,
            state="disabled"
        )
        self.convert_btn.pack(side="left", padx=5)
        
        # 設定ボタン
        settings_btn = ctk.CTkButton(
            button_frame,
            text="⚙️ 設定",
            command=self.open_settings,
            font=ctk.CTkFont(size=14),
            height=40,
            width=100
        )
        settings_btn.pack(side="left", padx=5)
    
    def select_image(self):
        """画像ファイルを選択"""
        filetypes = [
            ("画像ファイル", "*.jpg *.jpeg *.png"),
            ("JPEGファイル", "*.jpg *.jpeg"),
            ("PNGファイル", "*.png"),
            ("すべてのファイル", "*.*")
        ]
        
        filepath = filedialog.askopenfilename(
            title="清掃スケジュール画像を選択",
            filetypes=filetypes,
            initialdir=str(Path.home())
        )
        
        if filepath:
            self.image_path = filepath
            self.display_image(filepath)
            self.filename_label.configure(text=os.path.basename(filepath))
            self.convert_btn.configure(state="normal")
            self.status_label.configure(text="画像を選択しました")
    
    def display_image(self, filepath):
        """画像をプレビュー表示"""
        try:
            # 画像を読み込み
            img = Image.open(filepath)
            
            # プレビューサイズに調整（最大400x300）
            max_width, max_height = 400, 300
            img.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
            
            # PhotoImageに変換
            photo = ImageTk.PhotoImage(img)
            
            # 表示を更新
            self.preview_label.configure(image=photo, text="")
            self.preview_label.image = photo  # 参照を保持
            
        except Exception as e:
            messagebox.showerror("エラー", f"画像の読み込みに失敗しました:\n{e}")
    
    def open_settings(self):
        """設定ダイアログを開く"""
        settings_window = SettingsDialog(self, self.api_key)
        self.wait_window(settings_window)
        
        # 設定が更新された場合
        if hasattr(settings_window, 'result'):
            self.api_key = settings_window.result
            self.save_config()
    
    def start_conversion(self):
        """Excel変換を開始"""
        
        # APIキーチェック
        if not self.api_key:
            messagebox.showwarning(
                "APIキー未設定",
                "Claude APIキーが設定されていません。\n設定ボタンから設定してください。"
            )
            return
        
        # 画像チェック
        if not self.image_path or not os.path.exists(self.image_path):
            messagebox.showerror("エラー", "有効な画像が選択されていません")
            return
        
        # ボタンを無効化
        self.convert_btn.configure(state="disabled")
        
        # 別スレッドで処理を実行
        thread = threading.Thread(target=self.conversion_process, daemon=True)
        thread.start()
    
    def conversion_process(self):
        """変換処理（別スレッド）"""
        try:
            # ステップ1: 画像読み込み
            self.update_progress(0.1, "画像を読み込んでいます...")
            with open(self.image_path, 'rb') as f:
                image_data = base64.standard_b64encode(f.read()).decode('utf-8')
            
            # ステップ2: Claude APIで解析
            self.update_progress(0.3, "Claude APIで解析中...")
            table_data = self.analyze_with_claude(image_data)
            
            # ステップ3: Excel生成
            self.update_progress(0.7, "Excelファイルを生成中...")
            excel_path = self.generate_excel(table_data)
            
            # ステップ4: 完了
            self.update_progress(1.0, "完了しました！")
            
            # 完了メッセージ
            self.after(100, lambda: self.show_completion(excel_path))
            
        except Exception as e:
            error_msg = f"エラーが発生しました:\n{str(e)}"
            self.after(100, lambda: messagebox.showerror("エラー", error_msg))
            self.after(100, lambda: self.update_progress(0, "エラーが発生しました"))
        finally:
            self.after(100, lambda: self.convert_btn.configure(state="normal"))
    
    def analyze_with_claude(self, image_data):
        """Claude APIで画像を解析"""
        try:
            # Anthropic クライアントを初期化
            client = anthropic.Anthropic(
                api_key=self.api_key
            )
            
            message = client.messages.create(
                model="claude-opus-4-20250514",
                max_tokens=8000,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": "image/jpeg",
                                    "data": image_data,
                                },
                            },
                            {
                                "type": "text",
                                "text": """この清掃スケジュール表のすべてのデータを解析して、完全なJSON形式で出力してください。

【絶対厳守】
1. 表のすべての行を漏らさず出力してください（サンプルではなく全データ）
2. 説明文や注釈は一切含めず、純粋なJSONのみを出力してください
3. 「部分的な例」や「最初の数行のみ」といった省略は絶対にしないでください

【認識すべき項目】
- 左側の列：階、場所、作業品目、材質、作業面積、頻度/周期、年間回数
- 上部の日付列：1, 2, 3, ... 21（以上）
- 各セルの清掃頻度：1/D、2/W、3/M、1/月、空白など

【出力形式】
{
  "title": "清掃食器仕様書",
  "columns": ["階", "場所", "作業品目", "材質", "作業面積", "頻度/周期", "年間回数", "1/9", "2/9", "3/9", ...],
  "rows": [
    {"階": "1F", "場所": "東側", "作業品目": "事務室の清掃", "材質": "タイルカーペット", "作業面積": "49.48", "頻度/周期": "3/週", "年間回数": "", "1/9": "", "2/9": "", "3/9": "3/W", ...},
    {"階": "1F", "場所": "北側", ...},
    ...表のすべての行...
  ]
}

注意：JSON以外のテキスト（説明、コメント、注釈など）は一切含めないでください。"""
                            }
                        ]
                    }
                ]
            )
            
            # レスポンスからJSONを抽出
            response_text = message.content[0].text
            
            # デバッグ用：レスポンスをファイルに保存
            try:
                with open('claude_response.json', 'w', encoding='utf-8') as f:
                    f.write(response_text)
                print(f"Claude response saved to claude_response.json")
            except:
                pass
            
            # JSONの抽出（```json ``` で囲まれている場合に対応）
            if "```json" in response_text:
                json_start = response_text.find("```json") + 7
                json_end = response_text.find("```", json_start)
                response_text = response_text[json_start:json_end].strip()
            elif "```" in response_text:
                json_start = response_text.find("```") + 3
                json_end = response_text.find("```", json_start)
                response_text = response_text[json_start:json_end].strip()
            
            # JSONオブジェクトのみを抽出（余分なテキストを除去）
            # 最初の { から最後の } までを抽出
            json_start_bracket = response_text.find('{')
            json_end_bracket = response_text.rfind('}')
            
            if json_start_bracket != -1 and json_end_bracket != -1:
                response_text = response_text[json_start_bracket:json_end_bracket+1]
            
            table_data = json.loads(response_text)
            return table_data
            
        except Exception as e:
            raise Exception(f"Claude API解析エラー: {str(e)}")
    
    def generate_excel(self, table_data):
        """Excelファイルを生成"""
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "清掃スケジュール"
            
            current_row = 1
            
            # タイトル行（セル結合）
            if 'title' in table_data:
                num_cols = len(table_data.get('columns', []))
                if num_cols > 0:
                    ws['A1'] = table_data['title']
                    end_col = chr(64 + min(num_cols, 26))  # 最大Z列まで
                    ws.merge_cells(f'A1:{end_col}1')
                    ws['A1'].font = Font(size=14, bold=True, color='FFFFFF')
                    ws['A1'].fill = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
                    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
                    current_row += 1
            
            # ヘッダー行
            columns = table_data.get('columns', table_data.get('headers', []))
            if columns:
                for col_idx, header in enumerate(columns, start=1):
                    cell = ws.cell(row=current_row, column=col_idx, value=header)
                    cell.font = Font(bold=True, size=10)
                    cell.fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    cell.border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin')
                    )
                current_row += 1
            
            # データ行
            if 'rows' in table_data:
                for row_data in table_data['rows']:
                    for col_idx, header in enumerate(columns, start=1):
                        value = row_data.get(header, '')
                        cell = ws.cell(row=current_row, column=col_idx, value=value)
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        cell.border = Border(
                            left=Side(style='thin'),
                            right=Side(style='thin'),
                            top=Side(style='thin'),
                            bottom=Side(style='thin')
                        )
                    current_row += 1
            
            # 列幅を自動調整
            for column_cells in ws.columns:
                max_length = 0
                column_letter = None
                for cell in column_cells:
                    try:
                        # 結合されたセルをスキップ
                        if hasattr(cell, 'column_letter'):
                            if column_letter is None:
                                column_letter = cell.column_letter
                            if cell.value:
                                cell_length = len(str(cell.value))
                                if cell_length > max_length:
                                    max_length = cell_length
                    except:
                        pass
                
                # 列幅を設定
                if column_letter and max_length > 0:
                    adjusted_width = min(max_length + 2, 50)
                    ws.column_dimensions[column_letter].width = adjusted_width
            
            # 保存
            timestamp = Path(self.image_path).stem
            output_path = os.path.join(DEFAULT_SAVE_DIR, f"{timestamp}_変換結果.xlsx")
            
            # 同名ファイルがある場合は番号を追加
            counter = 1
            while os.path.exists(output_path):
                output_path = os.path.join(
                    DEFAULT_SAVE_DIR,
                    f"{timestamp}_変換結果_{counter}.xlsx"
                )
                counter += 1
            
            wb.save(output_path)
            return output_path
            
        except Exception as e:
            raise Exception(f"Excel生成エラー: {str(e)}")
    
    def update_progress(self, value, status_text):
        """進行状況を更新"""
        self.after(0, lambda: self.progress_bar.set(value))
        self.after(0, lambda: self.status_label.configure(text=status_text))
    
    def show_completion(self, excel_path):
        """完了ダイアログを表示"""
        result = messagebox.showinfo(
            "変換完了",
            f"Excelファイルを作成しました！\n\n保存先:\n{excel_path}\n\nフォルダを開きますか？"
        )
        
        # フォルダを開く
        if result:
            folder_path = os.path.dirname(excel_path)
            if sys.platform == 'win32':
                os.startfile(folder_path)
            elif sys.platform == 'darwin':
                os.system(f'open "{folder_path}"')
            else:
                os.system(f'xdg-open "{folder_path}"')


class SettingsDialog(ctk.CTkToplevel):
    """設定ダイアログ"""
    
    def __init__(self, parent, current_api_key):
        super().__init__(parent)
        
        self.title("設定")
        self.geometry("500x250")
        self.resizable(False, False)
        
        # モーダルにする
        self.transient(parent)
        self.grab_set()
        
        # 中央に配置
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")
        
        # UI構築
        self.create_widgets(current_api_key)
    
    def create_widgets(self, current_api_key):
        """設定画面のUI要素を作成"""
        
        frame = ctk.CTkFrame(self)
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # タイトル
        title = ctk.CTkLabel(
            frame,
            text="API設定",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title.pack(pady=(0, 20))
        
        # APIキー入力
        api_label = ctk.CTkLabel(
            frame,
            text="Claude API キー:",
            font=ctk.CTkFont(size=14)
        )
        api_label.pack(anchor="w", pady=(0, 5))
        
        self.api_entry = ctk.CTkEntry(
            frame,
            width=400,
            placeholder_text="sk-ant-api03-...",
            show="*"
        )
        self.api_entry.pack(pady=(0, 5))
        
        if current_api_key:
            self.api_entry.insert(0, current_api_key)
        
        # ヘルプテキスト
        help_text = ctk.CTkLabel(
            frame,
            text="APIキーは https://console.anthropic.com/ で取得できます",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        )
        help_text.pack(pady=(0, 20))
        
        # ボタンフレーム
        button_frame = ctk.CTkFrame(frame)
        button_frame.pack(pady=10)
        
        # 保存ボタン
        save_btn = ctk.CTkButton(
            button_frame,
            text="保存",
            command=self.save_settings,
            width=100
        )
        save_btn.pack(side="left", padx=5)
        
        # キャンセルボタン
        cancel_btn = ctk.CTkButton(
            button_frame,
            text="キャンセル",
            command=self.destroy,
            width=100,
            fg_color="gray"
        )
        cancel_btn.pack(side="left", padx=5)
    
    def save_settings(self):
        """設定を保存"""
        api_key = self.api_entry.get().strip()
        
        if not api_key:
            messagebox.showwarning("警告", "APIキーを入力してください")
            return
        
        if not api_key.startswith("sk-ant-"):
            messagebox.showwarning("警告", "Claude APIキーの形式が正しくありません")
            return
        
        self.result = api_key
        messagebox.showinfo("完了", "設定を保存しました")
        self.destroy()


def main():
    """メイン関数"""
    app = CleaningScheduleApp()
    app.mainloop()


if __name__ == "__main__":
    main()

