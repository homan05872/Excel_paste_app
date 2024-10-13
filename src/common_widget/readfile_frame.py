import customtkinter as ctk
import tkinter as tk
from tkinterdnd2 import DND_ALL

from typing import Any
import os
from pathlib import Path

from const.enums import ReadFileframe_Kinds
from windows import ConfirmWindow

class ReadFileframe(ctk.CTkFrame):
    ''' Excel読込用ウィジェット '''
    def __init__(self, master:Any, kind_flag, **kwargs):
        super().__init__(master, **kwargs)
        # 初期化
        self.style = master.style
        self.kind_flag = kind_flag
        
        # 背景色取得(CommonStyleから)
        self.bg_color = self.style.frame_bg['fg_color']
        
        # 画面表示用の文字列
        if self.kind_flag == ReadFileframe_Kinds.EXCEL :
            self.label = '元データファイル:'
            self.input_description = "Excelファイルを選択してください。" 
            self.drop_description = "Excelファイルをドラッグ&ドロップ"
        else :
            self.label = '出力フォルダ or ファイル:'
            self.input_description = "フォルダ or Excelファイルを選択してください。"
            self.drop_description = "フォルダ or Excelファイルをドラッグ&ドロップ"
            
        # UI生成
        self.build_ui()

    def build_ui(self):
        ''' UI生成 '''
        # 行方向のマスのレイアウトを設定する。
        self.grid_rowconfigure(5, weight=1)
        # 列方向のマスのレイアウトを設定する
        self.grid_columnconfigure(0, weight=1)
        
        # 入力
        self.input_label = ctk.CTkLabel(self, text=self.label, font=self.style.SMALL_DEFAULT)
        self.input_label.grid(row=0, column=0, pady=(30, 0), padx=20, sticky="nw")
        self.input = ctk.CTkEntry(self, placeholder_text=self.input_description)
        self.input.grid(row=1, column=0, pady=(0, 10), padx=20, sticky="nwe")
        self.select_btn = ctk.CTkButton(self, text='選択', width=70, command=self.open_picker,font=self.style.DEFAULT)
        self.select_btn.grid(row=1, column=1, padx=(0, 20), sticky="nwe")
        # エラーメッセージ
        self.error_label = ctk.CTkLabel(self, text='※選択されたファイルは対応していません。', text_color='red', font=self.style.SMALL_DEFAULT)
        self.error_label.grid(row=2, column=0, padx=20, sticky="nw", columnspan=2)
        self.error_label.grid_forget()
        # Canvasの作成
        self.drop_zone = tk.Canvas(self, bg=self.bg_color, highlightthickness=0, height=130)
        self.drop_zone.grid(row=3, column=0, pady=(10, 20), padx=20, sticky="new", columnspan=2)
        # シート選択
        self.sheet_select_list = ctk.CTkScrollableFrame(self)
        self.sheet_select_list.grid(row=5, column=0, padx=20, pady=10, columnspan=2, sticky="new")
        self.sheet_select_list.grid_columnconfigure(0, weight=1)
        self.no_choice_label = ctk.CTkLabel(self.sheet_select_list, text='選択できるシートがありません。', font=self.style.DEFAULT)
        self.no_choice_label.grid(row=0, column=0, padx=10, pady=(10, 0), sticky="w")
        
        # イベントへメソッドをバインド
        # 初期状態での点線と文字の描画
        self.drop_zone.bind("<Configure>", lambda event: self.draw_dashed_rectangle(event))     # drop_zoneのサイズ変化時
        # drop_zoneのへドラッグなどの操作イベントをバインド
        self.drop_zone.drop_target_register(DND_ALL)
        self.drop_zone.dnd_bind("<<Drop>>", lambda event: self.drop_file(event))        # ファイルドロップ
        self.drop_zone.dnd_bind("<<DropPosition>>", lambda event: self.on_drag(event))  # ファイルドラッグ
        self.drop_zone.dnd_bind("<<DropLeave>>", lambda event: self.on_drag_leave(event))   # ファイルドラッグが外れた時
        
    
    # ************************************************
    # 見た目に関するイベントメソッド
    # ************************************************
    def draw_dashed_rectangle(self, event=None, outline_color="gray", text_color="gray", file_name=None):
        ''' Drop_zoneの表示(点線の枠を描画する) & 再描画処理 '''
        # Canvasのサイズを取得
        if event:          # <Configure>イベントの時のみwidthとheightを取得
            self.width = event.width
            self.height = event.height
        
        if file_name:
            self.drop_description = self.drop_description.split('\n')[0]
            self.drop_description = self.drop_description + '\n添付：' + file_name
        
        # 既存の矩形を削除してから再描画
        self.drop_zone.delete("all")
        self.drop_zone.create_rectangle(10, 10, self.width - 10, self.height - 10, outline=outline_color, dash=(100, 20), width=4)
        # Frameの作成
        self.message_frame = ctk.CTkFrame(master=self, width=71, height=100)
        # テキストを中央に配置
        self.drop_zone.create_text(self.width/2, self.height/2, text=self.drop_description, font=("meiryo", 13, 'bold'), fill=text_color)

    def on_drag(self, event):
        ''' drop_zoneへファイルがドラッグされた時の処理 '''
        # 色を変更
        self.draw_dashed_rectangle(outline_color="white", text_color="white")

    def on_drag_leave(self, event):
        ''' drop_zoneからドラッグが外れた時の処理 '''
        # 元の色に戻す
        self.draw_dashed_rectangle()

    # ************************************************
    # ビジネスロジックに関連するイベントメソッド
    # ************************************************
    def open_picker(self):
        ''' ファイルピッカーを開きファイルパスを取得する処理 '''
        # 初期化
        choice = 'yes'
        path = None
        
        if self.kind_flag == ReadFileframe_Kinds.FOLDER:
            choice = tk.messagebox.askquestion("選択", "Excelファイルを選択しますか？フォルダを選択しますか？\n\nはい = Excelファイル, いいえ = フォルダ")

        if choice == 'yes':
            # ファイル選択ダイアログ (Excelファイル)
            path = tk.filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.xlsm *.xlsb")])
                
        elif choice == 'no' and self.kind_flag == ReadFileframe_Kinds.FOLDER:
            # フォルダ選択ダイアログ
            path = tk.filedialog.askdirectory()
            
        if path:
            self.display_excel_info(path)
    
    def drop_file(self, event:Any):
        ''' drop_zoneへファイルがドロップされた時の処理 '''
        # ファイルパスを取得し、正規化
        raw_path = event.data.strip('{}')
        file_path = os.path.normpath(raw_path)    # ファイルパスを正規化して取得
        self.display_excel_info(file_path)
    
    
    # ************************************************
    # その他のメソッド
    # ************************************************    
    def display_excel_info(self, file_path:str):
        ''' エクセルファイル読込処理 '''
        excel_flag = self.is_excel_path(file_path)
        
        # パスチェック
        if (excel_flag) or (self.kind_flag == ReadFileframe_Kinds.FOLDER and Path(file_path).is_dir()):
            # シート情報取得処理
            if excel_flag == False:
                # フォルダ内のExcelファイルのシート名を取得
                self.values = ["value", "value 2", "value 3", "value 4", "value 5", "value 6"]
                self.confirm_btns = []
            else:
                # 選択されたExcelファイル内のExcelファイルのシート名を取得
                self.values = ["value", "value 2", "value 3", "value 4", "value 5", "value 6"]
            
            # シート情報一覧の更新
            self.no_choice_label.destroy()
            self.checkboxes = []
            for i, value in enumerate(self.values):
                checkbox = ctk.CTkCheckBox(self.sheet_select_list, text=value)
                checkbox.grid(row=i, column=0, padx=10, pady=(10, 0), sticky="w")
                self.checkboxes.append(checkbox)
                # フォルダだった場合、確認ボタンを配置
                if excel_flag == False:
                    confirm_btn = ctk.CTkButton(self.sheet_select_list, width=30, text='確認', command=self.open_description,font=self.style.DEFAULT, **self.style.inline_btn)
                    confirm_btn.grid(row=i, column=1, padx=20, pady=(10, 0), sticky="e")
                    self.confirm_btns.append(confirm_btn)
            
            # inputの更新
            self.input.delete(0, ctk.END)  # 既存の値をクリア
            self.input.insert(0, file_path)  # インデックス0に挿入
            self.error_label.grid_forget()
                    
            # ドロップゾーンにファイル名表示
            file_name = os.path.basename(file_path)
            self.draw_dashed_rectangle(file_name=file_name)
            
        else :
            # 対応していないファイルが選択された場合のエラーメッセージ
            self.error_label.grid(row=2, column=0, padx=20, sticky="nw", columnspan=2)
            
    def open_description(self):
        ''' 対応するシート名があるExcelファイル一覧を表示する（別ウィンドウ） '''
        main_window = self.master.master
        confirm_window = ConfirmWindow(main_window)
        
    def is_excel_path(self, file_path):
        ''' Excelファイルであることを確認する処理 '''
        file_name = os.path.basename(file_path)
        excel_extensions = ['.xls', '.xlsx', '.xlsm', '.xlsb']
        _, extension = os.path.splitext(file_name)

        # 拡張子がExcelファイルのものであればTrueを返す
        return extension.lower() in excel_extensions