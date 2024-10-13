import customtkinter as ctk
from typing import Any

class HeaderFrame(ctk.CTkFrame):
    ''' ヘッダー用ウィジェット '''
    def __init__(self, master:Any, **kwargs):
        super().__init__(master, **kwargs)
        #初期化
        self.style = master.style
        # UI生成
        self.build_ui()

    def build_ui(self):
        ''' UI生成 '''
        # フレームのラベルを表示
        self.label = ctk.CTkLabel(self, text='Excelコピペくん', font=self.style.HEADER_TITLE)
        self.label.grid(row=0, column=0, padx=30, pady=15, sticky="w")
        
        self.submit = ctk.CTkButton(self, text='実行', width=100, font=('meiryo', 15))
        self.submit.grid(row=0, column=1, pady=15, sticky="e")