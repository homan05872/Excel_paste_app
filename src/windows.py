import customtkinter as ctk
from typing import Any, Tuple
from abc import ABC
from tkinterdnd2 import TkinterDnD, DND_ALL
    

class MainWindow(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, controllers: dict[str, Any], style: Any, **kwargs):
        super().__init__(**kwargs)
        
        # 初期化
        self.bg_theme = 'dark'              # 背景色テーマカラー
        self.ctl_theme = "blue"                 # 要素テーマカラー
        self.style = style                  # スタイルクラス
        self.controllers = controllers      # コントローラ群
        self.pages:dict[str] = {}           # ページのフレームを格納する辞書
        
        # テーマ設定
        self.bg_set(self.bg_theme)                   # Modes: system (default), light, dark
        ctk.set_default_color_theme(self.ctl_theme)  # Themes: blue (default), dark-blue, green
        
        # Gridレイアウトの設定
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        self.title("Excelコピペくん")
        self.geometry("860x550")
        
        # TkinterDnDのイベントを使用できるようにする設定
        self.TkdndVersion = TkinterDnD._require(self)
        
        
    def show_page(self, page_name:str) -> None:
        '''ページ切替を行うメソッド'''
        page = self.pages[page_name]
        page.tkraise()
        
    def page_set(self, pages:Any):
        ''' Pageクラスの配置を行うメソッド '''
        # ページクラス配置
        for PageClass in pages:
            page_name = PageClass.__name__
            page = PageClass(master=self, **self.style.base_bg)
            self.pages[page_name] = page
            page.grid(row=0, column=0, sticky="nsew")
        
    def bg_set(self, theme:str):
        ''' 背景テーマを設定するメソッド '''
        ctk.set_appearance_mode(theme)  # Modes: system (default), light, dark
        self.style.set_bg_theme(theme)


class ConfirmWindow(ctk.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("300x200")
        self.title("シートがあるファイル一覧")
        
        # 最前面に表示
        self.attributes('-topmost', True)
        
        self.grid_rowconfigure(0,weight=1)
        self.grid_columnconfigure(0,weight=1)
        
        self.excel_list = ctk.CTkScrollableFrame(self)
        self.excel_list.grid(row=0, column=0, padx=20, pady=10, sticky="nsew")
        
        values = ['Excel1','Excel2','Excel3']
        for i, value in enumerate(values):
                checkbox = ctk.CTkLabel(self.excel_list, text=value)
                checkbox.grid(row=i, column=0, padx=10, pady=(10, 0), sticky="w")