import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
from abc import ABC, abstractmethod
from common_widget.header_frame import HeaderFrame
from common_widget.readfile_frame import ReadFileframe
from const.enums import ReadFileframe_Kinds

class BasePage(ctk.CTkFrame, ABC):
    def __init__(self, master:ctk.CTk|tk.Tk, **kwargs) -> None:
        super().__init__(master, **kwargs)
        self.style = self.master.style
    
    @abstractmethod
    def build_ui(self) -> None:
        """UIを構築するための抽象メソッド"""
        ...
    
    def show_page(self, page_name:str) -> None:
        '''ページ遷移するメソッド'''
        self.master.show_page(page_name)

# **************************************
# メインページ
# **************************************        
class Main_Page(BasePage):
    def __init__(self, master:ctk.CTk, **kwargs) -> None:
        super().__init__(master, **kwargs)
        # UI生成
        self.build_ui()
        
    def build_ui(self) -> None:
        '''UI生成するメソッド'''
                
        # ウィンドウを3つの列でグリッドレイアウトに設定
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # ヘッダー
        self.header_frame = HeaderFrame(self, corner_radius=0, **self.style.frame_bg)
        self.header_frame.grid(row=0, column=0, columnspan=2, sticky="nsew")
        
        # 貼り付け元Excelファイル読み込み
        self.frame1 = ReadFileframe(self, ReadFileframe_Kinds.EXCEL, **self.style.frame_bg)
        self.frame1.grid(row=1, column=0, pady=20, padx=(30, 20), sticky="nsew")
        
        # 貼り付け先Excelファイルorフォルダ読み込み
        self.frame2 = ReadFileframe(self, ReadFileframe_Kinds.FOLDER, **self.style.frame_bg)
        self.frame2.grid(row=1, column=1, pady=20, padx=(0,30), sticky="nsew")



# **************************************
# 2ページ
# **************************************
class Page2(BasePage):
    def __init__(self, master:ctk.CTk, **kwargs) -> None:
        super().__init__(master, **kwargs)
        # UI生成
        self.build_ui()
        
    def build_ui(self) -> None:
        '''UI生成するメソッド'''
        # Gridレイアウト設定
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure((0,1), weight=1)
        
        self.label = ctk.CTkLabel(self, text="ページ２です。")
        self.label.grid(row=0, column=0, columnspan=2, pady=20)

        self.page_btn = ctk.CTkButton(self, text="ページ１へ", command=lambda: self.show_page("Main_Page"))
        self.page_btn.grid(row=1, column=0, padx=20, pady=40)
        
        self.msg_btn = ctk.CTkButton(self, text="メッセージ表示", command=lambda: self.msg_output(2), **self.style.inline_btn)
        self.msg_btn.grid(row=1, column=1, padx=(0,20), pady=40)
        
    def msg_output(self, page_num:int) -> None:
        '''メッセージを出力するメソッド'''
        messagebox.showinfo("Information", f"ページ{page_num}のメッセージです。")