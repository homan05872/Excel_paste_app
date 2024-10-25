import win32com.client as win32
import pywintypes
import os
from pathlib import Path
from typing import Any
import tkinter as tk

from const.enums import ReadFileframe_Kinds


class ExcelModel:
    def __init__(self) -> None:
        self._source_path:str = ""
        self._paste_path_list:str = []
        self._folder_sheet_info = {}
        self._excel_app:Any = None
    
    # ゲッター
    @property
    def source_path(self):
        return self._source_path
    
    @property
    def paste_path_list(self):
        return self._paste_path_list
    
    @property
    def folder_sheet_info(self):
        return self._folder_sheet_info
    
    @property
    def excel_app(self):
        return self._excel_app
    
    # セッター
    @source_path.setter
    def source_path(self, source_path):
        self._source_path = source_path
    
    @paste_path_list.setter
    def paste_path_list(self, paste_path_list):
        self._paste_path_list = paste_path_list
    
    @folder_sheet_info.setter
    def folder_sheet_info(self,folder_sheet_info):
        self._folder_sheet_info = folder_sheet_info
    
    @excel_app.setter
    def excel_app(self, excel_app):
        self._excel_app = excel_app
    
        
    def open(self, file_path:str) -> tuple[Any, bool]:
        ''' Excelファイルを開くメソッド '''
        
        open_flag = self.is_excel_file_open(file_path)
        
        if not self.excel_app:
            # Excelアプリケーションを起動
            self.excel_app = win32.Dispatch("Excel.Application")
        
        # Excelの描画を停止する
        self.excel_app.Visible = False           # ウィンドウの表示 停止
        self.excel_app.ScreenUpdating = False    # 画面の更新を停止
        self.excel_app.DisplayAlerts = False     # 警告ダイアログの非表示
        
        # Excelファイルを開く (フルパスを指定)
        workbook = self.excel_app.Workbooks.Open(file_path)
        
        return workbook, open_flag
    
    def close(self, workbook:Any, open_flag:bool) -> None:
        ''' Excelファイルを閉じるメソッド '''
        excel = self.excel_app
        # Excelを閉じる
        if workbook and open_flag == False:
            workbook.Close(SaveChanges=False)
            excel.Quit()
            
        elif excel and open_flag == True :
            # Excelがもともと開いていた場合は描画を再開
            excel.Visible = True           # ウィンドウの表示を再開
            excel.ScreenUpdating = True    # 画面の更新を再開
            excel.DisplayAlerts = True     # 警告ダイアログ表示の再開
        
        # リソース解放
        self.excel_app = None
    
    
    def get_sheet_names_folder(self, folder_path:str) -> list:
        ''' Excel情報、取得メソッド '''
        
        # 初期化
        sheet_names = []            # Excelファイルのシート名
        unique_sheet_names = []     # 戻り値（重複なしのシート名）
        
        # フォルダ内のExcelファイル名を取得
        folder_path = Path(folder_path)
        excel_files = list(folder_path.glob('*.xlsx')) + list(folder_path.glob('*.xls')) + list(folder_path.glob('*.xlsm')) + list(folder_path.glob('*.xlsb'))
        
        for excel_file in excel_files:
            # フォルダ内のExcelファイルパスを取得、保持
            excel_path = os.path.join(folder_path, excel_file)
            self.paste_path_list.append(excel_path)
            sheet_names = self.get_sheet_names(excel_path, ReadFileframe_Kinds.FOLDER)
            
            # 貼り付け先シートを取得
            for sheet_name in sheet_names:
                if not sheet_name in unique_sheet_names:
                    unique_sheet_names.append(sheet_name)
            
            # フォルダーの場合Sheetに紐づくExcelファイル名を保持
            self.set_sheet_info(excel_file, sheet_names)
                
        return unique_sheet_names
    
    
    def get_sheet_names(self, file_path:str, kind_flag:str=ReadFileframe_Kinds.EXCEL) -> list:
        ''' シート取得メソッド '''
        
        if kind_flag == ReadFileframe_Kinds.EXCEL:
            self.source_path = file_path
        
        # 初期化
        workbook = None
        open_flag = None
        
        try:
            # Excelファイルを開く
            workbook, open_flag = self.open(file_path)

            # シート名をリストに格納
            sheet_names = [sheet.Name for sheet in workbook.Sheets]

            return sheet_names
        
        except pywintypes.com_error as e:  # COMエラーをキャッチ
            print(f"ファイルを開く際にエラーが発生しました: {e}")  # エラーメッセージを表示
            
        except Exception as e:  # その他のエラーをキャッチ
            print(f"予期せぬ、エラーが発生しました: {e}")  # エラーメッセージを表示
            
        finally:
            # Excelを閉じる
            self.close(workbook, open_flag)
            
    def set_sheet_info(self, excel_name:str, sheet_list:list[str]):
        ''' ※の情報を保持。※シートを持っているExcelファイル'''
        for sheet in sheet_list:
            if sheet not in self.folder_sheet_info:
                self.folder_sheet_info[sheet] = []
            self.folder_sheet_info[sheet].append(excel_name)
        
    def paste(self) -> None:
        ''' データ貼り付けメソッド '''
        pass
    
    def is_excel_file_open(self, file_path):
        ''' 既存で指定のExcelが開かれているかチェック ※開く前に呼び出す'''
        # Excelアプリケーションのインスタンスを取得
        try:
            excel = win32.GetActiveObject("Excel.Application")
        except Exception:
            # Excelが起動していない場合のエラーハンドリング
            return False

        # 開いているすべてのワークブックを確認
        for workbook in excel.Workbooks:
            if workbook.FullName.lower() == file_path.lower():
                return True
        
        return False