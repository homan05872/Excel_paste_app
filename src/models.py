import win32com.client as win32
import pywintypes
import os
from pathlib import Path
from typing import Any

from const.enums import ReadFileframe_Kinds


class ExcelModel:
    def __init__(self) -> None:
        self._source_path:str = ""      # データ取得元のファイルパス
        self._paste_path_list:str = []  # データ貼り付け対象のファイルパス
        self._folder_sheet_info = {}    # フォルダー内のExcelのシート情報
        self._excel_app:Any = None      # Excel操作のためのApplicationオブジェクト
    
    # ***************************************************
    # ゲッター
    # ***************************************************
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
    
    # ***************************************************
    # セッター
    # ***************************************************
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
    
    
    # ***************************************************
    # クラス外で使用するメソッド
    # ***************************************************
    def get_sheet_names_folder(self, folder_path:str) -> list:
        """_summary_
        Excelのシート保持情報を取得する取得するメソッド。
        
        Args:
            folder_path (str): 選択されたフォルダーパス

        Returns:
            list: フォルダー内のExcelパスのリスト
        """

        # 初期化
        sheet_names = []            # Excelファイルのシート名
        unique_sheet_names = []     # 戻り値（重複なしのシート名）
        self.folder_sheet_info = {} # シート情報
        
        # フォルダ内のExcelファイル名を取得
        folder_path = Path(folder_path)
        excel_path_list = list(folder_path.glob('*.xlsx')) + list(folder_path.glob('*.xls')) + list(folder_path.glob('*.xlsm')) + list(folder_path.glob('*.xlsb'))
        
        for excel_path in excel_path_list:
            # フォルダ内のExcelファイルパスを取得、保持
            self.paste_path_list.append(excel_path)
            excel_name = os.path.basename(excel_path)
            sheet_names = self.get_sheet_names(excel_path, ReadFileframe_Kinds.FOLDER)
            
            # 貼り付け先シートを取得
            for sheet_name in sheet_names:
                if not sheet_name in unique_sheet_names:
                    unique_sheet_names.append(sheet_name)
            
            # フォルダーの場合Sheetに紐づくExcelファイル名を保持
            self._set_folder_sheet_info(excel_name, sheet_names)
                
        return unique_sheet_names
    
    
    def get_sheet_names(self, file_path:str, kind_flag:str=ReadFileframe_Kinds.EXCEL) -> list:
        """_summary_
        シート取得メソッド

        Args:
            file_path (str): Excelファイルパス
            kind_flag (str, optional): 取得するシートの種類. Defaults to ReadFileframe_Kinds.EXCEL.

        Returns:
            list: Excelが保持しているシートのリスト
        """
        
        if kind_flag == ReadFileframe_Kinds.EXCEL:
            self.source_path = file_path
        
        # 初期化
        wb = None
        
        try:
            # Excelファイルを開く
            wb, opened_pathlist = self._open(file_path)

            # シート名をリストに格納
            sheet_names = [sheet.Name for sheet in wb.Sheets]

            return sheet_names
        
        except pywintypes.com_error as e:  # COMエラーをキャッチ
            print(f"ファイルを開く際にエラーが発生しました: {e}")  # エラーメッセージを表示
            
        except Exception as e:  # その他のエラーをキャッチ
            print(f"予期せぬ、エラーが発生しました: {e}")  # エラーメッセージを表示
            
        finally:
            # Excelを閉じる
            self._close(wb, opened_pathlist)
    
    # ***************************************************
    # クラス内で使用するメソッド
    # ***************************************************
    def _open(self, file_path:str) -> tuple[Any, list]:
        """_summary_
        Excelファイルを開き、描画を停止するメソッド
        
        Args:
            file_path (str): 開くExcelファイルのパス

        Returns:
            tuple[Any, list]: 取得したWorkbookオブジェクト, 既存で開かれていたExcelファイルのパス
        """
        opened_pathlist = []
        
        if not self.excel_app:
            # Excelアプリケーションを起動
            self.excel_app = win32.Dispatch("Excel.Application")
            
            # 既に開かれているExcelファイルのパスを取得
            for wb in self.excel_app.Workbooks:
                opened_pathlist.append(wb.FullName)
            
            # Excelの描画を停止する
            self.excel_app.Visible = False           # ウィンドウの表示 停止
            self.excel_app.ScreenUpdating = False    # 画面の更新を停止
            self.excel_app.DisplayAlerts = False     # 警告ダイアログの非表示
        
        # Excelファイルを開く (フルパスを指定)
        wb = self.excel_app.Workbooks.Open(file_path)
        
        return wb, opened_pathlist
    
    def _close(self, wb:Any, opened_pathlist:list, saveflag:bool=False) -> None:
        """_summary_
        Excelファイルを閉じるメソッド(既存で開かれていた場合は閉じない) 
        
        Args:
            wb (Any): 開いているExcelファイルオブジェクト
            opened_pathlist (list): 既存で開かれていたExcelのパスリスト
            saveflag (bool, optional): Excelファイルの保存の可否。 Defaults to False.
        """
        self.excel_app
        # Excelを閉じる
        if wb and opened_pathlist == []:
            wb.Close(SaveChanges=saveflag)
            self.excel_app.Quit()
            
        elif self.excel_app and opened_pathlist != [] :
            # 既存で開枯れていなかったExcelファイルのみ閉じる
            if wb.Name not in opened_pathlist:
                wb.Close(SaveChanges=saveflag)
            elif saveflag == True:
                wb.Save()
                
            # Excelがもともと開いていた場合は描画を再開
            self.excel_app.Visible = True           # ウィンドウの表示を再開
            self.excel_app.ScreenUpdating = True    # 画面の更新を再開
            self.excel_app.DisplayAlerts = True     # 警告ダイアログ表示の再開
        
        # リソース解放
        self.excel_app = None
            
    def _set_folder_sheet_info(self, excel_name:str, sheet_list:list[str]):
        """_summary_
        シートがどのExcelファイルへ保持されているかの情報を「self.folder_sheet_info」へ保持する

        Args:
            excel_name (str): Excelファイルパス
            sheet_list (list[str]): _description_
        """
        for sheet in sheet_list:
            if sheet not in self.folder_sheet_info:
                self.folder_sheet_info[sheet] = []
            self.folder_sheet_info[sheet].append(excel_name)
        
    def paste(self) -> None:
        ''' データ貼り付けメソッド '''
        pass