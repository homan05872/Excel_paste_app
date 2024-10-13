from enum import Enum

FONT_TYPE = "meiryo"
BASE_COLOR_DARK = "#181818"
BASE_COLOR_LIGHT = "#fff"

class CommonStyle:
    ''' 共通のスタイルを定義するクラス '''
    def __init__(self) -> None:
        # 文字フォント設定
        self.HEADER_TITLE = (FONT_TYPE, 18, "bold")
        self.DEFAULT = (FONT_TYPE, 15)
        self.SMALL_DEFAULT = (FONT_TYPE, 13)
        
        # Windowの背景色
        self.base_bg = {
            "fg_color": BASE_COLOR_LIGHT,
        }
        
        self.frame_bg = {
            "fg_color": "#242424",
        }
        
        # インラインボタン
        self.inline_btn = {
            "text_color": ("gray10", "#DCE4EE"),
            "fg_color": "transparent",
            "border_width":2,
        }
        
    def set_bg_theme(self, theme:str):
        ''' windowsクラスの背景色テーマに合わせてデザイン変更 '''
        if theme == 'dark':
            self.base_bg["fg_color"] = BASE_COLOR_DARK
            self.frame_bg["fg_color"] = '#242424'
        elif theme == 'light':
            self.base_bg["fg_color"] = BASE_COLOR_LIGHT
            self.frame_bg["fg_color"] = '#DBDBDB'
        elif theme == 'system':
            self.base_bg["fg_color"] = BASE_COLOR_LIGHT
            self.frame_bg["fg_color"] = '#DBDBDB'