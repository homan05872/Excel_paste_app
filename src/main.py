from windows import MainWindow
from pages import Main_Page, Page2
from style_manager import CommonStyle

class App:
    def __init__(self) -> None:
        ''' 各クラスの連携(インスタンス生成) '''
        # モデルクラス群を生成
        
        # コントローラクラス群の辞書型定義
        self.main_controllers={}
        # スタイルクラス群を生成
        self.common_style=CommonStyle()     # MainWindowクラスで使用
        # ウィンドウクラスにコントローラークラス群を渡す
        self.main_window = MainWindow(self.main_controllers, self.common_style)
        # ページクラス配置
        self.main_window.page_set([Main_Page, Page2])      # ← 配置したいPageクラスを配列で渡す
        
    def run(self) -> None:
        ''' アプリ起動処理 '''
        self.main_window.show_page("Main_Page") # 最初に表示したいページクラス名を渡す
        self.main_window.mainloop()         # 起動


def main() -> None:
    # 各クラスのインスタンス生成
    app = App()
    # アプリ起動
    app.run()
    

if __name__ == "__main__":
    main()