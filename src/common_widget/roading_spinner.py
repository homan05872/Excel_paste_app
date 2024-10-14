import customtkinter as ctk
from PIL import Image
import os

class LoadingSpinner():
    def __init__(self, master, **kwargs):
        
        self.speed = 50
        self.angle = 0  # 角度を初期化
        # 現在のスクリプトのディレクトリから、read_img.pngの相対パスを生成
        current_dir = os.path.dirname(os.path.abspath(__file__))
        self.icon_path = os.path.join(current_dir, '..', 'img', 'read_img.png')

        # アイコン画像をロード
        self.image = Image.open(self.icon_path)
        self.image_size = (140, 140)  # アイコンのサイズを指定
        self.image = self.image.resize(self.image_size, Image.Resampling.LANCZOS)  # Pillowで画像をリサイズ

        # CTkImageを使用して画像をセット
        self.ctk_image = ctk.CTkImage(self.image, size=self.image_size)

        # 初期画像をLabelに表示
        self.image_label = ctk.CTkLabel(master, image=self.ctk_image, fg_color='white', text='Loading...', text_color='gray', font=('meiryo', 10, 'bold'))  
        # 中央に配置
        self.image_label.place(relx=0.5, rely=0.5, anchor="center")
        # アニメーションを開始
        self.animate()

    def animate(self):
        # 画像を回転させる
        self.angle = (self.angle - 10) % 360  # 角度を更新
        rotated_image = self.image.rotate(self.angle)  # PILを使って回転させる
        self.ctk_image = ctk.CTkImage(rotated_image, size=self.image_size)  # 再度画像をCTkImageで生成
        self.image_label.configure(image=self.ctk_image)  # Labelに再設定

        # afterメソッドで一定間隔で再度呼び出す
        self.image_label.after(self.speed, self.animate)

    def stop(self):
        # アニメーションを停止
        self.image_label.destroy()