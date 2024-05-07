import tkinter as tk
from tkinter import messagebox
import os 
import openpyxl
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("自動入力アプリ")
        self.geometry("300x150")

        self.label = tk.Label(self, text="自動入力を実行します")
        self.label.pack(pady=20)

        self.button = tk.Button(self, text="実行", command=self.run_automation)
        self.button.pack()

    def run_automation(self):
        # 確認ダイアログを表示
        answer = messagebox.askyesno("確認", "自動入力を実行しますか?")

        if answer:
            self.perform_automation()
            # 完了通知を表示
            messagebox.showinfo("完了", "自動入力が完了しました")
    def perform_automation(self):
        pass

if __name__ == "__main__":
    app = App()
    app.mainloop()