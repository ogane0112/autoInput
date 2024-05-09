import tkinter as tk
from tkinter import messagebox
import os 
import openpyxl
import pandas as pd

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("自動入力アプリ")
        self.geometry("300x150")

        self.label = tk.Label(self, text="自動入力を実行します")
        self.label.pack(pady=20)

        self.button = tk.Button(self, text="実行", command=self.run_automation)
        self.button.pack()
    def get_file_name(self):
        # 現在のディレクトリ内のファイルとフォルダを取得
        current_dir = os.getcwd()
        target_dir = f"{current_dir}\input"
        dir_contents = os.listdir(target_dir)

        print(f"{current_dir}\input")
        # inputフォルダー内のファイル名を取得する関数
        input_list = []
        for item in dir_contents:
            full_path = os.path.join(current_dir, item)
            if os.path.isdir(full_path):
                print(f"[ディレクトリ] {item}")
            else:
                print(f"[ファイル] {item}")
                input_list.append(item)
                
        return input_list

    def run_automation(self):
        # 確認ダイアログを表示
        answer = messagebox.askyesno("確認", "自動入力を実行しますか?")

        if answer:
            self.perform_automation()
            # 完了通知を表示
            messagebox.showinfo("完了", "自動入力が完了しました")
            
    def perform_automation(self):
        file_name = self.get_file_name()
        
        for i in file_name:
            # ファイル名==output.xlsxでなかったら次のループへ移動
            #部員の名前を取得する
            df_name_first = pd.read_excel(f"{i}.xlsx",usecols="D",skiprows=11,nrows=5,names=["氏名"])

            df_name_second = pd.read_excel(f"{i}.xlsx",usecols="D",skiprows=17,nrows=20,names=["氏名"])

            # df_name_three = pd.read_excel(f"{i}.xlsx",usecols="D",skiprows=17,nrows=20)

            df_name_four = pd.read_excel(f"{i}.xlsx",usecols="S",skiprows=17,nrows=20,names=["氏名"])

            df_name_five = pd.read_excel(f"{i}.xlsx",usecols="C",skiprows=41,nrows=35,names=["氏名"])

            df_name_six = pd.read_excel(f"{i}.xlsx",usecols="S",skiprows=41,nrows=35,names=["氏名"])

            df_name_five = pd.read_excel(f"{i}.xlsx",usecols="C",skiprows=79,nrows=35,names=["氏名"])

            df_name_six = pd.read_excel(f"{i}.xlsx",usecols="S",skiprows=79,nrows=35,names=["氏名"])



            #部員の学籍番号の取得
            df_num_first = pd.read_excel(f"{i}.xlsx",usecols="K",skiprows=11,nrows=5,names=["学籍番号"])

            df_num_second = pd.read_excel(f"{i}.xlsx",usecols="K",skiprows=17,nrows=20,names=["学籍番号"])

            # df_name_three = pd.read_excel("{i}.xlsx",usecols="K",skiprows=17,nrows=20)

            df_num_four = pd.read_excel(f"{i}.xlsx",usecols="AA",skiprows=17,nrows=20,names=["学籍番号"])

            df_num_five = pd.read_excel(f"{i}.xlsx",usecols="K",skiprows=41,nrows=35,names=["学籍番号"])

            df_num_six = pd.read_excel(f"{i}.xlsx",usecols="AA",skiprows=41,nrows=35,names=["学籍番号"])

            df_num_five = pd.read_excel(f"{i}.xlsx",usecols="K",skiprows=79,nrows=35,names=["学籍番号"])

            df_num_six = pd.read_excel(f"{i}.xlsx",usecols="AA",skiprows=79,nrows=35,names=["学籍番号"])



            # 名前のデータフレームを結合
            df_names = pd.concat([df_name_first, df_name_second, df_name_four,df_name_five,df_name_six ], ignore_index=True)

            #学籍番号のデータフレームを結合
            df_nums = pd.concat([df_num_first, df_num_second, df_num_four,df_num_five,df_num_six ], ignore_index=True)

            # 欠損値を含む行を削除
            df_names_perfect = df_names.dropna()

            df_nums_perfect = df_nums.dropna()
            
if __name__ == "__main__":
    app = App()
    app.mainloop()