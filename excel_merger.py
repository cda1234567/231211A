import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import datetime

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Excel 合併工具')
        self.file_list = []
        self.setup_ui()

    def setup_ui(self):
        frame = tk.Frame(self.root)
        frame.pack(padx=10, pady=10)

        self.listbox = tk.Listbox(frame, width=60, height=10, selectmode=tk.MULTIPLE)
        self.listbox.grid(row=0, column=0, columnspan=3, pady=5)
        self.listbox.bind('<Delete>', self.remove_selected_files)

        btn_add = tk.Button(frame, text='新增檔案', command=self.add_files)
        btn_add.grid(row=1, column=0, sticky='ew', padx=2)
        btn_remove = tk.Button(frame, text='移除選取', command=self.remove_selected_files)
        btn_remove.grid(row=1, column=1, sticky='ew', padx=2)
        btn_merge = tk.Button(frame, text='合併', command=self.merge_excels)
        btn_merge.grid(row=1, column=2, sticky='ew', padx=2)

        self.progress = ttk.Progressbar(frame, orient='horizontal', length=400, mode='determinate')
        self.progress.grid(row=2, column=0, columnspan=3, pady=10)

    def add_files(self):
        files = filedialog.askopenfilenames(
            title='選擇 Excel 檔案',
            filetypes=[('Excel 檔案', '*.xls *.xlsx *.xlsm *.xlsb')]
        )
        for f in files:
            if f not in self.file_list:
                self.file_list.append(f)
                self.listbox.insert(tk.END, f)

    def remove_selected_files(self, event=None):
        selected = list(self.listbox.curselection())
        for idx in reversed(selected):
            self.listbox.delete(idx)
            del self.file_list[idx]

    def merge_excels(self):
        if len(self.file_list) < 2:
            messagebox.showwarning('警告', '請至少選擇兩個檔案！')
            return
        try:
            folder = os.path.join(os.getcwd(), datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S'))
            os.makedirs(folder, exist_ok=True)
            self.progress['maximum'] = len(self.file_list) - 1
            self.progress['value'] = 0

            # 以第一個檔案為主檔
            main_file = self.file_list[0]
            main_df = pd.read_excel(main_file)
            for i, sec_file in enumerate(self.file_list[1:], 1):
                sec_df = pd.read_excel(sec_file)
                # 這裡僅示範合併，實際邏輯請根據 C# 版細節調整
                main_df = pd.concat([main_df, sec_df], ignore_index=True)
                sec_save = os.path.join(folder, f'secondary_{i}.xlsx')
                sec_df.to_excel(sec_save, index=False)
                self.progress['value'] = i
                self.root.update_idletasks()
            main_save = os.path.join(folder, 'main_merged.xlsx')
            main_df.to_excel(main_save, index=False)
            messagebox.showinfo('完成', f'合併完成！檔案儲存於：\n{folder}')
        except Exception as e:
            messagebox.showerror('錯誤', str(e))

if __name__ == '__main__':
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop() 