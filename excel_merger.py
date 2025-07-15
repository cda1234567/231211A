import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import datetime
import xlwings as xw

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Excel 合併工具')
        self.file_list = []
        self.setup_ui()

    def setup_ui(self):
        frame = tk.Frame(self.root)
        frame.pack(padx=10, pady=10)

        self.listbox = tk.Listbox(frame, width=90, height=10, selectmode=tk.MULTIPLE)  # 寬度由60改90
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
            filetypes=[('Excel 檔案', '*.xls *.xlsx')]
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

    def get_unique_filename(self, folder, base_name, ext, file_name_count):
        key = base_name + ext
        if key not in file_name_count:
            file_name_count[key] = 0
        file_name_count[key] += 1
        if file_name_count[key] == 1:
            return f"{base_name}{ext}"
        else:
            return f"{base_name}-{file_name_count[key]}{ext}"

    def get_unique_main_filename(self, folder, main_base_name, main_ext, file_name_count):
        key = main_base_name + "_main" + main_ext
        if key not in file_name_count:
            file_name_count[key] = 0
        file_name_count[key] += 1
        if file_name_count[key] == 1:
            return f"{main_base_name}_main{main_ext}"
        else:
            return f"{main_base_name}_main-{file_name_count[key]}{main_ext}"

    def merge_excels(self):
        if len(self.file_list) < 2:
            messagebox.showwarning('警告', '請至少選擇兩個 Excel 檔案！')
            return
        try:
            folder = os.path.join(os.getcwd(), datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S'))
            os.makedirs(folder, exist_ok=True)
            self.progress['maximum'] = len(self.file_list) - 1
            self.progress['value'] = 0

            file_name_count = {}
            main_file = self.file_list[0]
            main_base_name = os.path.splitext(os.path.basename(main_file))[0]
            main_ext = os.path.splitext(os.path.basename(main_file))[1]
            last_main_save_path = None

            app = xw.App(visible=False)
            main_wb = app.books.open(main_file)

            for i, sec_file in enumerate(self.file_list[1:], 1):
                sec_wb = app.books.open(sec_file)
                # 複製所有工作表到主檔（保留格式、巨集、群組等）
                for sht in sec_wb.sheets:
                    sht.api.Copy(Before=main_wb.sheets[0].api)
                # 副檔命名
                base_name = os.path.splitext(os.path.basename(sec_file))[0]
                ext = os.path.splitext(os.path.basename(sec_file))[1]
                save_name = self.get_unique_filename(folder, base_name, ext, file_name_count)
                sec_save = os.path.join(folder, save_name)
                sec_wb.save(sec_save)
                sec_wb.close()
                self.progress['value'] = i
                self.root.update_idletasks()
                # 主檔命名
                main_save_name = self.get_unique_main_filename(folder, main_base_name, main_ext, file_name_count)
                if last_main_save_path and os.path.exists(last_main_save_path):
                    try:
                        os.remove(last_main_save_path)
                    except Exception:
                        pass
                main_save_path = os.path.join(folder, main_save_name)
                main_wb.save(main_save_path)
                last_main_save_path = main_save_path
                # 重新載入主檔
                main_wb = app.books.open(main_save_path)
            main_wb.close()
            app.quit()
            messagebox.showinfo('完成', f'合併完成！檔案儲存於：\n{folder}')
        except Exception as e:
            messagebox.showerror('錯誤', str(e))

if __name__ == '__main__':
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop() 