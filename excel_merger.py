import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import datetime
import xlwings as xw
import pandas as pd
import numpy as np
from pathlib import Path
import subprocess
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    # 如果沒有安裝 tkinterdnd2，使用標準 tkinter
    TkinterDnD = None
    DND_FILES = None

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title('扣帳軟體')
        self.file_list = []
        self.setup_ui()

    def setup_ui(self):
        # 主框架
        frame = tk.Frame(self.root)
        frame.pack(padx=10, pady=10, fill='both', expand=True)

        # 檔案清單
        list_frame = tk.Frame(frame)
        list_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        self.listbox = tk.Listbox(list_frame, selectmode=tk.BROWSE, width=90, height=15)  # 單選模式，寬度放大
        self.listbox.pack(side='left', fill='both', expand=True)
        
        # 滾動條
        scrollbar = tk.Scrollbar(list_frame, orient='vertical', command=self.listbox.yview)
        scrollbar.pack(side='right', fill='y')
        self.listbox.config(yscrollcommand=scrollbar.set)

        # 支援鍵盤快捷鍵
        self.listbox.bind('<Delete>', lambda e: self.remove_selected_files())
        
        # 支援拖曳功能
        if TkinterDnD and DND_FILES:
            self.listbox.drop_target_register(DND_FILES)
            self.listbox.dnd_bind('<<Drop>>', self.on_file_drop)
        
        # 按鈕框架
        btn_frame = tk.Frame(frame)
        btn_frame.pack(fill='x', pady=(0, 10))
        
        btn_add = tk.Button(btn_frame, text='新增檔案', command=self.add_files)
        btn_add.pack(side='left', padx=(0, 5))
        
        btn_remove = tk.Button(btn_frame, text='移除選取', command=self.remove_selected_files)
        btn_remove.pack(side='left', padx=(0, 5))
        
        btn_up = tk.Button(btn_frame, text='往上', command=self.move_up)
        btn_up.pack(side='left', padx=(0, 5))
        
        btn_down = tk.Button(btn_frame, text='往下', command=self.move_down)
        btn_down.pack(side='left', padx=(0, 5))
        
        # 進度條
        self.progress = ttk.Progressbar(frame, orient='horizontal', mode='determinate', length=600)
        self.progress.pack(fill='x', pady=(0, 10))
        
        # 目前執行檔案顯示
        self.current_file_label = tk.Label(frame, text='目前執行到的檔案：', anchor='w', font=('Arial', 12))
        self.current_file_label.pack(fill='x', padx=5, pady=(0, 5), anchor='w')

        # 執行按鈕（放大）
        btn_execute = tk.Button(frame, text='執行', command=self.execute_merge, font=('Arial', 16, 'bold'), width=10, height=2)
        btn_execute.pack(side='right', padx=10, pady=10)
        
        # Rev3 標籤（右下角）
        rev3_frame = tk.Frame(frame)
        rev3_frame.pack(side='bottom', fill='x')
        rev3_label = tk.Label(rev3_frame, text='Rev3', font=('Arial', 10, 'bold'))
        rev3_label.pack(side='right', anchor='se', padx=5, pady=2)

    def on_click(self, event):
        self.drag_start = self.listbox.nearest(event.y)

    def on_drag(self, event):
        if hasattr(self, 'drag_start'):
            current = self.listbox.nearest(event.y)
            if current != self.drag_start:
                # 這裡可以實現拖曳排序功能
                pass

    def on_release(self, event):
        if hasattr(self, 'drag_start'):
            del self.drag_start

    def on_file_drop(self, event):
        """處理檔案拖曳"""
        files = event.data.split()
        for file_path in files:
            # 移除可能的引號和括號
            file_path = file_path.strip('"{}')
            if os.path.exists(file_path):
                ext = os.path.splitext(file_path)[1].lower()
                if ext in ['.xls', '.xlsx', '.xlsm', '.xlsb']:
                    if file_path not in self.file_list:
                        self.file_list.append(file_path)
                        self.listbox.insert(tk.END, file_path)

    def add_files(self):
        files = filedialog.askopenfilenames(
            title='選擇 Excel 檔案',
            filetypes=[('Excel 檔案', '*.xls *.xlsx *.xlsm *.xlsb')]
        )
        for f in files:
            if f not in self.file_list:
                self.file_list.append(f)
                self.listbox.insert(tk.END, f)  # 顯示完整路徑

    def remove_selected_files(self):
        selected = list(self.listbox.curselection())
        for idx in reversed(selected):
            self.listbox.delete(idx)
            del self.file_list[idx]

    def move_up(self):
        selected = list(self.listbox.curselection())
        if len(selected) != 1 or selected[0] == 0:
            return
        idx = selected[0]
        above = idx - 1
        self.file_list[above], self.file_list[idx] = self.file_list[idx], self.file_list[above]
        text = self.listbox.get(idx)
        self.listbox.delete(idx)
        self.listbox.insert(above, text)
        self.listbox.selection_set(above)
        self.listbox.selection_clear(idx)

    def move_down(self):
        selected = list(self.listbox.curselection())
        if len(selected) != 1 or selected[0] == self.listbox.size() - 1:
            return
        idx = selected[0]
        below = idx + 1
        self.file_list[below], self.file_list[idx] = self.file_list[idx], self.file_list[below]
        text = self.listbox.get(idx)
        self.listbox.delete(idx)
        self.listbox.insert(below, text)
        self.listbox.selection_set(below)
        self.listbox.selection_clear(idx)

    def find_last_non_empty_column_value_in_row(self, data_array, row_index):
        """找到指定行中最後一個非空欄位的數值"""
        for col in range(data_array.shape[1] - 1, -1, -1):
            value = data_array.iloc[row_index, col]
            if pd.notna(value):
                try:
                    return int(value)
                except (ValueError, TypeError):
                    continue
        return 0

    def execute_merge(self):
        if len(self.file_list) < 2:
            messagebox.showwarning('警告', '請至少選擇兩個 Excel 檔案！')
            return

        try:
            # 建立輸出資料夾 - 預設使用網路路徑，失敗則使用本地路徑
            timestamp = datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
            
            # 預設網路路徑
            default_network_path = Path(f"\\\\St-nas\\個人資料夾\\Andy\\excel\\{timestamp}")
            
            try:
                # 嘗試建立網路路徑資料夾
                default_network_path.mkdir(parents=True, exist_ok=True)
                folder_path = default_network_path
                print(f"使用網路路徑：{folder_path}")
            except Exception as network_error:
                # 如果網路路徑失敗，使用本地路徑
                local_folder = Path(f"./output_{timestamp}")
                local_folder.mkdir(parents=True, exist_ok=True)
                folder_path = local_folder
                print(f"網路路徑無法存取，使用本地路徑：{folder_path}")

            # 設定進度條
            self.progress['maximum'] = len(self.file_list) - 1
            self.progress['value'] = 0

            # 檔案名稱計數器
            file_name_count = {}
            last_main_save_path = None

            # 啟動 Excel 應用程式
            app = xw.App(visible=False)
            
            # 開啟主檔案
            main_file = self.file_list[0]
            main_wb = app.books.open(main_file)
            main_sheet = main_wb.sheets[0]
            
            # 讀取主檔案資料
            main_range = main_sheet.used_range
            main_data = main_range.options(pd.DataFrame, index=False, header=False).value

            # 處理每個次要檔案
            for i, secondary_file in enumerate(self.file_list[1:], 1):
                # 顯示目前執行到的檔案
                self.current_file_label.config(text=f'目前執行到的檔案：{os.path.basename(secondary_file)}')
                self.root.update_idletasks()
                # 開啟次要檔案
                sec_wb = app.books.open(secondary_file)
                sec_sheet = sec_wb.sheets[0]
                sec_range = sec_sheet.used_range
                sec_data = sec_range.options(pd.DataFrame, index=False, header=False).value

                # 資料處理邏輯
                for j in range(len(sec_data)):
                    for k in range(len(main_data)):
                        # 比對第3欄和第1欄
                        sec_value = str(sec_data.iloc[j, 2]).strip() if pd.notna(sec_data.iloc[j, 2]) else ""
                        main_value = str(main_data.iloc[k, 0]).strip() if pd.notna(main_data.iloc[k, 0]) else ""
                        
                        if sec_value == main_value:
                            # 找到最後一個有值的欄位
                            last_col = main_data.shape[1] - 1
                            for col in range(last_col, -1, -1):
                                if pd.notna(main_data.iloc[0, col]):
                                    # 處理第7欄的資料（取最後8位數字）
                                    fo_value = sec_data.iloc[0, 6] if pd.notna(sec_data.iloc[0, 6]) else ""
                                    g3 = str(fo_value)
                                    last_eight_digits = g3[-8:] if len(g3) >= 8 else g3

                                    # 設定標題欄位
                                    main_sheet.range(1, col + 4).value = sec_data.iloc[1, 2]  # 第2行第3欄
                                    main_sheet.range(1, col + 3).value = last_eight_digits

                                    # 設定字體格式
                                    main_sheet.range(1, col + 4).font.name = "Arial"
                                    main_sheet.range(1, col + 4).font.size = 9
                                    main_sheet.range(1, col + 4).api.WrapText = True
                                    
                                    main_sheet.range(1, col + 3).font.name = "Arial"
                                    main_sheet.range(1, col + 3).font.size = 9
                                    main_sheet.range(1, col + 3).api.WrapText = True

                                    # 找到主檔案中該行的最後非空欄位值
                                    last_value = self.find_last_non_empty_column_value_in_row(main_data, k)
                                    
                                    # 處理第7欄
                                    if pd.notna(sec_data.iloc[j, 6]) and str(sec_data.iloc[j, 6]).strip() == "-":
                                        break
                                    else:
                                        sec_sheet.range(j + 1, 7).value = last_value

                                    # 處理第6欄
                                    co_value = sec_data.iloc[j, 5]
                                    if pd.notna(co_value):
                                        f2 = int(co_value)
                                        main_sheet.range(k + 1, col + 3).value = f2

                                    # 處理第8欄
                                    add_value = sec_data.iloc[j, 7]
                                    if pd.notna(add_value) and str(add_value).strip() != "-":
                                        try:
                                            f3 = int(add_value)
                                            if f3 != 0:
                                                main_sheet.range(k + 1, col + 2).value = f3
                                        except (ValueError, TypeError):
                                            pass

                                    # 處理第10欄（四捨五入）
                                    original_value = sec_data.iloc[j, 9]
                                    if pd.notna(original_value):
                                        four_to_five = int(round(original_value))
                                        main_sheet.range(k + 1, col + 4).value = four_to_five

                                        # 負值標示為紅色
                                        if original_value < 0:
                                            main_sheet.range(k + 1, col + 4).color = (255, 0, 0)
                                    
                                    break

                # 儲存次要檔案
                base_name = Path(secondary_file).stem
                ext = Path(secondary_file).suffix
                save_name = base_name + ext

                if base_name not in file_name_count:
                    file_name_count[base_name] = 0
                file_name_count[base_name] += 1
                if file_name_count[base_name] > 1:
                    save_name = f"{base_name}-{file_name_count[base_name]}{ext}"

                secondary_save_path = folder_path / save_name
                sec_wb.save(str(secondary_save_path))
                sec_wb.close()

                # 儲存主檔案
                main_base_name = Path(main_file).stem
                main_ext = Path(main_file).suffix
                main_save_name = main_base_name + main_ext
                
                if f"{main_base_name}_main" not in file_name_count:
                    file_name_count[f"{main_base_name}_main"] = 0
                file_name_count[f"{main_base_name}_main"] += 1
                
                if file_name_count[f"{main_base_name}_main"] > 1:
                    main_save_name = f"{main_base_name}_main-{file_name_count[f'{main_base_name}_main']}{main_ext}"
                else:
                    main_save_name = f"{main_base_name}_main{main_ext}"

                # 刪除舊的主檔案
                if last_main_save_path and last_main_save_path.exists():
                    try:
                        last_main_save_path.unlink()
                    except:
                        pass

                main_save_path = folder_path / main_save_name
                main_wb.save(str(main_save_path))
                last_main_save_path = main_save_path

                # 重新載入主檔案
                main_wb.close()
                main_wb = app.books.open(str(main_save_path))
                main_sheet = main_wb.sheets[0]
                main_range = main_sheet.used_range
                main_data = main_range.options(pd.DataFrame, index=False, header=False).value

                # 更新進度條
                self.progress['value'] = i
                self.root.update_idletasks()

            # 最終儲存
            import time
            time.sleep(1)
            main_wb.save()
            main_wb.close()
            
            # 確保 Excel 程序完全關閉
            try:
                app.quit()
                time.sleep(2)  # 等待 Excel 完全關閉
            except:
                pass

            messagebox.showinfo('完成', '處理完成！')
            
            # 使用 destroy 而不是 quit 來確保視窗正確關閉
            self.root.after(1000, self.root.destroy)

        except Exception as e:
            # 強制關閉 Excel 程序
            try:
                subprocess.run(['taskkill', '/f', '/im', 'excel.exe'], 
                             capture_output=True, check=False)
                time.sleep(1)  # 等待程序關閉
            except:
                pass
            
            messagebox.showerror('錯誤', f'錯誤：{str(e)}')
            
            # 確保視窗可以正常關閉
            self.root.after(1000, lambda: self.root.destroy() if self.root.winfo_exists() else None)

if __name__ == '__main__':
    # 使用 TkinterDnD 如果可用，否則使用標準 tkinter
    if TkinterDnD:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    
    root.geometry('750x600')
    app = ExcelMergerApp(root)
    root.mainloop() 