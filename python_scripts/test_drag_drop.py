#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
測試拖曳功能
"""

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    print("✓ tkinterdnd2 可用")
    
    def test_drag_drop():
        root = TkinterDnD.Tk()
        root.title("拖曳測試")
        root.geometry("400x300")
        
        import tkinter as tk
        from tkinter import messagebox
        
        # 建立 Listbox
        listbox = tk.Listbox(root)
        listbox.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 註冊拖曳功能
        listbox.drop_target_register(DND_FILES)
        
        def on_drop(event):
            files = event.data.split()
            for file_path in files:
                file_path = file_path.strip('"{}')
                listbox.insert(tk.END, file_path)
        
        listbox.dnd_bind('<<Drop>>', on_drop)
        
        # 說明標籤
        label = tk.Label(root, text="請拖曳檔案到此處測試")
        label.pack(pady=10)
        
        root.mainloop()
    
    if __name__ == "__main__":
        test_drag_drop()
        
except ImportError:
    print("✗ tkinterdnd2 不可用，請執行：pip install tkinterdnd2")
except Exception as e:
    print(f"✗ 測試失敗：{e}") 