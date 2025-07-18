#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
測試腳本：驗證所有必要的模組都可以正常匯入
"""

def test_imports():
    """測試所有必要的模組匯入"""
    try:
        import tkinter as tk
        print("✓ tkinter 匯入成功")
        
        from tkinter import filedialog, messagebox, ttk
        print("✓ tkinter 子模組匯入成功")
        
        import os
        print("✓ os 匯入成功")
        
        import datetime
        print("✓ datetime 匯入成功")
        
        import xlwings as xw
        print("✓ xlwings 匯入成功")
        
        import pandas as pd
        print("✓ pandas 匯入成功")
        
        import numpy as np
        print("✓ numpy 匯入成功")
        
        from pathlib import Path
        print("✓ pathlib 匯入成功")
        
        print("\n所有模組匯入成功！程式可以正常執行。")
        return True
        
    except ImportError as e:
        print(f"✗ 匯入失敗：{e}")
        print("請執行：pip install -r requirements.txt")
        return False
    except Exception as e:
        print(f"✗ 其他錯誤：{e}")
        return False

def test_excel_connection():
    """測試 Excel 連線"""
    try:
        import xlwings as xw
        # 只測試是否可以匯入，不實際啟動 Excel
        print("✓ xlwings Excel 連線模組可用")
        return True
    except Exception as e:
        print(f"✗ Excel 連線測試失敗：{e}")
        print("請確保已安裝 Microsoft Excel")
        return False

if __name__ == "__main__":
    print("=== Excel 合併工具 - 模組測試 ===\n")
    
    # 測試模組匯入
    imports_ok = test_imports()
    
    if imports_ok:
        print("\n=== 測試 Excel 連線 ===")
        excel_ok = test_excel_connection()
        
        if excel_ok:
            print("\n🎉 所有測試通過！程式可以正常使用。")
        else:
            print("\n⚠️  Excel 連線有問題，但其他功能正常。")
    else:
        print("\n❌ 模組匯入失敗，請檢查安裝。") 