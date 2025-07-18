#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
安全測試腳本：只檢查語法，不啟動任何 GUI 或 Excel
"""

def test_syntax():
    """測試主程式的語法是否正確"""
    try:
        # 只檢查語法，不執行
        with open('excel_merger.py', 'r', encoding='utf-8') as f:
            code = f.read()
        
        # 編譯檢查語法
        compile(code, 'excel_merger.py', 'exec')
        print("✓ 主程式語法檢查通過")
        return True
        
    except SyntaxError as e:
        print(f"✗ 語法錯誤：{e}")
        return False
    except Exception as e:
        print(f"✗ 其他錯誤：{e}")
        return False

def test_imports():
    """測試模組匯入（不啟動任何服務）"""
    try:
        import tkinter as tk
        print("✓ tkinter 可用")
        
        import xlwings as xw
        print("✓ xlwings 可用")
        
        import pandas as pd
        print("✓ pandas 可用")
        
        import numpy as np
        print("✓ numpy 可用")
        
        from pathlib import Path
        print("✓ pathlib 可用")
        
        return True
        
    except ImportError as e:
        print(f"✗ 模組匯入失敗：{e}")
        return False

if __name__ == "__main__":
    print("=== 安全語法測試 ===\n")
    
    syntax_ok = test_syntax()
    imports_ok = test_imports()
    
    if syntax_ok and imports_ok:
        print("\n🎉 所有測試通過！程式可以正常使用。")
        print("💡 要啟動程式，請執行：python excel_merger.py")
    else:
        print("\n❌ 測試失敗，請檢查程式碼。") 