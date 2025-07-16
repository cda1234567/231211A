#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
測試路徑建立功能
"""

import os
import datetime
from pathlib import Path

def test_path_creation():
    """測試路徑建立功能"""
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
    
    print("=== 路徑建立測試 ===\n")
    
    # 預設網路路徑
    default_network_path = Path(f"\\\\St-nas\\個人資料夾\\Andy\\excel\\{timestamp}")
    print(f"預設網路路徑：{default_network_path}")
    
    # 備援本地路徑
    local_folder = Path(f"./output_{timestamp}")
    print(f"備援本地路徑：{local_folder}")
    
    # 測試網路路徑
    try:
        print("\n嘗試建立網路路徑...")
        # 設定超時，避免卡住
        import signal
        
        def timeout_handler(signum, frame):
            raise TimeoutError("網路路徑連線超時")
        
        # 在 Windows 上使用不同的超時方法
        import threading
        import time
        
        result = [None]
        exception = [None]
        
        def try_create_path():
            try:
                default_network_path.mkdir(parents=True, exist_ok=True)
                result[0] = True
            except Exception as e:
                exception[0] = e
        
        thread = threading.Thread(target=try_create_path)
        thread.daemon = True
        thread.start()
        thread.join(timeout=5)  # 5秒超時
        
        if thread.is_alive():
            print("✗ 網路路徑連線超時")
            raise TimeoutError("網路路徑連線超時")
        
        if exception[0]:
            raise exception[0]
        
        print("✓ 網路路徑建立成功")
        
        # 檢查是否真的可以寫入
        test_file = default_network_path / "test.txt"
        test_file.write_text("測試檔案")
        print("✓ 網路路徑可以寫入檔案")
        
        # 清理測試檔案
        test_file.unlink()
        print("✓ 測試檔案已清理")
        
        return True
        
    except Exception as e:
        print(f"✗ 網路路徑建立失敗：{e}")
        
        # 測試本地路徑
        try:
            print("\n嘗試建立本地路徑...")
            local_folder.mkdir(parents=True, exist_ok=True)
            print("✓ 本地路徑建立成功")
            
            # 檢查是否真的可以寫入
            test_file = local_folder / "test.txt"
            test_file.write_text("測試檔案")
            print("✓ 本地路徑可以寫入檔案")
            
            # 清理測試檔案和資料夾
            test_file.unlink()
            local_folder.rmdir()
            print("✓ 測試檔案和資料夾已清理")
            
            return True
            
        except Exception as e2:
            print(f"✗ 本地路徑也建立失敗：{e2}")
            return False

if __name__ == "__main__":
    success = test_path_creation()
    
    if success:
        print("\n🎉 路徑測試通過！程式可以正常使用。")
    else:
        print("\n❌ 路徑測試失敗，請檢查權限設定。") 