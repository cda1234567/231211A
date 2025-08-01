using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace PCBInventory
{
    public static class ExcelMergerApiNew
    {
        static ExcelMergerApiNew()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public static void MergeFiles(ListBox listBoxFiles, ProgressBar progressBar1, Label labelCurrentFile)
        {
            if (listBoxFiles.Items.Count < 2)
            {
                MessageBox.Show("請選擇至少兩個檔案");
                return;
            }

            string mainFileName = listBoxFiles.Items[0].ToString() ?? string.Empty;
            List<string> secondaryFileNames = new List<string>();
            for (int i = 1; i < listBoxFiles.Items.Count; i++)
                secondaryFileNames.Add(listBoxFiles.Items[i].ToString() ?? string.Empty);

            try
            {
                // 創建輸出資料夾
                string folderPath = @"\\St-nas\個人資料夾\Andy\excel\1\temp\" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm");
                Directory.CreateDirectory(folderPath);

                // 讀取主檔案
                var mainData = ReadExcelData(mainFileName);
                if (mainData == null || mainData.Count == 0)
                {
                    MessageBox.Show("無法讀取主檔案或主檔案為空");
                    return;
                }

                // 計算總進度
                int totalFiles = secondaryFileNames.Count;
                progressBar1.Maximum = totalFiles;
                progressBar1.Value = 0;

                var fileNameCount = new Dictionary<string, int>();

                for (int fileIndex = 0; fileIndex < secondaryFileNames.Count; fileIndex++)
                {
                    string secondaryFileName = secondaryFileNames[fileIndex];
                    labelCurrentFile.Text = $"目前執行到的檔案：{Path.GetFileName(secondaryFileName)} ({fileIndex + 1}/{totalFiles})";
                    Application.DoEvents();

                    if (!File.Exists(secondaryFileName))
                    {
                        MessageBox.Show($"檔案不存在: {secondaryFileName}");
                        continue;
                    }

                    try
                    {
                        // 讀取次要檔案
                        var secondaryData = ReadExcelData(secondaryFileName);
                        if (secondaryData == null || secondaryData.Count == 0)
                        {
                            continue;
                        }

                        // 處理合併邏輯 (簡化版本)
                        ProcessMerge(mainData, secondaryData);

                        // 儲存處理後的檔案
                        string baseName = Path.GetFileNameWithoutExtension(secondaryFileName);
                        string ext = ".xlsx"; // 統一使用 .xlsx
                        string saveName = baseName + ext;

                        if (!fileNameCount.ContainsKey(baseName))
                            fileNameCount[baseName] = 0;
                        fileNameCount[baseName]++;
                        if (fileNameCount[baseName] > 1)
                            saveName = $"{baseName}-{fileNameCount[baseName]}{ext}";

                        string secondarySavePath = Path.Combine(folderPath, saveName);
                        SaveExcelData(secondaryData, secondarySavePath);

                        // 更新進度
                        progressBar1.Value = fileIndex + 1;
                        Application.DoEvents();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"處理檔案 {Path.GetFileName(secondaryFileName)} 時發生錯誤：{ex.Message}");
                    }
                }

                // 儲存主檔案
                string mainBaseName = Path.GetFileNameWithoutExtension(mainFileName);
                string mainSaveName = $"{mainBaseName}_main.xlsx";
                string mainSavePath = Path.Combine(folderPath, mainSaveName);
                SaveExcelData(mainData, mainSavePath);

                labelCurrentFile.Text = "處理完成";
                MessageBox.Show($"處理完成！\n檔案已儲存到：{folderPath}");
                
                // 開啟結果資料夾
                System.Diagnostics.Process.Start("explorer.exe", folderPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"錯誤：{ex.Message}\n{ex.StackTrace}");
            }
        }

        private static List<List<object?>> ReadExcelData(string filePath)
        {
            var data = new List<List<object?>>();
            string extension = Path.GetExtension(filePath).ToLower();
            
            try
            {
                if (extension == ".xlsx" || extension == ".xlsm")
                {
                    // 使用 EPPlus 讀取新格式
                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        var dimension = worksheet.Dimension;
                        
                        if (dimension == null) return data;

                        for (int row = dimension.Start.Row; row <= dimension.End.Row; row++)
                        {
                            var rowData = new List<object?>();
                            for (int col = dimension.Start.Column; col <= dimension.End.Column; col++)
                            {
                                rowData.Add(worksheet.Cells[row, col].Value);
                            }
                            data.Add(rowData);
                        }
                    }
                }
                else if (extension == ".xls" || extension == ".xlsb")
                {
                    // 使用 ExcelDataReader 讀取舊格式
                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream))
                        {
                            do
                            {
                                while (reader.Read())
                                {
                                    var rowData = new List<object?>();
                                    for (int col = 0; col < reader.FieldCount; col++)
                                    {
                                        rowData.Add(reader.GetValue(col));
                                    }
                                    data.Add(rowData);
                                }
                            } while (reader.NextResult());
                        }
                    }
                }
                else
                {
                    MessageBox.Show($"不支援的檔案格式: {extension}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"讀取檔案 {Path.GetFileName(filePath)} 時發生錯誤：{ex.Message}");
            }
            
            return data;
        }

        private static void SaveExcelData(List<List<object?>> data, string filePath)
        {
            try
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    
                    for (int row = 0; row < data.Count; row++)
                    {
                        for (int col = 0; col < data[row].Count; col++)
                        {
                            worksheet.Cells[row + 1, col + 1].Value = data[row][col];
                        }
                    }
                    
                    // 自動調整列寬
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    
                    var file = new FileInfo(filePath);
                    package.SaveAs(file);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"儲存檔案 {Path.GetFileName(filePath)} 時發生錯誤：{ex.Message}");
            }
        }

        private static void ProcessMerge(List<List<object?>> mainData, List<List<object?>> secondaryData)
        {
            // 這裡實現簡化版的合併邏輯
            // 原本的合併邏輯相當複雜，這裡提供一個基本的框架
            // 您可以根據具體需求來調整這個方法

            try
            {
                // 基本的資料合併示例：
                // 假設第3欄是比對的關鍵字段
                for (int secRow = 1; secRow < secondaryData.Count; secRow++) // 跳過標題行
                {
                    if (secondaryData[secRow].Count > 2)
                    {
                        string? secKey = secondaryData[secRow][2]?.ToString()?.Trim();
                        
                        // 在主檔案中尋找相符的資料
                        for (int mainRow = 1; mainRow < mainData.Count; mainRow++)
                        {
                            if (mainData[mainRow].Count > 0)
                            {
                                string? mainKey = mainData[mainRow][0]?.ToString()?.Trim();
                                
                                if (!string.IsNullOrEmpty(secKey) && !string.IsNullOrEmpty(mainKey) && 
                                    string.Equals(secKey, mainKey, StringComparison.OrdinalIgnoreCase))
                                {
                                    // 找到相符的資料，進行合併處理
                                    // 這裡添加您的合併邏輯
                                    
                                    // 確保行有足夠的欄位
                                    while (mainData[mainRow].Count < mainData[0].Count + 3)
                                    {
                                        mainData[mainRow].Add(null);
                                    }
                                    
                                    // 簡單的資料複製示例
                                    if (secondaryData[secRow].Count > 5)
                                    {
                                        // 複製一些數據到主檔案的新欄位
                                        int targetCol = mainData[mainRow].Count - 3;
                                        if (targetCol >= 0)
                                        {
                                            mainData[mainRow][targetCol] = secondaryData[secRow][5]; // 數量
                                            if (secondaryData[secRow].Count > 6)
                                                mainData[mainRow][targetCol + 1] = secondaryData[secRow][6]; // 其他數據
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"合併處理時發生錯誤：{ex.Message}");
            }
        }
    }
}
