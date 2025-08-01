using OfficeOpenXml;
using ExcelDataReader;
using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace PCBInventory
{
    public static class ExcelHelper
    {
        static ExcelHelper()
        {
            // 設定 EPPlus 許可證
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // 註冊編碼提供者，以支援舊格式
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public static DataTable ReadExcelFile(string filePath)
        {
            var dataTable = new DataTable();
            string extension = Path.GetExtension(filePath).ToLower();
            
            try
            {
                if (extension == ".xlsx" || extension == ".xlsm")
                {
                    // 使用 EPPlus 讀取新格式
                    return ReadExcelWithEPPlus(filePath);
                }
                else if (extension == ".xls" || extension == ".xlsb")
                {
                    // 使用 ExcelDataReader 讀取舊格式
                    return ReadExcelWithDataReader(filePath);
                }
                else
                {
                    MessageBox.Show($"不支援的檔案格式: {extension}", "錯誤", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"讀取 Excel 檔案時發生錯誤：{ex.Message}", "錯誤", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
            return dataTable;
        }

        private static DataTable ReadExcelWithEPPlus(string filePath)
        {
            var dataTable = new DataTable();
            
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                
                // 獲取使用範圍
                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;
                
                // 添加列標題
                for (int col = start.Column; col <= end.Column; col++)
                {
                    var headerValue = worksheet.Cells[start.Row, col].Value;
                    string columnName = headerValue?.ToString() ?? $"Column{col}";
                    dataTable.Columns.Add(columnName);
                }
                
                // 添加數據行
                for (int row = start.Row + 1; row <= end.Row; row++)
                {
                    var dataRow = dataTable.NewRow();
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value;
                        dataRow[col - start.Column] = cellValue ?? DBNull.Value;
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }
            
            return dataTable;
        }

        private static DataTable ReadExcelWithDataReader(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });
                    
                    if (dataSet.Tables.Count > 0)
                    {
                        return dataSet.Tables[0];
                    }
                }
            }
            
            return new DataTable();
        }
        
        public static void ExportToExcel(DataTable dataTable, string filePath)
        {
            try
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("庫存報表");
                    
                    // 寫入標題
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        worksheet.Cells[1, col + 1].Value = dataTable.Columns[col].ColumnName;
                        worksheet.Cells[1, col + 1].Style.Font.Bold = true;
                    }
                    
                    // 寫入數據
                    for (int row = 0; row < dataTable.Rows.Count; row++)
                    {
                        for (int col = 0; col < dataTable.Columns.Count; col++)
                        {
                            worksheet.Cells[row + 2, col + 1].Value = dataTable.Rows[row][col];
                        }
                    }
                    
                    // 自動調整列寬
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    
                    // 保存檔案
                    var file = new FileInfo(filePath);
                    package.SaveAs(file);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"匯出 Excel 檔案時發生錯誤：{ex.Message}", "錯誤", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
