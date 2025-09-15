using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace _231211A
{
    /// <summary>
    /// 基準庫存快照管理 (SnapshotStock)
    /// 只在使用者載入後建立，不自動更新。
    /// </summary>
    public static class InventoryBaselineManager
    {
        public static DateTime? SnapshotTime { get; private set; }
        public static string? SnapshotSourceFile { get; private set; }
        private static readonly Dictionary<string, double> _snapshotStock = new(StringComparer.OrdinalIgnoreCase);

        public static IReadOnlyDictionary<string, double> SnapshotStock => _snapshotStock;

        /// <summary>
        /// 載入 Excel 第一張工作表：假設 A 欄 = 料號，其餘欄最後一個數字當庫存。
        /// </summary>
        public static void LoadSnapshot(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("找不到基準檔案", filePath);

            Excel.Application? app = null; Excel.Workbook? wb = null; Excel.Worksheet? ws = null; Excel.Range? used = null;
            var temp = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            try
            {
                app = new Excel.Application { Visible = false, DisplayAlerts = false };
                wb = app.Workbooks.Open(filePath);
                ws = (Excel.Worksheet)wb.Worksheets[1];
                used = ws.UsedRange;
                if (used?.Value is object[,] data)
                {
                    int rows = used.Rows.Count;
                    int cols = used.Columns.Count;
                    for (int r = 2; r <= rows; r++)
                    {
                        string part = data[r, 1]?.ToString()?.Trim() ?? string.Empty; // A欄
                        if (string.IsNullOrEmpty(part)) continue;
                        double lastNumeric = 0; bool found = false;
                        for (int c = cols; c >= 2; c--)
                        {
                            if (data[r, c] != null && double.TryParse(data[r, c].ToString(), out double val))
                            {
                                lastNumeric = val; found = true; break;
                            }
                        }
                        if (found) temp[part] = lastNumeric;
                    }
                }
            }
            finally
            {
                try { if (used != null) Marshal.ReleaseComObject(used); } catch { }
                try { if (ws != null) Marshal.ReleaseComObject(ws); } catch { }
                try { if (wb != null) { wb.Close(false); Marshal.ReleaseComObject(wb); } } catch { }
                try { if (app != null) { app.Quit(); Marshal.ReleaseComObject(app); } } catch { }
            }

            _snapshotStock.Clear();
            foreach (var kv in temp) _snapshotStock[kv.Key] = kv.Value;
            SnapshotTime = DateTime.Now;
            SnapshotSourceFile = filePath;
        }

        public static double GetSnapshotQty(string part)
        {
            if (string.IsNullOrWhiteSpace(part)) return 0;
            return _snapshotStock.TryGetValue(part, out var v) ? v : 0;
        }
    }
}
