using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace _231211A
{
    public class ReplenishmentItem
    {
        public int Row { get; set; }
        public string PartNumber { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public double CurrentStock { get; set; } // 主檔最終庫存（可能為負）
        public int TargetColumn { get; set; }
        public double ShortageAmount => Math.Abs(CurrentStock);
    }

    public class ReplenishmentDialogResult
    {
        public DialogResult DialogResult { get; set; }
        public double ReplenishmentQuantity { get; set; }
    }

    public static class ReplenishmentManager
    {
        private static readonly HashSet<string> _processedPartNumbers = new(StringComparer.OrdinalIgnoreCase);
        public static void ClearProcessedPartNumbers() => _processedPartNumbers.Clear();

        // 依副檔選取順序逐一處理，針對主檔最終庫存為負的料彈窗並回寫到該副檔 H/J；同時記錄到 ExcelMergerApi
        public static void ProcessMainFileNegativeInventory(string mainFilePath, string outputFolder, ProgressBar progressBar, Label statusLabel)
        {
            if (string.IsNullOrWhiteSpace(mainFilePath) || !File.Exists(mainFilePath))
            {
                statusLabel.Text = "主檔案不存在";
                return;
            }
            if (string.IsNullOrWhiteSpace(outputFolder) || !Directory.Exists(outputFolder))
            {
                statusLabel.Text = "輸出資料夾不存在";
                return;
            }

            try
            {
                statusLabel.Text = "讀取主檔最終庫存...";
                Application.DoEvents();
                var finalStockByPart = ReadMainFinalStock(mainFilePath);

                // 讀取第一次合併時保存的副檔順序（與選檔順序一致）
                var secondaryFiles = GetOrderedSecondaryFiles(outputFolder);
                if (secondaryFiles.Length == 0)
                {
                    statusLabel.Text = "找不到副檔";
                    return;
                }

                int fileIndex = 0;
                foreach (var filePath in secondaryFiles)
                {
                    fileIndex++;
                    statusLabel.Text = $"處理副檔({fileIndex}/{secondaryFiles.Length}): {Path.GetFileName(filePath)}";
                    Application.DoEvents();

                    ProcessSingleSecondaryFile(filePath, finalStockByPart, fileIndex, progressBar, statusLabel);
                    Application.DoEvents();
                }

                statusLabel.Text = "補料完成，請再執行一次扣帳以套用至主檔";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"發料處理錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                statusLabel.Text = "發料處理失敗";
            }
        }

        // 讀取主檔每個料號的最終庫存（最右側第一個數值）
        private static Dictionary<string, double> ReadMainFinalStock(string mainFilePath)
        {
            var map = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

            Excel.Application? app = null; Excel.Workbook? wb = null; Excel.Worksheet? ws = null; Excel.Range? used = null;
            try
            {
                app = new Excel.Application { Visible = false, DisplayAlerts = false };
                wb = app.Workbooks.Open(mainFilePath);
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
                        for (int c = cols; c >= 1; c--)
                        {
                            if (data[r, c] != null && double.TryParse(data[r, c].ToString(), out double val))
                            {
                                lastNumeric = val; found = true; break;
                            }
                        }
                        if (found) map[part] = lastNumeric;
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

            return map;
        }

        // 讀取 __order.txt 以維持與選檔一致的副檔順序
        private static string[] GetOrderedSecondaryFiles(string outputFolder)
        {
            var manifest = Path.Combine(outputFolder, "__order.txt");
            var files = new List<string>();
            if (File.Exists(manifest))
            {
                foreach (var line in File.ReadAllLines(manifest))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var fullPath = Path.Combine(outputFolder, line.Trim());
                    if (File.Exists(fullPath)) files.Add(fullPath);
                }
            }

            if (files.Count == 0)
            {
                // 後備：以建立時間排序（但優先使用 manifest）
                files = Directory.GetFiles(outputFolder, "*.xls*")
                    .Where(f => !Path.GetFileName(f).Contains("_main", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(f => File.GetCreationTime(f))
                    .ToList();
            }

            return files.ToArray();
        }

        // 處理單一副檔：找出本檔出現之料號，若其主檔最終庫存為負，彈窗並把補料數量寫回本檔 H/J，同時記錄到 API
        private static void ProcessSingleSecondaryFile(string filePath, Dictionary<string, double> finalStockByPart, int fileIndex, ProgressBar progressBar, Label statusLabel)
        {
            Excel.Application? app = null; Excel.Workbook? wb = null; Excel.Worksheet? ws = null; Excel.Range? used = null;

            try
            {
                app = new Excel.Application { Visible = false, DisplayAlerts = false };
                wb = app.Workbooks.Open(filePath);
                ws = (Excel.Worksheet)wb.Worksheets[1];
                used = ws.UsedRange;

                if (used?.Value is not object[,] data)
                    return;

                int rows = used.Rows.Count;
                int cols = used.Columns.Count;

                var candidateRows = new List<int>();
                for (int r = 2; r <= rows; r++)
                {
                    string part = cols >= 3 ? (data[r, 3]?.ToString()?.Trim() ?? string.Empty) : string.Empty; // C欄
                    if (string.IsNullOrEmpty(part)) continue;
                    if (!finalStockByPart.TryGetValue(part, out double finalStock)) continue;
                    if (finalStock < 0 && !_processedPartNumbers.Contains(part))
                        candidateRows.Add(r);
                }

                progressBar.Value = 0;
                progressBar.Maximum = candidateRows.Count;

                int processed = 0;
                foreach (int r in candidateRows)
                {
                    processed++;
                    progressBar.Value = processed;

                    string part = ws.Cells[r, 3].Value?.ToString()?.Trim() ?? string.Empty;
                    if (string.IsNullOrEmpty(part)) continue;
                    if (_processedPartNumbers.Contains(part)) continue;

                    double finalStock = finalStockByPart[part]; // 主檔最終庫存（負）
                    string description = cols >= 4 ? (ws.Cells[r, 4].Value?.ToString()?.Trim() ?? string.Empty) : string.Empty;

                    var item = new ReplenishmentItem
                    {
                        Row = r,
                        PartNumber = part,
                        Description = description,
                        CurrentStock = finalStock,
                        TargetColumn = 8
                    };

                    statusLabel.Text = $"處理發料({processed}/{candidateRows.Count})：{part}";
                    Application.DoEvents();

                    var result = ShowDispatchDialog(item, fileIndex);
                    if (result.DialogResult == DialogResult.Cancel)
                    {
                        statusLabel.Text = "發料處理已取消";
                        return;
                    }
                    if (result.DialogResult == DialogResult.Ignore)
                    {
                        _processedPartNumbers.Add(part);
                        continue;
                    }
                    if (result.DialogResult != DialogResult.OK) continue;

                    double qty = Math.Max(0, result.ReplenishmentQuantity);

                    // 寫回 H（發料）與 J（期初+發料）
                    double originalJ = 0;
                    if (cols >= 10 && ws.Cells[r, 10].Value != null)
                        double.TryParse(ws.Cells[r, 10].Value.ToString(), out originalJ);
                    double adjusted = originalJ + qty;

                    ws.Cells[r, 8].Value = qty;                                // H
                    ws.Cells[r, 10].Value = (int)Math.Round(adjusted);         // J

                    var hCell = (Excel.Range)ws.Cells[r, 8];
                    hCell.Font.Bold = true;
                    hCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);

                    var jCell = (Excel.Range)ws.Cells[r, 10];
                    jCell.Interior.Color = adjusted >= 0
                        ? System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen)
                        : System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightPink);

                    // 記錄到 API，供第二次扣帳用
                    ExcelMergerApi.SetDispatchData("main", part, qty);

                    _processedPartNumbers.Add(part);
                    statusLabel.Text = $"已記錄補料：{part} 數量：{qty}";
                    Application.DoEvents();
                }

                wb.Save();
            }
            finally
            {
                try { if (used != null) Marshal.ReleaseComObject(used); } catch { }
                try { if (ws != null) Marshal.ReleaseComObject(ws); } catch { }
                try { if (wb != null) { wb.Close(false); Marshal.ReleaseComObject(wb); } } catch { }
                try { if (app != null) { app.Quit(); Marshal.ReleaseComObject(app); } } catch { }
            }
        }

        private static ReplenishmentDialogResult ShowDispatchDialog(ReplenishmentItem item, int index)
        {
            using var dialog = new ReplenishmentDialog(item, index);
            var dr = dialog.ShowDialog();
            return new ReplenishmentDialogResult
            {
                DialogResult = dr,
                ReplenishmentQuantity = dialog.ReplenishmentQuantity
            };
        }

        public static void ProcessNegativeInventory(string mainFilePath, ProgressBar progressBar, Label statusLabel)
        {
            statusLabel.Text = "請使用新的 ProcessMainFileNegativeInventory";
        }
    }
}