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

        // 依副檔選取順序逐一處理，僅在本檔由非負→負時詢問補料；其他行自動把 H 補 0（空才補）
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

        // 依第一個字母(A→Z)排序詢問補料順序
        private static int FirstLetterRank(string part)
        {
            if (string.IsNullOrEmpty(part)) return int.MaxValue;
            char ch = char.ToUpperInvariant(part[0]);
            if (ch >= 'A' && ch <= 'Z') return ch - 'A';
            return 26 + ch; // 非英文字母排在最後
        }

        // 判斷是否為各種「-」符號
        private static bool IsDashLike(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            string t = s.Trim();
            return t == "-" || t == "－" || t == "?" || t == "—" || t == "–";
        }

        // 處理單一副檔：僅對本檔由非負→負的料號詢問補料；其他列自動把 H 補 0（空才補）
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

                // 先掃描本檔，收集候選與非候選
                var candidates = new List<(int Row, string Part, double G, double H, double J)>();
                var nonCandidates = new List<int>();

                for (int r = 2; r <= rows; r++)
                {
                    string part = cols >= 3 ? (data[r, 3]?.ToString()?.Trim() ?? string.Empty) : string.Empty; // C欄 料號
                    if (string.IsNullOrEmpty(part)) { nonCandidates.Add(r); continue; }
                    if (_processedPartNumbers.Contains(part)) { nonCandidates.Add(r); continue; }

                    string gText = cols >= 7 ? data[r, 7]?.ToString() : null;
                    string hText = cols >= 8 ? data[r, 8]?.ToString() : null;

                    // 若 G 或 H 為 dash-like，直接標記為非候選，且後續不要改動 H（保留其 "-"）
                    if (IsDashLike(gText) || IsDashLike(hText)) { nonCandidates.Add(r); continue; }

                    double g = 0, h = double.NaN, j = 0;
                    if (cols >= 7 && data[r, 7] != null) double.TryParse(data[r, 7].ToString(), out g);   // G
                    if (cols >= 8 && data[r, 8] != null) double.TryParse(data[r, 8].ToString(), out h);   // H
                    if (cols >= 10 && data[r, 10] != null) double.TryParse(data[r, 10].ToString(), out j); // J

                    // 本檔才轉負：G >= 0 且 J < 0
                    if (g >= 0 && j < 0)
                    {
                        candidates.Add((r, part, g, h, j));
                    }
                    else
                    {
                        nonCandidates.Add(r);
                    }
                }

                // 對非候選行：若 H 空白或非數值，補寫 0（不動 J）；但若 H 為 dash-like 則保留不動；且僅對有料號(C欄非空)的列做
                foreach (int r in nonCandidates)
                {
                    var partText = ws.Cells[r, 3].Value?.ToString()?.Trim();
                    if (string.IsNullOrEmpty(partText))
                    {
                        // 無料號的行完全不動
                        continue;
                    }

                    var hCellObj = ws.Cells[r, 8].Value;
                    string hText = hCellObj?.ToString();
                    if (IsDashLike(hText))
                    {
                        // 保留 "-"，不動
                        continue;
                    }

                    bool hasNumeric = false;
                    if (hCellObj != null)
                    {
                        double tmp; if (double.TryParse(hCellObj.ToString(), out tmp)) hasNumeric = true;
                    }
                    if (!hasNumeric)
                    {
                        var hCell = (Excel.Range)ws.Cells[r, 8];
                        hCell.Value = 0;
                    }
                }

                // 候選排序：首字 A→Z，再整個料號
                var ordered = candidates
                    .OrderBy(c => FirstLetterRank(c.Part))
                    .ThenBy(c => c.Part, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                progressBar.Value = 0;
                progressBar.Maximum = ordered.Count;

                int processed = 0;
                foreach (var cand in ordered)
                {
                    int r = cand.Row;
                    string part = cand.Part;
                    double j = cand.J;

                    if (_processedPartNumbers.Contains(part)) continue;

                    processed++;
                    progressBar.Value = processed;

                    string description = cols >= 4 ? (ws.Cells[r, 4].Value?.ToString()?.Trim() ?? string.Empty) : string.Empty;

                    // 對話框顯示以主檔最終值為準（讀不到則退回本檔 J）
                    double finalNeg = finalStockByPart.TryGetValue(part, out var finalVal) ? finalVal : j;

                    var item = new ReplenishmentItem
                    {
                        Row = r,
                        PartNumber = part,
                        Description = description,
                        CurrentStock = finalNeg,
                        TargetColumn = 8
                    };

                    statusLabel.Text = $"處理發料({processed}/{ordered.Count})：{part}";
                    Application.DoEvents();

                    var owner = statusLabel?.FindForm();
                    var result = ShowDispatchDialog(item, fileIndex, owner);
                    if (result.DialogResult == DialogResult.Cancel)
                    {
                        statusLabel.Text = "發料處理已取消";
                        return;
                    }
                    if (result.DialogResult == DialogResult.Ignore)
                    {
                        // 使用者略過：補 H=0（若尚未是數值且不是 dash-like）；僅對有料號的列
                        var partText2 = ws.Cells[r, 3].Value?.ToString()?.Trim();
                        if (!string.IsNullOrEmpty(partText2))
                        {
                            var hCellObj2 = ws.Cells[r, 8].Value;
                            string hText2 = hCellObj2?.ToString();
                            if (!IsDashLike(hText2))
                            {
                                bool hasNumeric2 = false;
                                if (hCellObj2 != null)
                                {
                                    double tmp2; if (double.TryParse(hCellObj2.ToString(), out tmp2)) hasNumeric2 = true;
                                }
                                if (!hasNumeric2)
                                {
                                    var hCell = (Excel.Range)ws.Cells[r, 8];
                                    hCell.Value = 0;
                                }
                            }
                        }
                        _processedPartNumbers.Add(part);
                        continue;
                    }
                    if (result.DialogResult != DialogResult.OK) continue;

                    double qty = Math.Max(0, result.ReplenishmentQuantity);

                    // 寫回 H（補料量）；J = 原 J + 補料（不做任何底色或粗體/格式變更）
                    double adjusted = j + qty;

                    ws.Cells[r, 8].Value = qty;                        // H
                    ws.Cells[r, 10].Value = (int)Math.Round(adjusted); // J

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

        private static ReplenishmentDialogResult ShowDispatchDialog(ReplenishmentItem item, int index, IWin32Window? owner)
        {
            using var dialog = new ReplenishmentDialog(item, index);
            var dr = owner != null ? dialog.ShowDialog(owner) : dialog.ShowDialog();
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