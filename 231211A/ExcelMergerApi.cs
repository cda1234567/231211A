using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace _231211A
{
    public static class ExcelMergerApi
    {
        public static string? CustomOutputBaseFolder { get; set; }

        // Track dispatch data (unchanged existing behavior)
        private static Dictionary<string, Dictionary<string, double>> _dispatchData = new();

        // New: detect first vs second merge in current process run
        private static bool _firstMergeDone = false;

        // New: aggregation data only gathered on second merge
        private class SummaryDetail
        {
            public string Part = string.Empty;
            public string Desc = string.Empty;
            public int Qty;
        }
        private static readonly List<(string FileName, List<SummaryDetail> Details)> _summaryData = new();

        public static string MergeFiles(ListBox listBoxFiles, ProgressBar progressBar1, Label labelCurrentFile)
        {
            bool isSecondMerge = _firstMergeDone; // second (or later) merge triggers summary
            if (!_firstMergeDone) _firstMergeDone = true; // mark after first call

            if (listBoxFiles.Items.Count < 2)
            {
                MessageBox.Show("請選擇至少兩個檔案");
                return string.Empty;
            }

            if (isSecondMerge)
                _summaryData.Clear(); // prepare for aggregation

            string mainFileName = listBoxFiles.Items[0].ToString();
            var secondaryFileNames = new List<string>();
            for (int i = 1; i < listBoxFiles.Items.Count; i++)
                secondaryFileNames.Add(listBoxFiles.Items[i].ToString());

            Excel.Application? excelApp = null;
            Excel.Workbook? mainWorkbook = null;
            var workbooks = new List<Excel.Workbook>();
            string folderPath = string.Empty;

            var orderedSavedSecondaryFiles = new List<string>();
            var dispatchedOnce = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var fileNameCount = new Dictionary<string, int>();

            try
            {
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;

                mainWorkbook = excelApp.Workbooks.Open(mainFileName);
                Excel.Worksheet mainWorksheet = mainWorkbook.Worksheets[1];
                Excel.Range mainExcelRange = mainWorksheet.UsedRange;
                object[,] mainDataArray = mainExcelRange.Value;

                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string baseFolder = CustomOutputBaseFolder;
                if (string.IsNullOrWhiteSpace(baseFolder))
                {
                    baseFolder = @"\\\St-nas\個人資料夾\Andy\excel";
                }
                else
                {
                    try { if (!Directory.Exists(baseFolder)) Directory.CreateDirectory(baseFolder); }
                    catch { baseFolder = @"\\\St-nas\個人資料夾\Andy\excel"; }
                }
                folderPath = Path.Combine(baseFolder, $"{Path.GetFileNameWithoutExtension(mainFileName)}_{timestamp}");
                Directory.CreateDirectory(folderPath);

                // Pre-calc total rows for progress
                int totalRows = 0;
                foreach (var secondaryFileName in secondaryFileNames)
                {
                    if (!File.Exists(secondaryFileName)) continue;
                    var workbookTmp = excelApp.Workbooks.Open(secondaryFileName);
                    Excel.Worksheet worksheetTmp = workbookTmp.Worksheets[1];
                    Excel.Range excelRangeTmp = worksheetTmp.UsedRange;
                    totalRows += excelRangeTmp.Rows.Count;
                    workbookTmp.Close(false);
                    Marshal.ReleaseComObject(worksheetTmp);
                    Marshal.ReleaseComObject(excelRangeTmp);
                }
                progressBar1.Maximum = Math.Max(totalRows, 1);
                progressBar1.Value = 0;
                int progressValue = 0;

                foreach (var secondaryFileName in secondaryFileNames)
                {
                    labelCurrentFile.Text = $"目前執行到的檔案：{Path.GetFileName(secondaryFileName)}";
                    Application.DoEvents();

                    if (!File.Exists(secondaryFileName))
                    {
                        MessageBox.Show($"檔案不存在: {secondaryFileName}");
                        continue;
                    }

                    var workbook = excelApp.Workbooks.Open(secondaryFileName);
                    workbooks.Add(workbook);
                    Excel.Worksheet worksheet = workbook.Worksheets[1];
                    Excel.Range excelRange = worksheet.UsedRange;
                    object[,] dataArray = excelRange.Value;

                    int lastRowMain = mainWorksheet.UsedRange.Rows.Count;
                    int lastRowSec = excelRange.Rows.Count;

                    int baseCol = ((Excel.Range)mainWorksheet.Cells[1, mainWorksheet.Columns.Count])
                        .get_End(Excel.XlDirection.xlToLeft).Column;

                    // 標題：不寫 Dispatch，產品欄自動換行，且不填任何底色
                    string orderRaw = null;
                    try { orderRaw = dataArray[1, 7]?.ToString(); } catch { orderRaw = null; }
                    if (string.IsNullOrEmpty(orderRaw))
                        try { orderRaw = dataArray[1, 8]?.ToString(); } catch { orderRaw = null; }
                    string orderLast8 = string.IsNullOrEmpty(orderRaw) ? string.Empty : (orderRaw.Length >= 8 ? orderRaw[^8..] : orderRaw);
                    string productName = string.Empty;
                    try { productName = dataArray[2, 3]?.ToString() ?? string.Empty; } catch { productName = string.Empty; }

                    mainWorksheet.Cells[1, baseCol + 1].Value = string.Empty; // Dispatch column empty
                    mainWorksheet.Cells[1, baseCol + 2].Value = orderLast8;
                    mainWorksheet.Cells[1, baseCol + 3].Value = productName;
                    var hdr1 = (Excel.Range)mainWorksheet.Cells[1, baseCol + 1];
                    var hdr2 = (Excel.Range)mainWorksheet.Cells[1, baseCol + 2];
                    var hdr3 = (Excel.Range)mainWorksheet.Cells[1, baseCol + 3];
                    hdr1.Font.Name = hdr2.Font.Name = hdr3.Font.Name = "Arial";
                    hdr1.Font.Size = hdr2.Font.Size = hdr3.Font.Size = 9;
                    hdr3.WrapText = true; hdr3.EntireColumn.WrapText = true; hdr3.EntireRow.AutoFit();
                    // 清除標題底色
                    ClearCellFill(hdr1); ClearCellFill(hdr2); ClearCellFill(hdr3);

                    for (int j = 2; j <= lastRowSec; j++)
                    {
                        bool skipRow = false;
                        if (dataArray.GetLength(1) >= 7 && dataArray[j, 7] != null && IsDashLike(dataArray[j, 7].ToString())) skipRow = true;
                        if (dataArray.GetLength(1) >= 8 && dataArray[j, 8] != null && IsDashLike(dataArray[j, 8].ToString())) skipRow = true;
                        if (skipRow)
                        {
                            progressValue++;
                            UpdateProgressBar(progressBar1, labelCurrentFile, Path.GetFileName(secondaryFileName), j, lastRowSec, progressValue);
                            continue;
                        }

                        string secPart = dataArray[j, 3]?.ToString()?.Trim() ?? string.Empty; // C column part number
                        if (string.IsNullOrEmpty(secPart))
                        {
                            progressValue++;
                            UpdateProgressBar(progressBar1, labelCurrentFile, Path.GetFileName(secondaryFileName), j, lastRowSec, progressValue);
                            continue;
                        }

                        int mainRowIndex = -1;
                        for (int k = 2; k <= lastRowMain; k++)
                        {
                            if (string.Equals(secPart, mainDataArray[k, 1]?.ToString()?.Trim(), StringComparison.OrdinalIgnoreCase))
                            { mainRowIndex = k; break; }
                        }
                        if (mainRowIndex == -1)
                        {
                            progressValue++;
                            UpdateProgressBar(progressBar1, labelCurrentFile, Path.GetFileName(secondaryFileName), j, lastRowSec, progressValue);
                            continue;
                        }

                        double f2 = 0;
                        if (dataArray.GetLength(1) >= 6 && dataArray[j, 6] != null && double.TryParse(dataArray[j, 6].ToString(), out double tmpF2))
                            f2 = Math.Round(tmpF2, MidpointRounding.AwayFromZero);
                        var midCell = (Excel.Range)mainWorksheet.Cells[mainRowIndex, baseCol + 2];
                        midCell.Value = f2;
                        ClearCellFill(midCell);

                        int prevFinal = FindLastNonEmptyColumnValueInRow(mainDataArray, mainRowIndex);
                        var gCell = (Excel.Range)worksheet.Cells[j, 7];
                        if (prevFinal != 0)
                        {
                            gCell.Value = prevFinal;
                            ApplySecondaryCellStyle(gCell);
                        }
                        else
                        {
                            object jRaw = worksheet.Cells[j, 10]?.Value;
                            if (jRaw != null && !string.IsNullOrWhiteSpace(jRaw.ToString()))
                            {
                                gCell.Value = 0;
                                ApplySecondaryCellStyle(gCell);
                            }
                        }
                        ClearCellFill(gCell);

                        double dispatchQtyPreset = GetDispatchQuantity("main", secPart);
                        if (dispatchQtyPreset > 0 && !dispatchedOnce.Contains(secPart))
                        {
                            var dispatchCell = (Excel.Range)mainWorksheet.Cells[mainRowIndex, baseCol + 1];
                            dispatchCell.Value = dispatchQtyPreset.ToString("F0");
                            ClearCellFill(dispatchCell);
                            dispatchedOnce.Add(secPart);
                        }

                        double jValue = 0;
                        if (worksheet.Cells[j, 10].Value != null)
                            double.TryParse(worksheet.Cells[j, 10].Value.ToString(), out jValue);
                        int finalRounded = (int)Math.Round(jValue, MidpointRounding.AwayFromZero);
                        var outCell = (Excel.Range)mainWorksheet.Cells[mainRowIndex, baseCol + 3];
                        outCell.Value = finalRounded;
                        ClearCellFill(outCell);
                        if (finalRounded < 0) ApplyNegativeFill(outCell);

                        progressValue++;
                        UpdateProgressBar(progressBar1, labelCurrentFile, Path.GetFileName(secondaryFileName), j, lastRowSec, progressValue);
                        Application.DoEvents();
                    }

                    // Post process secondary workbook & collect summary if needed
                    List<SummaryDetail>? collected = null;
                    if (isSecondMerge) collected = new List<SummaryDetail>();
                    try { PostProcessSecondaryWorkbook(workbook, collected); } catch { }
                    if (isSecondMerge && collected != null && collected.Count > 0)
                        _summaryData.Add((Path.GetFileName(secondaryFileName), collected));

                    string baseName = Path.GetFileNameWithoutExtension(secondaryFileName);
                    string ext = Path.GetExtension(secondaryFileName);
                    string saveName = baseName + ext;
                    if (!fileNameCount.ContainsKey(baseName)) fileNameCount[baseName] = 0;
                    fileNameCount[baseName]++;
                    if (fileNameCount[baseName] > 1) saveName = $"{baseName}-{fileNameCount[baseName]}{ext}";

                    string secondarySavePath = Path.Combine(folderPath, saveName);
                    workbook.SaveAs(secondarySavePath);
                    workbook.Close();
                    orderedSavedSecondaryFiles.Add(Path.GetFileName(secondarySavePath));

                    mainExcelRange = mainWorksheet.UsedRange;
                    mainDataArray = mainExcelRange.Value;
                }

                try { File.WriteAllLines(Path.Combine(folderPath, "__order.txt"), orderedSavedSecondaryFiles); } catch { }

                string mainSaveName = $"{Path.GetFileNameWithoutExtension(mainFileName)}_main{Path.GetExtension(mainFileName)}";
                string mainSavePath = Path.Combine(folderPath, mainSaveName);
                mainWorkbook.SaveAs(mainSavePath);

                // Write summary workbook only after second merge
                if (isSecondMerge && _summaryData.Count > 0)
                {
                    WriteSummaryWorkbook(excelApp, folderPath);
                }

                Thread.Sleep(200);
                mainWorkbook.Close();
                excelApp.Quit();
            }
            catch (FileNotFoundException fnfEx)
            {
                MessageBox.Show($"檔案找不到: {fnfEx.FileName}");
            }
            catch (COMException comEx)
            {
                MessageBox.Show($"Excel COM 錯誤: {comEx.Message}");
            }
            catch (Exception ex)
            {
                ExecuteCmdCommand("taskkill /f /im excel.exe");
                MessageBox.Show($"錯誤：{ex.Message}\n堆疊追蹤：{ex.StackTrace}");
            }
            finally
            {
                try { if (mainWorkbook != null) Marshal.ReleaseComObject(mainWorkbook); } catch { }
                foreach (var wb in workbooks) { try { if (wb != null) Marshal.ReleaseComObject(wb); } catch { } }
                try { if (excelApp != null) Marshal.ReleaseComObject(excelApp); } catch { }
            }
            return folderPath;
        }

        // Modified to optionally collect summary details (part, desc, qty)
        private static void PostProcessSecondaryWorkbook(Excel.Workbook workbook, List<SummaryDetail>? collect)
        {
            if (workbook == null) return;
            // Remove sheets after first
            try
            {
                for (int i = workbook.Worksheets.Count; i >= 2; i--)
                {
                    var wsDel = (Excel.Worksheet)workbook.Worksheets[i];
                    wsDel.Delete();
                }
            }
            catch { }

            var wsSource = (Excel.Worksheet)workbook.Worksheets[1];
            Excel.Worksheet wsTarget = (Excel.Worksheet)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            try { wsTarget.Name = "目標工作表名稱"; }
            catch { try { wsTarget.Name = "目標工作表名稱1"; } catch { } }

            long lastRow = wsSource.Cells[wsSource.Rows.Count, 8].End(Excel.XlDirection.xlUp).Row; // H column
            int outRow = 1;
            for (int i = 5; i <= lastRow; i++)
            {
                object hv = wsSource.Cells[i, 8].Value; // H
                if (hv == null) continue;
                string hs = hv.ToString()?.Trim() ?? string.Empty;
                if (hs == "-" || hs == "---" || hs == "0") continue;
                if (double.TryParse(hs, out double hd) && Math.Abs(hd) < double.Epsilon) continue;

                // Unmerge C:D if merged
                try
                {
                    Excel.Range rng = wsSource.Range[wsSource.Cells[i, 3], wsSource.Cells[i, 4]]; // C:D
                    if ((bool)rng.MergeCells) rng.UnMerge();
                }
                catch { }

                string part = wsSource.Cells[i, 3].Value?.ToString() ?? string.Empty; // C
                string desc = wsSource.Cells[i, 4].Value?.ToString() ?? string.Empty; // D
                int qty = 0;
                try
                {
                    object qtyObj = wsSource.Cells[i, 8].Value; // H
                    if (qtyObj != null && double.TryParse(qtyObj.ToString(), out double qd)) qty = (int)Math.Round(qd, MidpointRounding.AwayFromZero);
                }
                catch { }

                wsTarget.Cells[outRow, 1].Value = part;
                wsTarget.Cells[outRow, 2].Value = desc;
                wsTarget.Cells[outRow, 3].Value = qty;
                outRow++;

                if (collect != null)
                {
                    collect.Add(new SummaryDetail { Part = part, Desc = desc, Qty = qty });
                }
            }

            try
            {
                var used = wsTarget.UsedRange;
                used.WrapText = true;
                used.Rows.AutoFit();
            }
            catch { }

            try { wsSource.Activate(); } catch { }
        }

        private static void WriteSummaryWorkbook(Excel.Application excelApp, string folderPath)
        {
            Excel.Workbook summaryWb = excelApp.Workbooks.Add();
            Excel.Worksheet ws = summaryWb.Worksheets[1];
            int row = 1;

            foreach (var group in _summaryData)
            {
                // 檔名列 (合併 A:C)
                ws.Cells[row, 1].Value = group.FileName;
                try
                {
                    Excel.Range fr = ws.Range[ws.Cells[row, 1], ws.Cells[row, 3]];
                    fr.Merge();
                    fr.Font.Bold = true;
                    fr.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(233, 236, 239));
                }
                catch { }
                row++;

                // 料號 A-Z 排序（不分大小寫）。假設不會有重複料號。
                var ordered = group.Details
                    .OrderBy(d => d.Part ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                foreach (var d in ordered)
                {
                    ws.Cells[row, 1].Value = d.Part;
                    ws.Cells[row, 2].Value = d.Desc;
                    ws.Cells[row, 3].Value = d.Qty;
                    row++;
                }
            }
            try { ws.Columns.AutoFit(); } catch { }
            string summaryName = $"dispatch_summary_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            string savePath = Path.Combine(folderPath, summaryName);
            summaryWb.SaveAs(savePath);
            summaryWb.Close();
        }

        private static int FindLastNonEmptyColumnValueInRow(object[,] dataArray, int rowIndex)
        {
            for (int col = dataArray.GetLength(1); col >= 1; col--)
            {
                if (dataArray[rowIndex, col] != null)
                    if (int.TryParse(dataArray[rowIndex, col].ToString(), out int result)) return result;
            }
            return 0;
        }

        private static bool IsDashLike(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            string t = s.Trim();
            return t == "-" || t == "－" || t == "?" || t == "—" || t == "–";
        }

        private static void ApplySecondaryCellStyle(Excel.Range cell)
        { try { cell.Font.Name = "PMingLiU"; cell.Font.Bold = false; cell.Font.Size = 8; } catch { } }
        private static void ClearCellFill(Excel.Range cell)
        { try { var interior = cell.Interior; interior.Pattern = Excel.XlPattern.xlPatternNone; interior.TintAndShade = 0; interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone; } catch { } }
        private static void ApplyNegativeFill(Excel.Range cell)
        { try { var interior = cell.Interior; interior.Pattern = Excel.XlPattern.xlPatternSolid; interior.TintAndShade = 0; interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 199, 206)); } catch { } }

        private static void ExecuteCmdCommand(string command)
        {
            ProcessStartInfo processStartInfo = new ProcessStartInfo("cmd.exe", "/c " + command)
            { RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true };
            using (Process process = new Process())
            {
                process.StartInfo = processStartInfo;
                process.Start();
                string result = process.StandardOutput.ReadToEnd();
                process.WaitForExit();
                MessageBox.Show(result);
            }
        }

        public static void SetDispatchData(string fileName, string partNumber, double dispatchQuantity)
        {
            if (!_dispatchData.ContainsKey(fileName)) _dispatchData[fileName] = new Dictionary<string, double>();
            _dispatchData[fileName][partNumber] = dispatchQuantity;
        }
        public static void ClearDispatchData() => _dispatchData.Clear();
        private static double GetDispatchQuantity(string fileName, string partNumber)
            => _dispatchData.ContainsKey(fileName) && _dispatchData[fileName].ContainsKey(partNumber)
                ? _dispatchData[fileName][partNumber] : 0;
        public static void DebugDispatchData()
        {
            var debug = "發料數據內容：\n";
            foreach (var fileData in _dispatchData)
            {
                debug += $"檔案: {fileData.Key}\n";
                foreach (var partData in fileData.Value)
                    debug += $"  料號: {partData.Key} = {partData.Value}\n";
            }
            Debug.WriteLine(debug);
        }
        private static void UpdateProgressBar(ProgressBar progressBar1, Label labelCurrentFile, string fileName, int currentRow, int totalRows, int progressValue)
        { try { int max = Math.Max(progressBar1.Maximum, 1); progressBar1.Value = Math.Min(progressValue, max); int percent = (int)((double)progressBar1.Value / max * 100); labelCurrentFile.Text = $"目前執行到的檔案：{fileName} {currentRow}/{totalRows} ({percent}%)"; } catch { } }
    }
}
