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
        public double CurrentStock { get; set; } // �D�ɳ̲׮w�s�]�i�ର�t�^
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

        // �̰��ɿ�����ǳv�@�B�z�A�w��D�ɳ̲׮w�s���t���Ƽu���æ^�g��Ӱ��� H/J�F�P�ɰO���� ExcelMergerApi
        public static void ProcessMainFileNegativeInventory(string mainFilePath, string outputFolder, ProgressBar progressBar, Label statusLabel)
        {
            if (string.IsNullOrWhiteSpace(mainFilePath) || !File.Exists(mainFilePath))
            {
                statusLabel.Text = "�D�ɮפ��s�b";
                return;
            }
            if (string.IsNullOrWhiteSpace(outputFolder) || !Directory.Exists(outputFolder))
            {
                statusLabel.Text = "��X��Ƨ����s�b";
                return;
            }

            try
            {
                statusLabel.Text = "Ū���D�ɳ̲׮w�s...";
                Application.DoEvents();
                var finalStockByPart = ReadMainFinalStock(mainFilePath);

                // Ū���Ĥ@���X�֮ɫO�s�����ɶ��ǡ]�P���ɶ��Ǥ@�P�^
                var secondaryFiles = GetOrderedSecondaryFiles(outputFolder);
                if (secondaryFiles.Length == 0)
                {
                    statusLabel.Text = "�䤣�����";
                    return;
                }

                int fileIndex = 0;
                foreach (var filePath in secondaryFiles)
                {
                    fileIndex++;
                    statusLabel.Text = $"�B�z����({fileIndex}/{secondaryFiles.Length}): {Path.GetFileName(filePath)}";
                    Application.DoEvents();

                    ProcessSingleSecondaryFile(filePath, finalStockByPart, fileIndex, progressBar, statusLabel);
                    Application.DoEvents();
                }

                statusLabel.Text = "�ɮƧ����A�ЦA����@�����b�H�M�ΦܥD��";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"�o�ƳB�z���~�G{ex.Message}", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
                statusLabel.Text = "�o�ƳB�z����";
            }
        }

        // Ū���D�ɨC�ӮƸ����̲׮w�s�]�̥k���Ĥ@�Ӽƭȡ^
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
                        string part = data[r, 1]?.ToString()?.Trim() ?? string.Empty; // A��
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

        // Ū�� __order.txt �H�����P���ɤ@�P�����ɶ���
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
                // ��ơG�H�إ߮ɶ��Ƨǡ]���u���ϥ� manifest�^
                files = Directory.GetFiles(outputFolder, "*.xls*")
                    .Where(f => !Path.GetFileName(f).Contains("_main", StringComparison.OrdinalIgnoreCase))
                    .OrderBy(f => File.GetCreationTime(f))
                    .ToList();
            }

            return files.ToArray();
        }

        // �B�z��@���ɡG��X���ɥX�{���Ƹ��A�Y��D�ɳ̲׮w�s���t�A�u���ç�ɮƼƶq�g�^���� H/J�A�P�ɰO���� API
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
                    string part = cols >= 3 ? (data[r, 3]?.ToString()?.Trim() ?? string.Empty) : string.Empty; // C��
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

                    double finalStock = finalStockByPart[part]; // �D�ɳ̲׮w�s�]�t�^
                    string description = cols >= 4 ? (ws.Cells[r, 4].Value?.ToString()?.Trim() ?? string.Empty) : string.Empty;

                    var item = new ReplenishmentItem
                    {
                        Row = r,
                        PartNumber = part,
                        Description = description,
                        CurrentStock = finalStock,
                        TargetColumn = 8
                    };

                    statusLabel.Text = $"�B�z�o��({processed}/{candidateRows.Count})�G{part}";
                    Application.DoEvents();

                    var result = ShowDispatchDialog(item, fileIndex);
                    if (result.DialogResult == DialogResult.Cancel)
                    {
                        statusLabel.Text = "�o�ƳB�z�w����";
                        return;
                    }
                    if (result.DialogResult == DialogResult.Ignore)
                    {
                        _processedPartNumbers.Add(part);
                        continue;
                    }
                    if (result.DialogResult != DialogResult.OK) continue;

                    double qty = Math.Max(0, result.ReplenishmentQuantity);

                    // �g�^ H�]�o�ơ^�P J�]����+�o�ơ^
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

                    // �O���� API�A�ѲĤG�����b��
                    ExcelMergerApi.SetDispatchData("main", part, qty);

                    _processedPartNumbers.Add(part);
                    statusLabel.Text = $"�w�O���ɮơG{part} �ƶq�G{qty}";
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
            statusLabel.Text = "�Шϥηs�� ProcessMainFileNegativeInventory";
        }
    }
}