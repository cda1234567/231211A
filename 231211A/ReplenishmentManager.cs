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

        // �̰��ɿ�����ǳv�@�B�z�A�Ȧb���ɥѫD�t���t�ɸ߰ݸɮơF��L��۰ʧ� H �� 0�]�Ť~�ɡ^
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

        // �̲Ĥ@�Ӧr��(A��Z)�ƧǸ߰ݸɮƶ���
        private static int FirstLetterRank(string part)
        {
            if (string.IsNullOrEmpty(part)) return int.MaxValue;
            char ch = char.ToUpperInvariant(part[0]);
            if (ch >= 'A' && ch <= 'Z') return ch - 'A';
            return 26 + ch; // �D�^��r���Ʀb�̫�
        }

        // �P�_�O�_���U�ءu-�v�Ÿ�
        private static bool IsDashLike(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            string t = s.Trim();
            return t == "-" || t == "��" || t == "?" || t == "�X" || t == "�V";
        }

        // �B�z��@���ɡG�ȹ糧�ɥѫD�t���t���Ƹ��߰ݸɮơF��L�C�۰ʧ� H �� 0�]�Ť~�ɡ^
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

                // �����y���ɡA�����Կ�P�D�Կ�
                var candidates = new List<(int Row, string Part, double G, double H, double J)>();
                var nonCandidates = new List<int>();

                for (int r = 2; r <= rows; r++)
                {
                    string part = cols >= 3 ? (data[r, 3]?.ToString()?.Trim() ?? string.Empty) : string.Empty; // C�� �Ƹ�
                    if (string.IsNullOrEmpty(part)) { nonCandidates.Add(r); continue; }
                    if (_processedPartNumbers.Contains(part)) { nonCandidates.Add(r); continue; }

                    string gText = cols >= 7 ? data[r, 7]?.ToString() : null;
                    string hText = cols >= 8 ? data[r, 8]?.ToString() : null;

                    // �Y G �� H �� dash-like�A�����аO���D�Կ�A�B���򤣭n��� H�]�O�d�� "-"�^
                    if (IsDashLike(gText) || IsDashLike(hText)) { nonCandidates.Add(r); continue; }

                    double g = 0, h = double.NaN, j = 0;
                    if (cols >= 7 && data[r, 7] != null) double.TryParse(data[r, 7].ToString(), out g);   // G
                    if (cols >= 8 && data[r, 8] != null) double.TryParse(data[r, 8].ToString(), out h);   // H
                    if (cols >= 10 && data[r, 10] != null) double.TryParse(data[r, 10].ToString(), out j); // J

                    // ���ɤ~��t�GG >= 0 �B J < 0
                    if (g >= 0 && j < 0)
                    {
                        candidates.Add((r, part, g, h, j));
                    }
                    else
                    {
                        nonCandidates.Add(r);
                    }
                }

                // ��D�Կ��G�Y H �ťթΫD�ƭȡA�ɼg 0�]���� J�^�F���Y H �� dash-like �h�O�d���ʡF�B�ȹ靈�Ƹ�(C��D��)���C��
                foreach (int r in nonCandidates)
                {
                    var partText = ws.Cells[r, 3].Value?.ToString()?.Trim();
                    if (string.IsNullOrEmpty(partText))
                    {
                        // �L�Ƹ����槹������
                        continue;
                    }

                    var hCellObj = ws.Cells[r, 8].Value;
                    string hText = hCellObj?.ToString();
                    if (IsDashLike(hText))
                    {
                        // �O�d "-"�A����
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

                // �Կ�ƧǡG���r A��Z�A�A��ӮƸ�
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

                    // ��ܮ���ܥH�D�ɳ̲׭Ȭ��ǡ]Ū����h�h�^���� J�^
                    double finalNeg = finalStockByPart.TryGetValue(part, out var finalVal) ? finalVal : j;

                    var item = new ReplenishmentItem
                    {
                        Row = r,
                        PartNumber = part,
                        Description = description,
                        CurrentStock = finalNeg,
                        TargetColumn = 8
                    };

                    statusLabel.Text = $"�B�z�o��({processed}/{ordered.Count})�G{part}";
                    Application.DoEvents();

                    var owner = statusLabel?.FindForm();
                    var result = ShowDispatchDialog(item, fileIndex, owner);
                    if (result.DialogResult == DialogResult.Cancel)
                    {
                        statusLabel.Text = "�o�ƳB�z�w����";
                        return;
                    }
                    if (result.DialogResult == DialogResult.Ignore)
                    {
                        // �ϥΪ̲��L�G�� H=0�]�Y�|���O�ƭȥB���O dash-like�^�F�ȹ靈�Ƹ����C
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

                    // �g�^ H�]�ɮƶq�^�FJ = �� J + �ɮơ]�������󩳦�β���/�榡�ܧ�^
                    double adjusted = j + qty;

                    ws.Cells[r, 8].Value = qty;                        // H
                    ws.Cells[r, 10].Value = (int)Math.Round(adjusted); // J

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
            statusLabel.Text = "�Шϥηs�� ProcessMainFileNegativeInventory";
        }
    }
}