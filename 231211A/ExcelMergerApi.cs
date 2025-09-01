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
        private static Dictionary<string, Dictionary<string, double>> _dispatchData = new Dictionary<string, Dictionary<string, double>>();
        
        public static string MergeFiles(ListBox listBoxFiles, ProgressBar progressBar1, Label labelCurrentFile)
        {
            if (listBoxFiles.Items.Count < 2)
            {
                MessageBox.Show("�п�ܦܤ֨���ɮ�");
                return string.Empty;
            }

            string mainFileName = listBoxFiles.Items[0].ToString();
            List<string> secondaryFileNames = new List<string>();
            for (int i = 1; i < listBoxFiles.Items.Count; i++)
                secondaryFileNames.Add(listBoxFiles.Items[i].ToString());

            Excel.Application? excelApp = null;
            Excel.Workbook? mainWorkbook = null;
            List<Excel.Workbook> workbooks = new List<Excel.Workbook>();
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
                string baseFolder = @"\\St-nas\�ӤH��Ƨ�\Andy\excel";
                folderPath = Path.Combine(baseFolder, $"{Path.GetFileNameWithoutExtension(mainFileName)}_{timestamp}");
                Directory.CreateDirectory(folderPath);

                int totalRows = 0;
                foreach (var secondaryFileName in secondaryFileNames)
                {
                    if (!File.Exists(secondaryFileName))
                    {
                        MessageBox.Show($"�ɮפ��s�b: {secondaryFileName}");
                        continue;
                    }
                    var workbook = excelApp.Workbooks.Open(secondaryFileName);
                    Excel.Worksheet worksheet = workbook.Worksheets[1];
                    Excel.Range excelRange = worksheet.UsedRange;
                    int lastRow1 = excelRange.Rows.Count;
                    totalRows += lastRow1;
                    workbook.Close(false);
                    Marshal.ReleaseComObject(worksheet);
                    Marshal.ReleaseComObject(excelRange);
                }
                progressBar1.Maximum = Math.Max(totalRows, 1);
                progressBar1.Value = 0;
                int progressValue = 0;

                foreach (var secondaryFileName in secondaryFileNames)
                {
                    labelCurrentFile.Text = $"�ثe����쪺�ɮסG{Path.GetFileName(secondaryFileName)}";
                    Application.DoEvents();

                    if (!File.Exists(secondaryFileName))
                    {
                        MessageBox.Show($"�ɮפ��s�b: {secondaryFileName}");
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

                    // ���D�G���g Dispatch�A���~��۰ʴ���A�B������󩳦�
                    string orderRaw = null;
                    try { orderRaw = dataArray[1, 7]?.ToString(); } catch { orderRaw = null; }
                    if (string.IsNullOrEmpty(orderRaw))
                    {
                        try { orderRaw = dataArray[1, 8]?.ToString(); } catch { orderRaw = null; }
                    }
                    string orderLast8 = string.IsNullOrEmpty(orderRaw) ? string.Empty : (orderRaw.Length >= 8 ? orderRaw[^8..] : orderRaw);
                    string productName = string.Empty;
                    try { productName = dataArray[2, 3]?.ToString() ?? string.Empty; } catch { productName = string.Empty; }

                    mainWorksheet.Cells[1, baseCol + 1].Value = string.Empty; // ����� Dispatch
                    mainWorksheet.Cells[1, baseCol + 2].Value = orderLast8;
                    mainWorksheet.Cells[1, baseCol + 3].Value = productName;

                    var hdr1 = (Excel.Range)mainWorksheet.Cells[1, baseCol + 1];
                    var hdr2 = (Excel.Range)mainWorksheet.Cells[1, baseCol + 2];
                    var hdr3 = (Excel.Range)mainWorksheet.Cells[1, baseCol + 3];
                    hdr1.Font.Name = hdr2.Font.Name = hdr3.Font.Name = "Arial";
                    hdr1.Font.Size = hdr2.Font.Size = hdr3.Font.Size = 9;
                    hdr3.WrapText = true; hdr3.EntireColumn.WrapText = true; hdr3.EntireRow.AutoFit();

                    for (int j = 2; j <= lastRowSec; j++)
                    {
                        bool skipRow = false;
                        if (dataArray.GetLength(1) >= 7 && dataArray[j, 7] != null && dataArray[j, 7].ToString().Trim() == "-") skipRow = true;
                        if (dataArray.GetLength(1) >= 8 && dataArray[j, 8] != null && dataArray[j, 8].ToString().Trim() == "-") skipRow = true;
                        if (skipRow)
                        {
                            progressValue++;
                            UpdateProgressBar(progressBar1, labelCurrentFile, Path.GetFileName(secondaryFileName), j, lastRowSec, progressValue);
                            continue;
                        }

                        string secPart = dataArray[j, 3]?.ToString()?.Trim() ?? string.Empty; // C �� �Ƹ�
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

                        // �g���ɻݨD�Ȩ�D�ɤ�����
                        double f2 = 0;
                        if (dataArray.GetLength(1) >= 6 && dataArray[j, 6] != null && double.TryParse(dataArray[j, 6].ToString(), out double tmpF2))
                            f2 = Math.Round(tmpF2, MidpointRounding.AwayFromZero);
                        mainWorksheet.Cells[mainRowIndex, baseCol + 2].Value = f2;

                        // �N�D�ɦ��C�ثe�̥k�ȵ����� G ��A�����ɤ������� J ��
                        int prevFinal = FindLastNonEmptyColumnValueInRow(mainDataArray, mainRowIndex);
                        if (prevFinal != 0) worksheet.Cells[j, 7].Value = prevFinal; // G ��

                        // �Y�ĤG���X�֥B���e���o�ơ]�Ѽu���O���^�A�u�g�b�Ĥ@���X�{�ɡA�Ω���ɤ���
                        double dispatchQtyPreset = GetDispatchQuantity("main", secPart);
                        if (dispatchQtyPreset > 0 && !dispatchedOnce.Contains(secPart))
                        {
                            // ���ۦ�B���榡�ơA�ȼg�ȡ]�H�קK��/��/��������^
                            mainWorksheet.Cells[mainRowIndex, baseCol + 1].Value = dispatchQtyPreset.ToString("F0");
                            dispatchedOnce.Add(secPart);
                        }

                        // �q����Ū�^ J �浲��ȡA�^�g�D�ɥk��A���ۦ�
                        double jValue = 0;
                        if (worksheet.Cells[j, 10].Value != null)
                            double.TryParse(worksheet.Cells[j, 10].Value.ToString(), out jValue);
                        int finalRounded = (int)Math.Round(jValue, MidpointRounding.AwayFromZero);
                        var outCell = (Excel.Range)mainWorksheet.Cells[mainRowIndex, baseCol + 3];
                        outCell.Value = finalRounded;
                        outCell.Interior.ColorIndex = Excel.Constants.xlNone; // �������

                        progressValue++;
                        UpdateProgressBar(progressBar1, labelCurrentFile, Path.GetFileName(secondaryFileName), j, lastRowSec, progressValue);
                        Application.DoEvents();
                    }

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

                    // ��s main �d��
                    mainExcelRange = mainWorksheet.UsedRange;
                    mainDataArray = mainExcelRange.Value;
                }

                try
                {
                    var manifestPath = Path.Combine(folderPath, "__order.txt");
                    File.WriteAllLines(manifestPath, orderedSavedSecondaryFiles);
                }
                catch { }

                string mainSaveName = $"{Path.GetFileNameWithoutExtension(mainFileName)}_main{Path.GetExtension(mainFileName)}";
                string mainSavePath = Path.Combine(folderPath, mainSaveName);
                mainWorkbook.SaveAs(mainSavePath);
                Thread.Sleep(300);
                mainWorkbook.Close();
                excelApp.Quit();
            }
            catch (FileNotFoundException fnfEx)
            {
                MessageBox.Show($"�ɮק䤣��: {fnfEx.FileName}");
            }
            catch (COMException comEx)
            {
                MessageBox.Show($"Excel COM ���~: {comEx.Message}");
            }
            catch (Exception ex)
            {
                ExecuteCmdCommand("taskkill /f /im excel.exe");
                MessageBox.Show($"���~�G{ex.Message}\n���|�l�ܡG{ex.StackTrace}");
            }
            finally
            {
                try { if (mainWorkbook != null) Marshal.ReleaseComObject(mainWorkbook); } catch { }
                foreach (var wb in workbooks) { try { if (wb != null) Marshal.ReleaseComObject(wb); } catch { } }
                try { if (excelApp != null) Marshal.ReleaseComObject(excelApp); } catch { }
            }
            return folderPath;
        }

        private static int FindLastNonEmptyColumnValueInRow(object[,] dataArray, int rowIndex)
        {
            for (int col = dataArray.GetLength(1); col >= 1; col--)
            {
                if (dataArray[rowIndex, col] != null)
                {
                    if (int.TryParse(dataArray[rowIndex, col].ToString(), out int result))
                        return result;
                }
            }
            return 0;
        }

        private static void ExecuteCmdCommand(string command)
        {
            ProcessStartInfo processStartInfo = new ProcessStartInfo("cmd.exe", "/c " + command)
            {
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

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
            if (!_dispatchData.ContainsKey(fileName))
                _dispatchData[fileName] = new Dictionary<string, double>();
            _dispatchData[fileName][partNumber] = dispatchQuantity;
        }
        public static void ClearDispatchData() => _dispatchData.Clear();
        private static double GetDispatchQuantity(string fileName, string partNumber)
            => _dispatchData.ContainsKey(fileName) && _dispatchData[fileName].ContainsKey(partNumber)
                ? _dispatchData[fileName][partNumber] : 0;

        public static void DebugDispatchData()
        {
            var debug = "�o�Ƽƾڤ��e�G\n";
            foreach (var fileData in _dispatchData)
            {
                debug += $"�ɮ�: {fileData.Key}\n";
                foreach (var partData in fileData.Value)
                    debug += $"  �Ƹ�: {partData.Key} = {partData.Value}\n";
            }
            System.Diagnostics.Debug.WriteLine(debug);
        }

        private static void UpdateProgressBar(ProgressBar progressBar1, Label labelCurrentFile, string fileName, int currentRow, int totalRows, int progressValue)
        {
            try
            {
                int max = Math.Max(progressBar1.Maximum, 1);
                progressBar1.Value = Math.Min(progressValue, max);
                int percent = (int)((double)progressBar1.Value / max * 100);
                labelCurrentFile.Text = $"�ثe����쪺�ɮסG{fileName} {currentRow}/{totalRows} ({percent}%)";
            }
            catch { }
        }
    }
}
