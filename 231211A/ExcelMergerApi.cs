using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace _231211A
{
    public static class ExcelMergerApi
    {
        public static void MergeFiles(ListBox listBoxFiles, ProgressBar progressBar1, Label labelCurrentFile)
        {
            if (listBoxFiles.Items.Count < 2)
            {
                MessageBox.Show("請選擇至少兩個檔案");
                return;
            }

            string mainFileName = listBoxFiles.Items[0].ToString();
            List<string> secondaryFileNames = new List<string>();
            for (int i = 1; i < listBoxFiles.Items.Count; i++)
                secondaryFileNames.Add(listBoxFiles.Items[i].ToString());

            Excel.Application excelApp = null;
            Excel.Workbook mainWorkbook = null;
            List<Excel.Workbook> workbooks = new List<Excel.Workbook>();
            string lastMainSavePath = null;

            var fileNameCount = new Dictionary<string, int>();

            try
            {
                excelApp = new Excel.Application();
                mainWorkbook = excelApp.Workbooks.Open(mainFileName);
                Excel.Worksheet mainWorksheet = mainWorkbook.Worksheets[1];
                Excel.Range mainExcelRange = mainWorksheet.UsedRange;
                object[,] mainDataArray = mainExcelRange.Value;

                string folderPath = @"\\St-nas\個人資料夾\Andy\excel\" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm");
                Directory.CreateDirectory(folderPath);

                int totalRows = 0;
                foreach (var secondaryFileName in secondaryFileNames)
                {
                    var workbook = excelApp.Workbooks.Open(secondaryFileName);
                    Excel.Worksheet worksheet = workbook.Worksheets[1];
                    Excel.Range excelRange = worksheet.UsedRange;
                    int lastRow1 = excelRange.Rows.Count;
                    totalRows += lastRow1;
                    workbook.Close(false);
                    Marshal.ReleaseComObject(worksheet);
                    Marshal.ReleaseComObject(excelRange);
                }
                progressBar1.Maximum = totalRows;
                progressBar1.Value = 0;
                int progressValue = 0;

                foreach (var secondaryFileName in secondaryFileNames)
                {
                    labelCurrentFile.Text = $"目前執行到的檔案：{Path.GetFileName(secondaryFileName)}";
                    Application.DoEvents();

                    var workbook = excelApp.Workbooks.Open(secondaryFileName);
                    workbooks.Add(workbook);
                    Excel.Worksheet worksheet = workbook.Worksheets[1];
                    Excel.Range excelRange = worksheet.UsedRange;
                    object[,] dataArray = excelRange.Value;

                    int lastRow = mainExcelRange.Rows.Count;
                    int lastRow1 = excelRange.Rows.Count;

                    for (int j = 1; j <= lastRow1; j++)
                    {
                        for (int k = 1; k < lastRow; k++)
                        {
                            bool isequal = EqualityComparer<string>.Default.Equals(dataArray[j, 3]?.ToString().Trim(), mainDataArray[k, 1]?.ToString().Trim());
                            if (isequal)
                            {
                                int last = mainExcelRange.Columns.Count;
                                for (int col = last; col >= 1; col--)
                                {
                                    if (mainExcelRange.Cells[1, col].Value != null)
                                    {
                                        object fo = dataArray[1, 7];
                                        string g3 = (fo != null) ? fo.ToString() : string.Empty;
                                        string lastEightDigits = (g3.Length >= 8) ? g3.Substring(g3.Length - 8) : g3;

                                        mainWorksheet.Cells[1, col + 3].Value = dataArray[2, 3];

                                        Excel.Range cell = mainWorksheet.Cells[1, col + 3];
                                        Excel.Font font = cell.Font;
                                        font.Name = "Arial";
                                        font.Size = 9;
                                        cell.WrapText = true;

                                        mainWorksheet.Cells[1, col + 2].Value = lastEightDigits;

                                        Excel.Range cell1 = mainWorksheet.Cells[1, col + 2];
                                        Excel.Font font1 = cell1.Font;
                                        font1.Name = "Arial";
                                        font1.Size = 9;
                                        cell1.WrapText = true;

                                        int last1 = FindLastNonEmptyColumnValueInRow(mainDataArray, k);
                                        if (dataArray[j, 7] != null && dataArray[j, 7].ToString().Trim() == "-")
                                        {
                                            break;
                                        }
                                        else
                                        {
                                            worksheet.Cells[j, 7].Value = last1;
                                        }

                                        object co = dataArray[j, 6];
                                        int f2 = Convert.ToInt32(co);
                                        mainWorksheet.Cells[k, col + 2].Value = f2;

                                        object addob = dataArray[j, 8];
                                        if (addob != null && addob.ToString().Trim() != "-")
                                        {
                                            if (int.TryParse(addob?.ToString(), out int f3) && f3 != 0)
                                            {
                                                mainWorksheet.Cells[k, col + 1].Value = f3;
                                            }
                                        }

                                        double originalvalue = worksheet.Cells[j, 10].Value;
                                        int fourtofive = (int)Math.Round(originalvalue, MidpointRounding.AwayFromZero);
                                        mainWorksheet.Cells[k, col + 3].Value = fourtofive;

                                        if (originalvalue < 0)
                                        {
                                            mainWorksheet.Cells[k, col + 3].Interior.Color = ColorTranslator.ToOle(Color.Red);
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                        progressValue++;
                        if (progressValue <= progressBar1.Maximum)
                            progressBar1.Value = progressValue;
                        labelCurrentFile.Text = $"目前執行到的檔案：{Path.GetFileName(secondaryFileName)} {j}/{lastRow1}";
                        Application.DoEvents();
                    }

                    string baseName = Path.GetFileNameWithoutExtension(secondaryFileName);
                    string ext = Path.GetExtension(secondaryFileName);
                    string saveName = baseName + ext;

                    if (!fileNameCount.ContainsKey(baseName))
                        fileNameCount[baseName] = 0;
                    fileNameCount[baseName]++;
                    if (fileNameCount[baseName] > 1)
                        saveName = $"{baseName}-{fileNameCount[baseName]}{ext}";

                    string secondarySavePath = Path.Combine(folderPath, saveName);
                    workbook.SaveAs(secondarySavePath);
                    workbook.Close();

                    string mainBaseName = Path.GetFileNameWithoutExtension(mainFileName);
                    string mainExt = Path.GetExtension(mainFileName);
                    string mainSaveName = mainBaseName + mainExt;
                    if (!fileNameCount.ContainsKey(mainBaseName + "_main"))
                        fileNameCount[mainBaseName + "_main"] = 0;
                    fileNameCount[mainBaseName + "_main"]++;
                    if (fileNameCount[mainBaseName + "_main"] > 1)
                        mainSaveName = $"{mainBaseName}_main-{fileNameCount[mainBaseName + "_main"]}{mainExt}";
                    else
                        mainSaveName = $"{mainBaseName}_main{mainExt}";

                    if (lastMainSavePath != null && File.Exists(lastMainSavePath))
                    {
                        try { File.Delete(lastMainSavePath); } catch { }
                    }
                    string mainSavePath = Path.Combine(folderPath, mainSaveName);
                    mainWorkbook.SaveAs(mainSavePath);
                    lastMainSavePath = mainSavePath;

                    Marshal.ReleaseComObject(mainWorksheet);
                    Marshal.ReleaseComObject(mainExcelRange);
                    mainWorkbook.Close(false);
                    Marshal.ReleaseComObject(mainWorkbook);

                    mainWorkbook = excelApp.Workbooks.Open(mainSavePath);
                    mainWorksheet = mainWorkbook.Worksheets[1];
                    mainExcelRange = mainWorksheet.UsedRange;
                    mainDataArray = mainExcelRange.Value;

                    progressBar1.Value++;
                }

                Thread.Sleep(1000);
                mainWorkbook.Save();
                mainWorkbook.Close();

                excelApp.Quit();
                MessageBox.Show("執行完成");
                Application.Exit();
            }
            catch (Exception ex)
            {
                ExecuteCmdCommand("taskkill /f /im excel.exe");
                MessageBox.Show($"錯誤：{ex.Message}\n堆疊追蹤：{ex.StackTrace}");
            }
            finally
            {
                if (mainWorkbook != null) Marshal.ReleaseComObject(mainWorkbook);
                foreach (var workbook in workbooks)
                {
                    if (workbook != null) Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        private static int FindLastNonEmptyColumnValueInRow(object[,] dataArray, int rowIndex)
        {
            for (int col = dataArray.GetLength(1); col >= 1; col--)
            {
                if (dataArray[rowIndex, col] != null)
                {
                    if (int.TryParse(dataArray[rowIndex, col].ToString(), out int result))
                    {
                        return result;
                    }
                }
            }
            return 0;
        }

        private static void ExecuteCmdCommand(string command)
        {
            ProcessStartInfo processStartInfo = new ProcessStartInfo("cmd.exe", "/c " + command);
            processStartInfo.RedirectStandardOutput = true;
            processStartInfo.UseShellExecute = false;
            processStartInfo.CreateNoWindow = true;

            using (Process process = new Process())
            {
                process.StartInfo = processStartInfo;
                process.Start();

                string result = process.StandardOutput.ReadToEnd();
                process.WaitForExit();

                MessageBox.Show(result);
            }
        }
    }
}
