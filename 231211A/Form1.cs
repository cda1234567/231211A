using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;

namespace _231211A
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            // 拖曳檔案支援
            listBoxFiles.AllowDrop = true;
            listBoxFiles.DragEnter += listBoxFiles_DragEnter;
            listBoxFiles.DragDrop += listBoxFiles_DragDrop;
        }

        private void buttonAddFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel �ɮ� (*.xls;*.xlsx;*.xlsm;*.xlsb)|*.xls;*.xlsx;*.xlsm;*.xlsb";
            openFileDialog.Multiselect = true;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (var file in openFileDialog.FileNames)
                {
                    if (!listBoxFiles.Items.Contains(file))
                        listBoxFiles.Items.Add(file);
                }
            }
        }

        private void buttonRemoveFile_Click(object sender, EventArgs e)
        {
            while (listBoxFiles.SelectedItems.Count > 0)
                listBoxFiles.Items.Remove(listBoxFiles.SelectedItems[0]);
        }

        private int FindLastNonEmptyColumnValueInRow(object[,] dataArray, int rowIndex)
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
            return 0; // ���]�q�{�Ȭ� 0
        }

        private void ExecuteCmdCommand(string command)
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
        private void Form1_Load(object sender, EventArgs e)
        {
        }
        private void Form1_Load_1(object sender, EventArgs e)
        {
        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (listBoxFiles.Items.Count < 2)
            {
                MessageBox.Show("�п�ܦܤ֨���ɮ�");
                return;
            }

            string mainFileName = listBoxFiles.Items[0].ToString();
            List<string> secondaryFileNames = listBoxFiles.Items.Cast<string>().Skip(1).ToList();

            Excel.Application excelApp = null;
            Excel.Workbook mainWorkbook = null;
            List<Excel.Workbook> workbooks = new List<Excel.Workbook>();
            string lastMainSavePath = null;

            // �إ߰��ɦW���Ʀr��
            var fileNameCount = new Dictionary<string, int>();

            try
            {
                excelApp = new Excel.Application();
                mainWorkbook = excelApp.Workbooks.Open(mainFileName);
                Excel.Worksheet mainWorksheet = mainWorkbook.Worksheets[1];
                Excel.Range mainExcelRange = mainWorksheet.UsedRange;
                object[,] mainDataArray = mainExcelRange.Value;

                // �Ыظ�Ƨ�
                string folderPath = @"\\St-nas\�ӤH��Ƨ�\Andy\excel\\" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
                Directory.CreateDirectory(folderPath);

                // �]�m�i�ױ��̤j��
                progressBar1.Maximum = secondaryFileNames.Count;
                progressBar1.Value = 0;

                foreach (var secondaryFileName in secondaryFileNames)
                {
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
                            bool isequal = Comparer.Equals(dataArray[j, 3]?.ToString().Trim(), mainDataArray[k, 1]?.ToString().Trim());
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
                    }

                    // ���Ͱ����x�s�W�١]�۰ʥ[�y�����^
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

                    // �D���x�s�W�٤]�[�y�����]�קK���ơ^
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

                    // �R���¥D��
                    if (lastMainSavePath != null && File.Exists(lastMainSavePath))
                    {
                        try { File.Delete(lastMainSavePath); } catch { }
                    }
                    string mainSavePath = Path.Combine(folderPath, mainSaveName);
                    mainWorkbook.SaveAs(mainSavePath);
                    lastMainSavePath = mainSavePath;

                    // ����D�ɸ귽
                    Marshal.ReleaseComObject(mainWorksheet);
                    Marshal.ReleaseComObject(mainExcelRange);
                    mainWorkbook.Close(false);
                    Marshal.ReleaseComObject(mainWorkbook);

                    // �N�D�ɳ]���s���D��
                    mainWorkbook = excelApp.Workbooks.Open(mainSavePath);
                    mainWorksheet = mainWorkbook.Worksheets[1];
                    mainExcelRange = mainWorksheet.UsedRange;
                    mainDataArray = mainExcelRange.Value;

                    // ��s�i�ױ�
                    progressBar1.Value++;
                }

                // �̫�O�s�D�ɡ]�u�O�d�̫�@���^
                Thread.Sleep(1000);
                mainWorkbook.Save();
                mainWorkbook.Close();

                excelApp.Quit();
                MessageBox.Show("���槹��");
                Application.Exit();
            }
            catch (Exception ex)
            {
                // ���� taskkill ���O������ Excel
                ExecuteCmdCommand("taskkill /f /im excel.exe");

                // ��ܿ��~�H���M���|�l��
                MessageBox.Show($"���~�G{ex.Message}\n���|�l�ܡG{ex.StackTrace}");
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

        private void label1_Click(object sender, EventArgs e)
        {

        }

        // 拖曳檔案進入時，顯示允許拖曳效果
        private void listBoxFiles_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        // 拖曳檔案放下時，將檔案加入清單
        private void listBoxFiles_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (var file in files)
                {
                    string ext = Path.GetExtension(file).ToLower();
                    if ((ext == ".xls" || ext == ".xlsx" || ext == ".xlsm" || ext == ".xlsb") && !listBoxFiles.Items.Contains(file))
                    {
                        listBoxFiles.Items.Add(file);
                    }
                }
            }
        }
    }
}
