using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace _231211A
{
    public partial class Form1 : Form
    {
        private OpenFileDialog openFileDialog = new OpenFileDialog();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog.Filter = "Excel 檔案 (*.xls;*.xlsx)|*.xls;*.xlsx";
            openFileDialog.Multiselect = true;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = string.Join(";", openFileDialog.FileNames);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private int FindLastNonEmptyColumnValueInRow(object[,] dataArray, int rowIndex)
        {
            for (int col = dataArray.GetLength(1); col >= 1; col--)
            {
                if (dataArray[rowIndex, col] != null)
                {
                    return Convert.ToInt32(dataArray[rowIndex, col]);
                }
            }
            return 0; // 假設默認值為 0
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("請選擇至少一個檔案");
                return;
            }

            try
            {
                using (var excelApp = new Excel.Application())
                {
                    var fileNames = textBox1.Text.Split(';');
                    var workbooks = new Excel.Workbook[fileNames.Length];

                    for (int i = 0; i < fileNames.Length; i++)
                    {
                        workbooks[i] = excelApp.Workbooks.Open(fileNames[i]);
                    }

                    ProcessWorkbooks(workbooks);

                    SaveAndCloseWorkbooks(workbooks, excelApp);
                }

                MessageBox.Show("執行完成");
                Application.Exit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("錯誤：" + ex.Message);
            }
        }

        private void ProcessWorkbooks(Excel.Workbook[] workbooks)
        {
            var mainWorkbook = workbooks[0];
            var mainWorksheet = (Excel.Worksheet)mainWorkbook.Worksheets[1];
            var mainExcelRange = mainWorksheet.UsedRange;
            var mainDataArray = (object[,])mainExcelRange.Value;

            for (int k = 1; k < workbooks.Length; k++)
            {
                var workbook = workbooks[k];
                var worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                var excelRange = worksheet.UsedRange;
                var dataArray = (object[,])excelRange.Value;

                int lastRow = mainExcelRange.Rows.Count;
                int lastRow1 = excelRange.Rows.Count;

                for (int j = 1; j <= lastRow1; j++)
                {
                    for (int i = 1; i < lastRow; i++)
                    {
                        if (Comparer.Equals(dataArray[j, 3]?.ToString().Trim(), mainDataArray[i, 1]?.ToString().Trim()))
                        {
                            UpdateWorksheetCells(mainWorksheet, worksheet, mainDataArray, dataArray, i, j);
                            break;
                        }
                    }
                }
            }
        }

        private void UpdateWorksheetCells(Excel.Worksheet mainWorksheet, Excel.Worksheet worksheet, object[,] mainDataArray, object[,] dataArray, int i, int j)
        {
            int last = mainWorksheet.UsedRange.Columns.Count;
            for (int col = last; col >= 1; col--)
            {
                if (mainWorksheet.Cells[1, col].Value != null)
                {
                    string lastEightDigits = GetLastEightDigits(dataArray[1, 7]);
                    mainWorksheet.Cells[1, col + 3].Value = dataArray[2, 3];
                    SetCellStyle(mainWorksheet.Cells[1, col + 3]);

                    mainWorksheet.Cells[1, col + 2].Value = lastEightDigits;
                    SetCellStyle(mainWorksheet.Cells[1, col + 2]);

                    int last1 = FindLastNonEmptyColumnValueInRow(mainDataArray, i);
                    if (dataArray[j, 7] != null && dataArray[j, 7].ToString().Trim() == "-")
                    {
                        break;
                    }
                    else
                    {
                        worksheet.Cells[j, 7].Value = last1;
                    }

                    int f2 = Convert.ToInt32(dataArray[j, 6]);
                    mainWorksheet.Cells[i, col + 2].Value = f2;

                    double originalValue = worksheet.Cells[j, 10].Value;
                    int roundedValue = (int)Math.Round(originalValue, MidpointRounding.AwayFromZero);
                    mainWorksheet.Cells[i, col + 3].Value = roundedValue;

                    if (last1 - f2 < 0)
                    {
                        mainWorksheet.Cells[i, col + 3].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                    break;
                }
            }
        }

        private string GetLastEightDigits(object value)
        {
            string strValue = value?.ToString() ?? string.Empty;
            return (strValue.Length >= 8) ? strValue.Substring(strValue.Length - 8) : strValue;
        }

        private void SetCellStyle(Excel.Range cell)
        {
            cell.Font.Name = "Arial";
            cell.Font.Size = 9;
            cell.WrapText = true;
        }

        private void SaveAndCloseWorkbooks(Excel.Workbook[] workbooks, Excel.Application excelApp)
        {
            string folderPath = @"\\St-nas\個人資料夾\Andy\excel\\" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
            Directory.CreateDirectory(folderPath);

            foreach (var workbook in workbooks)
            {
                string savePath = Path.Combine(folderPath, workbook.Name);
                workbook.SaveAs(savePath);
                workbook.Close();
            }

            excelApp.Quit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }
    }
}
