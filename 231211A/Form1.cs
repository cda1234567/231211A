using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace _231211A
{
    public partial class Form1 : Form
    {
        private readonly OpenFileDialog openFileDialog1 = new OpenFileDialog();
        private readonly OpenFileDialog openFileDialog2 = new OpenFileDialog();
        private readonly List<string> subFilePaths = new();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel檔 (*.xls;*.xlsx)|*.xls;*.xlsx";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog2.Filter = "Excel檔 (*.xls;*.xlsx)|*.xls;*.xlsx";
            openFileDialog2.Multiselect = true;
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                foreach (var file in openFileDialog2.FileNames)
                {
                    if (!subFilePaths.Contains(file))
                    {
                        subFilePaths.Add(file);
                        listBoxSubFiles.Items.Add(file);
                    }
                }
            }
        }

        private void buttonRemove_Click(object sender, EventArgs e)
        {
            var selected = listBoxSubFiles.SelectedItems.Cast<string>().ToList();
            foreach (var file in selected)
            {
                subFilePaths.Remove(file);
                listBoxSubFiles.Items.Remove(file);
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
            return 0;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text) || subFilePaths.Count == 0)
            {
                MessageBox.Show("請選擇檔案");
                return;
            }

            Excel.Application? excelApp = null;
            Excel.Workbook? mainWorkbook = null;
            try
            {
                excelApp = new Excel.Application();
                mainWorkbook = excelApp.Workbooks.Open(textBox1.Text);
                Excel.Worksheet worksheet = mainWorkbook.Worksheets[1];

                string folderPath = @"\\St-nas\ӤHƧ\Andy\excel\\" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
                Directory.CreateDirectory(folderPath);

                foreach (string subPath in subFilePaths)
                {
                    Excel.Workbook? subWorkbook = null;
                    try
                    {
                        subWorkbook = excelApp.Workbooks.Open(subPath);
                        Excel.Worksheet worksheet1 = subWorkbook.Worksheets[1];

                        Excel.Range excelRange = worksheet.UsedRange;
                        Excel.Range excelRange1 = worksheet1.UsedRange;
                        int lastRow = excelRange.Rows.Count;
                        int lastRow1 = excelRange1.Rows.Count;
                        object[,] dataArray = excelRange.Value;
                        object[,] dataArray1 = excelRange1.Value;

                        for (int j = 1; j <= lastRow1; j++)
                        {
                            for (int i = 1; i < lastRow; i++)
                            {
                                bool isequal = Comparer.Equals(dataArray1[j, 3]?.ToString().Trim(), dataArray[i, 1]?.ToString().Trim());
                                if (isequal)
                                {
                                    int last = excelRange.Rows[1].Columns.Count;
                                    for (int col = last; col >= 1; col--)
                                    {
                                        if (excelRange.Cells[1, col].Value != null)
                                        {
                                            object fo = dataArray1[1, 7];
                                            string g3 = (fo != null) ? fo.ToString() : string.Empty;
                                            string lastEightDigits = (g3.Length >= 8) ? g3[^8..] : g3;

                                            worksheet.Cells[1, col + 3].Value = dataArray1[2, 3];
                                            Excel.Range cell = worksheet.Cells[1, col + 3];
                                            Excel.Font font = cell.Font;
                                            font.Name = "Arial";
                                            font.Size = 9;
                                            cell.WrapText = true;

                                            worksheet.Cells[1, col + 2].Value = lastEightDigits;
                                            Excel.Range cell1 = worksheet.Cells[1, col + 2];
                                            Excel.Font font1 = cell1.Font;
                                            font1.Name = "Arial";
                                            font1.Size = 9;
                                            cell1.WrapText = true;

                                            int last1 = FindLastNonEmptyColumnValueInRow(dataArray, i);
                                            if (dataArray1[j, 7] != null && dataArray1[j, 7].ToString().Trim() == "-")
                                            {
                                                break;
                                            }
                                            else
                                            {
                                                worksheet1.Cells[j, 7].Value = last1;
                                            }
                                            object co = dataArray1[j, 6];
                                            int f2 = Convert.ToInt32(co);
                                            worksheet.Cells[i, col + 2].Value = f2;

                                            double originalvalue = worksheet1.Cells[j, 10].Value;
                                            int fourtofive = (int)Math.Round(originalvalue, MidpointRounding.AwayFromZero);
                                            worksheet.Cells[i, col + 3].Value = fourtofive;

                                            if (last1 - f2 < 0)
                                            {
                                                worksheet.Cells[i, col + 3].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                            }
                                            break;
                                        }
                                    }
                                }
                            }
                        }

                        string savePathSub = Path.Combine(folderPath, subWorkbook.Name);
                        subWorkbook.SaveAs(savePathSub);
                        subWorkbook.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("處理副檔錯誤:" + ex.Message);
                    }
                    finally
                    {
                        if (subWorkbook != null) Marshal.ReleaseComObject(subWorkbook);
                    }
                }

                string savePathMain = Path.Combine(folderPath, mainWorkbook.Name);
                mainWorkbook.SaveAs(savePathMain);
                mainWorkbook.Close();
                excelApp.Quit();
                MessageBox.Show("完成");
                Application.Exit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("錯誤" + ex.Message);
            }
            finally
            {
                if (mainWorkbook != null) Marshal.ReleaseComObject(mainWorkbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}

