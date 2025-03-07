
using System.IO;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Runtime.InteropServices;
//taskkill -f -im excel.exe �R�����椤excel

namespace _231211A
{   

    public partial class Form1 : Form
    {
        OpenFileDialog openFileDialog1 = new OpenFileDialog();
        OpenFileDialog openFileDialog2 = new OpenFileDialog();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel �ɮ� (*.xls;*.xlsx)|*.xls;*.xlsx";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog2.Filter = "Excel �ɮ� (*.xls;*.xlsx)|*.xls;*.xlsx";
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = openFileDialog2.FileName;
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

            // �p�G��泣�O�ťաA�i��ݭn��^�q�{�ȩζi�@�B�B�z
            return 0; // ���]�q�{�Ȭ� 0
        }


        private void button2_Click(object sender, EventArgs e)
        {   

            if (string.IsNullOrEmpty(textBox1.Text) || string.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("�D�ɩΰ��ɨS��");
            }

                Excel.Application excelApp = null;
                Excel.Workbook workbook = null;
                Excel.Workbook workbook1 = null;
            try
            {
                //�}���ɮ�1,2
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(openFileDialog1.FileName);
                Excel.Worksheet worksheet = workbook.Worksheets[1];
                workbook1 = excelApp.Workbooks.Open(openFileDialog2.FileName);
                Excel.Worksheet worksheet1 = workbook1.Worksheets[1];

                Excel.Range excelRange = worksheet.UsedRange;
                Excel.Range excelRange1 = worksheet1.UsedRange;
                int lastRow = excelRange.Rows.Count;
                int lastRow1 = excelRange1.Rows.Count;
                object[,] dataArray = excelRange.Value;
                object[,] dataArray1 = excelRange1.Value;

                /*
                int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                int lastRow1 = worksheet1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                */
                for (int j = 1; j <= lastRow1; j++)
                {

                    for (int i = 1; i < lastRow; i++)
                    {
                        bool isequal = Comparer.Equals(dataArray1[j, 3]?.ToString().Trim(), dataArray[i, 1]?.ToString().Trim());
                        if (isequal)
                        {
                            int last = excelRange.Rows[1].columns.Count;
                            for (int col = last; col >= 1; col--)
                            {
                                if (excelRange.Cells[1, col].Value != null)
                                {
                                    object fo = dataArray1[1, 7];
                                    string g3 = (fo != null) ? fo.ToString() : string.Empty;
                                    string lastEightDigits = (g3.Length >= 8) ? g3.Substring(g3.Length - 8) : g3;
                                    //�q�渹�X
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

                                    //�W�ܾl��
                                    int last1 = FindLastNonEmptyColumnValueInRow(dataArray, i);
                                    //�W�ܾl�Ƽg��o�Ƴ�
                                    if (dataArray1[j, 7] != null && dataArray1[j, 7].ToString().Trim() == "-")
                                    {
                                        break;
                                    }
                                    else
                                    {
                                        worksheet1.Cells[j, 7].value = last1;
                                    }
                                    //�ζq
                                    object co = dataArray1[j, 6];
                                    int f2 = Convert.ToInt32(co);
                                    //�ϥζq�g�^�D��
                                    worksheet.Cells[i, col + 2].Value = f2;
                                    //���s�Ƽ�                                  

                                    double originalvalue = worksheet1.Cells[j, 10].value;
                                    int fourtofive = (int)Math.Round(originalvalue , MidpointRounding.AwayFromZero);
                                    worksheet.Cells[i, col + 3].Value = fourtofive;

                                    //
                                    if(last1 - f2 < 0)
                                    {
                                        worksheet.Cells[i, col + 3].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                    }
                                    break;
                                }
                            }           
                        }
                    }
                }



                //�H�U���s�ɳ��� �������A���ק�
                string folderPath = @"\\St-nas\�ӤH��Ƨ�\Andy\excel\\" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
                Directory.CreateDirectory(folderPath);
                string savePath = folderPath + "\\" + workbook.Name;
                string savePath1 = folderPath + "\\" + workbook1.Name;
                workbook.SaveAs(savePath);
                workbook.Close();
                workbook1.SaveAs(savePath1);
                workbook1.Close();
                excelApp.Quit();
                MessageBox.Show("���槹��");
                System.Windows.Forms.Application.Exit();
            }   
            catch (Exception ex)
            {
                MessageBox.Show("���~�G" + ex.Message);
            }
            finally
            {
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (workbook1 != null) Marshal.ReleaseComObject(workbook1);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }
            }

            }



