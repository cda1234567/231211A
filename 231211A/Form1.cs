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
            openFileDialog.Filter = "Excel 檔案 (*.xls;*.xlsx;*.xlsm;*.xlsb)|*.xls;*.xlsx;*.xlsm;*.xlsb";
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
            return 0; // 預設值為 0
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


        #region 保護區塊: 請勿修改
private void button2_Click(object sender, EventArgs e)
{
    ExcelMergerApi.MergeFiles(listBoxFiles, progressBar1, labelCurrentFile);
}
#endregion

        private void label1_Click(object sender, EventArgs e)
        {

        }

        
        private void buttonMoveUp_Click(object sender, EventArgs e)
        {
            if (listBoxFiles.SelectedItem == null || listBoxFiles.SelectedIndex <= 0)
                return;
            int index = listBoxFiles.SelectedIndex;
            var item = listBoxFiles.SelectedItem;
            
            listBoxFiles.Items.RemoveAt(index);
            listBoxFiles.Items.Insert(index -1 , item);
            listBoxFiles.SelectedIndex = index -1;
            
        }

        private void buttonMoveDown_Click(object sender, EventArgs e)
        {
            if (listBoxFiles.SelectedItem == null || listBoxFiles.SelectedIndex < 0 || listBoxFiles.SelectedIndex >= listBoxFiles.Items.Count - 1)
                return;
            int index = listBoxFiles.SelectedIndex;
            var item = listBoxFiles.SelectedItem;
            listBoxFiles.Items.RemoveAt(index);
            listBoxFiles.Items.Insert(index + 1, item);
            listBoxFiles.SelectedIndex = index + 1;
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
