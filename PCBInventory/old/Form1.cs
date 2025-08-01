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
using System.Data;
using System.Data;

namespace _231211A
{
    public partial class Form1 : Form
    {
        private DataGridView inventoryGridView = null!;
        private Button previewButton = null!;
        private TextBox searchBox = null!;
        private Button exportButton = null!;
        
        public Form1()
        {
            InitializeComponent();
            // 拖曳檔案支援
            listBoxFiles.AllowDrop = true;
            listBoxFiles.DragEnter += listBoxFiles_DragEnter;
            listBoxFiles.DragDrop += listBoxFiles_DragDrop;
            
            // 初始化庫存預覽控件
            InitializeInventoryControls();
            
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
        #region 保護區塊: 請勿修改
private void button2_Click(object sender, EventArgs e)
{
    ExcelMergerApi.MergeFiles(listBoxFiles, progressBar1, labelCurrentFile);
}
#endregion

private void label1_Click(object sender, EventArgs e)
{
    // 保留空事件，避免編譯錯誤
}

        private void InitializeInventoryControls()
        {
            // 創建預覽按鈕
            previewButton = new Button
            {
                Text = "預覽庫存",
                Location = new Point(420, 20),
                Size = new Size(100, 30)
            };
            previewButton.Click += PreviewButton_Click;
            this.Controls.Add(previewButton);

            // 創建搜尋框
            searchBox = new TextBox
            {
                Location = new Point(420, 60),
                Size = new Size(200, 25),
                PlaceholderText = "搜尋庫存..."
            };
            searchBox.TextChanged += SearchBox_TextChanged;
            this.Controls.Add(searchBox);

            // 創建匯出按鈕
            exportButton = new Button
            {
                Text = "匯出庫存",
                Location = new Point(530, 20),
                Size = new Size(100, 30),
                Enabled = false
            };
            exportButton.Click += ExportButton_Click;
            this.Controls.Add(exportButton);

            // 創建 DataGridView
            inventoryGridView = new DataGridView
            {
                Location = new Point(420, 100),
                Size = new Size(500, 300),
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom,
                AllowUserToAddRows = false,
                ReadOnly = true,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Visible = false
            };
            this.Controls.Add(inventoryGridView);

            // 調整窗體大小以容納新控件
            this.Width = 950;
        }

        private void PreviewButton_Click(object? sender, EventArgs e)
        {
            if (listBoxFiles.Items.Count == 0)
            {
                MessageBox.Show("請先選擇Excel檔案", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                string mainFileName = listBoxFiles.Items[0].ToString() ?? string.Empty;
                if (string.IsNullOrEmpty(mainFileName))
                {
                    MessageBox.Show("無效的檔案", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 清空舊數據
                inventoryGridView.DataSource = null;
                inventoryGridView.Columns.Clear();

                // 讀取Excel檔案
                using (var excelApp = new Excel.Application())
                {
                    excelApp.Visible = false;
                    var workbook = excelApp.Workbooks.Open(mainFileName);
                    var worksheet = workbook.Worksheets[1];
                    var range = worksheet.UsedRange;

                    // 獲取數據
                    var data = range.Value;
                    int rowCount = range.Rows.Count;
                    int colCount = range.Columns.Count;

                    // 創建DataTable
                    var dt = new System.Data.DataTable();

                    // 添加列
                    for (int i = 1; i <= colCount; i++)
                    {
                        string columnName = ((object[,])data)[1, i]?.ToString() ?? $"Column{i}";
                        dt.Columns.Add(columnName);
                    }

                    // 添加數據
                    for (int r = 2; r <= rowCount; r++)
                    {
                        var row = dt.NewRow();
                        for (int c = 1; c <= colCount; c++)
                        {
                            row[c - 1] = ((object[,])data)[r, c] ?? DBNull.Value;
                        }
                        dt.Rows.Add(row);
                    }

                    // 設置DataGridView數據源
                    inventoryGridView.DataSource = dt;
                    inventoryGridView.Visible = true;
                    exportButton.Enabled = true;

                    // 釋放資源
                    workbook.Close(false);
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"預覽時發生錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SearchBox_TextChanged(object? sender, EventArgs e)
        {
            if (inventoryGridView.DataSource is not System.Data.DataTable dt)
                return;

            string searchText = searchBox.Text.ToLower();
            if (string.IsNullOrWhiteSpace(searchText))
            {
                // 清除篩選
                (inventoryGridView.DataSource as System.Data.DataTable)?.DefaultView.RowFilter = "";
                return;
            }

            // 建立篩選條件
            var filterConditions = new List<string>();
            foreach (System.Data.DataColumn column in dt.Columns)
            {
                filterConditions.Add($"Convert({column.ColumnName}, 'System.String') LIKE '%{searchText}%'");
            }

            // 套用篩選
            dt.DefaultView.RowFilter = string.Join(" OR ", filterConditions);
        }

        private void ExportButton_Click(object? sender, EventArgs e)
        {
            if (inventoryGridView.DataSource == null)
                return;

            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel 檔案|*.xlsx",
                FileName = "庫存報表_" + DateTime.Now.ToString("yyyyMMdd_HHmmss")
            };

            if (saveDialog.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                var excelApp = new Excel.Application();
                var workbook = excelApp.Workbooks.Add();
                var worksheet = workbook.Worksheets[1];

                // 寫入標題
                for (int i = 0; i < inventoryGridView.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1] = inventoryGridView.Columns[i].HeaderText;
                }

                // 寫入數據
                for (int r = 0; r < inventoryGridView.Rows.Count; r++)
                {
                    for (int c = 0; c < inventoryGridView.Columns.Count; c++)
                    {
                        worksheet.Cells[r + 2, c + 1] = inventoryGridView.Rows[r].Cells[c].Value ?? "";
                    }
                }

                // 儲存並關閉
                workbook.SaveAs(saveDialog.FileName);
                workbook.Close();
                excelApp.Quit();

                MessageBox.Show("匯出成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"匯出時發生錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
