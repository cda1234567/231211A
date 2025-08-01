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

namespace PCBInventory
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
            {
                var selectedItem = listBoxFiles.SelectedItems[0];
                if (selectedItem != null)
                    listBoxFiles.Items.Remove(selectedItem);
            }
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
        private void listBoxFiles_DragEnter(object? sender, DragEventArgs e)
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        // 拖曳檔案放下時，將檔案加入清單
        private void listBoxFiles_DragDrop(object? sender, DragEventArgs e)
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
            {
                string[]? files = (string[]?)e.Data.GetData(DataFormats.FileDrop);
                if (files != null)
                {
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
        
        #region 保護區塊: 請勿修改
        private void button2_Click(object sender, EventArgs e)
        {
            // 使用新的 EPPlus 版本，不依賴 Office Interop
            ExcelMergerApiNew.MergeFiles(listBoxFiles, progressBar1, labelCurrentFile);
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

                // 使用 EPPlus 讀取 Excel 檔案 (更穩定的方法)
                DataTable dt = ExcelHelper.ReadExcelFile(mainFileName);
                
                if (dt.Rows.Count > 0)
                {
                    // 設置DataGridView數據源
                    inventoryGridView.DataSource = dt;
                    inventoryGridView.Visible = true;
                    exportButton.Enabled = true;
                    
                    MessageBox.Show($"成功載入 {dt.Rows.Count} 行數據", "成功", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("檔案中沒有找到數據", "提示", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                dt.DefaultView.RowFilter = "";
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
                DataTable dataTable = (DataTable)inventoryGridView.DataSource;
                ExcelHelper.ExportToExcel(dataTable, saveDialog.FileName);
                MessageBox.Show("匯出成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"匯出時發生錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
