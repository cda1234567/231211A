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

namespace _231211A
{
    public partial class Form1 : Form
    {
        private DataGridView inventoryGridView = null!;
        private Button previewButton = null!;
        private TextBox searchBox = null!;
        private Button exportButton = null!;
        private Label inventoryLabel = null!;
        private Panel inventoryPanel = null!;
        private ComboBox filterComboBox = null!;
        private Label summaryLabel = null!;
        private Button setMainFileButton = null!;
        private Label mainFileLabel = null!;
        private Button backupButton = null!;

        // 主檔案路徑 - 綁定在程式中（此路徑現用作：公司庫存檔，僅供預覽顯示）
        private string mainFilePath = string.Empty;
        private const string MAIN_FILE_CONFIG = "mainfile.config";

        public Form1()
        {
            InitializeComponent();

            try
            {
                // 拖曳檔案支援
                listBoxFiles.AllowDrop = true;
                listBoxFiles.DragEnter += listBoxFiles_DragEnter;
                listBoxFiles.DragDrop += listBoxFiles_DragDrop;

                // 初始化庫存管理控件
                InitializeInventoryControls();

                // 載入主檔案設定（現作為公司庫存檔用於預覽）
                LoadMainFileConfig();

                // 顯示載入完成訊息（測試用）
                this.Text = "PCB 扣帳系統 - 庫存管理 (已載入)";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
            listBoxFiles.Items.Insert(index - 1, item);
            listBoxFiles.SelectedIndex = index - 1;

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
            string firstMergeFolder = string.Empty; // 記錄第一次合併的資料夾
            string secondMergeFolder = string.Empty; // 記錄第二次合併的資料夾

            // 執行期間鎖住「執行」按鈕，避免重複點擊
            buttonExecute.Enabled = false;
            buttonExecute.Text = "執行中...";
            Cursor previousCursor = this.Cursor;
            this.Cursor = Cursors.WaitCursor;

            try
            {
                // 自動備份主檔案（目前備份的是 mainFilePath 所指之公司預覽檔，如需改為備份清單第一個檔案可再調整）
                if (!string.IsNullOrEmpty(mainFilePath) && File.Exists(mainFilePath))
                {
                    CreateBackup();
                }

                // 清除之前的發料數據
                ExcelMergerApi.ClearDispatchData();

                // 清除已處理的料號記錄
                ReplenishmentManager.ClearProcessedPartNumbers();

                // 準備檔案清單（以清單第 1 個檔案為主檔）
                var tempListBox = new ListBox();
                foreach (var item in listBoxFiles.Items)
                {
                    tempListBox.Items.Add(item);
                }

                // 第一次執行合併作業（臨時處理）
                firstMergeFolder = ExcelMergerApi.MergeFiles(tempListBox, progressBar1, labelCurrentFile);

                // 扣帳完成後自動進入發料處理
                // 第二次合併與發料處理沿用第一次清單的第 1 個檔案作為主檔
                string selectedMainPath = tempListBox.Items.Count > 0 ? tempListBox.Items[0]?.ToString() ?? string.Empty : string.Empty;
                if (!string.IsNullOrEmpty(selectedMainPath) && File.Exists(selectedMainPath))
                {
                    // 自動開始發料處理，不顯示提示
                    ProcessReplenishment(firstMergeFolder);

                    // 準備第二次合併的檔案清單：主檔改用「第一次清單的第 1 個檔案」，副檔改用「第一次輸出的副檔（含補料）"
                    var secondListBox = new ListBox();

                    // 主檔沿用第一次清單第 1 個檔案
                    secondListBox.Items.Add(selectedMainPath);

                    // 副檔：優先依據 __order.txt 使用第一次輸出檔；找不到則使用資料夾內所有非 *_main 的 Excel
                    var orderPath = Path.Combine(firstMergeFolder, "__order.txt");
                    if (File.Exists(orderPath))
                    {
                        foreach (var line in File.ReadAllLines(orderPath))
                        {
                            if (string.IsNullOrWhiteSpace(line)) continue;
                            var full = Path.Combine(firstMergeFolder, line.Trim());
                            if (File.Exists(full)) secondListBox.Items.Add(full);
                        }
                    }
                    else if (Directory.Exists(firstMergeFolder))
                    {
                        var seconds = Directory.GetFiles(firstMergeFolder, "*.xls*", SearchOption.TopDirectoryOnly)
                            .Where(f => !Path.GetFileName(f).Contains("_main", System.StringComparison.OrdinalIgnoreCase))
                            .OrderBy(File.GetCreationTime)
                            .ToArray();
                        foreach (var f in seconds) secondListBox.Items.Add(f);
                    }

                    // 發料處理完成後，重新執行扣帳以應用發料數據（主檔=第一次清單第 1 個，副檔=第一次輸出）
                    secondMergeFolder = ExcelMergerApi.MergeFiles(secondListBox, progressBar1, labelCurrentFile);

                    // 刪除第一次產生的臨時資料夾
                    if (!string.IsNullOrEmpty(firstMergeFolder) && Directory.Exists(firstMergeFolder))
                    {
                        try
                        {
                            Directory.Delete(firstMergeFolder, true);
                        }
                        catch { } // 靜默處理刪除失敗
                    }

                    // 全部完成
                    labelCurrentFile.Text = "完成";
                    progressBar1.Value = progressBar1.Maximum;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"處理過程中發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                labelCurrentFile.Text = "處理失敗";
            }
            finally
            {
                // 還原按鈕與游標狀態
                buttonExecute.Enabled = true;
                buttonExecute.Text = "執行";
                this.Cursor = previousCursor;
            }
        }
        #endregion

        private void label1_Click(object sender, EventArgs e)
        {
            // 保留空事件，避免編譯錯誤
        }

        #region 庫存管理功能

        /// <summary>
        /// 初始化庫存管理控件
        /// </summary>
        private void InitializeInventoryControls()
        {
            // 創建庫存管理面板
            inventoryPanel = new Panel
            {
                Location = new Point(920, 12),
                Size = new Size(480, 450),
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.FromArgb(248, 249, 250)
            };
            this.Controls.Add(inventoryPanel);

            // 庫存管理標題
            inventoryLabel = new Label
            {
                Text = "📦 庫存管理系統",
                Location = new Point(10, 10),
                Size = new Size(200, 25),
                Font = new Font("Microsoft YaHei", 12, FontStyle.Bold),
                ForeColor = Color.FromArgb(33, 37, 41)
            };
            inventoryPanel.Controls.Add(inventoryLabel);

            // 主檔案設定區域（此處按鈕改為設定公司庫存檔，僅供預覽）
            var mainFileGroupBox = new GroupBox
            {
                Text = "主檔案設定",
                Location = new Point(10, 40),
                Size = new Size(460, 80),
                Font = new Font("Microsoft YaHei", 9)
            };
            inventoryPanel.Controls.Add(mainFileGroupBox);

            setMainFileButton = new Button
            {
                Text = "設定公司庫存檔",
                Location = new Point(10, 20),
                Size = new Size(100, 30),
                BackColor = Color.FromArgb(0, 123, 255),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            setMainFileButton.Click += SetMainFileButton_Click;
            mainFileGroupBox.Controls.Add(setMainFileButton);

            mainFileLabel = new Label
            {
                Text = "尚未設定公司庫存檔",
                Location = new Point(120, 25),
                Size = new Size(320, 20),
                ForeColor = Color.FromArgb(108, 117, 125)
            };
            mainFileGroupBox.Controls.Add(mainFileLabel);

            backupButton = new Button
            {
                Text = "手動備份",
                Location = new Point(10, 50),
                Size = new Size(100, 25),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            backupButton.Click += BackupButton_Click;
            mainFileGroupBox.Controls.Add(backupButton);

            // 庫存預覽按鈕
            previewButton = new Button
            {
                Text = "預覽庫存",
                Location = new Point(10, 130),
                Size = new Size(100, 30),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            previewButton.Click += PreviewButton_Click;
            inventoryPanel.Controls.Add(previewButton);

            // 搜尋功能
            var searchLabel = new Label
            {
                Text = "搜尋料號：",
                Location = new Point(120, 135),
                Size = new Size(80, 20)
            };
            inventoryPanel.Controls.Add(searchLabel);

            searchBox = new TextBox
            {
                Location = new Point(200, 133),
                Size = new Size(150, 25),
                PlaceholderText = "輸入料號或名稱..."
            };
            searchBox.TextChanged += SearchBox_TextChanged;
            inventoryPanel.Controls.Add(searchBox);

            // 篩選下拉選單
            filterComboBox = new ComboBox
            {
                Location = new Point(360, 133),
                Size = new Size(100, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            filterComboBox.Items.AddRange(new[] { "全部", "低庫存", "零庫存", "負庫存" });
            filterComboBox.SelectedIndex = 0;
            filterComboBox.SelectedIndexChanged += FilterComboBox_SelectedIndexChanged;
            inventoryPanel.Controls.Add(filterComboBox);

            // 庫存表格
            inventoryGridView = new DataGridView
            {
                Location = new Point(10, 170),
                Size = new Size(460, 220),
                AllowUserToAddRows = false,
                ReadOnly = true,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(233, 236, 239),
                    ForeColor = Color.FromArgb(33, 37, 41),
                    Font = new Font("Microsoft YaHei", 9, FontStyle.Bold)
                }
            };
            inventoryPanel.Controls.Add(inventoryGridView);

            // 匯出按鈕
            exportButton = new Button
            {
                Text = "匯出庫存報表",
                Location = new Point(10, 400),
                Size = new Size(120, 30),
                BackColor = Color.FromArgb(220, 53, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Enabled = false
            };
            exportButton.Click += ExportButton_Click;
            inventoryPanel.Controls.Add(exportButton);

            // 統計資訊
            summaryLabel = new Label
            {
                Text = "庫存統計：0 項目",
                Location = new Point(140, 405),
                Size = new Size(200, 20),
                ForeColor = Color.FromArgb(108, 117, 125)
            };
            inventoryPanel.Controls.Add(summaryLabel);

            // 調整主窗體大小
            this.Width = 1420;
        }

        /// <summary>
        /// 設定公司庫存檔（僅供預覽）
        /// </summary>
        private void SetMainFileButton_Click(object? sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel 檔案 (*.xls;*.xlsx;*.xlsm;*.xlsb)|*.xls;*.xlsx;*.xlsm;*.xlsb";
            openFileDialog.Title = "選擇公司庫存檔（預覽用）";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                mainFilePath = openFileDialog.FileName;
                SaveMainFileConfig();
                UpdateMainFileLabel();
                MessageBox.Show("公司庫存檔設定成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// 載入主檔案設定（現作為公司庫存檔路徑）
        /// </summary>
        private void LoadMainFileConfig()
        {
            try
            {
                if (File.Exists(MAIN_FILE_CONFIG))
                {
                    mainFilePath = File.ReadAllText(MAIN_FILE_CONFIG);
                    UpdateMainFileLabel();
                }
            }
            catch { }
        }

        /// <summary>
        /// 儲存公司庫存檔設定
        /// </summary>
        private void SaveMainFileConfig()
        {
            try
            {
                File.WriteAllText(MAIN_FILE_CONFIG, mainFilePath);
            }
            catch { }
        }

        /// <summary>
        /// 更新公司庫存檔標籤
        /// </summary>
        private void UpdateMainFileLabel()
        {
            if (!string.IsNullOrEmpty(mainFilePath) && File.Exists(mainFilePath))
            {
                mainFileLabel.Text = $"✓ {Path.GetFileName(mainFilePath)}";
                mainFileLabel.ForeColor = Color.FromArgb(40, 167, 69);
            }
            else
            {
                mainFileLabel.Text = "❌ 公司庫存檔不存在或未設定";
                mainFileLabel.ForeColor = Color.FromArgb(220, 53, 69);
            }
        }

        /// <summary>
        /// 創建備份
        /// </summary>
        private void CreateBackup()
        {
            try
            {
                if (string.IsNullOrEmpty(mainFilePath) || !File.Exists(mainFilePath))
                    return;

                string backupFolder = Path.Combine(Path.GetDirectoryName(mainFilePath)!, "Backups");
                Directory.CreateDirectory(backupFolder);

                string fileName = Path.GetFileNameWithoutExtension(mainFilePath);
                string extension = Path.GetExtension(mainFilePath);
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                string backupPath = Path.Combine(backupFolder, $"{fileName}_備份_{timestamp}{extension}");

                File.Copy(mainFilePath, backupPath, true);
            }
            catch (Exception ex)
            {
                // 靜默處理備份錯誤
            }
        }

        /// <summary>
        /// 手動備份按鈕事件
        /// </summary>
        private void BackupButton_Click(object? sender, EventArgs e)
        {
            CreateBackup();
        }

        /// <summary>
        /// 預覽庫存（僅讀公司庫存檔）
        /// </summary>
        private void PreviewButton_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(mainFilePath) || !File.Exists(mainFilePath))
            {
                MessageBox.Show("請先設定公司庫存檔", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Excel.Application? excelApp = null;
            Excel.Workbook? workbook = null;
            Excel.Worksheet? worksheet = null;

            try
            {
                // 清空舊數據
                inventoryGridView.DataSource = null;
                inventoryGridView.Columns.Clear();

                // 讀取Excel檔案
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                workbook = excelApp.Workbooks.Open(mainFilePath);
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                Excel.Range range = worksheet.UsedRange;

                // 獲取數據
                object[,] data = (object[,])range.Value;
                int rowCount = range.Rows.Count;
                int colCount = range.Columns.Count;

                // 創建DataTable
                var dt = new DataTable();

                // 用於記錄已使用的欄位名稱
                var usedColumnNames = new HashSet<string>();

                // 添加列
                for (int i = 1; i <= colCount; i++)
                {
                    string originalColumnName = data[1, i]?.ToString()?.Trim() ?? $"Column{i}";
                    string columnName = originalColumnName;

                    // 處理重複的欄位名稱
                    int counter = 1;
                    while (usedColumnNames.Contains(columnName))
                    {
                        columnName = $"{originalColumnName}_{counter}";
                        counter++;
                    }

                    // 確保欄位名稱不為空
                    if (string.IsNullOrWhiteSpace(columnName))
                    {
                        columnName = $"Column{i}";
                    }

                    usedColumnNames.Add(columnName);
                    dt.Columns.Add(columnName);
                }

                // 添加數據行
                for (int r = 2; r <= rowCount; r++)
                {
                    var row = dt.NewRow();
                    for (int c = 1; c <= colCount; c++)
                    {
                        try
                        {
                            row[c - 1] = data[r, c] ?? DBNull.Value;
                        }
                        catch
                        {
                            row[c - 1] = DBNull.Value;
                        }
                    }
                    dt.Rows.Add(row);
                }

                // 設置數據源
                inventoryGridView.DataSource = dt;
                exportButton.Enabled = true;

                // 更新統計
                UpdateInventorySummary(dt);

                // 設置庫存數量列的顏色
                SetInventoryColors();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"預覽時發生錯誤：{ex.Message}\n\n詳細錯誤：{ex.StackTrace}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 釋放資源
                try
                {
                    if (worksheet != null)
                    {
                        Marshal.ReleaseComObject(worksheet);
                        worksheet = null;
                    }
                    if (workbook != null)
                    {
                        workbook.Close(false);
                        Marshal.ReleaseComObject(workbook);
                        workbook = null;
                    }
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                        excelApp = null;
                    }
                }
                catch { }
            }
        }

        /// <summary>
        /// 搜尋功能
        /// </summary>
        private void SearchBox_TextChanged(object? sender, EventArgs e)
        {
            if (inventoryGridView.DataSource is not DataTable dt)
                return;

            string searchText = searchBox.Text.ToLower();
            ApplyFilters(dt, searchText, filterComboBox.SelectedItem?.ToString() ?? "全部");
        }

        /// <summary>
        /// 篩選功能
        /// </summary>
        private void FilterComboBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (inventoryGridView.DataSource is not DataTable dt)
                return;

            string searchText = searchBox.Text.ToLower();
            ApplyFilters(dt, searchText, filterComboBox.SelectedItem?.ToString() ?? "全部");
        }

        /// <summary>
        /// 應用篩選條件
        /// </summary>
        private void ApplyFilters(DataTable dt, string searchText, string filterType)
        {
            try
            {
                var filterConditions = new List<string>();

                // 搜尋條件
                if (!string.IsNullOrWhiteSpace(searchText))
                {
                    foreach (DataColumn column in dt.Columns)
                    {
                        // 避免使用特殊字元的欄位名稱
                        string columnName = column.ColumnName.Replace("'", "''").Replace("[", "").Replace("]", "");
                        filterConditions.Add($"Convert([{columnName}], 'System.String') LIKE '%{searchText.Replace("'", "''")}%");
                    }
                }

                // 庫存狀態篩選
                var stockFilters = new List<string>();
                if (dt.Columns.Count > 0)
                {
                    // 尋找可能的庫存欄位（通常是數字型態的最後幾欄）
                    string stockColumnName = "";
                    for (int i = dt.Columns.Count - 1; i >= 0; i--)
                    {
                        var column = dt.Columns[i];
                        bool hasNumericData = false;

                        // 檢查這一欄是否包含數字資料
                        foreach (DataRow row in dt.Rows)
                        {
                            if (row[i] != null && row[i] != DBNull.Value)
                            {
                                if (double.TryParse(row[i].ToString(), out _))
                                {
                                    hasNumericData = true;
                                    break;
                                }
                            }
                        }

                        if (hasNumericData)
                        {
                            stockColumnName = column.ColumnName.Replace("'", "''").Replace("[", "").Replace("]", "");
                            break;
                        }
                    }

                    if (!string.IsNullOrEmpty(stockColumnName))
                    {
                        switch (filterType)
                        {
                            case "低庫存":
                                stockFilters.Add($"(ISNULL([{stockColumnName}], 0) < 10 AND ISNULL([{stockColumnName}], 0) > 0)");
                                break;
                            case "零庫存":
                                stockFilters.Add($"ISNULL([{stockColumnName}], 0) = 0");
                                break;
                            case "負庫存":
                                stockFilters.Add($"ISNULL([{stockColumnName}], 0) < 0");
                                break;
                        }
                    }
                }

                // 組合條件
                string finalFilter = "";
                if (filterConditions.Count > 0 && stockFilters.Count > 0)
                {
                    finalFilter = $"({string.Join(" OR ", filterConditions)}) AND ({string.Join(" OR ", stockFilters)})";
                }
                else if (filterConditions.Count > 0)
                {
                    finalFilter = string.Join(" OR ", filterConditions);
                }
                else if (stockFilters.Count > 0)
                {
                    finalFilter = string.Join(" OR ", stockFilters);
                }

                dt.DefaultView.RowFilter = finalFilter;
                UpdateInventorySummary(dt);
            }
            catch (Exception ex)
            {
                // 如果篩選失敗，清除篩選條件
                try
                {
                    dt.DefaultView.RowFilter = "";
                }
                catch { }

                MessageBox.Show($"篩選時發生錯誤：{ex.Message}", "篩選錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// 設置庫存顏色
        /// </summary>
        private void SetInventoryColors()
        {
            foreach (DataGridViewRow row in inventoryGridView.Rows)
            {
                if (row.Cells.Count > 0)
                {
                    var lastCell = row.Cells[row.Cells.Count - 1];
                    if (double.TryParse(lastCell.Value?.ToString(), out double stock))
                    {
                        if (stock < 0)
                        {
                            row.DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 238); // 淺紅色
                            row.DefaultCellStyle.ForeColor = Color.FromArgb(220, 53, 69);
                        }
                        else if (stock == 0)
                        {
                            row.DefaultCellStyle.BackColor = Color.FromArgb(255, 243, 205); // 淺黃色
                            row.DefaultCellStyle.ForeColor = Color.FromArgb(255, 193, 7);
                        }
                        else if (stock < 10)
                        {
                            row.DefaultCellStyle.BackColor = Color.FromArgb(255, 248, 225); // 淺橙色
                            row.DefaultCellStyle.ForeColor = Color.FromArgb(253, 126, 20);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 更新庫存統計
        /// </summary>
        private void UpdateInventorySummary(DataTable dt)
        {
            int totalItems = dt.DefaultView.Count;
            int lowStock = 0, zeroStock = 0, negativeStock = 0;

            foreach (DataRowView rowView in dt.DefaultView)
            {
                var row = rowView.Row;
                if (row.ItemArray.Length > 0)
                {
                    var lastValue = row.ItemArray[row.ItemArray.Length - 1];
                    if (double.TryParse(lastValue?.ToString(), out double stock))
                    {
                        if (stock < 0) negativeStock++;
                        else if (stock == 0) zeroStock++;
                        else if (stock < 10) lowStock++;
                    }
                }
            }

            summaryLabel.Text = $"總計：{totalItems} | 低庫存：{lowStock} | 零庫存：{zeroStock} | 負庫存：{negativeStock}";
        }

        /// <summary>
        /// 匯出庫存報表
        /// </summary>
        private void ExportButton_Click(object? sender, EventArgs e)
        {
            if (inventoryGridView.DataSource == null)
                return;

            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel 檔案|*.xlsx",
                FileName = "庫存報表_" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss")
            };

            if (saveDialog.ShowDialog() != DialogResult.OK)
                return;

            Excel.Application? excelApp = null;
            Excel.Workbook? workbook = null;
            Excel.Worksheet? worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];

                // 寫入標題
                for (int i = 0; i < inventoryGridView.Columns.Count; i++)
                {
                    Excel.Range headerCell = (Excel.Range)worksheet.Cells[1, i + 1];
                    headerCell.Value = inventoryGridView.Columns[i].HeaderText;
                    headerCell.Font.Bold = true;
                    headerCell.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                }

                // 寫入數據
                for (int r = 0; r < inventoryGridView.Rows.Count; r++)
                {
                    for (int c = 0; c < inventoryGridView.Columns.Count; c++)
                    {
                        Excel.Range dataCell = (Excel.Range)worksheet.Cells[r + 2, c + 1];
                        dataCell.Value = inventoryGridView.Rows[r].Cells[c].Value ?? "";
                    }
                }

                // 自動調整列寬
                worksheet.Columns.AutoFit();

                // 儲存並關閉
                workbook.SaveAs(saveDialog.FileName);

                MessageBox.Show("匯出成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"匯出時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                try
                {
                    if (worksheet != null)
                    {
                        Marshal.ReleaseComObject(worksheet);
                        worksheet = null;
                    }
                    if (workbook != null)
                    {
                        workbook.Close();
                        Marshal.ReleaseComObject(workbook);
                        workbook = null;
                    }
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                        excelApp = null;
                    }
                }
                catch { }
            }
        }

        /// <summary>
        /// 執行發料處理 - 檢查主檔案的最終負庫存
        /// </summary>
        private void ProcessReplenishment(string outputFolder)
        {
            if (string.IsNullOrEmpty(mainFilePath) || !File.Exists(mainFilePath))
            {
                return; // 沒有公司庫存檔設定不影響發料，僅不處理預覽相關
            }

            try
            {
                labelCurrentFile.Text = "正在檢查主檔案的最終負庫存...";

                // 取得最新的輸出資料夾中的主檔案
                string mainFileToProcess = mainFilePath; // 預設使用原始主檔案（如輸出有 *_main 則覆蓋）

                if (!string.IsNullOrEmpty(outputFolder) && Directory.Exists(outputFolder))
                {
                    // 尋找輸出資料夾中的主檔案
                    var mainFiles = Directory.GetFiles(outputFolder, "*_main*")
                        .Where(f => f.EndsWith(".xlsx", System.StringComparison.OrdinalIgnoreCase) || f.EndsWith(".xls", System.StringComparison.OrdinalIgnoreCase))
                        .OrderByDescending(f => File.GetLastWriteTime(f))
                        .ToArray();

                    if (mainFiles.Length > 0)
                    {
                        mainFileToProcess = mainFiles[0]; // 使用最新的主檔案
                    }
                }

                // 處理主檔案的最終負庫存，找出對應的副檔案進行發料
                labelCurrentFile.Text = $"正在檢查主檔案負庫存：{Path.GetFileName(mainFileToProcess)}";
                ReplenishmentManager.ProcessMainFileNegativeInventory(mainFileToProcess, outputFolder, progressBar1, labelCurrentFile);

                // 發料處理完成後的狀態顯示
                labelCurrentFile.Text = "發料處理完成";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"發料處理時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                labelCurrentFile.Text = "發料處理失敗";
            }
            finally
            {
                // 恢復狀態
                if (progressBar1.Maximum > 0)
                {
                    progressBar1.Value = progressBar1.Maximum;
                }
            }
        }

        /// <summary>
        /// 取得今天的資料夾路徑
        /// </summary>
        private string GetTodayFolderPath()
        {
            return @"\\St-nas\個人資料夾\Andy\excel\" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm");
        }

        /// <summary>
        /// 取得最新的輸出資料夾
        /// </summary>
        private string GetLatestOutputFolder()
        {
            try
            {
                string baseFolder = @"\\St-nas\個人資料夾\Andy\excel\";
                if (Directory.Exists(baseFolder))
                {
                    var todayFolders = Directory.GetDirectories(baseFolder)
                        .Where(d => Path.GetFileName(d).StartsWith(DateTime.Now.ToString("yyyy-MM-dd")))
                        .OrderByDescending(d => d)
                        .ToArray();

                    if (todayFolders.Length > 0)
                    {
                        return todayFolders[0];
                    }
                }

                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        #endregion

        private void labelProgress_Click(object sender, EventArgs e)
        {

        }
    }
}
