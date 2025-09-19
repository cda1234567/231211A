using System;
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
using System.Drawing.Drawing2D;

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
        private Button btnLoadSnapshot = null!;
        private Button btnExportRequirement = null!;
        private Label lblSnapshotInfo = null!;
        private ToolTip tips = new ToolTip();
        
        // 新增：完成摘要面板
        private Panel completionPanel = null!;
        private Label completionTitle = null!;
        private Label completionDetails = null!;
        private Button openFolderButton = null!;
        private Button closeCompletionButton = null!;
        private System.Windows.Forms.Timer fadeTimer = null!;

        // 主檔案路徑 - 綁定在程式中（此路徑現用作：公司庫存檔，僅供預覽顯示）
        private string mainFilePath = string.Empty;
        private const string MAIN_FILE_CONFIG = "mainfile.config";
        private bool baselineWarnedThisRun = false; // 新增：避免重複提醒

        private string outputFolderPath = string.Empty;
        private const string OUTPUT_CONFIG = "output.config";
        
        // 執行統計
        private string lastOutputFolder = string.Empty;
        private int filesProcessed = 0;
        private DateTime executionStartTime;

        // 新增：記錄按鈕原始尺寸，避免放大後無法還原
        private readonly Dictionary<Control, Size> _originalButtonSizes = new();

        public Form1()
        {
            InitializeComponent();

            try
            {
                // 設定緊緻的最小尺寸
                this.MinimumSize = new Size(1380, 520);
                this.StartPosition = FormStartPosition.CenterScreen;
                
                UiStyle.ApplyTheme(this);
                SetupGradientBackground();
                
                // 拖曱檔案支援
                listBoxFiles.AllowDrop = true;
                listBoxFiles.DragEnter += listBoxFiles_DragEnter;
                listBoxFiles.DragDrop += listBoxFiles_DragDrop;

                // 初始化庫存管理控件
                InitializeInventoryControls();
                
                // 初始化完成摘要面板
                InitializeCompletionPanel();

                // 套用左側元件樣式
                TryStyleLeftControls();

                // 載入主檔案設定（現作為公司庫存檔用於預覽）
                LoadMainFileConfig();

                // 標題與圖示美化
                this.Text = "📋 PCB 扣帳系統 - 庫存管理";
                SetupAnimations(); // 使用更新後的動畫註冊
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化時發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetupGradientBackground()
        {
            this.Paint += (s, e) =>
            {
                using var brush = new LinearGradientBrush(
                    this.ClientRectangle,
                    Color.FromArgb(250, 251, 252),
                    Color.FromArgb(241, 245, 249),
                    90f);
                e.Graphics.FillRectangle(brush, this.ClientRectangle);
            };
        }

        private void SetupAnimations()
        {
            // 保留淡入面板 timer
            if (fadeTimer == null)
                fadeTimer = new System.Windows.Forms.Timer { Interval = 50 };

            // 為所有（包含巢狀）Button 記錄原始尺寸並掛載事件
            foreach (var btn in GetAllButtons(this))
            {
                if (!_originalButtonSizes.ContainsKey(btn))
                    _originalButtonSizes[btn] = btn.Size;

                btn.MouseEnter -= Button_MouseEnter;
                btn.MouseLeave -= Button_MouseLeave;
                btn.MouseEnter += Button_MouseEnter;
                btn.MouseLeave += Button_MouseLeave;
            }
        }

        // 遞迴取得所有 Button
        private IEnumerable<Button> GetAllButtons(Control root)
        {
            foreach (Control c in root.Controls)
            {
                if (c is Button b)
                    yield return b;
                if (c.HasChildren)
                {
                    foreach (var inner in GetAllButtons(c))
                        yield return inner;
                }
            }
        }

        private void Button_MouseEnter(object? sender, EventArgs e)
        {
            if (sender is Control ctl)
                AnimateScale(ctl, true);
        }
        private void Button_MouseLeave(object? sender, EventArgs e)
        {
            if (sender is Control ctl)
                AnimateScale(ctl, false);
        }

        // 重新實作：以原始尺寸為基準放大/還原，不累積誤差
        private void AnimateScale(Control control, bool enlarge)
        {
            if (!_originalButtonSizes.TryGetValue(control, out var baseSize))
                baseSize = control.Size; // fallback

            if (enlarge)
            {
                const float factor = 1.05f; // 放大 5%
                int targetW = (int)(baseSize.Width * factor);
                int targetH = (int)(baseSize.Height * factor);
                int dx = (targetW - baseSize.Width) / 2;
                int dy = (targetH - baseSize.Height) / 2;
                control.SuspendLayout();
                control.Size = new Size(targetW, targetH);
                control.Location = new Point(control.Location.X - dx, control.Location.Y - dy);
                control.ResumeLayout();
            }
            else
            {
                int dx = (control.Width - baseSize.Width) / 2;
                int dy = (control.Height - baseSize.Height) / 2;
                control.SuspendLayout();
                control.Size = baseSize;
                control.Location = new Point(control.Location.X + dx, control.Location.Y + dy);
                control.ResumeLayout();
            }
        }

        private void InitializeCompletionPanel()
        {
            completionPanel = new Panel
            {
                Size = new Size(400, 200),
                Location = new Point(this.Width / 2 - 200, this.Height / 2 - 100),
                BackColor = Color.FromArgb(248, 250, 252),
                BorderStyle = BorderStyle.None,
                Visible = false,
                Anchor = AnchorStyles.None
            };
            
            // 添加陰影效果
            completionPanel.Paint += (s, e) =>
            {
                var rect = completionPanel.ClientRectangle;
                using var path = CreateRoundedRectangle(rect, 12);
                using var shadowBrush = new SolidBrush(Color.FromArgb(50, 0, 0, 0));
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                
                // 繪製陰影
                using var shadowPath = CreateRoundedRectangle(new Rectangle(rect.X + 3, rect.Y + 3, rect.Width, rect.Height), 12);
                e.Graphics.FillPath(shadowBrush, shadowPath);
                
                // 繪製主背景
                using var bgBrush = new LinearGradientBrush(rect, Color.White, Color.FromArgb(248, 250, 252), 45f);
                e.Graphics.FillPath(bgBrush, path);
                
                // 繪製邊框
                using var borderPen = new Pen(Color.FromArgb(226, 232, 240), 1);
                e.Graphics.DrawPath(borderPen, path);
            };

            completionTitle = new Label
            {
                Text = "✅ 執行完成！",
                Location = new Point(20, 20),
                Size = new Size(360, 30),
                Font = new Font(UiStyle.BaseFont.FontFamily, 14F, FontStyle.Bold),
                ForeColor = Color.FromArgb(22, 163, 74),
                TextAlign = ContentAlignment.MiddleCenter
            };

            completionDetails = new Label
            {
                Location = new Point(20, 60),
                Size = new Size(360, 60),
                Font = UiStyle.BaseFont,
                ForeColor = Color.FromArgb(71, 85, 105),
                TextAlign = ContentAlignment.TopCenter
            };

            openFolderButton = new Button
            {
                Text = "📁 開啟輸出資料夾",
                Location = new Point(20, 130),
                Size = new Size(150, 35),
                Font = UiStyle.BaseFont,
                BackColor = UiStyle.Primary,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 0 }
            };
            openFolderButton.Click += OpenFolderButton_Click;

            closeCompletionButton = new Button
            {
                Text = "❌",
                Location = new Point(360, 10),
                Size = new Size(30, 30),
                Font = new Font(UiStyle.BaseFont.FontFamily, 10F),
                BackColor = Color.FromArgb(239, 68, 68),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 0 }
            };
            closeCompletionButton.Click += (s, e) => HideCompletionPanel();

            completionPanel.Controls.AddRange(new Control[] 
            { 
                completionTitle, completionDetails, openFolderButton, closeCompletionButton 
            });
            this.Controls.Add(completionPanel);
            completionPanel.BringToFront();
        }

        private GraphicsPath CreateRoundedRectangle(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            int diameter = radius * 2;
            
            path.AddArc(rect.X, rect.Y, diameter, diameter, 180, 90);
            path.AddArc(rect.Right - diameter, rect.Y, diameter, diameter, 270, 90);
            path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(rect.X, rect.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();
            
            return path;
        }

        private void ShowCompletionPanel(string folder, int processed, TimeSpan duration)
        {
            lastOutputFolder = folder;
            filesProcessed = processed;
            
            completionDetails.Text = $"處理檔案：{processed} 個\n" +
                                   $"耗時：{duration.TotalSeconds:F1} 秒\n" +
                                   $"輸出位置：{Path.GetFileName(folder)}";
            
            // 重新計算位置
            completionPanel.Location = new Point(
                (this.ClientSize.Width - completionPanel.Width) / 2,
                (this.ClientSize.Height - completionPanel.Height) / 2
            );
            
            completionPanel.Visible = true;
            completionPanel.BringToFront();
            
            // 淡入動畫
            completionPanel.BackColor = Color.FromArgb(0, 248, 250, 252);
            fadeTimer.Tag = "in";
            fadeTimer.Tick += FadeAnimation;
            fadeTimer.Start();
        }

        private void HideCompletionPanel()
        {
            fadeTimer.Tag = "out";
            fadeTimer.Tick += FadeAnimation;
            fadeTimer.Start();
        }

        private void FadeAnimation(object? sender, EventArgs e)
        {
            if (fadeTimer.Tag?.ToString() == "in")
            {
                var alpha = completionPanel.BackColor.A + 15;
                if (alpha >= 255)
                {
                    fadeTimer.Stop();
                    fadeTimer.Tick -= FadeAnimation;
                    return;
                }
                completionPanel.BackColor = Color.FromArgb(alpha, 248, 250, 252);
            }
            else if (fadeTimer.Tag?.ToString() == "out")
            {
                var alpha = completionPanel.BackColor.A - 15;
                if (alpha <= 0)
                {
                    completionPanel.Visible = false;
                    fadeTimer.Stop();
                    fadeTimer.Tick -= FadeAnimation;
                    return;
                }
                completionPanel.BackColor = Color.FromArgb(alpha, 248, 250, 252);
            }
        }

        private void OpenFolderButton_Click(object? sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(lastOutputFolder) && Directory.Exists(lastOutputFolder))
            {
                try
                {
                    Process.Start("explorer.exe", lastOutputFolder);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"無法開啟資料夾：{ex.Message}", "錯誤");
                }
            }
        }

        private void TryStyleLeftControls()
        {
            try
            {
                if (listBoxFiles != null) 
                {
                    UiStyle.StyleListBox(listBoxFiles);
                    // 增加拖放視覺提示
                    listBoxFiles.BackColor = Color.FromArgb(249, 250, 251);
                    listBoxFiles.Font = new Font(UiStyle.BaseFont.FontFamily, 9F);
                }
                
                if (progressBar1 != null) 
                {
                    UiStyle.StyleProgressBar(progressBar1);
                    progressBar1.Height = 8; // 更現代的細進度條
                }
                
                // 按鈕加上圖示和間距
                if (buttonAddFile != null) 
                {
                    UiStyle.StyleButtonPrimary(buttonAddFile);
                    buttonAddFile.Text = "➕ 新增檔案";
                    buttonAddFile.Height = 32;
                }
                if (buttonRemoveFile != null) 
                {
                    UiStyle.StyleButtonDanger(buttonRemoveFile);
                    buttonRemoveFile.Text = "🗑️ 移除選取";
                    buttonRemoveFile.Height = 32;
                }
                if (buttonMoveUp != null) 
                {
                    UiStyle.StyleButtonNeutral(buttonMoveUp);
                    buttonMoveUp.Text = "⬆️";
                    buttonMoveUp.Width = 36;
                    buttonMoveUp.Height = 32;
                }
                if (buttonMoveDown != null) 
                {
                    UiStyle.StyleButtonNeutral(buttonMoveDown);
                    buttonMoveDown.Text = "⬇️";
                    buttonMoveDown.Width = 36;
                    buttonMoveDown.Height = 32;
                }
                if (buttonExecute != null) 
                {
                    UiStyle.StyleButtonSuccess(buttonExecute);
                    buttonExecute.Text = "🚀 執行";
                    buttonExecute.Height = 40;
                    buttonExecute.Font = new Font(UiStyle.BaseFont.FontFamily, 10F, FontStyle.Bold);
                }
                if (buttonBrowseOutput != null) 
                {
                    UiStyle.StyleButtonNeutral(buttonBrowseOutput);
                    buttonBrowseOutput.Text = "📁 選擇路徑";
                    buttonBrowseOutput.Height = 32;
                }
                
                if (labelOutputFolder != null) UiStyle.StyleLabel(labelOutputFolder, subtle: true);
                if (labelCurrentFile != null) UiStyle.StyleLabel(labelCurrentFile, subtle: true);
            }
            catch { }
        }

        // 新增：以目前公司庫存檔作為基準庫存（缺料判斷用）
        private void SetSnapshotFromMainFile()
        {
            try
            {
                if (!string.IsNullOrEmpty(mainFilePath) && File.Exists(mainFilePath))
                {
                    InventoryBaselineManager.LoadSnapshot(mainFilePath);
                    var name = Path.GetFileName(mainFilePath);
                    lblSnapshotInfo.Text = $"基準: {InventoryBaselineManager.SnapshotTime:MM-dd HH:mm} {name}";
                    lblSnapshotInfo.ForeColor = Color.FromArgb(40, 167, 69);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入基準失敗：{ex.Message}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                // 拖入時高亮
                listBoxFiles.BackColor = Color.FromArgb(219, 234, 254);
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void listBoxFiles_DragDrop(object sender, DragEventArgs e)
        {
            listBoxFiles.BackColor = Color.FromArgb(249, 250, 251); // 恢復原色
            
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
            executionStartTime = DateTime.Now;
            string firstMergeFolder = string.Empty;
            string secondMergeFolder = string.Empty;

            buttonExecute.Enabled = false;
            buttonExecute.Text = "🔄 執行中...";
            Cursor previousCursor = this.Cursor;
            this.Cursor = Cursors.WaitCursor;

            var toToggle = new Control[]
            {
                buttonAddFile, buttonRemoveFile, buttonMoveUp, buttonMoveDown, listBoxFiles,
                setMainFileButton, btnLoadSnapshot, previewButton, exportButton, btnExportRequirement,
                buttonBrowseOutput
            };
            foreach (var c in toToggle) if (c != null) c.Enabled = false;

            try
            {
                if (!string.IsNullOrEmpty(mainFilePath) && File.Exists(mainFilePath))
                {
                    CreateBackup();
                }

                if (!baselineWarnedThisRun && (InventoryBaselineManager.SnapshotTime == null || InventoryBaselineManager.SnapshotStock.Count == 0))
                {
                    baselineWarnedThisRun = true;
                    var dr = MessageBox.Show("尚未載入基準庫存，是否仍要繼續？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (dr != DialogResult.OK)
                    {
                        return;
                    }
                }

                ExcelMergerApi.ClearDispatchData();
                ReplenishmentManager.ClearProcessedPartNumbers();

                var tempListBox = new ListBox();
                foreach (var item in listBoxFiles.Items)
                {
                    tempListBox.Items.Add(item);
                }
                
                filesProcessed = tempListBox.Items.Count;

                firstMergeFolder = ExcelMergerApi.MergeFiles(tempListBox, progressBar1, labelCurrentFile);

                string selectedMainPath = tempListBox.Items.Count > 0 ? tempListBox.Items[0]?.ToString() ?? string.Empty : string.Empty;
                if (!string.IsNullOrEmpty(selectedMainPath) && File.Exists(selectedMainPath))
                {
                    ProcessReplenishment(firstMergeFolder);

                    var secondListBox = new ListBox();
                    secondListBox.Items.Add(selectedMainPath);

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
                            .Where(f => !Path.GetFileName(f).Contains("_main", StringComparison.OrdinalIgnoreCase))
                            .OrderBy(File.GetCreationTime)
                            .ToArray();
                        foreach (var f in seconds) secondListBox.Items.Add(f);
                    }

                    secondMergeFolder = ExcelMergerApi.MergeFiles(secondListBox, progressBar1, labelCurrentFile);

                    if (!string.IsNullOrEmpty(firstMergeFolder) && Directory.Exists(firstMergeFolder))
                    {
                        try { Directory.Delete(firstMergeFolder, true); } catch { }
                    }

                    labelCurrentFile.Text = "✅ 完成";
                    progressBar1.Value = progressBar1.Maximum;
                    
                    // 顯示完成摘要
                    var duration = DateTime.Now - executionStartTime;
                    ShowCompletionPanel(secondMergeFolder, filesProcessed, duration);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"處理過程中發生錯誤：{ex.Message}", "錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                labelCurrentFile.Text = "❌ 處理失敗";
            }
            finally
            {
                buttonExecute.Enabled = true;
                buttonExecute.Text = "🚀 執行";
                this.Cursor = previousCursor;
                foreach (var c in toToggle) if (c != null) c.Enabled = true;
            }
        }
        #endregion

        private void label1_Click(object sender, EventArgs e) { }

        #region 庫存管理功能
        private void InitializeInventoryControls()
        {
            inventoryPanel = new Panel
            {
                Location = new Point(890, 20),  // 從 650 調回 890，恢復原來比例
                Size = new Size(500, 440),      // 從 820 調回 500，恢復原來寬度
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom,
                BorderStyle = BorderStyle.None,
                BackColor = Color.White
            };
            
            // 為右側面板添加圓角和陰影
            inventoryPanel.Paint += (s, e) =>
            {
                var rect = inventoryPanel.ClientRectangle;
                using var path = CreateRoundedRectangle(rect, 16);
                using var shadowBrush = new SolidBrush(Color.FromArgb(30, 0, 0, 0));
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                
                // 陰影
                using var shadowPath = CreateRoundedRectangle(new Rectangle(rect.X + 2, rect.Y + 2, rect.Width, rect.Height), 16);
                e.Graphics.FillPath(shadowBrush, shadowPath);
                
                // 主背景
                using var bgBrush = new LinearGradientBrush(rect, Color.White, Color.FromArgb(252, 253, 254), 45f);
                e.Graphics.FillPath(bgBrush, path);
                
                // 邊框
                using var borderPen = new Pen(Color.FromArgb(226, 232, 240), 1);
                e.Graphics.DrawPath(borderPen, path);
            };
            
            this.Controls.Add(inventoryPanel);

            inventoryLabel = new Label
            {
                Text = "📦 庫存管理系統",
                Location = new Point(20, 20),
                Size = new Size(300, 30),
                Font = new Font(UiStyle.BaseFont.FontFamily, 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(51, 65, 85)
            };
            inventoryPanel.Controls.Add(inventoryLabel);

            var mainFileGroupBox = new GroupBox
            {
                Text = "🔧 主檔案設定",
                Location = new Point(10, 60),
                Size = new Size(480, 110),  // 調整到適合的寬度
                Font = new Font(UiStyle.BaseFont.FontFamily, 10F, FontStyle.Bold),
                ForeColor = Color.FromArgb(71, 85, 105)
            };
            inventoryPanel.Controls.Add(mainFileGroupBox);

            setMainFileButton = new Button
            {
                Text = "🏢 設定公司庫存",
                Location = new Point(10, 25),
                Size = new Size(110, 30)
            };
            setMainFileButton.Click += SetMainFileButton_Click;
            mainFileGroupBox.Controls.Add(setMainFileButton); 
            UiStyle.StyleButtonPrimary(setMainFileButton);

            backupButton = new Button
            {
                Text = "💾 手動備份",
                Location = new Point(10, 60),
                Size = new Size(110, 25)
            };
            backupButton.Click += BackupButton_Click;
            mainFileGroupBox.Controls.Add(backupButton); 
            UiStyle.StyleButtonSuccess(backupButton);

            btnLoadSnapshot = new Button
            {
                Text = "📊 載入基準庫存",
                Location = new Point(130, 25),
                Size = new Size(120, 30)
            };
            btnLoadSnapshot.Click += BtnLoadSnapshot_Click;
            mainFileGroupBox.Controls.Add(btnLoadSnapshot); 
            UiStyle.StyleButtonWarning(btnLoadSnapshot);

            mainFileLabel = new Label
            {
                Text = "尚未設定公司庫存檔",
                Location = new Point(260, 30),
                Size = new Size(210, 20),  // 調整寬度
                Font = new Font(UiStyle.BaseFont.FontFamily, 8.5F),
                ForeColor = Color.FromArgb(107, 114, 128)
            };
            mainFileGroupBox.Controls.Add(mainFileLabel);

            lblSnapshotInfo = new Label
            {
                Text = "尚未載入基準",
                Location = new Point(260, 65),
                Size = new Size(210, 20),  // 調整寬度
                Font = new Font(UiStyle.BaseFont.FontFamily, 8.5F),
                ForeColor = Color.FromArgb(107, 114, 128)
            };
            mainFileGroupBox.Controls.Add(lblSnapshotInfo);

            previewButton = new Button
            {
                Text = "👁️ 預覽庫存",
                Location = new Point(10, 180),
                Size = new Size(100, 32)
            };
            previewButton.Click += PreviewButton_Click;
            inventoryPanel.Controls.Add(previewButton); 
            UiStyle.StyleButtonNeutral(previewButton);

            var searchLabel = new Label
            {
                Text = "🔍 搜尋料號：",
                Location = new Point(120, 185),
                Size = new Size(80, 22),
                Font = new Font(UiStyle.BaseFont.FontFamily, 9F),
                ForeColor = Color.FromArgb(75, 85, 99)
            };
            inventoryPanel.Controls.Add(searchLabel);

            searchBox = new TextBox
            {
                Location = new Point(200, 183),
                Size = new Size(150, 26),
                Font = UiStyle.BaseFont,
                PlaceholderText = "輸入關鍵字..."
            };
            searchBox.TextChanged += SearchBox_TextChanged;
            inventoryPanel.Controls.Add(searchBox); 
            UiStyle.StyleTextBox(searchBox);

            filterComboBox = new ComboBox
            {
                Location = new Point(360, 183),
                Size = new Size(100, 26),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            filterComboBox.Items.AddRange(new[] { "全部", "低庫存", "零庫存", "負庫存" });
            filterComboBox.SelectedIndex = 0;
            filterComboBox.SelectedIndexChanged += FilterComboBox_SelectedIndexChanged;
            inventoryPanel.Controls.Add(filterComboBox); 
            UiStyle.StyleComboBox(filterComboBox);

            inventoryGridView = new DataGridView
            {
                Location = new Point(10, 220),
                Size = new Size(480, 170),  // 調整寬度配合面板
                AllowUserToAddRows = false,
                ReadOnly = true,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            };
            inventoryPanel.Controls.Add(inventoryGridView); 
            UiStyle.StyleDataGridView(inventoryGridView);

            exportButton = new Button
            {
                Text = "📋 匯出庫存報表",
                Location = new Point(10, 400),
                Size = new Size(120, 32)
            };
            exportButton.Click += ExportButton_Click;
            inventoryPanel.Controls.Add(exportButton); 
            UiStyle.StyleButtonDanger(exportButton);

            summaryLabel = new Label
            {
                Text = "庫存統計：0 項目",
                Location = new Point(140, 407),
                Size = new Size(200, 20),
                Font = new Font(UiStyle.BaseFont.FontFamily, 8.5F),
                ForeColor = Color.FromArgb(107, 114, 128)
            };
            inventoryPanel.Controls.Add(summaryLabel);

            btnExportRequirement = new Button
            {
                Text = "🛒 匯出缺料清單",
                Location = new Point(350, 400),  // 調整位置
                Size = new Size(120, 32)
            };
            btnExportRequirement.Click += BtnExportRequirement_Click;
            inventoryPanel.Controls.Add(btnExportRequirement); 
            UiStyle.StyleButtonDanger(btnExportRequirement);

            this.Width = 1420;  // 恢復原來的總寬度
        }

        /// <summary>
        /// 設定公司庫存檔（僅供預覽）
        /// </summary>
        private void SetMainFileButton_Click(object? sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel 檔案 (*.xls;*.xlsx;*.xlsm;*.xlsb)|*.xls;*.xlsx;*.xlsm;*.xlsb";
            openFileDialog.Title = "選擇公司庫存檔（預覽用/基準）";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                mainFilePath = openFileDialog.FileName;
                SaveMainFileConfig();
                UpdateMainFileLabel();

                try
                {
                    InventoryBaselineManager.LoadSnapshot(mainFilePath);
                    var name = Path.GetFileName(mainFilePath);
                    lblSnapshotInfo.Text = $"基準: {InventoryBaselineManager.SnapshotTime:MM-dd HH:mm} {name}";
                    lblSnapshotInfo.ForeColor = Color.FromArgb(40, 167, 69);
                    MessageBox.Show("公司庫存檔設定成功，並已作為缺料判斷基準！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"載入基準失敗：{ex.Message}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
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
                    SetSnapshotFromMainFile();
                }
            }
            catch { }
        }

        /// <summary>
        /// 儲存公司庫存檔設定
        /// </summary>
        private void SaveMainFileConfig()
        {
            try { File.WriteAllText(MAIN_FILE_CONFIG, mainFilePath); } catch { }
        }

        /// <summary>
        /// 更新公司庫存檔標籤
        /// </summary>
        private void UpdateMainFileLabel()
        {
            if (!string.IsNullOrEmpty(mainFilePath) && File.Exists(mainFilePath))
            {
                mainFileLabel.Text = $"✅ {Path.GetFileName(mainFilePath)}";
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
            catch (Exception ex) { }
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
                inventoryGridView.DataSource = null;
                inventoryGridView.Columns.Clear();

                excelApp = new Excel.Application();
                excelApp.Visible = false;
                workbook = excelApp.Workbooks.Open(mainFilePath);
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                Excel.Range range = worksheet.UsedRange;

                object[,] data = (object[,])range.Value;
                int rowCount = range.Rows.Count;
                int colCount = range.Columns.Count;

                var dt = new DataTable();
                var usedColumnNames = new HashSet<string>();

                for (int i = 1; i <= colCount; i++)
                {
                    string originalColumnName = data[1, i]?.ToString()?.Trim() ?? $"Column{i}";
                    string columnName = originalColumnName;

                    int counter = 1;
                    while (usedColumnNames.Contains(columnName))
                    {
                        columnName = $"{originalColumnName}_{counter}";
                        counter++;
                    }

                    if (string.IsNullOrWhiteSpace(columnName))
                        columnName = $"Column{i}";

                    usedColumnNames.Add(columnName);
                    dt.Columns.Add(columnName);
                }

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

                inventoryGridView.DataSource = dt;
                exportButton.Enabled = true;

                UpdateInventorySummary(dt);
                SetInventoryColors();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"預覽時發生錯誤：{ex.Message}\n\n詳細錯誤：{ex.StackTrace}", "錯誤",
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
                        filterConditions.Add($"Convert([{columnName}], 'System.String') LIKE '%{searchText.Replace("'", "''")}%'");
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
                            row.DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 238);
                            row.DefaultCellStyle.ForeColor = Color.FromArgb(220, 53, 69);
                        }
                        else if (stock == 0)
                        {
                            row.DefaultCellStyle.BackColor = Color.FromArgb(255, 243, 205);
                            row.DefaultCellStyle.ForeColor = Color.FromArgb(255, 193, 7);
                        }
                        else if (stock < 10)
                        {
                            row.DefaultCellStyle.BackColor = Color.FromArgb(255, 248, 225);
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

            summaryLabel.Text = $"📊 縂計：{totalItems} | 🔻 低庫存：{lowStock} | 🔴 零庫存：{zeroStock} | ⚠️ 負庫存：{negativeStock}";
        }

        /// <summary>
        /// 匯出庫存報表
        /// </summary>
        private void ExportButton_Click(object? sender, EventArgs e)
        {
            if (inventoryGridView.DataSource == null) return;

            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel 檔案|*.xlsx",
                FileName = "庫存報表_" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss")
            };

            if (saveDialog.ShowDialog() != DialogResult.OK) return;

            Excel.Application? excelApp = null;
            Excel.Workbook? workbook = null;
            Excel.Worksheet? worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Add();
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];

                for (int i = 0; i < inventoryGridView.Columns.Count; i++)
                {
                    Excel.Range headerCell = (Excel.Range)worksheet.Cells[1, i + 1];
                    headerCell.Value = inventoryGridView.Columns[i].HeaderText;
                    headerCell.Font.Bold = true;
                    headerCell.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                }

                for (int r = 0; r < inventoryGridView.Rows.Count; r++)
                {
                    for (int c = 0; c < inventoryGridView.Columns.Count; c++)
                    {
                        Excel.Range dataCell = (Excel.Range)worksheet.Cells[r + 2, c + 1];
                        dataCell.Value = inventoryGridView.Rows[r].Cells[c].Value ?? "";
                    }
                }

                worksheet.Columns.AutoFit();
                workbook.SaveAs(saveDialog.FileName);
                MessageBox.Show("匯出成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"匯出時發生錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                try
                {
                    if (worksheet != null) { Marshal.ReleaseComObject(worksheet); worksheet = null; }
                    if (workbook != null) { workbook.Close(); Marshal.ReleaseComObject(workbook); workbook = null; }
                    if (excelApp != null) { excelApp.Quit(); Marshal.ReleaseComObject(excelApp); excelApp = null; }
                }
                catch { }
            }
        }

        private void ProcessReplenishment(string outputFolder)
        {
            if (string.IsNullOrEmpty(mainFilePath) || !File.Exists(mainFilePath)) return;

            try
            {
                labelCurrentFile.Text = "🔍 正在檢查主檔案的最終負庫存...";

                string mainFileToProcess = mainFilePath;

                if (!string.IsNullOrEmpty(outputFolder) && Directory.Exists(outputFolder))
                {
                    var mainFiles = Directory.GetFiles(outputFolder, "*_main*")
                        .Where(f => f.EndsWith(".xlsx", System.StringComparison.OrdinalIgnoreCase) || f.EndsWith(".xls", System.StringComparison.OrdinalIgnoreCase))
                        .OrderByDescending(f => File.GetLastWriteTime(f))
                        .ToArray();

                    if (mainFiles.Length > 0)
                        mainFileToProcess = mainFiles[0];
                }

                labelCurrentFile.Text = $"⚡ 正在檢查主檔案負庫存：{Path.GetFileName(mainFileToProcess)}";
                ReplenishmentManager.ProcessMainFileNegativeInventory(mainFileToProcess, outputFolder, progressBar1, labelCurrentFile);

                labelCurrentFile.Text = "✅ 發料處理完成";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"發料處理時發生錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                labelCurrentFile.Text = "❌ 發料處理失敗";
            }
            finally
            {
                if (progressBar1.Maximum > 0)
                    progressBar1.Value = progressBar1.Maximum;
            }
        }

        private string GetTodayFolderPath()
        {
            return @"\\St-nas\個人資料夾\Andy\excel\" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm");
        }

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
                        return todayFolders[0];
                }
                return string.Empty;
            }
            catch { return string.Empty; }
        }

        private void BtnLoadSnapshot_Click(object? sender, EventArgs e)
        {
            try
            {
                using var ofd = new OpenFileDialog
                {
                    Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx",
                    Title = "選擇基準庫存檔（缺料判斷用）"
                };
                if (ofd.ShowDialog() != DialogResult.OK) return;
                InventoryBaselineManager.LoadSnapshot(ofd.FileName);
                var name = Path.GetFileName(InventoryBaselineManager.SnapshotSourceFile ?? ofd.FileName);
                lblSnapshotInfo.Text = $"基準: {InventoryBaselineManager.SnapshotTime:MM-dd HH:mm} {name}";
                lblSnapshotInfo.ForeColor = Color.FromArgb(40, 167, 69);
                MessageBox.Show("基準庫存載入完成", "成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入失敗: {ex.Message}");
            }
        }

        private void BtnExportRequirement_Click(object? sender, EventArgs e)
        {
            try
            {
                PurchaseRequirementManager.ExportCsv();
                MessageBox.Show("已匯出 purchase_requirements.csv", "完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"匯出失敗: {ex.Message}");
            }
        }

        private void LoadOutputFolderConfig()
        {
            try
            {
                if (File.Exists(OUTPUT_CONFIG))
                {
                    var path = File.ReadAllText(OUTPUT_CONFIG).Trim();
                    if (!string.IsNullOrWhiteSpace(path))
                    {
                        outputFolderPath = path;
                        ExcelMergerApi.CustomOutputBaseFolder = outputFolderPath;
                    }
                }
            }
            catch { }
            UpdateOutputFolderLabel();
        }

        private void SaveOutputFolderConfig()
        {
            try { File.WriteAllText(OUTPUT_CONFIG, outputFolderPath); } catch { }
        }

        private void UpdateOutputFolderLabel()
        {
            if (!string.IsNullOrWhiteSpace(outputFolderPath))
            {
                if (labelOutputFolder != null)
                {
                    labelOutputFolder.Text = $"📁 輸出路徑：{outputFolderPath}";
                    labelOutputFolder.ForeColor = Color.FromArgb(40, 167, 69);
                    toolTip1.SetToolTip(labelOutputFolder, outputFolderPath);
                }
            }
            else
            {
                if (labelOutputFolder != null)
                {
                    labelOutputFolder.Text = "📁 輸出路徑： (未設定，使用預設)";
                    labelOutputFolder.ForeColor = Color.FromArgb(108, 117, 125);
                }
            }
        }

        private void buttonBrowseOutput_Click(object? sender, EventArgs e)
        {
            using var fbd = new FolderBrowserDialog
            {
                Description = "選擇輸出根目錄 (將建立時間戳子資料夾)",
                ShowNewFolderButton = true
            };
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                outputFolderPath = fbd.SelectedPath;
                ExcelMergerApi.CustomOutputBaseFolder = outputFolderPath;
                SaveOutputFolderConfig();
                UpdateOutputFolderLabel();
            }
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            LoadOutputFolderConfig();
            if (buttonBrowseOutput != null)
            {
                buttonBrowseOutput.Click -= buttonBrowseOutput_Click;
                buttonBrowseOutput.Click += buttonBrowseOutput_Click;
            }
        }
        
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            // 重新計算完成面板位置
            if (completionPanel != null && completionPanel.Visible)
            {
                completionPanel.Location = new Point(
                    (this.ClientSize.Width - completionPanel.Width) / 2,
                    (this.ClientSize.Height - completionPanel.Height) / 2
                );
            }
        }
        #endregion
    }
}
