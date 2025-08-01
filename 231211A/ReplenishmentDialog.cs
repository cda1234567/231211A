using System;
using System.Drawing;
using System.Windows.Forms;

namespace _231211A
{
    /// <summary>
    /// 發料確認對話框
    /// </summary>
    public partial class ReplenishmentDialog : Form
    {
        private Label lblTitle = null!;
        private Label lblPartNumber = null!;
        private Label lblDescription = null!;
        private Label lblCurrentStock = null!;
        private Label lblRequiredAmount = null!;
        private Label lblFileIndex = null!;
        private Label lblDispatchAmount = null!;
        private TextBox txtDispatchAmount = null!;
        private Button btnOK = null!;
        private Button btnSkip = null!;
        private Button btnCancel = null!;
        
        private readonly ReplenishmentItem _item;
        private readonly int _fileIndex;

        public double ReplenishmentQuantity { get; private set; }

        public ReplenishmentDialog(ReplenishmentItem item, int fileIndex)
        {
            _item = item;
            _fileIndex = fileIndex;
            InitializeComponent();
            LoadData();
        }

        private void InitializeComponent()
        {
            this.Text = "發料確認";
            this.Size = new Size(450, 350);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.FromArgb(248, 249, 250);
            this.TopMost = false; // 改為非全域置頂，僅隨 Owner 置於上層

            // 標題
            lblTitle = new Label
            {
                Text = "需要發料項目",
                Location = new Point(20, 20),
                Size = new Size(400, 30),
                Font = new Font("Microsoft YaHei", 14, FontStyle.Bold),
                ForeColor = Color.FromArgb(220, 53, 69)
            };
            this.Controls.Add(lblTitle);

            // 檔案索引
            lblFileIndex = new Label
            {
                Text = $"目前處理：第 {_fileIndex} 個負庫存項目",
                Location = new Point(20, 55),
                Size = new Size(400, 20),
                Font = new Font("Microsoft YaHei", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 123, 255)
            };
            this.Controls.Add(lblFileIndex);

            // 料號
            lblPartNumber = new Label
            {
                Text = $"料號：{_item.PartNumber}",
                Location = new Point(20, 85),
                Size = new Size(400, 25),
                Font = new Font("Microsoft YaHei", 10),
                ForeColor = Color.FromArgb(33, 37, 41)
            };
            this.Controls.Add(lblPartNumber);

            // 說明
            lblDescription = new Label
            {
                Text = $"說明：{_item.Description}",
                Location = new Point(20, 115),
                Size = new Size(400, 25),
                Font = new Font("Microsoft YaHei", 10),
                ForeColor = Color.FromArgb(33, 37, 41)
            };
            this.Controls.Add(lblDescription);

            // 目前庫存
            lblCurrentStock = new Label
            {
                Text = $"目前庫存：{_item.CurrentStock:F2}",
                Location = new Point(20, 145),
                Size = new Size(400, 25),
                Font = new Font("Microsoft YaHei", 10),
                ForeColor = Color.FromArgb(220, 53, 69)
            };
            this.Controls.Add(lblCurrentStock);

            // 需要發料數量
            lblRequiredAmount = new Label
            {
                Text = $"需要發料數量：{_item.ShortageAmount:F2}",
                Location = new Point(20, 175),
                Size = new Size(400, 25),
                Font = new Font("Microsoft YaHei", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(220, 53, 69)
            };
            this.Controls.Add(lblRequiredAmount);

            // 我要發料數量標籤
            lblDispatchAmount = new Label
            {
                Text = "我要發料數量：",
                Location = new Point(20, 210),
                Size = new Size(150, 25),
                Font = new Font("Microsoft YaHei", 10),
                ForeColor = Color.FromArgb(33, 37, 41)
            };
            this.Controls.Add(lblDispatchAmount);

            // 我要發料數量輸入框
            txtDispatchAmount = new TextBox
            {
                Location = new Point(180, 208),
                Size = new Size(120, 25),
                Font = new Font("Microsoft YaHei", 10),
                Text = _item.ShortageAmount.ToString("F2")
            };
            txtDispatchAmount.KeyPress += TxtDispatchAmount_KeyPress;
            this.Controls.Add(txtDispatchAmount);

            // 確認按鈕
            btnOK = new Button
            {
                Text = "確認發料",
                Location = new Point(80, 260),
                Size = new Size(90, 35),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei", 9)
            };
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            // 跳過按鈕
            btnSkip = new Button
            {
                Text = "跳過此項",
                Location = new Point(180, 260),
                Size = new Size(90, 35),
                BackColor = Color.FromArgb(255, 193, 7),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei", 9)
            };
            btnSkip.Click += BtnSkip_Click;
            this.Controls.Add(btnSkip);

            // 取消按鈕
            btnCancel = new Button
            {
                Text = "取消全部",
                Location = new Point(280, 260),
                Size = new Size(90, 35),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei", 9)
            };
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);

            // 設定預設按鈕
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void LoadData()
        {
            // 更新標題顯示目前處理項目和料號
            lblTitle.Text = $"需要發料項目：{_item.PartNumber}";
            
            // 更新檔案索引資訊
            lblFileIndex.Text = $"目前處理：第 {_fileIndex} 個負庫存項目";
            
            // 預設發料數量就是需要的數量（負庫存絕對值）
            // _item.CurrentStock 是主檔案最右邊的負庫存數量
            // _item.ShortageAmount 是負庫存的絕對值（需要補的最少數量）
            double suggestedAmount = _item.ShortageAmount;
            
            txtDispatchAmount.Text = suggestedAmount.ToString("F0");
            
            // 將輸入焦點設在輸入框並選擇所有文字
            txtDispatchAmount.Focus();
            txtDispatchAmount.SelectAll();
        }

        private void TxtDispatchAmount_KeyPress(object? sender, KeyPressEventArgs e)
        {
            // 只允許數字、小數點和控制鍵
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
            }

            // 只允許一個小數點
            if (e.KeyChar == '.' && txtDispatchAmount.Text.Contains('.'))
            {
                e.Handled = true;
            }
        }

        private void BtnOK_Click(object? sender, EventArgs e)
        {
            if (double.TryParse(txtDispatchAmount.Text, out double amount) && amount > 0)
            {
                ReplenishmentQuantity = amount;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show("請輸入有效的發料數量（必須大於0）", "輸入錯誤", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtDispatchAmount.Focus();
                txtDispatchAmount.SelectAll();
            }
        }

        private void BtnSkip_Click(object? sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Ignore;
            this.Close();
        }

        private void BtnCancel_Click(object? sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}