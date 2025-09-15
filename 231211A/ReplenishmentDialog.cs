using System;
using System.Drawing;
using System.Windows.Forms;

namespace _231211A
{
    /// <summary>
    /// 發料確認對話框（按確認後自動檢查缺料，顯示基準庫存）
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
        private Label lblSnapshot = null!;
        private Label lblSnapshotInfo = null!;
        private Label lblMOQ = null!;
        private TextBox txtDispatchAmount = null!;
        private Button btnOK = null!;
        private Button btnSkip = null!;
        private Button btnCancel = null!;

        private readonly ReplenishmentItem _item;
        private readonly int _fileIndex;

        public double ReplenishmentQuantity { get; private set; }
        public bool IsShortageConfirmed { get; private set; }
        public ShortageDecision Decision { get; private set; } = ShortageDecision.None;
        public double SnapshotQty { get; set; } // 由外部注入的基準庫存
        public double MOQ { get; set; } // 新增：由外部注入的 MOQ

        public ReplenishmentDialog(ReplenishmentItem item, int fileIndex)
        {
            _item = item;
            _fileIndex = fileIndex;
            InitializeComponent();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            LoadData();
        }

        private void InitializeComponent()
        {
            this.Text = "發料確認";
            this.Size = new Size(520, 420);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.FromArgb(248, 249, 250);

            lblTitle = new Label { Text = "需要發料項目", Location = new Point(20, 15), Size = new Size(460, 28), Font = new Font("Microsoft YaHei", 14, FontStyle.Bold), ForeColor = Color.FromArgb(220, 53, 69) };
            this.Controls.Add(lblTitle);

            lblFileIndex = new Label { Text = $"目前處理：第 {_fileIndex} 個負庫存項目", Location = new Point(20, 48), Size = new Size(460, 18), Font = new Font("Microsoft YaHei", 9, FontStyle.Bold), ForeColor = Color.FromArgb(0, 123, 255) };
            this.Controls.Add(lblFileIndex);

            lblPartNumber = new Label { Text = $"料號：{_item.PartNumber}", Location = new Point(20, 75), Size = new Size(460, 22) };
            this.Controls.Add(lblPartNumber);
            lblDescription = new Label { Text = $"說明：{_item.Description}", Location = new Point(20, 100), Size = new Size(460, 22) };
            this.Controls.Add(lblDescription);
            lblCurrentStock = new Label { Text = $"目前庫存：{_item.CurrentStock:F2}", Location = new Point(20, 125), Size = new Size(460, 22), ForeColor = Color.FromArgb(220, 53, 69) };
            this.Controls.Add(lblCurrentStock);
            lblRequiredAmount = new Label { Text = $"缺口(建議補)數量：{_item.ShortageAmount:F2}", Location = new Point(20, 150), Size = new Size(460, 22), Font = new Font("Microsoft YaHei", 10, FontStyle.Bold), ForeColor = Color.FromArgb(220, 53, 69) };
            this.Controls.Add(lblRequiredAmount);

            // 基準庫存顯示
            lblSnapshot = new Label { Text = "基準庫存(快照)：—", Location = new Point(20, 175), Size = new Size(460, 20), ForeColor = Color.FromArgb(33, 37, 41) };
            this.Controls.Add(lblSnapshot);
            lblSnapshotInfo = new Label { Text = "基準來源/時間：—", Location = new Point(20, 195), Size = new Size(460, 18), ForeColor = Color.FromArgb(108, 117, 125) };
            this.Controls.Add(lblSnapshotInfo);

            // MOQ 顯示
            lblMOQ = new Label { Text = "MOQ：—", Location = new Point(20, 215), Size = new Size(460, 20), ForeColor = Color.FromArgb(33, 37, 41) };
            this.Controls.Add(lblMOQ);

            lblDispatchAmount = new Label { Text = "本次補(發)料數量：", Location = new Point(20, 245), Size = new Size(150, 24) };
            this.Controls.Add(lblDispatchAmount);

            txtDispatchAmount = new TextBox { Location = new Point(180, 243), Size = new Size(120, 25), Text = _item.ShortageAmount.ToString("F0") };
            txtDispatchAmount.KeyPress += TxtDispatchAmount_KeyPress;
            this.Controls.Add(txtDispatchAmount);

            btnOK = new Button { Text = "確認", Location = new Point(70, 310), Size = new Size(90, 35), BackColor = Color.FromArgb(40, 167, 69), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);
            btnSkip = new Button { Text = "跳過", Location = new Point(180, 310), Size = new Size(90, 35), BackColor = Color.FromArgb(108, 117, 125), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnSkip.Click += BtnSkip_Click;
            this.Controls.Add(btnSkip);
            btnCancel = new Button { Text = "取消全部", Location = new Point(290, 310), Size = new Size(90, 35), BackColor = Color.FromArgb(73, 80, 87), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void LoadData()
        {
            lblTitle.Text = $"需要發料項目：{_item.PartNumber}";
            lblFileIndex.Text = $"目前處理：第 {_fileIndex} 個負庫存項目";
            // 顯示基準資訊
            lblSnapshot.Text = $"基準庫存(快照)：{SnapshotQty:F0}";
            var src = InventoryBaselineManager.SnapshotSourceFile != null ? System.IO.Path.GetFileName(InventoryBaselineManager.SnapshotSourceFile) : "—";
            var timeStr = InventoryBaselineManager.SnapshotTime.HasValue ? InventoryBaselineManager.SnapshotTime.Value.ToString("MM-dd HH:mm") : "—";
            lblSnapshotInfo.Text = $"基準來源/時間：{src} / {timeStr}";
            // MOQ
            lblMOQ.Text = $"MOQ：{MOQ:F0}";

            txtDispatchAmount.Focus();
            txtDispatchAmount.SelectAll();
        }

        private void TxtDispatchAmount_KeyPress(object? sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.') e.Handled = true;
            if (e.KeyChar == '.' && txtDispatchAmount.Text.Contains('.')) e.Handled = true;
        }

        private void BtnOK_Click(object? sender, System.EventArgs e)
        {
            if (!double.TryParse(txtDispatchAmount.Text, out double amount) || amount <= 0)
            {
                MessageBox.Show("請輸入有效數量 (>0)");
                txtDispatchAmount.Focus();
                txtDispatchAmount.SelectAll();
                return;
            }

            // 按確認後自動檢查缺料
            double snap = SnapshotQty;
            if (amount > snap)
            {
                double shortage = amount; // 依你的定義：需求量=本次輸入量
                using var dlg = new ShortageConfirmDialog(_item.PartNumber, _item.Description, snap, amount, shortage);
                var dr = dlg.ShowDialog(this);
                if (dr != DialogResult.OK)
                {
                    // 使用者取消缺料處理，回到輸入
                    return;
                }
                Decision = dlg.Decision;
                IsShortageConfirmed = true;
                if (Decision == ShortageDecision.ReInput)
                {
                    // 回到輸入，不關閉
                    txtDispatchAmount.Focus();
                    txtDispatchAmount.SelectAll();
                    return;
                }
                if (Decision == ShortageDecision.MarkHasPO)
                {
                    PurchaseRequirementManager.MarkHasPO(_item.PartNumber);
                }
                else if (Decision == ShortageDecision.CreateRequirement)
                {
                    PurchaseRequirementManager.AddRequirement(_item.PartNumber, _item.Description, shortage, PurchaseRequirementStatus.Open);
                }
                else if (Decision == ShortageDecision.IgnoreOnce)
                {
                    PurchaseRequirementManager.IgnoreOnce(_item.PartNumber, shortage);
                }
            }

            // 不缺或缺料處理已確認 → 回傳 OK
            ReplenishmentQuantity = amount;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BtnSkip_Click(object? sender, System.EventArgs e)
        {
            this.DialogResult = DialogResult.Ignore;
            this.Close();
        }

        private void BtnCancel_Click(object? sender, System.EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }

    public enum ShortageDecision
    {
        None,
        MarkHasPO,
        CreateRequirement,
        ReInput,
        IgnoreOnce
    }
}