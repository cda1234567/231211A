using System;
using System.Drawing;
using System.Windows.Forms;

namespace _231211A
{
    /// <summary>
    /// 缺料確認視窗
    /// </summary>
    public class ShortageConfirmDialog : Form
    {
        private readonly string _part;
        private readonly string _desc;
        private readonly double _snapshot;
        private readonly double _request;
        private readonly double _shortage;

        private Label lblTitle = null!;
        private Label lblPartNumber = null!;
        private Label lblDescription = null!;
        private Label lblStockInfo = null!;
        private Label lblShortage = null!;
        private Button btnMarkPO = null!;
        private Button btnCreateReq = null!;
        private Button btnReInput = null!;
        private Button btnIgnoreOnce = null!;

        public ShortageDecision Decision { get; private set; } = ShortageDecision.None;

        public ShortageConfirmDialog(string part, string desc, double snapshotQty, double requestQty, double shortageQty)
        {
            _part = part; _desc = desc; _snapshot = snapshotQty; _request = requestQty; _shortage = shortageQty;
            Build();
        }

        private void Build()
        {
            Text = "缺料確認";
            Size = new Size(500, 300);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;

            // Apply Apple-style theme
            UiStyle.ApplyTheme(this);

            lblTitle = new Label { Text = "庫存不足", Location = new Point(20, 20), Size = new Size(460, 28), Font = UiStyle.TitleFont, ForeColor = UiStyle.DestructiveColor };
            this.Controls.Add(lblTitle);

            lblPartNumber = new Label { Text = $"料號：{_part}", Location = new Point(20, 60), Size = new Size(460, 22) };
            this.Controls.Add(lblPartNumber);
            lblDescription = new Label { Text = $"說明：{_desc}", Location = new Point(20, 85), Size = new Size(460, 22) };
            this.Controls.Add(lblDescription);
            lblStockInfo = new Label { Text = $"基準庫存: {_snapshot:F0} / 需求: {_request:F0}", Location = new Point(20, 110), Size = new Size(460, 22) };
            this.Controls.Add(lblStockInfo);
            lblShortage = new Label { Text = $"缺料數量：{_shortage:F0}", Location = new Point(20, 135), Size = new Size(460, 22), Font = UiStyle.BoldFont, ForeColor = UiStyle.DestructiveColor };
            this.Controls.Add(lblShortage);

            btnMarkPO = new Button { Name = "MarkPO", Text = "標記已有PO", Location = new Point(20, 180), Size = new Size(120, 35) };
            btnMarkPO.Click += (s, e) => { Decision = ShortageDecision.MarkHasPO; this.DialogResult = DialogResult.OK; this.Close(); };
            this.Controls.Add(btnMarkPO);

            btnCreateReq = new Button { Name = "CreateReq", Text = "產生需求", Location = new Point(150, 180), Size = new Size(120, 35) };
            btnCreateReq.Click += (s, e) => { Decision = ShortageDecision.CreateRequirement; this.DialogResult = DialogResult.OK; this.Close(); };
            this.Controls.Add(btnCreateReq);

            btnReInput = new Button { Name = "ReInput", Text = "重新輸入", Location = new Point(280, 180), Size = new Size(90, 35) };
            btnReInput.Click += (s, e) => { Decision = ShortageDecision.ReInput; this.DialogResult = DialogResult.OK; this.Close(); };
            this.Controls.Add(btnReInput);

            btnIgnoreOnce = new Button { Name = "IgnoreOnce", Text = "忽略本次", Location = new Point(380, 180), Size = new Size(90, 35) };
            btnIgnoreOnce.Click += (s, e) => { Decision = ShortageDecision.IgnoreOnce; this.DialogResult = DialogResult.OK; this.Close(); };
            this.Controls.Add(btnIgnoreOnce);
        }
    }
}
