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

        private RadioButton rbHasPO = null!;
        private RadioButton rbCreateReq = null!;
        private RadioButton rbReInput = null!;
        private RadioButton rbIgnore = null!;
        private Button btnOk = null!;
        private Button btnCancel = null!;

        public ShortageDecision Decision { get; private set; } = ShortageDecision.None;

        public ShortageConfirmDialog(string part, string desc, double snapshotQty, double requestQty, double shortageQty)
        {
            _part = part; _desc = desc; _snapshot = snapshotQty; _request = requestQty; _shortage = shortageQty;
            Build();
        }

        private void Build()
        {
            Text = "缺料確認";
            Size = new Size(430, 340);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;

            var lbl = new Label
            {
                Text = $"料號: {_part}\n描述: {_desc}\n基準庫存: {_snapshot}\n本次輸入: {_request}\n判定缺料數量: {_shortage}",
                Location = new Point(15, 10),
                Size = new Size(380, 90)
            };
            Controls.Add(lbl);

            rbHasPO = new RadioButton { Text = "已有採購未到 (標記，不列入缺料)", Location = new Point(20, 110), Size = new Size(360, 22) };
            rbCreateReq = new RadioButton { Text = "沒有採購 → 建立缺料需求", Location = new Point(20, 135), Size = new Size(360, 22) };
            rbReInput = new RadioButton { Text = "我輸入錯了 → 返回重輸", Location = new Point(20, 160), Size = new Size(360, 22) };
            rbIgnore = new RadioButton { Text = "忽略一次 (紀錄但不建立需求)", Location = new Point(20, 185), Size = new Size(360, 22) };

            rbCreateReq.Checked = true;

            Controls.Add(rbHasPO);
            Controls.Add(rbCreateReq);
            Controls.Add(rbReInput);
            Controls.Add(rbIgnore);

            btnOk = new Button { Text = "確定", Location = new Point(90, 240), Size = new Size(100, 30), BackColor = Color.FromArgb(40,167,69), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, DialogResult = DialogResult.OK };
            btnCancel = new Button { Text = "取消", Location = new Point(210, 240), Size = new Size(100, 30), BackColor = Color.FromArgb(108,117,125), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, DialogResult = DialogResult.Cancel };

            btnOk.Click += (_, __) =>
            {
                if (rbHasPO.Checked) Decision = ShortageDecision.MarkHasPO;
                else if (rbCreateReq.Checked) Decision = ShortageDecision.CreateRequirement;
                else if (rbReInput.Checked) Decision = ShortageDecision.ReInput;
                else if (rbIgnore.Checked) Decision = ShortageDecision.IgnoreOnce;
                // DialogResult 已預設為 OK
                Close();
            };
            btnCancel.Click += (_, __) => { /* DialogResult 已預設為 Cancel */ Close(); };

            Controls.Add(btnOk);
            Controls.Add(btnCancel);

            // 讓 Enter = 確定、Esc = 取消
            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;
        }
    }
}
