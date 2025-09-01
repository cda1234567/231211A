using System;
using System.Drawing;
using System.Windows.Forms;

namespace _231211A
{
    /// <summary>
    /// �o�ƽT�{��ܮ�
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
            this.Text = "�o�ƽT�{";
            this.Size = new Size(450, 350);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.FromArgb(248, 249, 250);
            this.TopMost = false; // �אּ�D����m���A���H Owner �m��W�h

            // ���D
            lblTitle = new Label
            {
                Text = "�ݭn�o�ƶ���",
                Location = new Point(20, 20),
                Size = new Size(400, 30),
                Font = new Font("Microsoft YaHei", 14, FontStyle.Bold),
                ForeColor = Color.FromArgb(220, 53, 69)
            };
            this.Controls.Add(lblTitle);

            // �ɮׯ���
            lblFileIndex = new Label
            {
                Text = $"�ثe�B�z�G�� {_fileIndex} �ӭt�w�s����",
                Location = new Point(20, 55),
                Size = new Size(400, 20),
                Font = new Font("Microsoft YaHei", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 123, 255)
            };
            this.Controls.Add(lblFileIndex);

            // �Ƹ�
            lblPartNumber = new Label
            {
                Text = $"�Ƹ��G{_item.PartNumber}",
                Location = new Point(20, 85),
                Size = new Size(400, 25),
                Font = new Font("Microsoft YaHei", 10),
                ForeColor = Color.FromArgb(33, 37, 41)
            };
            this.Controls.Add(lblPartNumber);

            // ����
            lblDescription = new Label
            {
                Text = $"�����G{_item.Description}",
                Location = new Point(20, 115),
                Size = new Size(400, 25),
                Font = new Font("Microsoft YaHei", 10),
                ForeColor = Color.FromArgb(33, 37, 41)
            };
            this.Controls.Add(lblDescription);

            // �ثe�w�s
            lblCurrentStock = new Label
            {
                Text = $"�ثe�w�s�G{_item.CurrentStock:F2}",
                Location = new Point(20, 145),
                Size = new Size(400, 25),
                Font = new Font("Microsoft YaHei", 10),
                ForeColor = Color.FromArgb(220, 53, 69)
            };
            this.Controls.Add(lblCurrentStock);

            // �ݭn�o�Ƽƶq
            lblRequiredAmount = new Label
            {
                Text = $"�ݭn�o�Ƽƶq�G{_item.ShortageAmount:F2}",
                Location = new Point(20, 175),
                Size = new Size(400, 25),
                Font = new Font("Microsoft YaHei", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(220, 53, 69)
            };
            this.Controls.Add(lblRequiredAmount);

            // �ڭn�o�Ƽƶq����
            lblDispatchAmount = new Label
            {
                Text = "�ڭn�o�Ƽƶq�G",
                Location = new Point(20, 210),
                Size = new Size(150, 25),
                Font = new Font("Microsoft YaHei", 10),
                ForeColor = Color.FromArgb(33, 37, 41)
            };
            this.Controls.Add(lblDispatchAmount);

            // �ڭn�o�Ƽƶq��J��
            txtDispatchAmount = new TextBox
            {
                Location = new Point(180, 208),
                Size = new Size(120, 25),
                Font = new Font("Microsoft YaHei", 10),
                Text = _item.ShortageAmount.ToString("F2")
            };
            txtDispatchAmount.KeyPress += TxtDispatchAmount_KeyPress;
            this.Controls.Add(txtDispatchAmount);

            // �T�{���s
            btnOK = new Button
            {
                Text = "�T�{�o��",
                Location = new Point(80, 260),
                Size = new Size(90, 35),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei", 9)
            };
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            // ���L���s
            btnSkip = new Button
            {
                Text = "���L����",
                Location = new Point(180, 260),
                Size = new Size(90, 35),
                BackColor = Color.FromArgb(255, 193, 7),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei", 9)
            };
            btnSkip.Click += BtnSkip_Click;
            this.Controls.Add(btnSkip);

            // �������s
            btnCancel = new Button
            {
                Text = "��������",
                Location = new Point(280, 260),
                Size = new Size(90, 35),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei", 9)
            };
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);

            // �]�w�w�]���s
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void LoadData()
        {
            // ��s���D��ܥثe�B�z���ةM�Ƹ�
            lblTitle.Text = $"�ݭn�o�ƶ��ءG{_item.PartNumber}";
            
            // ��s�ɮׯ��޸�T
            lblFileIndex.Text = $"�ثe�B�z�G�� {_fileIndex} �ӭt�w�s����";
            
            // �w�]�o�Ƽƶq�N�O�ݭn���ƶq�]�t�w�s����ȡ^
            // _item.CurrentStock �O�D�ɮ׳̥k�䪺�t�w�s�ƶq
            // _item.ShortageAmount �O�t�w�s������ȡ]�ݭn�ɪ��ּ̤ƶq�^
            double suggestedAmount = _item.ShortageAmount;
            
            txtDispatchAmount.Text = suggestedAmount.ToString("F0");
            
            // �N��J�J�I�]�b��J�بÿ�ܩҦ���r
            txtDispatchAmount.Focus();
            txtDispatchAmount.SelectAll();
        }

        private void TxtDispatchAmount_KeyPress(object? sender, KeyPressEventArgs e)
        {
            // �u���\�Ʀr�B�p���I�M������
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
            }

            // �u���\�@�Ӥp���I
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
                MessageBox.Show("�п�J���Ī��o�Ƽƶq�]�����j��0�^", "��J���~", 
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