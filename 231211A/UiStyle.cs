using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Drawing.Drawing2D;

namespace _231211A
{
    internal static class UiStyle
    {
        public static bool DarkMode { get; private set; } = false;

        public static readonly Color Primary = Color.FromArgb(0, 123, 255);
        public static readonly Color PrimaryHover = Color.FromArgb(0, 105, 217);
        public static readonly Color Success = Color.FromArgb(40, 167, 69);
        public static readonly Color Warning = Color.FromArgb(255, 193, 7);
        public static readonly Color Danger = Color.FromArgb(220, 53, 69);
        public static readonly Color Accent = Color.FromArgb(23, 162, 184);
        public static readonly Color Neutral = Color.FromArgb(108, 117, 125);
        public static readonly Color PanelBg = Color.FromArgb(248, 249, 250);
        public static readonly Color Border = Color.FromArgb(222, 226, 230);

        private static readonly Color DarkBack = Color.FromArgb(32, 34, 37);
        private static readonly Color DarkPanel = Color.FromArgb(44, 47, 51);
        private static readonly Color DarkBorder = Color.FromArgb(64, 68, 75);
        private static readonly Color DarkText = Color.FromArgb(236, 239, 244);
        private static readonly Color DarkSubtle = Color.FromArgb(153, 159, 170);

        private static Font? _baseFont;
        public static Font BaseFont => _baseFont ??= BuildBaseFont();

        private static Font BuildBaseFont()
        {
            string[] preferred = { "Segoe UI", "Microsoft JhengHei UI", "Microsoft JhengHei", "·L³n¥¿¶ÂÅé", "Arial" };
            foreach (var name in preferred)
            {
                try
                {
                    using var f = new Font(name, 9F, FontStyle.Regular, GraphicsUnit.Point);
                    if (string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase))
                        return new Font(f, FontStyle.Regular);
                }
                catch { }
            }
            return new Font(SystemFonts.DefaultFont.FontFamily, 9F, FontStyle.Regular);
        }

        public static void ApplyTheme(Form form)
        {
            form.Font = BaseFont;
            form.BackColor = DarkMode ? DarkBack : Color.White;
            form.DoubleBuffered(true);
            ApplyRecursive(form.Controls);
        }

        private static void ApplyRecursive(Control.ControlCollection controls)
        {
            foreach (Control c in controls)
            {
                switch (c)
                {
                    case Button b:
                        if (b.BackColor.A == 0 || b.BackColor == SystemColors.Control)
                            StyleButtonNeutral(b);
                        b.ForeColor = DarkMode ? DarkText : Color.White;
                        break;
                    case Panel p:
                        p.BackColor = DarkMode ? DarkPanel : PanelBg;
                        ApplyRounded(p, 10);
                        break;
                    case GroupBox g:
                        g.ForeColor = DarkMode ? DarkText : Color.Black;
                        g.Font = BaseFont;
                        break;
                    case Label l:
                        if (l.ForeColor == SystemColors.ControlText || l.ForeColor == Color.Black)
                            l.ForeColor = DarkMode ? DarkText : Color.FromArgb(33, 37, 41);
                        break;
                    case TextBox tb:
                        StyleTextBox(tb);
                        break;
                    case ComboBox cb:
                        StyleComboBox(cb);
                        break;
                    case DataGridView dgv:
                        StyleDataGridView(dgv);
                        break;
                    case ListBox lb:
                        StyleListBox(lb);
                        break;
                }
                if (c.HasChildren) ApplyRecursive(c.Controls);
            }
        }

        public static void ToggleDarkMode(Form? form = null)
        {
            DarkMode = !DarkMode;
            if (form != null) ApplyTheme(form);
        }

        public static void StylePanel(Panel panel)
        {
            panel.BackColor = DarkMode ? DarkPanel : PanelBg;
            panel.BorderStyle = BorderStyle.FixedSingle;
            ApplyRounded(panel, 10);
        }

        public static void StyleDataGridView(DataGridView dgv)
        {
            dgv.EnableHeadersVisualStyles = false;
            dgv.BackgroundColor = DarkMode ? DarkPanel : Color.White;
            dgv.BorderStyle = BorderStyle.FixedSingle;
            dgv.GridColor = DarkMode ? DarkBorder : Border;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = DarkMode ? DarkBorder : Color.FromArgb(233, 236, 239);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = DarkMode ? DarkText : Color.FromArgb(33, 37, 41);
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font(BaseFont, FontStyle.Bold);
            dgv.DefaultCellStyle.Font = BaseFont;
            dgv.DefaultCellStyle.BackColor = DarkMode ? DarkPanel : Color.White;
            dgv.DefaultCellStyle.ForeColor = DarkMode ? DarkText : Color.Black;
            dgv.DefaultCellStyle.SelectionBackColor = Primary;
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = DarkMode ? Color.FromArgb(52, 55, 59) : Color.FromArgb(245, 247, 250);
            dgv.AlternatingRowsDefaultCellStyle.ForeColor = DarkMode ? DarkText : Color.Black;
            dgv.RowHeadersVisible = false;
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
        }

        public static void StyleListBox(ListBox list)
        {
            list.Font = BaseFont;
            list.BorderStyle = BorderStyle.FixedSingle;
            list.BackColor = DarkMode ? DarkPanel : Color.White;
            list.ForeColor = DarkMode ? DarkText : Color.Black;
            list.IntegralHeight = false;
        }

        public static void StyleComboBox(ComboBox cb)
        {
            cb.Font = BaseFont;
            cb.BackColor = DarkMode ? DarkPanel : Color.White;
            cb.ForeColor = DarkMode ? DarkText : Color.Black;
            cb.FlatStyle = FlatStyle.Standard;
            ApplyRounded(cb, 6);
        }

        public static void StyleTextBox(TextBox tb)
        {
            tb.Font = BaseFont;
            tb.BorderStyle = BorderStyle.FixedSingle;
            tb.BackColor = DarkMode ? DarkPanel : Color.White;
            tb.ForeColor = DarkMode ? DarkText : Color.Black;
            ApplyRounded(tb, 6);
        }

        public static void StyleLabel(Label lbl, bool title = false, bool subtle = false)
        {
            if (title)
            {
                lbl.Font = new Font(BaseFont.FontFamily, 12F, FontStyle.Bold);
                lbl.ForeColor = DarkMode ? DarkText : Color.FromArgb(33, 37, 41);
            }
            else if (subtle)
            {
                lbl.ForeColor = DarkMode ? DarkSubtle : Neutral;
            }
            else
            {
                lbl.ForeColor = DarkMode ? DarkText : Color.FromArgb(33, 37, 41);
            }
        }

        public static void StyleProgressBar(ProgressBar pb)
        {
            pb.ForeColor = Primary;
        }

        public static void StyleButtonPrimary(Button b) => StyleButtonBase(b, Primary);
        public static void StyleButtonSuccess(Button b) => StyleButtonBase(b, Success);
        public static void StyleButtonDanger(Button b) => StyleButtonBase(b, Danger);
        public static void StyleButtonWarning(Button b)
        {
            StyleButtonBase(b, Warning);
            b.ForeColor = Color.Black;
        }
        public static void StyleButtonNeutral(Button b) => StyleButtonBase(b, Neutral);

        private static void StyleButtonBase(Button b, Color baseColor)
        {
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 0;
            b.BackColor = baseColor;
            b.ForeColor = Color.White;
            b.Font = BaseFont;
            b.Height = Math.Max(30, b.Height);
            b.Padding = new Padding(8, 4, 8, 4);
            ApplyRounded(b, 8);
            var original = baseColor;
            b.MouseEnter += (_, _) => b.BackColor = Darken(original, 0.08f);
            b.MouseLeave += (_, _) => b.BackColor = original;
            b.MouseDown += (_, _) => b.BackColor = Darken(original, 0.15f);
            b.MouseUp += (_, _) => b.BackColor = Darken(original, 0.08f);
        }

        private static void ApplyRounded(Control c, int radius)
        {
            if (radius <= 0) return;
            c.Region = null;
            c.HandleCreated += (_, _) => SetRoundRegion(c, radius);
            c.SizeChanged += (_, _) => SetRoundRegion(c, radius);
            if (c.IsHandleCreated) SetRoundRegion(c, radius);
        }

        private static void SetRoundRegion(Control c, int radius)
        {
            try
            {
                var rect = c.ClientRectangle;
                if (rect.Width <= 0 || rect.Height <= 0) return;
                using GraphicsPath path = new GraphicsPath();
                int d = radius * 2;
                path.StartFigure();
                path.AddArc(rect.X, rect.Y, d, d, 180, 90);
                path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
                path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
                path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                c.Region = new Region(path);
            }
            catch { }
        }

        private static Color Darken(Color c, float ratio)
        {
            int r = (int)(c.R * (1 - ratio));
            int g = (int)(c.G * (1 - ratio));
            int b = (int)(c.B * (1 - ratio));
            return Color.FromArgb(r, g, b);
        }

        public static void DoubleBuffered(this Control control, bool enable)
        {
            try
            {
                var prop = control.GetType().GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                prop?.SetValue(control, enable, null);
            }
            catch { }
        }
    }
}
