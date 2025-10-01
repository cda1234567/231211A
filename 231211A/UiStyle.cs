using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace _231211A
{
    public static class UiStyle
    {
        // Apple-style Color Palette
        public static readonly Color WindowBackground = Color.FromArgb(242, 242, 247); // Light Gray
        public static readonly Color ControlBackground = Color.White;
        public static readonly Color TextColor = Color.FromArgb(28, 28, 30); // Near Black
        public static readonly Color SecondaryTextColor = Color.FromArgb(142, 142, 147); // Gray
        public static readonly Color AccentColor = Color.FromArgb(0, 122, 255); // Blue
        public static readonly Color DestructiveColor = Color.FromArgb(255, 59, 48); // Red
        public static readonly Color SeparatorColor = Color.FromArgb(229, 229, 234);

        // Fonts
        public static readonly Font BaseFont = new Font("Segoe UI", 9F, FontStyle.Regular);
        public static readonly Font BoldFont = new Font("Segoe UI", 9F, FontStyle.Bold);
        public static readonly Font TitleFont = new Font("Segoe UI", 14F, FontStyle.Bold);

        public static void ApplyTheme(Form form)
        {
            form.BackColor = WindowBackground;
            form.Font = BaseFont;
            form.ForeColor = TextColor;

            ApplyThemeToControls(form.Controls);
        }

        private static void ApplyThemeToControls(Control.ControlCollection controls)
        {
            foreach (Control control in controls)
            {
                if (control is Button button)
                {
                    button.FlatStyle = FlatStyle.Flat;
                    button.FlatAppearance.BorderSize = 0;
                    button.Font = BoldFont;
                    button.ForeColor = Color.White;
                    button.Padding = new Padding(5);
                    button.MinimumSize = new Size(0, 30);

                    // Differentiate by name or tag for specific styling
                    if (button.Name.Contains("Export") || button.Name.Contains("OK") || button.Name.Contains("Execute"))
                    {
                        button.BackColor = AccentColor; // Primary action
                    }
                    else if (button.Name.Contains("Remove") || button.Name.Contains("Cancel"))
                    {
                        button.BackColor = DestructiveColor; // Destructive action
                    }
                    else
                    {
                        button.BackColor = SecondaryTextColor; // Secondary action
                        button.ForeColor = TextColor;
                    }
                    ApplyRoundedCorners(button, 8);
                }
                else if (control is Label label)
                {
                    label.BackColor = Color.Transparent;
                    label.Font = BaseFont;
                    label.ForeColor = TextColor;
                }
                else if (control is TextBox textBox)
                {
                    textBox.BackColor = ControlBackground;
                    textBox.ForeColor = TextColor;
                    textBox.Font = BaseFont;
                    textBox.BorderStyle = BorderStyle.FixedSingle;
                }
                else if (control is ListBox listBox)
                {
                    listBox.BackColor = ControlBackground;
                    listBox.ForeColor = TextColor;
                    listBox.Font = BaseFont;
                    listBox.BorderStyle = BorderStyle.None;
                }
                else if (control is DataGridView dgv)
                {
                    dgv.BackgroundColor = ControlBackground;
                    dgv.BorderStyle = BorderStyle.None;
                    dgv.GridColor = SeparatorColor;
                    dgv.ColumnHeadersDefaultCellStyle.BackColor = WindowBackground;
                    dgv.ColumnHeadersDefaultCellStyle.ForeColor = TextColor;
                    dgv.ColumnHeadersDefaultCellStyle.Font = BoldFont;
                    dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
                    dgv.EnableHeadersVisualStyles = false;
                    dgv.DefaultCellStyle.BackColor = ControlBackground;
                    dgv.DefaultCellStyle.ForeColor = TextColor;
                    dgv.DefaultCellStyle.Font = BaseFont;
                    dgv.AlternatingRowsDefaultCellStyle.BackColor = WindowBackground;
                }
                else if (control is Panel panel)
                {
                    panel.BackColor = Color.Transparent;
                }
                else if (control is ComboBox comboBox)
                {
                    comboBox.BackColor = ControlBackground;
                    comboBox.ForeColor = TextColor;
                    comboBox.Font = BaseFont;
                    comboBox.FlatStyle = FlatStyle.Flat;
                }

                if (control.HasChildren)
                {
                    ApplyThemeToControls(control.Controls);
                }
            }
        }

        private static void ApplyRoundedCorners(Control control, int radius)
        {
            // Use a GraphicsPath to create a rounded rectangle region for the control
            control.Paint += (sender, e) =>
            {
                if (sender is Control c)
                {
                    using (var path = new GraphicsPath())
                    {
                        path.AddArc(0, 0, radius, radius, 180, 90);
                        path.AddArc(c.Width - radius, 0, radius, radius, 270, 90);
                        path.AddArc(c.Width - radius, c.Height - radius, radius, radius, 0, 90);
                        path.AddArc(0, c.Height - radius, radius, radius, 90, 90);
                        path.CloseFigure();
                        c.Region = new Region(path);
                    }
                }
            };
        }
    }
}
