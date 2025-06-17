namespace _231211A
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            button1 = new Button();
            button2 = new Button();
            textBox1 = new TextBox();
            listBoxSubFiles = new ListBox();
            button3 = new Button();
            buttonRemove = new Button();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(136, 63);
            button1.Name = "button1";
            button1.Size = new Size(75, 23);
            button1.TabIndex = 0;
            button1.Text = "選擇主檔";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button2
            // 
            button2.Location = new Point(365, 387);
            button2.Name = "button2";
            button2.Size = new Size(75, 23);
            button2.TabIndex = 1;
            button2.Text = "執行";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // textBox1
            //
            textBox1.Location = new Point(39, 94);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(458, 23);
            textBox1.TabIndex = 2;
            textBox1.TextChanged += textBox1_TextChanged;
            //
            // listBoxSubFiles
            //
            listBoxSubFiles.FormattingEnabled = true;
            listBoxSubFiles.ItemHeight = 15;
            listBoxSubFiles.Location = new Point(39, 164);
            listBoxSubFiles.Name = "listBoxSubFiles";
            listBoxSubFiles.Size = new Size(458, 109);
            listBoxSubFiles.TabIndex = 3;
            //
            // button3
            //
            button3.Location = new Point(136, 135);
            button3.Name = "button3";
            button3.Size = new Size(75, 23);
            button3.TabIndex = 4;
            button3.Text = "新增副檔";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            //
            // buttonRemove
            //
            buttonRemove.Location = new Point(231, 135);
            buttonRemove.Name = "buttonRemove";
            buttonRemove.Size = new Size(75, 23);
            buttonRemove.TabIndex = 5;
            buttonRemove.Text = "移除副檔";
            buttonRemove.UseVisualStyleBackColor = true;
            buttonRemove.Click += buttonRemove_Click;
            //
            // Form1
            //
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(buttonRemove);
            Controls.Add(button3);
            Controls.Add(listBoxSubFiles);
            Controls.Add(textBox1);
            Controls.Add(button2);
            Controls.Add(button1);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private Button button2;
        private TextBox textBox1;
        private ListBox listBoxSubFiles;
        private Button button3;
        private Button buttonRemove;
    }
}
