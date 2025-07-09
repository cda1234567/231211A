namespace _231211A
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ListBox listBoxFiles;
        private System.Windows.Forms.Button buttonAddFile;
        private System.Windows.Forms.Button buttonRemoveFile;
        private System.Windows.Forms.Button button2; // 執行
        private System.Windows.Forms.Label label1;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            progressBar1 = new ProgressBar();
            listBoxFiles = new ListBox();
            buttonAddFile = new Button();
            buttonRemoveFile = new Button();
            button2 = new Button();
            label1 = new Label();
            SuspendLayout();
            // 
            // progressBar1
            // 
            progressBar1.Location = new Point(12, 400);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(494, 23);
            progressBar1.TabIndex = 0;
            // 
            // listBoxFiles
            // 
            listBoxFiles.FormattingEnabled = true;
            listBoxFiles.ItemHeight = 15;
            listBoxFiles.Location = new Point(29, 25);
            listBoxFiles.Name = "listBoxFiles";
            listBoxFiles.Size = new Size(458, 319);
            listBoxFiles.TabIndex = 1;
            // 
            // buttonAddFile
            // 
            buttonAddFile.Location = new Point(29, 355);
            buttonAddFile.Name = "buttonAddFile";
            buttonAddFile.Size = new Size(75, 23);
            buttonAddFile.TabIndex = 2;
            buttonAddFile.Text = "新增檔案";
            buttonAddFile.UseVisualStyleBackColor = true;
            buttonAddFile.Click += buttonAddFile_Click;
            // 
            // buttonRemoveFile
            // 
            buttonRemoveFile.Location = new Point(110, 355);
            buttonRemoveFile.Name = "buttonRemoveFile";
            buttonRemoveFile.Size = new Size(75, 23);
            buttonRemoveFile.TabIndex = 3;
            buttonRemoveFile.Text = "移除選取";
            buttonRemoveFile.UseVisualStyleBackColor = true;
            buttonRemoveFile.Click += buttonRemoveFile_Click;
            // 
            // button2
            // 
            button2.Location = new Point(412, 443);
            button2.Name = "button2";
            button2.Size = new Size(75, 23);
            button2.TabIndex = 4;
            button2.Text = "執行";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(0, 464);
            label1.Name = "label1";
            label1.Size = new Size(35, 15);
            label1.TabIndex = 5;
            label1.Text = "Rev2";
            label1.Click += label1_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(518, 478);
            Controls.Add(label1);
            Controls.Add(button2);
            Controls.Add(buttonRemoveFile);
            Controls.Add(buttonAddFile);
            Controls.Add(listBoxFiles);
            Controls.Add(progressBar1);
            Name = "Form1";
            Text = "扣帳軟體";
            ResumeLayout(false);
            PerformLayout();
        }
    }
}
