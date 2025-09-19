namespace _231211A
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ListBox listBoxFiles;
        private System.Windows.Forms.Button buttonAddFile;
        private System.Windows.Forms.Button buttonRemoveFile;
        private System.Windows.Forms.Button buttonExecute; // 執行
        private System.Windows.Forms.Label labelProgress;
        private System.Windows.Forms.Button buttonMoveUp;
        private System.Windows.Forms.Button buttonMoveDown;
        private System.Windows.Forms.Label labelCurrentFile;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label labelOutputFolder; // 新增：輸出路徊顯示
        private System.Windows.Forms.Button buttonBrowseOutput; // 新增：選擇輸出路徊

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
            components = new System.ComponentModel.Container();
            toolTip1 = new ToolTip(components);
            listBoxFiles = new ListBox();
            buttonAddFile = new Button();
            buttonRemoveFile = new Button();
            buttonExecute = new Button();
            buttonMoveUp = new Button();
            buttonMoveDown = new Button();
            labelOutputFolder = new Label();
            buttonBrowseOutput = new Button();
            progressBar1 = new ProgressBar();
            labelProgress = new Label();
            labelCurrentFile = new Label();
            label1 = new Label();
            SuspendLayout();
            // 
            // listBoxFiles
            // 
            listBoxFiles.FormattingEnabled = true;
            listBoxFiles.ItemHeight = 15;
            listBoxFiles.Location = new Point(20, 25);
            listBoxFiles.Name = "listBoxFiles";
            listBoxFiles.Size = new Size(850, 274);
            listBoxFiles.TabIndex = 1;
            toolTip1.SetToolTip(listBoxFiles, "拖曳檔案到此處以加入清單");
            // 
            // buttonAddFile
            // 
            buttonAddFile.Location = new Point(20, 315);
            buttonAddFile.Name = "buttonAddFile";
            buttonAddFile.Size = new Size(100, 32);
            buttonAddFile.TabIndex = 2;
            buttonAddFile.Text = "➕ 新增檔案";
            toolTip1.SetToolTip(buttonAddFile, "點擊以選擇檔案加入清單");
            buttonAddFile.UseVisualStyleBackColor = true;
            buttonAddFile.Click += buttonAddFile_Click;
            // 
            // buttonRemoveFile
            // 
            buttonRemoveFile.Location = new Point(153, 315);
            buttonRemoveFile.Name = "buttonRemoveFile";
            buttonRemoveFile.Size = new Size(105, 32);
            buttonRemoveFile.TabIndex = 3;
            buttonRemoveFile.Text = "🗑️ 移除選取";
            toolTip1.SetToolTip(buttonRemoveFile, "移除清單中選取的檔案");
            buttonRemoveFile.UseVisualStyleBackColor = true;
            buttonRemoveFile.Click += buttonRemoveFile_Click;
            // 
            // buttonExecute
            // 
            buttonExecute.Location = new Point(750, 395);
            buttonExecute.Name = "buttonExecute";
            buttonExecute.Size = new Size(120, 40);
            buttonExecute.TabIndex = 4;
            buttonExecute.Text = "🚀 執行";
            toolTip1.SetToolTip(buttonExecute, "開始執行檔案合併");
            buttonExecute.UseVisualStyleBackColor = true;
            buttonExecute.Click += button2_Click;
            // 
            // buttonMoveUp
            // 
            buttonMoveUp.Location = new Point(800, 315);
            buttonMoveUp.Name = "buttonMoveUp";
            buttonMoveUp.Size = new Size(32, 32);
            buttonMoveUp.TabIndex = 6;
            buttonMoveUp.Text = "⬆️";
            toolTip1.SetToolTip(buttonMoveUp, "將選取的檔案向上移動");
            buttonMoveUp.UseVisualStyleBackColor = true;
            buttonMoveUp.Click += buttonMoveUp_Click;
            // 
            // buttonMoveDown
            // 
            buttonMoveDown.Location = new Point(838, 315);
            buttonMoveDown.Name = "buttonMoveDown";
            buttonMoveDown.Size = new Size(32, 32);
            buttonMoveDown.TabIndex = 7;
            buttonMoveDown.Text = "⬇️";
            toolTip1.SetToolTip(buttonMoveDown, "將選取的檔案向下移動");
            buttonMoveDown.UseVisualStyleBackColor = true;
            buttonMoveDown.Click += buttonMoveDown_Click;
            // 
            // labelOutputFolder
            // 
            labelOutputFolder.AutoEllipsis = true;
            labelOutputFolder.Location = new Point(20, 355);
            labelOutputFolder.Name = "labelOutputFolder";
            labelOutputFolder.Size = new Size(650, 20);
            labelOutputFolder.TabIndex = 10;
            labelOutputFolder.Text = "📁 輸出路徑： (未設定，使用預設)";
            toolTip1.SetToolTip(labelOutputFolder, "檔案輸出根目錄");
            // 
            // buttonBrowseOutput
            // 
            buttonBrowseOutput.Location = new Point(20, 375);
            buttonBrowseOutput.Name = "buttonBrowseOutput";
            buttonBrowseOutput.Size = new Size(100, 32);
            buttonBrowseOutput.TabIndex = 11;
            buttonBrowseOutput.Text = "📁 選擇路徑";
            toolTip1.SetToolTip(buttonBrowseOutput, "設定自訂輸出根目錄");
            buttonBrowseOutput.UseVisualStyleBackColor = true;
            // 
            // progressBar1
            // 
            progressBar1.Location = new Point(20, 415);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(714, 8);
            progressBar1.TabIndex = 0;
            // 
            // labelProgress
            // 
            labelProgress.AutoSize = true;
            labelProgress.Location = new Point(20, 430);
            labelProgress.Name = "labelProgress";
            labelProgress.Size = new Size(0, 15);
            labelProgress.TabIndex = 5;
            // 
            // labelCurrentFile
            // 
            labelCurrentFile.AutoSize = true;
            labelCurrentFile.Location = new Point(20, 445);
            labelCurrentFile.Name = "labelCurrentFile";
            labelCurrentFile.Size = new Size(124, 15);
            labelCurrentFile.TabIndex = 8;
            labelCurrentFile.Text = "目前執行到的檔案：";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(780, 457);
            label1.Name = "label1";
            label1.Size = new Size(74, 15);
            label1.TabIndex = 8;
            label1.Text = "Rev2025.9.15";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(1420, 481);
            Controls.Add(buttonBrowseOutput);
            Controls.Add(labelOutputFolder);
            Controls.Add(label1);
            Controls.Add(labelCurrentFile);
            Controls.Add(labelProgress);
            Controls.Add(buttonExecute);
            Controls.Add(buttonRemoveFile);
            Controls.Add(buttonAddFile);
            Controls.Add(listBoxFiles);
            Controls.Add(progressBar1);
            Controls.Add(buttonMoveDown);
            Controls.Add(buttonMoveUp);
            Font = new Font("Segoe UI", 9F);
            MinimumSize = new Size(1380, 520);
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "📋 PCB 扣帳系統 - 庫存管理";
            ResumeLayout(false);
            PerformLayout();
        }
    }
}