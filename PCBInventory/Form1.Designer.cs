namespace PCBInventory;

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
    private System.Windows.Forms.Label labelRev4; // 新增的Label元件
    private Label label1;

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
        progressBar1 = new ProgressBar();
        labelProgress = new Label();
        labelCurrentFile = new Label();
        labelRev4 = new Label();
        label1 = new Label();
        SuspendLayout();
        // 
        // listBoxFiles
        // 
        listBoxFiles.FormattingEnabled = true;
        listBoxFiles.ItemHeight = 15;
        listBoxFiles.Location = new Point(12, 25);
        listBoxFiles.Name = "listBoxFiles";
        listBoxFiles.Size = new Size(868, 319);
        listBoxFiles.TabIndex = 1;
        toolTip1.SetToolTip(listBoxFiles, "拖曳檔案到此處以加入清單");
        // 
        // buttonAddFile
        // 
        buttonAddFile.Location = new Point(29, 355);
        buttonAddFile.Name = "buttonAddFile";
        buttonAddFile.Size = new Size(75, 23);
        buttonAddFile.TabIndex = 2;
        buttonAddFile.Text = "新增檔案";
        toolTip1.SetToolTip(buttonAddFile, "點擊以選擇檔案加入清單");
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
        toolTip1.SetToolTip(buttonRemoveFile, "移除清單中選取的檔案");
        buttonRemoveFile.UseVisualStyleBackColor = true;
        buttonRemoveFile.Click += buttonRemoveFile_Click;
        // 
        // buttonExecute
        // 
        buttonExecute.Location = new Point(762, 429);
        buttonExecute.Name = "buttonExecute";
        buttonExecute.Size = new Size(90, 37);
        buttonExecute.TabIndex = 4;
        buttonExecute.Text = "執行";
        toolTip1.SetToolTip(buttonExecute, "開始執行檔案合併");
        buttonExecute.UseVisualStyleBackColor = true;
        buttonExecute.Click += button2_Click;
        // 
        // buttonMoveUp
        // 
        buttonMoveUp.Location = new Point(810, 350);
        buttonMoveUp.Name = "buttonMoveUp";
        buttonMoveUp.Size = new Size(32, 23);
        buttonMoveUp.TabIndex = 6;
        buttonMoveUp.Text = "▲";
        toolTip1.SetToolTip(buttonMoveUp, "將選取的檔案向上移動");
        buttonMoveUp.UseVisualStyleBackColor = true;
        buttonMoveUp.Click += buttonMoveUp_Click;
        // 
        // buttonMoveDown
        // 
        buttonMoveDown.Location = new Point(848, 350);
        buttonMoveDown.Name = "buttonMoveDown";
        buttonMoveDown.Size = new Size(32, 23);
        buttonMoveDown.TabIndex = 7;
        buttonMoveDown.Text = "▼";
        toolTip1.SetToolTip(buttonMoveDown, "將選取的檔案向下移動");
        buttonMoveDown.UseVisualStyleBackColor = true;
        buttonMoveDown.Click += buttonMoveDown_Click;
        // 
        // progressBar1
        // 
        progressBar1.Location = new Point(12, 400);
        progressBar1.Name = "progressBar1";
        progressBar1.Size = new Size(868, 23);
        progressBar1.TabIndex = 0;
        // 
        // labelProgress
        // 
        labelProgress.AutoSize = true;
        labelProgress.Location = new Point(12, 430);
        labelProgress.Name = "labelProgress";
        labelProgress.Size = new Size(61, 15);
        labelProgress.TabIndex = 5;
        labelProgress.Text = "進度：0%";
        // 
        // labelCurrentFile
        // 
        labelCurrentFile.AutoSize = true;
        labelCurrentFile.Location = new Point(12, 460);
        labelCurrentFile.Name = "labelCurrentFile";
        labelCurrentFile.Size = new Size(115, 15);
        labelCurrentFile.TabIndex = 8;
        labelCurrentFile.Text = "目前執行到的檔案：";
        // 
        // labelRev4
        // 
        labelRev4.AutoSize = true;
        labelRev4.Location = new Point(284, 261);
        labelRev4.Name = "labelRev4";
        labelRev4.Size = new Size(35, 15);
        labelRev4.TabIndex = 9;
        labelRev4.Text = "Rev4";
        // 
        // label1
        // 
        label1.AutoSize = true;
        label1.Location = new Point(858, 460);
        label1.Name = "label1";
        label1.Size = new Size(35, 15);
        label1.TabIndex = 8;
        label1.Text = "Rev4";
        label1.Click += label1_Click;
        // 
        // Form1
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(950, 478);
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
        Controls.Add(labelRev4);
        Name = "Form1";
        Text = "PCB庫存管理系統";
        ResumeLayout(false);
        PerformLayout();
    }
}
