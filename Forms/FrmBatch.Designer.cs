namespace Office_File_Explorer.Forms
{
    partial class FrmBatch
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmBatch));
            this.label1 = new System.Windows.Forms.Label();
            this.TxbDirectoryPath = new System.Windows.Forms.TextBox();
            this.BtnBrowseDirectory = new System.Windows.Forms.Button();
            this.lstOutput = new System.Windows.Forms.ListBox();
            this.BtnChangeCustomProps = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdoPowerPoint = new System.Windows.Forms.RadioButton();
            this.rdoExcel = new System.Windows.Forms.RadioButton();
            this.rdoWord = new System.Windows.Forms.RadioButton();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.BtnChangeTheme = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 34);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(113, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "Directory Path:";
            // 
            // TxbDirectoryPath
            // 
            this.TxbDirectoryPath.Location = new System.Drawing.Point(134, 28);
            this.TxbDirectoryPath.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.TxbDirectoryPath.Name = "TxbDirectoryPath";
            this.TxbDirectoryPath.Size = new System.Drawing.Size(523, 26);
            this.TxbDirectoryPath.TabIndex = 1;
            // 
            // BtnBrowseDirectory
            // 
            this.BtnBrowseDirectory.Location = new System.Drawing.Point(668, 26);
            this.BtnBrowseDirectory.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.BtnBrowseDirectory.Name = "BtnBrowseDirectory";
            this.BtnBrowseDirectory.Size = new System.Drawing.Size(159, 35);
            this.BtnBrowseDirectory.TabIndex = 2;
            this.BtnBrowseDirectory.Text = "...Choose Location";
            this.BtnBrowseDirectory.UseVisualStyleBackColor = true;
            this.BtnBrowseDirectory.Click += new System.EventHandler(this.BtnBrowseDirectory_Click);
            // 
            // lstOutput
            // 
            this.lstOutput.FormattingEnabled = true;
            this.lstOutput.ItemHeight = 20;
            this.lstOutput.Location = new System.Drawing.Point(10, 29);
            this.lstOutput.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.lstOutput.Name = "lstOutput";
            this.lstOutput.Size = new System.Drawing.Size(1134, 524);
            this.lstOutput.TabIndex = 3;
            // 
            // BtnChangeCustomProps
            // 
            this.BtnChangeCustomProps.Location = new System.Drawing.Point(10, 29);
            this.BtnChangeCustomProps.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.BtnChangeCustomProps.Name = "BtnChangeCustomProps";
            this.BtnChangeCustomProps.Size = new System.Drawing.Size(212, 35);
            this.BtnChangeCustomProps.TabIndex = 4;
            this.BtnChangeCustomProps.Text = "Change Custom Props";
            this.BtnChangeCustomProps.UseVisualStyleBackColor = true;
            this.BtnChangeCustomProps.Click += new System.EventHandler(this.BtnChangeCustomProps_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rdoPowerPoint);
            this.groupBox1.Controls.Add(this.rdoExcel);
            this.groupBox1.Controls.Add(this.rdoWord);
            this.groupBox1.Location = new System.Drawing.Point(22, 18);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Size = new System.Drawing.Size(310, 74);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "File Type:";
            // 
            // rdoPowerPoint
            // 
            this.rdoPowerPoint.AutoSize = true;
            this.rdoPowerPoint.Location = new System.Drawing.Point(180, 29);
            this.rdoPowerPoint.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.rdoPowerPoint.Name = "rdoPowerPoint";
            this.rdoPowerPoint.Size = new System.Drawing.Size(114, 24);
            this.rdoPowerPoint.TabIndex = 2;
            this.rdoPowerPoint.Text = "PowerPoint";
            this.rdoPowerPoint.UseVisualStyleBackColor = true;
            this.rdoPowerPoint.CheckedChanged += new System.EventHandler(this.rdoPowerPoint_CheckedChanged);
            // 
            // rdoExcel
            // 
            this.rdoExcel.AutoSize = true;
            this.rdoExcel.Location = new System.Drawing.Point(94, 29);
            this.rdoExcel.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.rdoExcel.Name = "rdoExcel";
            this.rdoExcel.Size = new System.Drawing.Size(72, 24);
            this.rdoExcel.TabIndex = 1;
            this.rdoExcel.Text = "Excel";
            this.rdoExcel.UseVisualStyleBackColor = true;
            this.rdoExcel.CheckedChanged += new System.EventHandler(this.rdoExcel_CheckedChanged);
            // 
            // rdoWord
            // 
            this.rdoWord.AutoSize = true;
            this.rdoWord.Checked = true;
            this.rdoWord.Location = new System.Drawing.Point(9, 29);
            this.rdoWord.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.rdoWord.Name = "rdoWord";
            this.rdoWord.Size = new System.Drawing.Size(72, 24);
            this.rdoWord.TabIndex = 0;
            this.rdoWord.TabStop = true;
            this.rdoWord.Text = "Word";
            this.rdoWord.UseVisualStyleBackColor = true;
            this.rdoWord.CheckedChanged += new System.EventHandler(this.rdoWord_CheckedChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.TxbDirectoryPath);
            this.groupBox2.Controls.Add(this.BtnBrowseDirectory);
            this.groupBox2.Location = new System.Drawing.Point(342, 18);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox2.Size = new System.Drawing.Size(840, 74);
            this.groupBox2.TabIndex = 6;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "File Location:";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.BtnChangeTheme);
            this.groupBox3.Controls.Add(this.BtnChangeCustomProps);
            this.groupBox3.Location = new System.Drawing.Point(22, 697);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox3.Size = new System.Drawing.Size(1160, 91);
            this.groupBox3.TabIndex = 7;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Batch Commands";
            // 
            // BtnChangeTheme
            // 
            this.BtnChangeTheme.Location = new System.Drawing.Point(231, 29);
            this.BtnChangeTheme.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.BtnChangeTheme.Name = "BtnChangeTheme";
            this.BtnChangeTheme.Size = new System.Drawing.Size(147, 35);
            this.BtnChangeTheme.TabIndex = 5;
            this.BtnChangeTheme.Text = "Change Theme";
            this.BtnChangeTheme.UseVisualStyleBackColor = true;
            this.BtnChangeTheme.Click += new System.EventHandler(this.BtnChangeTheme_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.lstOutput);
            this.groupBox4.Location = new System.Drawing.Point(22, 102);
            this.groupBox4.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox4.Size = new System.Drawing.Size(1160, 585);
            this.groupBox4.TabIndex = 8;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Files";
            // 
            // FrmBatch
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1200, 802);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.Name = "FrmBatch";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Batch File Processing";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox TxbDirectoryPath;
        private System.Windows.Forms.Button BtnBrowseDirectory;
        private System.Windows.Forms.ListBox lstOutput;
        private System.Windows.Forms.Button BtnChangeCustomProps;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rdoPowerPoint;
        private System.Windows.Forms.RadioButton rdoExcel;
        private System.Windows.Forms.RadioButton rdoWord;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button BtnChangeTheme;
    }
}