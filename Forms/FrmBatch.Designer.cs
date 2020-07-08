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
            this.BtnConvertStrict = new System.Windows.Forms.Button();
            this.BtnPPTResetPII = new System.Windows.Forms.Button();
            this.BtnFixCorruptRevisions = new System.Windows.Forms.Button();
            this.BtnFixCorruptBookmarks = new System.Windows.Forms.Button();
            this.BtnRemovePII = new System.Windows.Forms.Button();
            this.BtnFixNotesPageSize = new System.Windows.Forms.Button();
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
            this.label1.Location = new System.Drawing.Point(6, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Directory Path:";
            // 
            // TxbDirectoryPath
            // 
            this.TxbDirectoryPath.Location = new System.Drawing.Point(89, 18);
            this.TxbDirectoryPath.Name = "TxbDirectoryPath";
            this.TxbDirectoryPath.Size = new System.Drawing.Size(350, 20);
            this.TxbDirectoryPath.TabIndex = 1;
            this.TxbDirectoryPath.TextChanged += new System.EventHandler(this.TxbDirectoryPath_TextChanged);
            // 
            // BtnBrowseDirectory
            // 
            this.BtnBrowseDirectory.Location = new System.Drawing.Point(445, 17);
            this.BtnBrowseDirectory.Name = "BtnBrowseDirectory";
            this.BtnBrowseDirectory.Size = new System.Drawing.Size(106, 23);
            this.BtnBrowseDirectory.TabIndex = 2;
            this.BtnBrowseDirectory.Text = "...Choose Location";
            this.BtnBrowseDirectory.UseVisualStyleBackColor = true;
            this.BtnBrowseDirectory.Click += new System.EventHandler(this.BtnBrowseDirectory_Click);
            // 
            // lstOutput
            // 
            this.lstOutput.FormattingEnabled = true;
            this.lstOutput.Location = new System.Drawing.Point(7, 19);
            this.lstOutput.Name = "lstOutput";
            this.lstOutput.Size = new System.Drawing.Size(757, 342);
            this.lstOutput.TabIndex = 3;
            // 
            // BtnChangeCustomProps
            // 
            this.BtnChangeCustomProps.Location = new System.Drawing.Point(7, 19);
            this.BtnChangeCustomProps.Name = "BtnChangeCustomProps";
            this.BtnChangeCustomProps.Size = new System.Drawing.Size(120, 23);
            this.BtnChangeCustomProps.TabIndex = 4;
            this.BtnChangeCustomProps.Text = "Add Custom Props";
            this.BtnChangeCustomProps.UseVisualStyleBackColor = true;
            this.BtnChangeCustomProps.Click += new System.EventHandler(this.BtnChangeCustomProps_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rdoPowerPoint);
            this.groupBox1.Controls.Add(this.rdoExcel);
            this.groupBox1.Controls.Add(this.rdoWord);
            this.groupBox1.Location = new System.Drawing.Point(15, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(207, 48);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "File Type:";
            // 
            // rdoPowerPoint
            // 
            this.rdoPowerPoint.AutoSize = true;
            this.rdoPowerPoint.Location = new System.Drawing.Point(120, 19);
            this.rdoPowerPoint.Name = "rdoPowerPoint";
            this.rdoPowerPoint.Size = new System.Drawing.Size(79, 17);
            this.rdoPowerPoint.TabIndex = 2;
            this.rdoPowerPoint.Text = "PowerPoint";
            this.rdoPowerPoint.UseVisualStyleBackColor = true;
            this.rdoPowerPoint.CheckedChanged += new System.EventHandler(this.rdoPowerPoint_CheckedChanged);
            // 
            // rdoExcel
            // 
            this.rdoExcel.AutoSize = true;
            this.rdoExcel.Location = new System.Drawing.Point(63, 19);
            this.rdoExcel.Name = "rdoExcel";
            this.rdoExcel.Size = new System.Drawing.Size(51, 17);
            this.rdoExcel.TabIndex = 1;
            this.rdoExcel.Text = "Excel";
            this.rdoExcel.UseVisualStyleBackColor = true;
            this.rdoExcel.CheckedChanged += new System.EventHandler(this.rdoExcel_CheckedChanged);
            // 
            // rdoWord
            // 
            this.rdoWord.AutoSize = true;
            this.rdoWord.Checked = true;
            this.rdoWord.Location = new System.Drawing.Point(6, 19);
            this.rdoWord.Name = "rdoWord";
            this.rdoWord.Size = new System.Drawing.Size(51, 17);
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
            this.groupBox2.Location = new System.Drawing.Point(228, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(560, 48);
            this.groupBox2.TabIndex = 6;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "File Location:";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.BtnConvertStrict);
            this.groupBox3.Controls.Add(this.BtnPPTResetPII);
            this.groupBox3.Controls.Add(this.BtnFixCorruptRevisions);
            this.groupBox3.Controls.Add(this.BtnFixCorruptBookmarks);
            this.groupBox3.Controls.Add(this.BtnRemovePII);
            this.groupBox3.Controls.Add(this.BtnFixNotesPageSize);
            this.groupBox3.Controls.Add(this.BtnChangeTheme);
            this.groupBox3.Controls.Add(this.BtnChangeCustomProps);
            this.groupBox3.Location = new System.Drawing.Point(15, 453);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(773, 79);
            this.groupBox3.TabIndex = 7;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Batch Commands";
            // 
            // BtnConvertStrict
            // 
            this.BtnConvertStrict.Location = new System.Drawing.Point(606, 48);
            this.BtnConvertStrict.Name = "BtnConvertStrict";
            this.BtnConvertStrict.Size = new System.Drawing.Size(158, 23);
            this.BtnConvertStrict.TabIndex = 12;
            this.BtnConvertStrict.Text = "Convert Strict To Non-Strict";
            this.BtnConvertStrict.UseVisualStyleBackColor = true;
            this.BtnConvertStrict.Click += new System.EventHandler(this.BtnConvertStrict_Click);
            // 
            // BtnPPTResetPII
            // 
            this.BtnPPTResetPII.Location = new System.Drawing.Point(7, 48);
            this.BtnPPTResetPII.Name = "BtnPPTResetPII";
            this.BtnPPTResetPII.Size = new System.Drawing.Size(120, 23);
            this.BtnPPTResetPII.TabIndex = 11;
            this.BtnPPTResetPII.Text = "Reset PII On Save";
            this.BtnPPTResetPII.UseVisualStyleBackColor = true;
            this.BtnPPTResetPII.Click += new System.EventHandler(this.BtnPPTResetPII_Click);
            // 
            // BtnFixCorruptRevisions
            // 
            this.BtnFixCorruptRevisions.Location = new System.Drawing.Point(606, 19);
            this.BtnFixCorruptRevisions.Name = "BtnFixCorruptRevisions";
            this.BtnFixCorruptRevisions.Size = new System.Drawing.Size(158, 23);
            this.BtnFixCorruptRevisions.TabIndex = 9;
            this.BtnFixCorruptRevisions.Text = "Fix Corrupt Revisions";
            this.BtnFixCorruptRevisions.UseVisualStyleBackColor = true;
            this.BtnFixCorruptRevisions.Click += new System.EventHandler(this.BtnFixCorruptRevisions_Click);
            // 
            // BtnFixCorruptBookmarks
            // 
            this.BtnFixCorruptBookmarks.Location = new System.Drawing.Point(470, 19);
            this.BtnFixCorruptBookmarks.Margin = new System.Windows.Forms.Padding(2);
            this.BtnFixCorruptBookmarks.Name = "BtnFixCorruptBookmarks";
            this.BtnFixCorruptBookmarks.Size = new System.Drawing.Size(131, 23);
            this.BtnFixCorruptBookmarks.TabIndex = 9;
            this.BtnFixCorruptBookmarks.Text = "Fix Corrupt Bookmarks";
            this.BtnFixCorruptBookmarks.UseVisualStyleBackColor = true;
            this.BtnFixCorruptBookmarks.Click += new System.EventHandler(this.BtnFixCorruptBookmarks_Click);
            // 
            // BtnRemovePII
            // 
            this.BtnRemovePII.Location = new System.Drawing.Point(364, 19);
            this.BtnRemovePII.Name = "BtnRemovePII";
            this.BtnRemovePII.Size = new System.Drawing.Size(101, 23);
            this.BtnRemovePII.TabIndex = 10;
            this.BtnRemovePII.Text = "Remove PII";
            this.BtnRemovePII.UseVisualStyleBackColor = true;
            this.BtnRemovePII.Click += new System.EventHandler(this.BtnRemovePII_Click);
            // 
            // BtnFixNotesPageSize
            // 
            this.BtnFixNotesPageSize.Location = new System.Drawing.Point(236, 19);
            this.BtnFixNotesPageSize.Margin = new System.Windows.Forms.Padding(2);
            this.BtnFixNotesPageSize.Name = "BtnFixNotesPageSize";
            this.BtnFixNotesPageSize.Size = new System.Drawing.Size(123, 23);
            this.BtnFixNotesPageSize.TabIndex = 9;
            this.BtnFixNotesPageSize.Text = "Fix Notes Page Size";
            this.BtnFixNotesPageSize.UseVisualStyleBackColor = true;
            this.BtnFixNotesPageSize.Click += new System.EventHandler(this.BtnFixNotesPageSize_Click);
            // 
            // BtnChangeTheme
            // 
            this.BtnChangeTheme.Location = new System.Drawing.Point(133, 19);
            this.BtnChangeTheme.Name = "BtnChangeTheme";
            this.BtnChangeTheme.Size = new System.Drawing.Size(98, 23);
            this.BtnChangeTheme.TabIndex = 5;
            this.BtnChangeTheme.Text = "Change Theme";
            this.BtnChangeTheme.UseVisualStyleBackColor = true;
            this.BtnChangeTheme.Click += new System.EventHandler(this.BtnChangeTheme_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.lstOutput);
            this.groupBox4.Location = new System.Drawing.Point(15, 66);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(773, 380);
            this.groupBox4.TabIndex = 8;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Files";
            // 
            // FrmBatch
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 544);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
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
        private System.Windows.Forms.Button BtnFixNotesPageSize;
        private System.Windows.Forms.Button BtnRemovePII;
        private System.Windows.Forms.Button BtnFixCorruptBookmarks;
        private System.Windows.Forms.Button BtnFixCorruptRevisions;
        private System.Windows.Forms.Button BtnPPTResetPII;
        private System.Windows.Forms.Button BtnConvertStrict;
    }
}