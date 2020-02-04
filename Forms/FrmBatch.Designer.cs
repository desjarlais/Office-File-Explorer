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
            this.rdoWord = new System.Windows.Forms.RadioButton();
            this.rdoExcel = new System.Windows.Forms.RadioButton();
            this.rdoPowerPoint = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 56);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Directory Path:";
            // 
            // TxbDirectoryPath
            // 
            this.TxbDirectoryPath.Location = new System.Drawing.Point(95, 53);
            this.TxbDirectoryPath.Name = "TxbDirectoryPath";
            this.TxbDirectoryPath.Size = new System.Drawing.Size(566, 20);
            this.TxbDirectoryPath.TabIndex = 1;
            // 
            // BtnBrowseDirectory
            // 
            this.BtnBrowseDirectory.Location = new System.Drawing.Point(667, 51);
            this.BtnBrowseDirectory.Name = "BtnBrowseDirectory";
            this.BtnBrowseDirectory.Size = new System.Drawing.Size(121, 23);
            this.BtnBrowseDirectory.TabIndex = 2;
            this.BtnBrowseDirectory.Text = "...Choose Location";
            this.BtnBrowseDirectory.UseVisualStyleBackColor = true;
            this.BtnBrowseDirectory.Click += new System.EventHandler(this.BtnBrowseDirectory_Click);
            // 
            // lstOutput
            // 
            this.lstOutput.FormattingEnabled = true;
            this.lstOutput.Location = new System.Drawing.Point(15, 90);
            this.lstOutput.Name = "lstOutput";
            this.lstOutput.Size = new System.Drawing.Size(773, 316);
            this.lstOutput.TabIndex = 3;
            // 
            // BtnChangeCustomProps
            // 
            this.BtnChangeCustomProps.Location = new System.Drawing.Point(647, 415);
            this.BtnChangeCustomProps.Name = "BtnChangeCustomProps";
            this.BtnChangeCustomProps.Size = new System.Drawing.Size(141, 23);
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
            this.groupBox1.Location = new System.Drawing.Point(15, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(773, 33);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "File Type:";
            // 
            // rdoWord
            // 
            this.rdoWord.AutoSize = true;
            this.rdoWord.Checked = true;
            this.rdoWord.Location = new System.Drawing.Point(68, 10);
            this.rdoWord.Name = "rdoWord";
            this.rdoWord.Size = new System.Drawing.Size(51, 17);
            this.rdoWord.TabIndex = 0;
            this.rdoWord.TabStop = true;
            this.rdoWord.Text = "Word";
            this.rdoWord.UseVisualStyleBackColor = true;
            // 
            // rdoExcel
            // 
            this.rdoExcel.AutoSize = true;
            this.rdoExcel.Location = new System.Drawing.Point(137, 10);
            this.rdoExcel.Name = "rdoExcel";
            this.rdoExcel.Size = new System.Drawing.Size(51, 17);
            this.rdoExcel.TabIndex = 1;
            this.rdoExcel.Text = "Excel";
            this.rdoExcel.UseVisualStyleBackColor = true;
            // 
            // rdoPowerPoint
            // 
            this.rdoPowerPoint.AutoSize = true;
            this.rdoPowerPoint.Location = new System.Drawing.Point(206, 10);
            this.rdoPowerPoint.Name = "rdoPowerPoint";
            this.rdoPowerPoint.Size = new System.Drawing.Size(79, 17);
            this.rdoPowerPoint.TabIndex = 2;
            this.rdoPowerPoint.Text = "PowerPoint";
            this.rdoPowerPoint.UseVisualStyleBackColor = true;
            // 
            // FrmBatch
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.BtnChangeCustomProps);
            this.Controls.Add(this.lstOutput);
            this.Controls.Add(this.BtnBrowseDirectory);
            this.Controls.Add(this.TxbDirectoryPath);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmBatch";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Batch File Processing";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

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
    }
}