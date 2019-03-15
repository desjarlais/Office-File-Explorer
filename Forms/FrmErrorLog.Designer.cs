namespace Office_File_Explorer.Forms
{
    partial class FrmErrorLog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmErrorLog));
            this.LstErrorLog = new System.Windows.Forms.ListBox();
            this.BtnClearLog = new System.Windows.Forms.Button();
            this.BtnCopyResults = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // LstErrorLog
            // 
            this.LstErrorLog.FormattingEnabled = true;
            this.LstErrorLog.HorizontalScrollbar = true;
            this.LstErrorLog.ItemHeight = 20;
            this.LstErrorLog.Location = new System.Drawing.Point(12, 12);
            this.LstErrorLog.Name = "LstErrorLog";
            this.LstErrorLog.Size = new System.Drawing.Size(1124, 704);
            this.LstErrorLog.TabIndex = 0;
            // 
            // BtnClearLog
            // 
            this.BtnClearLog.Location = new System.Drawing.Point(1022, 728);
            this.BtnClearLog.Name = "BtnClearLog";
            this.BtnClearLog.Size = new System.Drawing.Size(114, 34);
            this.BtnClearLog.TabIndex = 1;
            this.BtnClearLog.Text = "Clear Log";
            this.BtnClearLog.UseVisualStyleBackColor = true;
            this.BtnClearLog.Click += new System.EventHandler(this.BtnClearLog_Click);
            // 
            // BtnCopyResults
            // 
            this.BtnCopyResults.Location = new System.Drawing.Point(884, 728);
            this.BtnCopyResults.Name = "BtnCopyResults";
            this.BtnCopyResults.Size = new System.Drawing.Size(132, 34);
            this.BtnCopyResults.TabIndex = 2;
            this.BtnCopyResults.Text = "Copy Results";
            this.BtnCopyResults.UseVisualStyleBackColor = true;
            this.BtnCopyResults.Click += new System.EventHandler(this.BtnCopyResults_Click);
            // 
            // FrmErrorLog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1148, 774);
            this.Controls.Add(this.BtnCopyResults);
            this.Controls.Add(this.BtnClearLog);
            this.Controls.Add(this.LstErrorLog);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmErrorLog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "FrmErrorLog";
            this.Load += new System.EventHandler(this.FrmErrorLog_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox LstErrorLog;
        private System.Windows.Forms.Button BtnClearLog;
        private System.Windows.Forms.Button BtnCopyResults;
    }
}