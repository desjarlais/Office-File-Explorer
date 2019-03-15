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
            this.LstErrorLog.Location = new System.Drawing.Point(8, 8);
            this.LstErrorLog.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.LstErrorLog.Name = "LstErrorLog";
            this.LstErrorLog.Size = new System.Drawing.Size(751, 459);
            this.LstErrorLog.TabIndex = 0;
            // 
            // BtnClearLog
            // 
            this.BtnClearLog.Location = new System.Drawing.Point(681, 473);
            this.BtnClearLog.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.BtnClearLog.Name = "BtnClearLog";
            this.BtnClearLog.Size = new System.Drawing.Size(76, 22);
            this.BtnClearLog.TabIndex = 1;
            this.BtnClearLog.Text = "Clear Log";
            this.BtnClearLog.UseVisualStyleBackColor = true;
            this.BtnClearLog.Click += new System.EventHandler(this.BtnClearLog_Click);
            // 
            // BtnCopyResults
            // 
            this.BtnCopyResults.Location = new System.Drawing.Point(589, 473);
            this.BtnCopyResults.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.BtnCopyResults.Name = "BtnCopyResults";
            this.BtnCopyResults.Size = new System.Drawing.Size(88, 22);
            this.BtnCopyResults.TabIndex = 2;
            this.BtnCopyResults.Text = "Copy Results";
            this.BtnCopyResults.UseVisualStyleBackColor = true;
            this.BtnCopyResults.Click += new System.EventHandler(this.BtnCopyResults_Click);
            // 
            // FrmErrorLog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(765, 503);
            this.Controls.Add(this.BtnCopyResults);
            this.Controls.Add(this.BtnClearLog);
            this.Controls.Add(this.LstErrorLog);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmErrorLog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Error Log";
            this.Load += new System.EventHandler(this.FrmErrorLog_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox LstErrorLog;
        private System.Windows.Forms.Button BtnClearLog;
        private System.Windows.Forms.Button BtnCopyResults;
    }
}