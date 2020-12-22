namespace Office_File_Explorer.Forms
{
    partial class FrmPrinterSettings
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmPrinterSettings));
            this.label1 = new System.Windows.Forms.Label();
            this.CboPrinters = new System.Windows.Forms.ComboBox();
            this.LstDisplay = new System.Windows.Forms.ListBox();
            this.BtnCopy = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Printers:";
            // 
            // CboPrinters
            // 
            this.CboPrinters.FormattingEnabled = true;
            this.CboPrinters.Location = new System.Drawing.Point(63, 6);
            this.CboPrinters.Name = "CboPrinters";
            this.CboPrinters.Size = new System.Drawing.Size(504, 21);
            this.CboPrinters.TabIndex = 1;
            this.CboPrinters.SelectedIndexChanged += new System.EventHandler(this.CboPrinters_SelectedIndexChanged_1);
            // 
            // LstDisplay
            // 
            this.LstDisplay.FormattingEnabled = true;
            this.LstDisplay.Location = new System.Drawing.Point(15, 33);
            this.LstDisplay.Name = "LstDisplay";
            this.LstDisplay.Size = new System.Drawing.Size(552, 420);
            this.LstDisplay.TabIndex = 2;
            // 
            // BtnCopy
            // 
            this.BtnCopy.Location = new System.Drawing.Point(12, 459);
            this.BtnCopy.Name = "BtnCopy";
            this.BtnCopy.Size = new System.Drawing.Size(117, 23);
            this.BtnCopy.TabIndex = 3;
            this.BtnCopy.Text = "Copy Print Settings";
            this.BtnCopy.UseVisualStyleBackColor = true;
            // 
            // FrmPrinterSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(579, 489);
            this.Controls.Add(this.BtnCopy);
            this.Controls.Add(this.LstDisplay);
            this.Controls.Add(this.CboPrinters);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmPrinterSettings";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Printer Settings";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox CboPrinters;
        private System.Windows.Forms.ListBox LstDisplay;
        private System.Windows.Forms.Button BtnCopy;
    }
}