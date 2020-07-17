namespace Office_File_Explorer.Forms
{
    partial class FrmFontDetails
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmFontDetails));
            this.LstFontInfo = new System.Windows.Forms.ListBox();
            this.pBoxAlias = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pBoxAlias)).BeginInit();
            this.SuspendLayout();
            // 
            // LstFontInfo
            // 
            this.LstFontInfo.FormattingEnabled = true;
            this.LstFontInfo.Location = new System.Drawing.Point(12, 12);
            this.LstFontInfo.Name = "LstFontInfo";
            this.LstFontInfo.Size = new System.Drawing.Size(290, 199);
            this.LstFontInfo.TabIndex = 0;
            // 
            // pBoxAlias
            // 
            this.pBoxAlias.Location = new System.Drawing.Point(12, 217);
            this.pBoxAlias.Name = "pBoxAlias";
            this.pBoxAlias.Size = new System.Drawing.Size(290, 141);
            this.pBoxAlias.TabIndex = 1;
            this.pBoxAlias.TabStop = false;
            this.pBoxAlias.Paint += new System.Windows.Forms.PaintEventHandler(this.PBoxAlias_Paint);
            // 
            // FrmFontDetails
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(316, 370);
            this.Controls.Add(this.pBoxAlias);
            this.Controls.Add(this.LstFontInfo);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmFontDetails";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Font Details";
            ((System.ComponentModel.ISupportInitialize)(this.pBoxAlias)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox LstFontInfo;
        private System.Windows.Forms.PictureBox pBoxAlias;
    }
}