namespace Office_File_Explorer.Forms
{
    partial class FrmFontViewer
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmFontViewer));
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.RdoDrawText = new System.Windows.Forms.RadioButton();
            this.RdoRenderText = new System.Windows.Forms.RadioButton();
            this.RdoDrawString = new System.Windows.Forms.RadioButton();
            this.label5 = new System.Windows.Forms.Label();
            this.BtnFontDetails = new System.Windows.Forms.Button();
            this.BtnColorDlg = new System.Windows.Forms.Button();
            this.NudFontSize = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.LblColor = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.CboFonts = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txbInput = new System.Windows.Forms.TextBox();
            this.pBoxFont = new System.Windows.Forms.PictureBox();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.label3 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.NudFontSize)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pBoxFont)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.BtnFontDetails);
            this.panel1.Controls.Add(this.BtnColorDlg);
            this.panel1.Controls.Add(this.NudFontSize);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.LblColor);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.CboFonts);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(12, 27);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(271, 157);
            this.panel1.TabIndex = 0;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.RdoDrawText);
            this.panel2.Controls.Add(this.RdoRenderText);
            this.panel2.Controls.Add(this.RdoDrawString);
            this.panel2.Location = new System.Drawing.Point(136, 59);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(120, 82);
            this.panel2.TabIndex = 9;
            // 
            // RdoDrawText
            // 
            this.RdoDrawText.AutoSize = true;
            this.RdoDrawText.Location = new System.Drawing.Point(8, 59);
            this.RdoDrawText.Name = "RdoDrawText";
            this.RdoDrawText.Size = new System.Drawing.Size(71, 17);
            this.RdoDrawText.TabIndex = 2;
            this.RdoDrawText.Text = "DrawText";
            this.RdoDrawText.UseVisualStyleBackColor = true;
            this.RdoDrawText.CheckedChanged += new System.EventHandler(this.RdoDrawText_CheckedChanged);
            // 
            // RdoRenderText
            // 
            this.RdoRenderText.AutoSize = true;
            this.RdoRenderText.Location = new System.Drawing.Point(8, 33);
            this.RdoRenderText.Name = "RdoRenderText";
            this.RdoRenderText.Size = new System.Drawing.Size(81, 17);
            this.RdoRenderText.TabIndex = 1;
            this.RdoRenderText.Text = "RenderText";
            this.RdoRenderText.UseVisualStyleBackColor = true;
            this.RdoRenderText.CheckedChanged += new System.EventHandler(this.RdoRenderText_CheckedChanged);
            // 
            // RdoDrawString
            // 
            this.RdoDrawString.AutoSize = true;
            this.RdoDrawString.Checked = true;
            this.RdoDrawString.Location = new System.Drawing.Point(8, 7);
            this.RdoDrawString.Name = "RdoDrawString";
            this.RdoDrawString.Size = new System.Drawing.Size(77, 17);
            this.RdoDrawString.TabIndex = 0;
            this.RdoDrawString.TabStop = true;
            this.RdoDrawString.Text = "DrawString";
            this.RdoDrawString.UseVisualStyleBackColor = true;
            this.RdoDrawString.CheckedChanged += new System.EventHandler(this.RdoDrawString_CheckedChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(133, 43);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(79, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Rendering API:";
            // 
            // BtnFontDetails
            // 
            this.BtnFontDetails.Location = new System.Drawing.Point(15, 126);
            this.BtnFontDetails.Name = "BtnFontDetails";
            this.BtnFontDetails.Size = new System.Drawing.Size(105, 23);
            this.BtnFontDetails.TabIndex = 7;
            this.BtnFontDetails.Text = "Font Details";
            this.BtnFontDetails.UseVisualStyleBackColor = true;
            this.BtnFontDetails.Click += new System.EventHandler(this.BtnFontDetails_Click_1);
            // 
            // BtnColorDlg
            // 
            this.BtnColorDlg.Location = new System.Drawing.Point(15, 97);
            this.BtnColorDlg.Name = "BtnColorDlg";
            this.BtnColorDlg.Size = new System.Drawing.Size(105, 23);
            this.BtnColorDlg.TabIndex = 6;
            this.BtnColorDlg.Text = "Change Font Color";
            this.BtnColorDlg.UseVisualStyleBackColor = true;
            this.BtnColorDlg.Click += new System.EventHandler(this.BtnColorDlg_Click_1);
            // 
            // NudFontSize
            // 
            this.NudFontSize.Location = new System.Drawing.Point(72, 65);
            this.NudFontSize.Maximum = new decimal(new int[] {
            72,
            0,
            0,
            0});
            this.NudFontSize.Minimum = new decimal(new int[] {
            8,
            0,
            0,
            0});
            this.NudFontSize.Name = "NudFontSize";
            this.NudFontSize.Size = new System.Drawing.Size(48, 20);
            this.NudFontSize.TabIndex = 5;
            this.NudFontSize.Value = new decimal(new int[] {
            24,
            0,
            0,
            0});
            this.NudFontSize.ValueChanged += new System.EventHandler(this.NudFontSize_ValueChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 67);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(54, 13);
            this.label4.TabIndex = 4;
            this.label4.Text = "Font Size:";
            // 
            // LblColor
            // 
            this.LblColor.AutoSize = true;
            this.LblColor.Location = new System.Drawing.Point(76, 42);
            this.LblColor.Name = "LblColor";
            this.LblColor.Size = new System.Drawing.Size(34, 13);
            this.LblColor.TabIndex = 3;
            this.LblColor.Text = "Black";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 42);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(58, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Font Color:";
            // 
            // CboFonts
            // 
            this.CboFonts.FormattingEnabled = true;
            this.CboFonts.Location = new System.Drawing.Point(49, 6);
            this.CboFonts.Name = "CboFonts";
            this.CboFonts.Size = new System.Drawing.Size(207, 21);
            this.CboFonts.TabIndex = 1;
            this.CboFonts.SelectedIndexChanged += new System.EventHandler(this.CboFonts_SelectedIndexChanged_1);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(31, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Font:";
            // 
            // txbInput
            // 
            this.txbInput.Location = new System.Drawing.Point(289, 27);
            this.txbInput.Multiline = true;
            this.txbInput.Name = "txbInput";
            this.txbInput.Size = new System.Drawing.Size(746, 157);
            this.txbInput.TabIndex = 1;
            this.txbInput.TextChanged += new System.EventHandler(this.TxbInput_TextChanged);
            // 
            // pBoxFont
            // 
            this.pBoxFont.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pBoxFont.Location = new System.Drawing.Point(12, 208);
            this.pBoxFont.Name = "pBoxFont";
            this.pBoxFont.Size = new System.Drawing.Size(1023, 281);
            this.pBoxFont.TabIndex = 2;
            this.pBoxFont.TabStop = false;
            this.pBoxFont.Paint += new System.Windows.Forms.PaintEventHandler(this.PBoxFont_Paint);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 188);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(44, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Display:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(12, 8);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(86, 13);
            this.label6.TabIndex = 4;
            this.label6.Text = "Font Information:";
            // 
            // FrmFontViewer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1050, 501);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.pBoxFont);
            this.Controls.Add(this.txbInput);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmFontViewer";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Font Viewer";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.NudFontSize)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pBoxFont)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.RadioButton RdoDrawText;
        private System.Windows.Forms.RadioButton RdoRenderText;
        private System.Windows.Forms.RadioButton RdoDrawString;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button BtnFontDetails;
        private System.Windows.Forms.Button BtnColorDlg;
        private System.Windows.Forms.NumericUpDown NudFontSize;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label LblColor;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox CboFonts;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txbInput;
        private System.Windows.Forms.PictureBox pBoxFont;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label6;
    }
}