namespace Office_File_Explorer.Forms
{
    partial class FrmParagraphs
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmParagraphs));
            this.label1 = new System.Windows.Forms.Label();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cbParagraphs = new System.Windows.Forms.ComboBox();
            this.lblParaCount = new System.Windows.Forms.Label();
            this.PBoxFont = new System.Windows.Forms.PictureBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.LblStyleName = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.LblFontSize = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.LblFontColor = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.PBoxFont)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Paragraphs:";
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(12, 130);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(748, 149);
            this.richTextBox1.TabIndex = 2;
            this.richTextBox1.Text = "";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 114);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Paragraph Text:";
            // 
            // cbParagraphs
            // 
            this.cbParagraphs.FormattingEnabled = true;
            this.cbParagraphs.Location = new System.Drawing.Point(79, 6);
            this.cbParagraphs.Name = "cbParagraphs";
            this.cbParagraphs.Size = new System.Drawing.Size(207, 21);
            this.cbParagraphs.TabIndex = 4;
            this.cbParagraphs.SelectedIndexChanged += new System.EventHandler(this.CbParagraphs_SelectedIndexChanged);
            // 
            // lblParaCount
            // 
            this.lblParaCount.AutoSize = true;
            this.lblParaCount.Location = new System.Drawing.Point(297, 10);
            this.lblParaCount.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblParaCount.Name = "lblParaCount";
            this.lblParaCount.Size = new System.Drawing.Size(99, 13);
            this.lblParaCount.TabIndex = 5;
            this.lblParaCount.Text = "Paragraph Count = ";
            // 
            // PBoxFont
            // 
            this.PBoxFont.Location = new System.Drawing.Point(12, 298);
            this.PBoxFont.Name = "PBoxFont";
            this.PBoxFont.Size = new System.Drawing.Size(748, 235);
            this.PBoxFont.TabIndex = 6;
            this.PBoxFont.TabStop = false;
            this.PBoxFont.Paint += new System.Windows.Forms.PaintEventHandler(this.PBoxFont_Paint_1);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 282);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(83, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Text Rendering:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 37);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(64, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Style Name:";
            // 
            // LblStyleName
            // 
            this.LblStyleName.AutoSize = true;
            this.LblStyleName.Location = new System.Drawing.Point(82, 37);
            this.LblStyleName.Name = "LblStyleName";
            this.LblStyleName.Size = new System.Drawing.Size(73, 13);
            this.LblStyleName.TabIndex = 9;
            this.LblStyleName.Text = "<Style Name>";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 60);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(54, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Font Size:";
            // 
            // LblFontSize
            // 
            this.LblFontSize.AutoSize = true;
            this.LblFontSize.Location = new System.Drawing.Point(82, 60);
            this.LblFontSize.Name = "LblFontSize";
            this.LblFontSize.Size = new System.Drawing.Size(63, 13);
            this.LblFontSize.TabIndex = 11;
            this.LblFontSize.Text = "<Font Size>";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(12, 84);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(61, 13);
            this.label6.TabIndex = 12;
            this.label6.Text = "Font Color: ";
            // 
            // LblFontColor
            // 
            this.LblFontColor.AutoSize = true;
            this.LblFontColor.Location = new System.Drawing.Point(82, 84);
            this.LblFontColor.Name = "LblFontColor";
            this.LblFontColor.Size = new System.Drawing.Size(67, 13);
            this.LblFontColor.TabIndex = 13;
            this.LblFontColor.Text = "<Font Color>";
            // 
            // FrmParagraphs
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(765, 545);
            this.Controls.Add(this.LblFontColor);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.LblFontSize);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.LblStyleName);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.PBoxFont);
            this.Controls.Add(this.lblParaCount);
            this.Controls.Add(this.cbParagraphs);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmParagraphs";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Paragraphs";
            ((System.ComponentModel.ISupportInitialize)(this.PBoxFont)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cbParagraphs;
        private System.Windows.Forms.Label lblParaCount;
        private System.Windows.Forms.PictureBox PBoxFont;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label LblStyleName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label LblFontSize;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label LblFontColor;
    }
}