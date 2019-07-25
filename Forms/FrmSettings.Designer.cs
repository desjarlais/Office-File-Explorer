namespace Office_File_Explorer.Forms
{
    partial class FrmSettings
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmSettings));
            this.ckRemoveFallback = new System.Windows.Forms.CheckBox();
            this.ckOpenInWord = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.BtnOK = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // ckRemoveFallback
            // 
            this.ckRemoveFallback.AutoSize = true;
            this.ckRemoveFallback.Location = new System.Drawing.Point(12, 37);
            this.ckRemoveFallback.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.ckRemoveFallback.Name = "ckRemoveFallback";
            this.ckRemoveFallback.Size = new System.Drawing.Size(294, 29);
            this.ckRemoveFallback.TabIndex = 0;
            this.ckRemoveFallback.Text = "Remove All Fallback Tags";
            this.ckRemoveFallback.UseVisualStyleBackColor = true;
            // 
            // ckOpenInWord
            // 
            this.ckOpenInWord.AutoSize = true;
            this.ckOpenInWord.Location = new System.Drawing.Point(12, 81);
            this.ckOpenInWord.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.ckOpenInWord.Name = "ckOpenInWord";
            this.ckOpenInWord.Size = new System.Drawing.Size(320, 29);
            this.ckOpenInWord.TabIndex = 1;
            this.ckOpenInWord.Text = "Open file in Word after repair";
            this.ckOpenInWord.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ckRemoveFallback);
            this.groupBox1.Controls.Add(this.ckOpenInWord);
            this.groupBox1.Location = new System.Drawing.Point(24, 23);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.groupBox1.Size = new System.Drawing.Size(454, 152);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Corrupt Document Settings";
            // 
            // BtnCancel
            // 
            this.BtnCancel.Location = new System.Drawing.Point(321, 187);
            this.BtnCancel.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(150, 44);
            this.BtnCancel.TabIndex = 2;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // BtnOK
            // 
            this.BtnOK.Location = new System.Drawing.Point(179, 187);
            this.BtnOK.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.BtnOK.Name = "BtnOK";
            this.BtnOK.Size = new System.Drawing.Size(130, 44);
            this.BtnOK.TabIndex = 3;
            this.BtnOK.Text = "Ok";
            this.BtnOK.UseVisualStyleBackColor = true;
            this.BtnOK.Click += new System.EventHandler(this.BtnOK_Click);
            // 
            // FrmSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(508, 254);
            this.Controls.Add(this.BtnOK);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.BtnCancel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmSettings";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Settings";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.CheckBox ckRemoveFallback;
        private System.Windows.Forms.CheckBox ckOpenInWord;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.Button BtnOK;
    }
}