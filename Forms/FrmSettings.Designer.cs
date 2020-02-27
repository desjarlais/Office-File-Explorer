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
            this.ckGroupShapeFix = new System.Windows.Forms.CheckBox();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.BtnOK = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // ckRemoveFallback
            // 
            this.ckRemoveFallback.AutoSize = true;
            this.ckRemoveFallback.Location = new System.Drawing.Point(6, 19);
            this.ckRemoveFallback.Name = "ckRemoveFallback";
            this.ckRemoveFallback.Size = new System.Drawing.Size(150, 17);
            this.ckRemoveFallback.TabIndex = 0;
            this.ckRemoveFallback.Text = "Remove All Fallback Tags";
            this.ckRemoveFallback.UseVisualStyleBackColor = true;
            // 
            // ckOpenInWord
            // 
            this.ckOpenInWord.AutoSize = true;
            this.ckOpenInWord.Location = new System.Drawing.Point(6, 42);
            this.ckOpenInWord.Name = "ckOpenInWord";
            this.ckOpenInWord.Size = new System.Drawing.Size(161, 17);
            this.ckOpenInWord.TabIndex = 1;
            this.ckOpenInWord.Text = "Open file in Word after repair";
            this.ckOpenInWord.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ckGroupShapeFix);
            this.groupBox1.Controls.Add(this.ckRemoveFallback);
            this.groupBox1.Controls.Add(this.ckOpenInWord);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(227, 101);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Corrupt Document Settings";
            // 
            // ckGroupShapeFix
            // 
            this.ckGroupShapeFix.AutoSize = true;
            this.ckGroupShapeFix.Location = new System.Drawing.Point(6, 65);
            this.ckGroupShapeFix.Name = "ckGroupShapeFix";
            this.ckGroupShapeFix.Size = new System.Drawing.Size(118, 17);
            this.ckGroupShapeFix.TabIndex = 2;
            this.ckGroupShapeFix.Text = "Fix grouped shapes";
            this.ckGroupShapeFix.UseVisualStyleBackColor = true;
            // 
            // BtnCancel
            // 
            this.BtnCancel.Location = new System.Drawing.Point(173, 119);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(63, 23);
            this.BtnCancel.TabIndex = 2;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // BtnOK
            // 
            this.BtnOK.Location = new System.Drawing.Point(108, 119);
            this.BtnOK.Name = "BtnOK";
            this.BtnOK.Size = new System.Drawing.Size(60, 23);
            this.BtnOK.TabIndex = 3;
            this.BtnOK.Text = "Ok";
            this.BtnOK.UseVisualStyleBackColor = true;
            this.BtnOK.Click += new System.EventHandler(this.BtnOK_Click);
            // 
            // FrmSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(248, 154);
            this.Controls.Add(this.BtnOK);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.BtnCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmSettings";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
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
        private System.Windows.Forms.CheckBox ckGroupShapeFix;
    }
}