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
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.ckResetNotesMaster = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // ckRemoveFallback
            // 
            this.ckRemoveFallback.AutoSize = true;
            this.ckRemoveFallback.Location = new System.Drawing.Point(9, 29);
            this.ckRemoveFallback.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.ckRemoveFallback.Name = "ckRemoveFallback";
            this.ckRemoveFallback.Size = new System.Drawing.Size(217, 24);
            this.ckRemoveFallback.TabIndex = 0;
            this.ckRemoveFallback.Text = "Remove All Fallback Tags";
            this.ckRemoveFallback.UseVisualStyleBackColor = true;
            // 
            // ckOpenInWord
            // 
            this.ckOpenInWord.AutoSize = true;
            this.ckOpenInWord.Location = new System.Drawing.Point(9, 65);
            this.ckOpenInWord.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.ckOpenInWord.Name = "ckOpenInWord";
            this.ckOpenInWord.Size = new System.Drawing.Size(237, 24);
            this.ckOpenInWord.TabIndex = 1;
            this.ckOpenInWord.Text = "Open file in Word after repair";
            this.ckOpenInWord.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ckGroupShapeFix);
            this.groupBox1.Controls.Add(this.ckRemoveFallback);
            this.groupBox1.Controls.Add(this.ckOpenInWord);
            this.groupBox1.Location = new System.Drawing.Point(18, 18);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Size = new System.Drawing.Size(340, 155);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Word Corrupt Document";
            // 
            // ckGroupShapeFix
            // 
            this.ckGroupShapeFix.AutoSize = true;
            this.ckGroupShapeFix.Location = new System.Drawing.Point(9, 100);
            this.ckGroupShapeFix.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.ckGroupShapeFix.Name = "ckGroupShapeFix";
            this.ckGroupShapeFix.Size = new System.Drawing.Size(174, 24);
            this.ckGroupShapeFix.TabIndex = 2;
            this.ckGroupShapeFix.Text = "Fix grouped shapes";
            this.ckGroupShapeFix.UseVisualStyleBackColor = true;
            // 
            // BtnCancel
            // 
            this.BtnCancel.Location = new System.Drawing.Point(595, 181);
            this.BtnCancel.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(94, 35);
            this.BtnCancel.TabIndex = 2;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // BtnOK
            // 
            this.BtnOK.Location = new System.Drawing.Point(497, 181);
            this.BtnOK.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.BtnOK.Name = "BtnOK";
            this.BtnOK.Size = new System.Drawing.Size(90, 35);
            this.BtnOK.TabIndex = 3;
            this.BtnOK.Text = "OK";
            this.BtnOK.UseVisualStyleBackColor = true;
            this.BtnOK.Click += new System.EventHandler(this.BtnOK_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.ckResetNotesMaster);
            this.groupBox2.Location = new System.Drawing.Point(365, 18);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(324, 155);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "PowerPoint Notes Page Fix";
            // 
            // ckResetNotesMaster
            // 
            this.ckResetNotesMaster.AutoSize = true;
            this.ckResetNotesMaster.Location = new System.Drawing.Point(15, 29);
            this.ckResetNotesMaster.Name = "ckResetNotesMaster";
            this.ckResetNotesMaster.Size = new System.Drawing.Size(301, 24);
            this.ckResetNotesMaster.TabIndex = 0;
            this.ckResetNotesMaster.Text = "Reset Notes Slides and Notes Master";
            this.ckResetNotesMaster.UseVisualStyleBackColor = true;
            // 
            // FrmSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(719, 233);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.BtnOK);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.BtnCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmSettings";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Settings";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.CheckBox ckRemoveFallback;
        private System.Windows.Forms.CheckBox ckOpenInWord;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.Button BtnOK;
        private System.Windows.Forms.CheckBox ckGroupShapeFix;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox ckResetNotesMaster;
    }
}