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
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.rdoDom = new System.Windows.Forms.RadioButton();
            this.rdoSax = new System.Windows.Forms.RadioButton();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.ckDeleteCopies = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
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
            this.groupBox1.Size = new System.Drawing.Size(175, 101);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Word Corrupt Document";
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
            this.BtnCancel.Location = new System.Drawing.Point(347, 225);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(63, 23);
            this.BtnCancel.TabIndex = 2;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // BtnOK
            // 
            this.BtnOK.Location = new System.Drawing.Point(282, 225);
            this.BtnOK.Name = "BtnOK";
            this.BtnOK.Size = new System.Drawing.Size(60, 23);
            this.BtnOK.TabIndex = 3;
            this.BtnOK.Text = "OK";
            this.BtnOK.UseVisualStyleBackColor = true;
            this.BtnOK.Click += new System.EventHandler(this.BtnOK_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.ckResetNotesMaster);
            this.groupBox2.Location = new System.Drawing.Point(192, 12);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox2.Size = new System.Drawing.Size(218, 101);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "PowerPoint Options";
            // 
            // ckResetNotesMaster
            // 
            this.ckResetNotesMaster.AutoSize = true;
            this.ckResetNotesMaster.Location = new System.Drawing.Point(10, 19);
            this.ckResetNotesMaster.Margin = new System.Windows.Forms.Padding(2);
            this.ckResetNotesMaster.Name = "ckResetNotesMaster";
            this.ckResetNotesMaster.Size = new System.Drawing.Size(203, 17);
            this.ckResetNotesMaster.TabIndex = 0;
            this.ckResetNotesMaster.Text = "Reset Notes Slides and Notes Master";
            this.ckResetNotesMaster.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.rdoDom);
            this.groupBox3.Controls.Add(this.rdoSax);
            this.groupBox3.Location = new System.Drawing.Point(250, 119);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(160, 100);
            this.groupBox3.TabIndex = 5;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Excel List Cell Value Options";
            // 
            // rdoDom
            // 
            this.rdoDom.AutoSize = true;
            this.rdoDom.Location = new System.Drawing.Point(6, 39);
            this.rdoDom.Name = "rdoDom";
            this.rdoDom.Size = new System.Drawing.Size(76, 17);
            this.rdoDom.TabIndex = 1;
            this.rdoDom.Text = "DOM Style";
            this.rdoDom.UseVisualStyleBackColor = true;
            // 
            // rdoSax
            // 
            this.rdoSax.AutoSize = true;
            this.rdoSax.Checked = true;
            this.rdoSax.Location = new System.Drawing.Point(6, 16);
            this.rdoSax.Name = "rdoSax";
            this.rdoSax.Size = new System.Drawing.Size(72, 17);
            this.rdoSax.TabIndex = 0;
            this.rdoSax.TabStop = true;
            this.rdoSax.Text = "SAX Style";
            this.rdoSax.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.ckDeleteCopies);
            this.groupBox4.Location = new System.Drawing.Point(12, 119);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(232, 100);
            this.groupBox4.TabIndex = 6;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "App Settings";
            // 
            // ckDeleteCopies
            // 
            this.ckDeleteCopies.AutoSize = true;
            this.ckDeleteCopies.Location = new System.Drawing.Point(6, 19);
            this.ckDeleteCopies.Name = "ckDeleteCopies";
            this.ckDeleteCopies.Size = new System.Drawing.Size(154, 17);
            this.ckDeleteCopies.TabIndex = 0;
            this.ckDeleteCopies.Text = "Delete Copied Files On Exit";
            this.ckDeleteCopies.UseVisualStyleBackColor = true;
            // 
            // FrmSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(418, 256);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox2);
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
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
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
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.RadioButton rdoDom;
        private System.Windows.Forms.RadioButton rdoSax;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.CheckBox ckDeleteCopies;
    }
}