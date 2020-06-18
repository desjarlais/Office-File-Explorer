namespace Office_File_Explorer.Forms
{
    partial class FrmFixDocument
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmFixDocument));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.RdoLT = new System.Windows.Forms.RadioButton();
            this.RdoEndnotes = new System.Windows.Forms.RadioButton();
            this.RdoRev = new System.Windows.Forms.RadioButton();
            this.RdoBK = new System.Windows.Forms.RadioButton();
            this.BtnOk = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.RdoNotes = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.RdoLT);
            this.groupBox1.Controls.Add(this.RdoEndnotes);
            this.groupBox1.Controls.Add(this.RdoRev);
            this.groupBox1.Controls.Add(this.RdoBK);
            this.groupBox1.Location = new System.Drawing.Point(12, 11);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(195, 150);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Word Document Fixes";
            // 
            // RdoLT
            // 
            this.RdoLT.AutoSize = true;
            this.RdoLT.Enabled = false;
            this.RdoLT.Location = new System.Drawing.Point(6, 87);
            this.RdoLT.Name = "RdoLT";
            this.RdoLT.Size = new System.Drawing.Size(106, 17);
            this.RdoLT.TabIndex = 3;
            this.RdoLT.Text = "Fix ListTemplates";
            this.RdoLT.UseVisualStyleBackColor = true;
            // 
            // RdoEndnotes
            // 
            this.RdoEndnotes.AutoSize = true;
            this.RdoEndnotes.Enabled = false;
            this.RdoEndnotes.Location = new System.Drawing.Point(6, 64);
            this.RdoEndnotes.Name = "RdoEndnotes";
            this.RdoEndnotes.Size = new System.Drawing.Size(86, 17);
            this.RdoEndnotes.TabIndex = 2;
            this.RdoEndnotes.Text = "Fix Endnotes";
            this.RdoEndnotes.UseVisualStyleBackColor = true;
            // 
            // RdoRev
            // 
            this.RdoRev.AutoSize = true;
            this.RdoRev.Enabled = false;
            this.RdoRev.Location = new System.Drawing.Point(6, 41);
            this.RdoRev.Name = "RdoRev";
            this.RdoRev.Size = new System.Drawing.Size(124, 17);
            this.RdoRev.TabIndex = 1;
            this.RdoRev.Text = "Fix Corrupt Revisions";
            this.RdoRev.UseVisualStyleBackColor = true;
            // 
            // RdoBK
            // 
            this.RdoBK.AutoSize = true;
            this.RdoBK.Enabled = false;
            this.RdoBK.Location = new System.Drawing.Point(6, 19);
            this.RdoBK.Name = "RdoBK";
            this.RdoBK.Size = new System.Drawing.Size(131, 17);
            this.RdoBK.TabIndex = 0;
            this.RdoBK.Text = "Fix Corrupt Bookmarks";
            this.RdoBK.UseVisualStyleBackColor = true;
            // 
            // BtnOk
            // 
            this.BtnOk.Location = new System.Drawing.Point(232, 167);
            this.BtnOk.Name = "BtnOk";
            this.BtnOk.Size = new System.Drawing.Size(75, 23);
            this.BtnOk.TabIndex = 0;
            this.BtnOk.Text = "OK";
            this.BtnOk.UseVisualStyleBackColor = true;
            this.BtnOk.Click += new System.EventHandler(this.BtnOk_Click);
            // 
            // BtnCancel
            // 
            this.BtnCancel.Location = new System.Drawing.Point(313, 167);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(75, 23);
            this.BtnCancel.TabIndex = 1;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.RdoNotes);
            this.groupBox2.Location = new System.Drawing.Point(213, 15);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(175, 146);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "PowerPoint Document Fixes";
            // 
            // RdoNotes
            // 
            this.RdoNotes.AutoSize = true;
            this.RdoNotes.Enabled = false;
            this.RdoNotes.Location = new System.Drawing.Point(19, 19);
            this.RdoNotes.Name = "RdoNotes";
            this.RdoNotes.Size = new System.Drawing.Size(120, 17);
            this.RdoNotes.TabIndex = 0;
            this.RdoNotes.Text = "Fix Notes Page Size";
            this.RdoNotes.UseVisualStyleBackColor = true;
            // 
            // FrmFixDocument
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(400, 202);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.BtnOk);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmFixDocument";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Fix Document";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton RdoLT;
        private System.Windows.Forms.RadioButton RdoEndnotes;
        private System.Windows.Forms.RadioButton RdoRev;
        private System.Windows.Forms.RadioButton RdoBK;
        private System.Windows.Forms.Button BtnOk;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton RdoNotes;
    }
}