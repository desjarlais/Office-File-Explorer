namespace Office_File_Explorer.Forms
{
    partial class FrmDeleteXmlNode
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmDeleteXmlNode));
            this.label1 = new System.Windows.Forms.Label();
            this.cboNodes = new System.Windows.Forms.ComboBox();
            this.BtnDeleteNode = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(36, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Node:";
            // 
            // cboNodes
            // 
            this.cboNodes.FormattingEnabled = true;
            this.cboNodes.Location = new System.Drawing.Point(54, 12);
            this.cboNodes.Name = "cboNodes";
            this.cboNodes.Size = new System.Drawing.Size(358, 21);
            this.cboNodes.TabIndex = 1;
            // 
            // BtnDeleteNode
            // 
            this.BtnDeleteNode.Location = new System.Drawing.Point(232, 39);
            this.BtnDeleteNode.Name = "BtnDeleteNode";
            this.BtnDeleteNode.Size = new System.Drawing.Size(99, 23);
            this.BtnDeleteNode.TabIndex = 2;
            this.BtnDeleteNode.Text = "Delete Node";
            this.BtnDeleteNode.UseVisualStyleBackColor = true;
            this.BtnDeleteNode.Click += new System.EventHandler(this.BtnDeleteNode_Click);
            // 
            // BtnCancel
            // 
            this.BtnCancel.Location = new System.Drawing.Point(337, 39);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(75, 23);
            this.BtnCancel.TabIndex = 3;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // FrmDeleteXmlNode
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(428, 71);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.BtnDeleteNode);
            this.Controls.Add(this.cboNodes);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmDeleteXmlNode";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Delete Xml Node";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cboNodes;
        private System.Windows.Forms.Button BtnDeleteNode;
        private System.Windows.Forms.Button BtnCancel;
    }
}