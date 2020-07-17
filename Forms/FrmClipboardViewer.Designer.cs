namespace Office_File_Explorer.Forms
{
    partial class FrmClipboardViewer
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmClipboardViewer));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.clipboardToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.refreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ownerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.clearToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.autoRefreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.showRichTextToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.showMemoryInHexToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.showPicturesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.lbClipFormats = new System.Windows.Forms.ListBox();
            this.pbClipData = new System.Windows.Forms.PictureBox();
            this.rtbClipData = new System.Windows.Forms.RichTextBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbClipData)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.clipboardToolStripMenuItem,
            this.viewToolStripMenuItem,
            this.toolStripMenuItem1});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(800, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // clipboardToolStripMenuItem
            // 
            this.clipboardToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.refreshToolStripMenuItem,
            this.ownerToolStripMenuItem,
            this.clearToolStripMenuItem,
            this.saveAsToolStripMenuItem});
            this.clipboardToolStripMenuItem.Name = "clipboardToolStripMenuItem";
            this.clipboardToolStripMenuItem.Size = new System.Drawing.Size(71, 20);
            this.clipboardToolStripMenuItem.Text = "Clipboard";
            // 
            // refreshToolStripMenuItem
            // 
            this.refreshToolStripMenuItem.Name = "refreshToolStripMenuItem";
            this.refreshToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.refreshToolStripMenuItem.Text = "Refresh";
            this.refreshToolStripMenuItem.Click += new System.EventHandler(this.RefreshToolStripMenuItem_Click);
            // 
            // ownerToolStripMenuItem
            // 
            this.ownerToolStripMenuItem.Name = "ownerToolStripMenuItem";
            this.ownerToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.ownerToolStripMenuItem.Text = "Owner";
            this.ownerToolStripMenuItem.Click += new System.EventHandler(this.OwnerToolStripMenuItem_Click);
            // 
            // clearToolStripMenuItem
            // 
            this.clearToolStripMenuItem.Name = "clearToolStripMenuItem";
            this.clearToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.clearToolStripMenuItem.Text = "Clear";
            this.clearToolStripMenuItem.Click += new System.EventHandler(this.ClearToolStripMenuItem_Click);
            // 
            // saveAsToolStripMenuItem
            // 
            this.saveAsToolStripMenuItem.Name = "saveAsToolStripMenuItem";
            this.saveAsToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.saveAsToolStripMenuItem.Text = "Save As";
            this.saveAsToolStripMenuItem.Click += new System.EventHandler(this.SaveAsToolStripMenuItem_Click);
            // 
            // viewToolStripMenuItem
            // 
            this.viewToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.autoRefreshToolStripMenuItem,
            this.showRichTextToolStripMenuItem,
            this.showMemoryInHexToolStripMenuItem,
            this.showPicturesToolStripMenuItem});
            this.viewToolStripMenuItem.Name = "viewToolStripMenuItem";
            this.viewToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.viewToolStripMenuItem.Text = "View";
            // 
            // autoRefreshToolStripMenuItem
            // 
            this.autoRefreshToolStripMenuItem.Checked = true;
            this.autoRefreshToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked;
            this.autoRefreshToolStripMenuItem.Name = "autoRefreshToolStripMenuItem";
            this.autoRefreshToolStripMenuItem.Size = new System.Drawing.Size(188, 22);
            this.autoRefreshToolStripMenuItem.Text = "Auto Refresh";
            this.autoRefreshToolStripMenuItem.Click += new System.EventHandler(this.AutoRefreshToolStripMenuItem_Click);
            // 
            // showRichTextToolStripMenuItem
            // 
            this.showRichTextToolStripMenuItem.Name = "showRichTextToolStripMenuItem";
            this.showRichTextToolStripMenuItem.Size = new System.Drawing.Size(188, 22);
            this.showRichTextToolStripMenuItem.Text = "Show Rich Text";
            this.showRichTextToolStripMenuItem.Click += new System.EventHandler(this.ShowRichTextToolStripMenuItem_Click);
            // 
            // showMemoryInHexToolStripMenuItem
            // 
            this.showMemoryInHexToolStripMenuItem.Name = "showMemoryInHexToolStripMenuItem";
            this.showMemoryInHexToolStripMenuItem.Size = new System.Drawing.Size(188, 22);
            this.showMemoryInHexToolStripMenuItem.Text = "Show Memory in Hex";
            this.showMemoryInHexToolStripMenuItem.Click += new System.EventHandler(this.ShowMemoryInHexToolStripMenuItem_Click);
            // 
            // showPicturesToolStripMenuItem
            // 
            this.showPicturesToolStripMenuItem.Name = "showPicturesToolStripMenuItem";
            this.showPicturesToolStripMenuItem.Size = new System.Drawing.Size(188, 22);
            this.showPicturesToolStripMenuItem.Text = "Show Pictures";
            this.showPicturesToolStripMenuItem.Click += new System.EventHandler(this.ShowPicturesToolStripMenuItem_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(12, 20);
            // 
            // lbClipFormats
            // 
            this.lbClipFormats.Dock = System.Windows.Forms.DockStyle.Left;
            this.lbClipFormats.FormattingEnabled = true;
            this.lbClipFormats.Location = new System.Drawing.Point(0, 24);
            this.lbClipFormats.Name = "lbClipFormats";
            this.lbClipFormats.Size = new System.Drawing.Size(327, 426);
            this.lbClipFormats.TabIndex = 1;
            this.lbClipFormats.SelectedIndexChanged += new System.EventHandler(this.LbClipFormats_SelectedIndexChanged);
            // 
            // pbClipData
            // 
            this.pbClipData.Dock = System.Windows.Forms.DockStyle.Right;
            this.pbClipData.Location = new System.Drawing.Point(333, 24);
            this.pbClipData.Name = "pbClipData";
            this.pbClipData.Size = new System.Drawing.Size(467, 426);
            this.pbClipData.TabIndex = 2;
            this.pbClipData.TabStop = false;
            this.pbClipData.Visible = false;
            // 
            // rtbClipData
            // 
            this.rtbClipData.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rtbClipData.Location = new System.Drawing.Point(327, 24);
            this.rtbClipData.Name = "rtbClipData";
            this.rtbClipData.Size = new System.Drawing.Size(6, 426);
            this.rtbClipData.TabIndex = 3;
            this.rtbClipData.Text = "";
            // 
            // FrmClipboardViewer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.rtbClipData);
            this.Controls.Add(this.pbClipData);
            this.Controls.Add(this.lbClipFormats);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmClipboardViewer";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Clipboard Viewer";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FrmClipboardViewer_FormClosed);
            this.Shown += new System.EventHandler(this.FrmClipboardViewer_Shown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbClipData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ListBox lbClipFormats;
        private System.Windows.Forms.PictureBox pbClipData;
        private System.Windows.Forms.ToolStripMenuItem clipboardToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewToolStripMenuItem;
        private System.Windows.Forms.RichTextBox rtbClipData;
        private System.Windows.Forms.ToolStripMenuItem refreshToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ownerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem clearToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveAsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem autoRefreshToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem showRichTextToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem showMemoryInHexToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem showPicturesToolStripMenuItem;
    }
}