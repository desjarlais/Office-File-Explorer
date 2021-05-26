using Office_File_Explorer.App_Helpers;
using Office_File_Explorer.Forms;
using System;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmChangeTemplate : Form
    {
        public FrmChangeTemplate(string templatePath)
        {
            InitializeComponent();
            tbOldPath.Text = templatePath;
        }

        public FrmChangeTemplate()
        {
            InitializeComponent();
            tbOldPath.Text = string.Empty;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (tbNewPath.Text.Length > 0)
            {
                if (Owner is FrmBatch f)
                {
                    if (tbNewPath.Text != "Normal")
                    {
                        f.DefaultTemplate = FileUtilities.ConvertFilePathToUri(tbNewPath.Text);
                    }
                    else
                    {
                        f.DefaultTemplate = "Normal";
                    }
                }
                else if (Owner is FrmMain fm)
                {
                    if (tbNewPath.Text != "Normal")
                    {
                        fm.DefaultTemplate = FileUtilities.ConvertFilePathToUri(tbNewPath.Text);
                    }
                    else
                    {
                        fm.DefaultTemplate = "Normal";
                    }
                }
            }

            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            if (Owner is FrmBatch f)
            {
                f.DefaultTemplate = "Cancel";
            }
            else if (Owner is FrmMain fm)
            {
                fm.DefaultTemplate = "Cancel";
            }
            Close();
        }
    }
}
