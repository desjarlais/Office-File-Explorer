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
            lblCurrentPath.Text = templatePath;
        }

        public FrmChangeTemplate()
        {
            InitializeComponent();
            lblCurrentPath.Text = string.Empty;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (Owner is FrmBatch f && tbNewPath.Text.Length > 0)
            {
                f.DefaultTemplate = FileUtilities.ConvertFilePathToUri(tbNewPath.Text); ;
            }

            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
