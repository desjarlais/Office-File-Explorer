using Office_File_Explorer.App_Helpers;
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

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (Owner is FrmMain f && tbNewPath.Text.Length > 0)
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
