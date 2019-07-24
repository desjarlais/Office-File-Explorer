using System;
using System.Windows.Forms;

namespace Office_File_Explorer.Forms
{
    public partial class FrmSettings : Form
    {
        public FrmSettings()
        {
            InitializeComponent();
            ckRemoveFallback.Checked = Properties.Settings.Default.RemoveFallback == "true";
            ckOpenInWord.Checked = Properties.Settings.Default.OpenInWord == "true";
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.RemoveFallback = ckRemoveFallback.Checked ? "true" : "false";
            Properties.Settings.Default.OpenInWord = ckOpenInWord.Checked ? "true" : "false";
            Properties.Settings.Default.Save();
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
