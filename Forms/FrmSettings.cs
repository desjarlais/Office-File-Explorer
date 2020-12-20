using System;
using System.Windows.Forms;

namespace Office_File_Explorer.Forms
{
    public partial class FrmSettings : Form
    {
        public FrmSettings()
        {
            InitializeComponent();

            // populate checkboxes from settings
            if (Properties.Settings.Default.RemoveFallback == "true")
            {
                ckRemoveFallback.Checked = true;
            }
            
            if (Properties.Settings.Default.OpenInWord == "true")
            {
                ckOpenInWord.Checked = true;
            }
             
            if (Properties.Settings.Default.FixGroupedShapes == "true")
            {
                ckGroupShapeFix.Checked = true;
            }
            
            if (Properties.Settings.Default.ResetNotesMaster == "true")
            {
                ckResetNotesMaster.Checked = true;
            }

            if (Properties.Settings.Default.DeleteCopiesOnExit == true)
            {
                ckDeleteCopies.Checked = true;
            }

            if (Properties.Settings.Default.ListCellValuesSax == "true")
            {
                rdoSax.Checked = true;
            }
            else
            {
                rdoDom.Checked = true;
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.RemoveFallback = ckRemoveFallback.Checked ? "true" : "false";
            Properties.Settings.Default.OpenInWord = ckOpenInWord.Checked ? "true" : "false";
            Properties.Settings.Default.FixGroupedShapes = ckGroupShapeFix.Checked ? "true" : "false";
            Properties.Settings.Default.ResetNotesMaster = ckResetNotesMaster.Checked ? "true" : "false";
            Properties.Settings.Default.DeleteCopiesOnExit = ckDeleteCopies.Checked ? true : false;
            
            if (rdoSax.Checked == true)
            {
                Properties.Settings.Default.ListCellValuesSax = "true";
            }
            else
            {
                Properties.Settings.Default.ListCellValuesSax = "false";
            }
            
            Properties.Settings.Default.Save();
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
