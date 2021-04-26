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
            if (Properties.Settings.Default.RemoveFallback == true)
            {
                ckRemoveFallback.Checked = true;
            }
            
            if (Properties.Settings.Default.OpenInWord == true)
            {
                ckOpenInWord.Checked = true;
            }
             
            if (Properties.Settings.Default.FixGroupedShapes == true)
            {
                ckGroupShapeFix.Checked = true;
            }
            
            if (Properties.Settings.Default.ResetNotesMaster == true)
            {
                ckResetNotesMaster.Checked = true;
            }

            if (Properties.Settings.Default.DeleteCopiesOnExit == true)
            {
                ckDeleteCopies.Checked = true;
            }

            if (Properties.Settings.Default.RemoveCorruptAtMentions == true)
            {
                ckRemoveCorruptAtMentions.Checked = true;
            }

            if (Properties.Settings.Default.ListCellValuesSax == true)
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
            Properties.Settings.Default.RemoveFallback = ckRemoveFallback.Checked;
            Properties.Settings.Default.OpenInWord = ckOpenInWord.Checked;
            Properties.Settings.Default.FixGroupedShapes = ckGroupShapeFix.Checked;
            Properties.Settings.Default.ResetNotesMaster = ckResetNotesMaster.Checked;
            Properties.Settings.Default.DeleteCopiesOnExit = ckDeleteCopies.Checked;
            Properties.Settings.Default.RemoveCorruptAtMentions = ckRemoveCorruptAtMentions.Checked;
            
            if (rdoSax.Checked == true)
            {
                Properties.Settings.Default.ListCellValuesSax = true;
            }
            else
            {
                Properties.Settings.Default.ListCellValuesSax = false;
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
