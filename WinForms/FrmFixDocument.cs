using Office_File_Explorer.App_Helpers;
using System;
using System.Windows.Forms;

namespace Office_File_Explorer.Forms
{
    public partial class FrmFixDocument : Form
    {
        public string OfficeApp;

        public string OptionSelected { get; set; }

        public FrmFixDocument(string app)
        {
            InitializeComponent();
            OfficeApp = app;

            if (OfficeApp == StringResources.wWord)
            {
                EnableWordUI();
                RdoBK.Checked = true;
            }
            else
            {
                EnablePPTUI();
                RdoNotes.Checked = true;
            }
        }

        /// <summary>
        /// Reset all checkboxes to false
        /// </summary>
        public void ResetCheckboxes()
        {
            RdoBK.Checked = false;
            RdoLT.Checked = false;
            RdoRev.Checked = false;
            RdoEndnotes.Checked = false;
            RdoNotes.Checked = false;
            RdoFixNotesPageWithFile.Checked = false;
            RdoTblGrid.Checked = false;
            RdoFixComments.Checked = false;
            RdoFixHyperlinks.Checked = false;
            RdoFixCoAuthHyperlinks.Checked = false;
        }

        /// <summary>
        /// Reset all checkboxes, then enable all Word options
        /// </summary>
        public void EnableWordUI()
        {
            ResetCheckboxes();
            RdoBK.Enabled = true;
            RdoLT.Enabled = true;
            RdoRev.Enabled = true;
            RdoEndnotes.Enabled = true;
            RdoTblGrid.Enabled = true;
            RdoFixComments.Enabled = true;
            RdoFixHyperlinks.Enabled = true;
            RdoFixCoAuthHyperlinks.Enabled = true;
        }

        /// <summary>
        /// Reset all checkboxes, then enable all PPT options
        /// </summary>
        public void EnablePPTUI()
        {
            ResetCheckboxes();
            RdoNotes.Enabled = true;
            RdoFixNotesPageWithFile.Enabled = true;
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            if (RdoBK.Checked)
            {
                OptionSelected = "Bookmark";
            }
            else if (RdoEndnotes.Checked)
            {
                OptionSelected = "Endnote";
            }
            else if (RdoLT.Checked)
            {
                OptionSelected = "ListTemplates";
            }
            else if (RdoRev.Checked)
            {
                OptionSelected = "Revision";
            }
            else if (RdoNotes.Checked)
            {
                OptionSelected = "Notes";
            }
            else if (RdoTblGrid.Checked)
            {
                OptionSelected = "TblGrid";
            }
            else if (RdoFixNotesPageWithFile.Checked)
            {
                OptionSelected = "NotesWithFile";
            }
            else if (RdoFixComments.Checked)
            {
                OptionSelected = "FixComments";
            }
            else if (RdoFixHyperlinks.Checked)
            {
                OptionSelected = "FixHyperlinks";
            }
            else if (RdoFixCoAuthHyperlinks.Checked)
            {
                OptionSelected = "FixCoAuthHyperlinks";
            }
            
            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            OptionSelected = "Cancel";
            Close();
        }
    }
}
