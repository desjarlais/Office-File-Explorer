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

            if (OfficeApp == "Word")
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
        }

        /// <summary>
        /// Reset all checkboxes, then enable all PPT options
        /// </summary>
        public void EnablePPTUI()
        {
            ResetCheckboxes();
            RdoNotes.Enabled = true;
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
