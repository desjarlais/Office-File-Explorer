using DocumentFormat.OpenXml.Packaging;
using Office_File_Explorer.App_Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace Office_File_Explorer.Forms
{
    public partial class FrmBatch : Form
    {
        public List<string> files = new List<string>();
        public string fileType = "";
        public string fType = "";

        public FrmBatch()
        {
            InitializeComponent();
            BtnFixNotesPageSize.Enabled = false;
        }

        public string GetFileExtension()
        {
            if (rdoWord.Checked == true)
            {
                fileType = "*.docx";
                fType = StringResources.word;
            }
            else if (rdoExcel.Checked == true)
            {
                fileType = "*.xlsx";
                fType = StringResources.excel;
            }
            else if (rdoPowerPoint.Checked == true)
            {
                fileType = "*.pptx";
                fType = StringResources.powerpoint;
            }

            return fileType;
        }

        public void PopulateAndDisplayFiles()
        {
            try
            {
                lstOutput.Items.Clear();
                files.Clear();

                // get all the file paths for .docx files in the folder
                DirectoryInfo dir = new DirectoryInfo(TxbDirectoryPath.Text);
                foreach (FileInfo f in dir.GetFiles(GetFileExtension()))
                {
                    if (f.Name.StartsWith("~"))
                    {
                        // we don't want to change temp files
                        continue;
                    }
                    else
                    {
                        // populate the list of file paths
                        files.Add(f.FullName);
                        lstOutput.Items.Add(f.FullName);
                    }
                }
            }
            catch (ArgumentException ae)
            {
                LoggingHelper.Log("BtnPopulateAndDisplayFiles Error: " + ae.Message);
                lstOutput.Items.Add("** Invalid folder path **");
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("PopulateAndDisplayFiles Error: " + ex.Message);
            }
        }

        private void BtnBrowseDirectory_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult result = folderBrowserDialog1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    TxbDirectoryPath.Text = folderBrowserDialog1.SelectedPath;
                    PopulateAndDisplayFiles();
                }
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("BtnBrowseDirectory Error: " + ex.Message);
            }
        }

        private void BtnChangeCustomProps_Click(object sender, EventArgs e)
        {
            FrmCustomProperties cFrm = new FrmCustomProperties(files, fType)
            {
                Owner = this
            };
            cFrm.ShowDialog();

            lstOutput.Items.Clear();
            lstOutput.Items.Add("** Batch Processing done **");
        }

        private void rdoPowerPoint_CheckedChanged(object sender, EventArgs e)
        {
            // only need to run if the change was to enable the button
            // disabling one of the other options causes multiple events
            // so we just want to run the populate function once
            if (rdoPowerPoint.Checked)
            {
                PopulateAndDisplayFiles();
                BtnFixNotesPageSize.Enabled = true;
            }
        }

        private void rdoExcel_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoExcel.Checked)
            {
                PopulateAndDisplayFiles();
                BtnFixNotesPageSize.Enabled = false;
            }
        }

        private void rdoWord_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoWord.Checked)
            {
                PopulateAndDisplayFiles();
                BtnFixNotesPageSize.Enabled = false;
            }
        }

        private void BtnChangeTheme_Click(object sender, EventArgs e)
        {
            string sThemeFilePath = StringResources.emptyString;
            lstOutput.Items.Clear();

            OpenFileDialog fDialog = new OpenFileDialog
            {
                Title = "Select Office Theme File.",
                Filter = "Open XML Theme File | *.xml",
                RestoreDirectory = true,
                InitialDirectory = @"%userprofile%"
            };

            if (fDialog.ShowDialog() == DialogResult.OK)
            {
                sThemeFilePath = fDialog.FileName.ToString();

                foreach (string f in files)
                {
                    try
                    {
                        // call the replace function using the theme file provided
                        OfficeHelpers.ReplaceTheme(f, sThemeFilePath, fType);
                        LoggingHelper.Log(f + ": Theme Replaced.");
                    }
                    catch (Exception ex)
                    {
                        LoggingHelper.Log(f + " : Failed to replace theme : Error = " + ex.Message);
                    }
                }

                lstOutput.Items.Add("** Themes Changed ** ");
            }
            else
            {
                return;
            }
        }

        private void BtnFixNotesPageSize_Click(object sender, EventArgs e)
        {
            lstOutput.Items.Clear();

            foreach (string f in files)
            {
                try
                {
                    using (PresentationDocument document = PresentationDocument.Open(f, true))
                    {
                        PowerPoint_Helpers.PowerPointOpenXml.ChangeNotesPageSize(document);
                        lstOutput.Items.Add(f + ": Notes Page Size Reset.");
                        LoggingHelper.Log(f + ": Notes Page Size Reset.");
                    }
                }
                catch (Exception ex)
                {
                    lstOutput.Items.Add(f + " not saved : error = " + ex.Message);
                    LoggingHelper.Log(f + " not saved : error = " + ex.Message);
                }
            }
        }
    }
}
