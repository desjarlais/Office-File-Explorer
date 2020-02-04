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
        }

        public string GetFileExtension()
        {
            if (rdoWord.Checked == true)
            {
                fileType = "*.docx";
                fType = "Word";
            }
            else if (rdoExcel.Checked == true)
            {
                fileType = "*.xlsx";
                fType = "Excel";
            }
            else if (rdoPowerPoint.Checked == true)
            {
                fileType = "*.pptx";
                fType = "PowerPoint";
            }

            return fileType;
        }

        private void BtnBrowseDirectory_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult result = folderBrowserDialog1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    TxbDirectoryPath.Text = folderBrowserDialog1.SelectedPath;

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
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
            lstOutput.Items.Add("Batch Processing done");
        }
    }
}
