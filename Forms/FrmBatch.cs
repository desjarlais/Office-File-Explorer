using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Office_File_Explorer.App_Helpers;
using Office_File_Explorer.Word_Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
            DisableUI();
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

        public void DisableUI()
        {
            // disable all buttons
            BtnFixNotesPageSize.Enabled = false;
            BtnChangeTheme.Enabled = false;
            BtnChangeCustomProps.Enabled = false;
            BtnRemovePII.Enabled = false;
            BtnFixCorruptBookmarks.Enabled = false;

            // disable all radio buttons
            rdoExcel.Enabled = false;
            rdoPowerPoint.Enabled = false;
            rdoWord.Enabled = false;

            lstOutput.Items.Clear();
        }

        public void EnableUI()
        {
            // enable buttons that work for each app
            BtnChangeTheme.Enabled = true;
            BtnChangeCustomProps.Enabled = true;
            BtnRemovePII.Enabled = true;

            // enable the radio buttons
            rdoExcel.Enabled = true;
            rdoPowerPoint.Enabled = true;
            rdoWord.Enabled = true;

            // now check which radio button is selected and light up appropriate buttons
            if (rdoWord.Checked == true)
            {
                BtnFixCorruptBookmarks.Enabled = true;
                BtnFixNotesPageSize.Enabled = false;
            }

            if (rdoPowerPoint.Checked == true)
            {
                BtnFixNotesPageSize.Enabled = true;
                BtnFixCorruptBookmarks.Enabled = false;
            }

            if (rdoExcel.Checked == true)
            {
                BtnFixNotesPageSize.Enabled = false;
                BtnFixCorruptBookmarks.Enabled = false;
            }
        }

        public void PopulateAndDisplayFiles()
        {
            try
            {
                lstOutput.Items.Clear();
                files.Clear();
                int fCount = 0;

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
                        fCount++;
                    }
                }

                if (fCount == 0)
                {
                    lstOutput.Items.Add("** No Files **");
                }
            }
            catch (ArgumentException ae)
            {
                LoggingHelper.Log("BtnPopulateAndDisplayFiles Error: " + ae.Message);
                lstOutput.Items.Add("** Invalid folder path **");
            }
            catch (DirectoryNotFoundException dnfe)
            {
                LoggingHelper.Log("BtnPopulateAndDisplayFiles Error: " + dnfe.Message);
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
                    EnableUI();
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
            if (rdoPowerPoint.Checked)
            {
                PopulateAndDisplayFiles();
                EnableUI();
            }
        }

        private void rdoExcel_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoExcel.Checked)
            {
                PopulateAndDisplayFiles();
                EnableUI();
            }
        }

        private void rdoWord_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoWord.Checked)
            {
                PopulateAndDisplayFiles();
                EnableUI();
            }
        }

        private void BtnChangeTheme_Click(object sender, EventArgs e)
        {
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
                string sThemeFilePath = fDialog.FileName.ToString();

                foreach (string f in files)
                {
                    try
                    {
                        // call the replace function using the theme file provided
                        OfficeHelpers.ReplaceTheme(f, sThemeFilePath, fType);
                        LoggingHelper.Log(f + "--> Theme Replaced.");
                        lstOutput.Items.Add(f + "--> Theme Replaced.");
                    }
                    catch (Exception ex)
                    {
                        LoggingHelper.Log(f + " --> Failed to replace theme : Error = " + ex.Message);
                        lstOutput.Items.Add(f + " --> Failed to replace theme : Error = " + ex.Message);
                    }
                }
            }
            else
            {
                return;
            }
        }

        private void BtnFixNotesPageSize_Click(object sender, EventArgs e)
        {
            lstOutput.Items.Clear();
            Cursor = Cursors.WaitCursor;

            foreach (string f in files)
            {
                try
                {
                    using (PresentationDocument document = PresentationDocument.Open(f, true))
                    {
                        PowerPoint_Helpers.PowerPointOpenXml.ChangeNotesPageSize(document);
                        lstOutput.Items.Add(f + StringResources.arrow + StringResources.pptNotesSizeReset);
                        LoggingHelper.Log(f + StringResources.arrow + StringResources.pptNotesSizeReset);
                    }
                }
                catch (NullReferenceException nre)
                {
                    lstOutput.Items.Add(f + StringResources.arrow + "** Document does not contain Notes Master **");
                    LoggingHelper.Log(f + StringResources.arrow + "error = " + nre.Message);
                }
                catch (Exception ex)
                {
                    lstOutput.Items.Add(f + StringResources.arrow + "error = " + ex.Message);
                    LoggingHelper.Log(f + StringResources.arrow + "error = " + ex.Message);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
        }

        private void TxbDirectoryPath_TextChanged(object sender, EventArgs e)
        {
            // Word is enabled by default, once a text change happens we can enable UI
            if (rdoWord.Enabled == false)
            {
                EnableUI();
            }

            // if someone deletes the text, disable ui
            if (TxbDirectoryPath.Text.Length == 0)
            {
                DisableUI();
            }
        }

        private void BtnRemovePII_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                foreach (string f in files)
                {
                    using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                    {
                        if (WordExtensionClass.HasPersonalInfo(document) == true)
                        {
                            WordExtensionClass.RemovePersonalInfo(document);
                            LoggingHelper.Log(f + " : PII removed from file.");
                        }
                        else
                        {
                            LoggingHelper.Log(f + " : does not contain PII.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("BtnRemovePII Error: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnFixCorruptBookmarks_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                lstOutput.Items.Clear();
                foreach (string f in files)
                {
                    using (WordprocessingDocument package = WordprocessingDocument.Open(f, true))
                    {
                        IEnumerable<BookmarkStart> bkStartList = package.MainDocumentPart.Document.Descendants<BookmarkStart>();
                        IEnumerable<BookmarkEnd> bkEndList = package.MainDocumentPart.Document.Descendants<BookmarkEnd>();
                        List<string> removedBookmarkIds = new List<string>();

                        if (bkStartList.Count() > 0)
                        {
                            foreach (BookmarkStart bk in bkStartList)
                            {
                                var cElem = bk.Parent;
                                var pElem = bk.Parent;
                                bool endLoop = false;

                                do
                                {
                                    if (cElem.Parent.ToString().Contains("DocumentFormat.OpenXml.Wordprocessing.Sdt"))
                                    {
                                        foreach (OpenXmlElement oxe in cElem.Parent.ChildElements)
                                        {
                                            if (oxe.GetType().Name == "SdtProperties")
                                            {
                                                foreach (OpenXmlElement oxeSdtAlias in oxe)
                                                {
                                                    if (oxeSdtAlias.GetType().Name == "SdtContentText")
                                                    {
                                                        // if the parent is a content control, bookmark is only allowed in rich text
                                                        // if this is a plain text control, it is invalid
                                                        // add the id to the list of bookmarks that need to be deleted
                                                        removedBookmarkIds.Add(bk.Id);
                                                        endLoop = true;
                                                    }
                                                }
                                            }
                                        }

                                        pElem = cElem.Parent;
                                        cElem = pElem;
                                    }
                                    else
                                    {
                                        pElem = cElem.Parent;
                                        cElem = pElem;

                                        // if the parent is body, we can stop looping up
                                        // otherwise, set cElem to the parent so we can continue moving up the element chain
                                        if (pElem.ToString() == "DocumentFormat.OpenXml.Wordprocessing.Body")
                                        {
                                            endLoop = true;
                                        }
                                    }
                                } while (endLoop == false);
                            }

                            // now that we have the list of bookmark id's to be removed
                            // loop each list and delete any bookmark that has a matching id
                            foreach (var o in removedBookmarkIds)
                            {
                                foreach (BookmarkStart bkStart in bkStartList)
                                {
                                    if (bkStart.Id == o)
                                    {
                                        bkStart.Remove();
                                    }
                                }

                                foreach (BookmarkEnd bkEnd in bkEndList)
                                {
                                    if (bkEnd.Id == o)
                                    {
                                        bkEnd.Remove();
                                    }
                                }
                            }

                            // save the part
                            package.MainDocumentPart.Document.Save();
                            if (removedBookmarkIds.Count > 0)
                            {
                                lstOutput.Items.Add(f + " : " + "bookmarks fixed.");
                            }
                            else
                            {
                                lstOutput.Items.Add(f + " : " + "does not contain any corrupt bookmarks.");
                            }
                        }
                        else
                        {
                            lstOutput.Items.Add(f + " : " + " does not contain any bookmarks.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add(StringResources.errorText + ex.Message);
                LoggingHelper.Log("BtnListBookmarks: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }
    }
}