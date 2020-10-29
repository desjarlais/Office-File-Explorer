// Open Xml SDK refs
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// app refs
using Office_File_Explorer.App_Helpers;
using Office_File_Explorer.Word_Helpers;

//.NET refs
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using System.Xml.XPath;

// namespace refs
using O = DocumentFormat.OpenXml;

namespace Office_File_Explorer.Forms
{
    public partial class FrmBatch : Form
    {
        public List<string> files = new List<string>();
        public string fileType = "";
        public string fType = "";
        public bool nodeDeleted = false;

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
            BtnPPTResetPII.Enabled = false;
            BtnRemovePII.Enabled = false;
            BtnFixCorruptBookmarks.Enabled = false;
            BtnFixCorruptRevisions.Enabled = false;
            BtnConvertStrict.Enabled = false;
            BtnDeleteProps.Enabled = false;
            BtnFixTableProps.Enabled = false;
            BtnDeleteRequestStatus.Enabled = false;

            // disable all radio buttons
            rdoExcel.Enabled = false;
            rdoPowerPoint.Enabled = false;
            rdoWord.Enabled = false;

            // disable checkbox
            ckSearchSubfolders.Enabled = false;

            lstOutput.Items.Clear();
        }

        public void EnableUI()
        {
            // enable buttons that work for each app
            BtnChangeTheme.Enabled = true;
            BtnChangeCustomProps.Enabled = true;
            BtnRemovePII.Enabled = true;
            BtnDeleteProps.Enabled = true;
            BtnDeleteRequestStatus.Enabled = true;

            // enable the radio buttons
            rdoExcel.Enabled = true;
            rdoPowerPoint.Enabled = true;
            rdoWord.Enabled = true;

            // enable checkbox
            ckSearchSubfolders.Enabled = true;

            // now check which radio button is selected and light up appropriate buttons
            if (rdoWord.Checked == true)
            {
                BtnFixCorruptBookmarks.Enabled = true;
                BtnFixCorruptRevisions.Enabled = true;
                BtnFixTableProps.Enabled = true;
                BtnRemovePII.Enabled = true;

                BtnFixNotesPageSize.Enabled = false;
                BtnPPTResetPII.Enabled = false;
                BtnConvertStrict.Enabled = false;
            }

            if (rdoPowerPoint.Checked == true)
            {
                BtnRemovePII.Enabled = true;
                BtnPPTResetPII.Enabled = true;
                BtnFixNotesPageSize.Enabled = true;

                BtnFixCorruptBookmarks.Enabled = false;
                BtnFixCorruptRevisions.Enabled = false;
                BtnFixTableProps.Enabled = false;
                BtnConvertStrict.Enabled = false;
            }

            if (rdoExcel.Checked == true)
            {
                BtnConvertStrict.Enabled = true;

                BtnFixNotesPageSize.Enabled = false;
                BtnFixCorruptBookmarks.Enabled = false;
                BtnFixCorruptRevisions.Enabled = false;
                BtnFixTableProps.Enabled = false;
                BtnRemovePII.Enabled = false;
                BtnPPTResetPII.Enabled = false;
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
                if (ckSearchSubfolders.Checked == true)
                {
                    foreach (FileInfo f in dir.GetFiles(GetFileExtension(), SearchOption.AllDirectories))
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
                }
                else
                {
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
                lstOutput.Items.Clear();

                foreach (string f in files)
                {
                    using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                    {
                        if (WordExtensionClass.HasPersonalInfo(document) == true)
                        {
                            WordExtensionClass.RemovePersonalInfo(document);
                            lstOutput.Items.Add(f + " : PII removed from file.");
                            LoggingHelper.Log(f + " : PII removed from file.");
                        }
                        else
                        {
                            lstOutput.Items.Add(f + " : does not contain PII.");
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
                    if (WordOpenXml.RemoveMissingBookmarkTags(f) == true || WordOpenXml.RemovePlainTextCcFromBookmark(f) == true)
                    {
                        lstOutput.Items.Add(f + " : Fixed Corrupt Bookmarks");
                    }
                    else
                    {
                        lstOutput.Items.Add(f + " : No Corrupt Bookmarks Found");
                    }
                }
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add(StringResources.errorText + ex.Message);
                LoggingHelper.Log("BtnFixCorruptBookmarks: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnFixCorruptRevisions_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                lstOutput.Items.Clear();
                foreach (string f in files)
                {
                    bool isFixed = false;
                    
                    if (FileUtilities.IsZipArchiveFile(f) == false)
                    {
                        lstOutput.Items.Add(f + " : Not A Valid Office File");
                        return;
                    }

                    using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                    {
                        if (WordOpenXml.IsPartNull(document, "DeletedRun") == false)
                        {
                            var deleted = document.MainDocumentPart.Document.Descendants<DeletedRun>().ToList();

                            // loop each DeletedRun
                            foreach (DeletedRun dr in deleted)
                            {
                                foreach (OpenXmlElement oxedr in dr)
                                {
                                    // if we have a Run, we need to look for Text tags
                                    if (oxedr.GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Run")
                                    {
                                        Run r = (Run)oxedr;
                                        foreach (OpenXmlElement oxe in oxedr.ChildElements)
                                        {
                                            // you can't have a Text tag inside a DeletedRun
                                            if (oxe.GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Text")
                                            {
                                                // create a DeletedText object so we can replace it with the Text tag
                                                DeletedText dt = new DeletedText();

                                                // check for attributes
                                                if (oxe.HasAttributes)
                                                {
                                                    if (oxe.GetAttributes().Count > 0)
                                                    {
                                                        dt.SetAttributes(oxe.GetAttributes());
                                                    }
                                                }

                                                // set the text value
                                                dt.Text = oxe.InnerText;

                                                // replace the Text with new DeletedText
                                                r.ReplaceChild(dt, oxe);
                                                isFixed = true;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        
                        // now save the file if we have changes
                        if (isFixed == true)
                        {
                            document.MainDocumentPart.Document.Save();
                            lstOutput.Items.Add(f + ": Fixed Corrupt Revisions");
                        }
                        else
                        {
                            lstOutput.Items.Add(f + ": No Corrupt Revisions Found");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add(StringResources.errorText + ex.Message);
                LoggingHelper.Log("BtnFixCorruptRevisions: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnPPTResetPII_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                lstOutput.Items.Clear();
                foreach (string f in files)
                {
                    bool isFixed = false;
                    Cursor = Cursors.WaitCursor;
                    using (PresentationDocument document = PresentationDocument.Open(f, true))
                    {
                        document.PresentationPart.Presentation.RemovePersonalInfoOnSave = false;
                        document.PresentationPart.Presentation.Save();
                    }

                    if (isFixed)
                    {
                        lstOutput.Items.Add(f + ": PII Reset");
                    }
                    else
                    {
                        lstOutput.Items.Add(f + ": PII Not Reset");
                    }
                }
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add(StringResources.errorText + ex.Message);
                LoggingHelper.Log("BtnPPTResetPII: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnConvertStrict_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                lstOutput.Items.Clear();
                foreach (string f in files)
                {
                    // check if the excelcnv.exe exists
                    string excelcnvPath;

                    if (File.Exists(StringResources.sameBitnessO365))
                    {
                        excelcnvPath = StringResources.sameBitnessO365;
                    }
                    else if (File.Exists(StringResources.x86OfficeO365))
                    {
                        excelcnvPath = StringResources.x86OfficeO365;
                    }
                    else if (File.Exists(StringResources.sameBitnessMSI2016))
                    {
                        excelcnvPath = StringResources.sameBitnessMSI2016;
                    }
                    else if (File.Exists(StringResources.x86OfficeMSI2016))
                    {
                        excelcnvPath = StringResources.x86OfficeMSI2016;
                    }
                    else if (File.Exists(StringResources.sameBitnessMSI2013))
                    {
                        excelcnvPath = StringResources.sameBitnessMSI2013;
                    }
                    else if (File.Exists(StringResources.x86OfficeMSI2013))
                    {
                        excelcnvPath = StringResources.x86OfficeMSI2013;
                    }
                    else
                    {
                        excelcnvPath = StringResources.emptyString;
                    }

                    // check if the file is strict
                    bool isStrict = false;

                    using (Package package = Package.Open(f, FileMode.Open, FileAccess.Read))
                    {
                        foreach (PackagePart part in package.GetParts())
                        {
                            if (part.Uri.ToString() == "/xl/workbook.xml")
                            {
                                try
                                {
                                    string docText = null;
                                    using (StreamReader sr = new StreamReader(part.GetStream()))
                                    {
                                        docText = sr.ReadToEnd();
                                        if (docText.Contains(@"conformance=""strict"""))
                                        {
                                            isStrict = true;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    LoggingHelper.Log(ex.Message);
                                }
                            }
                        }
                    }

                    if (isStrict == true && excelcnvPath != StringResources.emptyString)
                    {
                        // setup destination file path
                        string strOriginalFile = f;
                        string strOutputPath = Path.GetDirectoryName(strOriginalFile) + "\\";
                        string strFileExtension = Path.GetExtension(strOriginalFile);
                        string strOutputFileName = strOutputPath + Path.GetFileNameWithoutExtension(strOriginalFile) + "(Fixed)" + strFileExtension;

                        // run the command to convert the file "excelcnv.exe -nme -oice "file-path" "converted-file-path""
                        string cParams = " -nme -oice " + '"' + f + '"' + " " + '"' + strOutputFileName + '"';
                        var proc = Process.Start(excelcnvPath, cParams);
                        proc.Close();
                        lstOutput.Items.Add(f + " : Converted Successfully");
                        lstOutput.Items.Add("   File Location: " + strOutputFileName);
                    }
                    else
                    {
                        lstOutput.Items.Add(f + " : Is Not Strict Open Xml Format");
                    }
                }
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add(StringResources.errorText + ex.Message);
                LoggingHelper.Log("BtnConvertStrict: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnDeleteProps_Click(object sender, EventArgs e)
        {
            try
            {
                string propNameToDelete = "";
                lstOutput.Items.Clear();

                if (fType == StringResources.word)
                {
                    using (var fm = new FrmBatchDeleteCustomProps())
                    {
                        fm.ShowDialog();
                        propNameToDelete = fm.PropName;
                    }
                    
                    foreach (string f in files)
                    {
                        using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                        {
                            if (propNameToDelete == "Cancel")
                            {
                                return;
                            }
                            else
                            {
                                if (document.CustomFilePropertiesPart != null)
                                {
                                    bool customPropFound = false;

                                    foreach (CustomDocumentProperty cdp in document.CustomFilePropertiesPart.RootElement)
                                    {
                                        if (propNameToDelete == cdp.Name)
                                        {
                                            cdp.Remove();
                                            lstOutput.Items.Add(f + " : " + propNameToDelete + " deleted");
                                            customPropFound = true;
                                        }
                                    }

                                    if (customPropFound == false)
                                    {
                                        lstOutput.Items.Add(f + " : Property Does Not Exist");
                                    }
                                }
                                else
                                {
                                    lstOutput.Items.Add(f + " : Property Does Not Exist");
                                }
                            }
                        }
                    }
                }
                else if (fType == StringResources.excel)
                {
                    using (var fm = new FrmBatchDeleteCustomProps())
                    {
                        fm.ShowDialog();
                        propNameToDelete = fm.PropName;
                    }

                    foreach (string f in files)
                    {
                        using (SpreadsheetDocument document = SpreadsheetDocument.Open(f, true))
                        {
                            if (propNameToDelete == "Cancel")
                            {
                                return;
                            }
                            else
                            {
                                if (document.CustomFilePropertiesPart != null)
                                {
                                    foreach (CustomDocumentProperty cdp in document.CustomFilePropertiesPart.RootElement)
                                    {
                                        if (propNameToDelete == cdp.Name)
                                        {
                                            cdp.Remove();
                                            lstOutput.Items.Add(f + " : " + propNameToDelete + " deleted");
                                        }
                                        else
                                        {
                                            lstOutput.Items.Add(f + " : Property Does Not Exist");
                                        }
                                    }
                                }
                                else
                                {
                                    lstOutput.Items.Add(f + " : Property Does Not Exist");
                                }
                            }
                        }
                    }
                }
                else if (fType == StringResources.powerpoint)
                {
                    using (var fm = new FrmBatchDeleteCustomProps())
                    {
                        fm.ShowDialog();
                        propNameToDelete = fm.PropName;
                    }

                    foreach (string f in files)
                    {
                        using (PresentationDocument document = PresentationDocument.Open(f, true))
                        {
                            if (propNameToDelete == "Cancel")
                            {
                                return;
                            }
                            else
                            {
                                if (document.CustomFilePropertiesPart != null)
                                {
                                    foreach (CustomDocumentProperty cdp in document.CustomFilePropertiesPart.RootElement)
                                    {
                                        if (propNameToDelete == cdp.Name)
                                        {
                                            cdp.Remove();
                                            lstOutput.Items.Add(f + " : " + propNameToDelete + " deleted");
                                        }
                                        else
                                        {
                                            lstOutput.Items.Add(f + " : Property Does Not Exist");
                                        }
                                    }
                                }
                                else
                                {
                                    lstOutput.Items.Add(f + " : Property Does Not Exist");
                                }
                            }
                        }
                    }
                }
                else
                {
                    return;
                }
            }
            catch (IOException ioe)
            {
                LoggingHelper.Log("BtnListCustomProps Error: " + ioe.Message);
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("BtnListCustomProps Error: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnFixTableProps_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                lstOutput.Items.Clear();
                foreach (string f in files)
                {
                    if (FileUtilities.IsZipArchiveFile(f) == false)
                    {
                        lstOutput.Items.Add(f + " : Not A Valid Office File");
                        return;
                    }

                    using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                    {
                        // "global" document variables
                        bool tblModified = false;
                        OpenXmlElement tgClone = null;

                        // get the list of tables in the document
                        if (WordOpenXml.IsPartNull(document, "Table") == false)
                        {
                            List<Table> tbls = document.MainDocumentPart.Document.Descendants<Table>().ToList();

                            foreach (Table tbl in tbls)
                            {
                                // you can have only one tblGrid per table, including nested tables
                                // it needs to be before any row elements so sequence is
                                // 1. check if the tblGrid element is before any trow
                                // 2. check for multiple tblGrid elements
                                bool tRowFound = false;
                                bool tGridBeforeRowFound = false;
                                int tGridCount = 0;

                                foreach (OpenXmlElement oxe in tbl.Elements())
                                {
                                    // flag if we found a table row, once we find 1, the rest do not matter
                                    if (oxe.GetType().Name == "TableRow")
                                    {
                                        tRowFound = true;
                                    }

                                    // when we get to a tablegrid, we have a few things to check
                                    // 1. have we found a table row
                                    // 2. only one table grid can exist in the table, if there are multiple, delete the extras
                                    if (oxe.GetType().Name == "TableGrid")
                                    {
                                        // increment the tg counter
                                        tGridCount++;

                                        // if we have a table row and no table grid has been found yet, we need to save out this table grid
                                        // then move it in front of the table row later
                                        if (tRowFound == true && tGridCount == 1)
                                        {
                                            tGridBeforeRowFound = true;
                                            tgClone = oxe.CloneNode(true);
                                            oxe.Remove();
                                        }

                                        // if we have multiple table grids, delete the extras
                                        if (tGridCount > 1)
                                        {
                                            oxe.Remove();
                                            tblModified = true;
                                        }
                                    }
                                }

                                // if we had a table grid before a row, move it before the first row
                                if (tGridBeforeRowFound == true)
                                {
                                    tbl.InsertBefore(tgClone, tbl.GetFirstChild<TableRow>());
                                    tblModified = true;
                                }
                            }
                        }

                        if (tblModified == true)
                        {
                            document.MainDocumentPart.Document.Save();
                            lstOutput.Items.Add(f + " : Table Fix Completed");
                        }
                        else
                        {
                            lstOutput.Items.Add(f + " : No Corrupt Table Found");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add(StringResources.errorText + ex.Message);
                LoggingHelper.Log("BtnFixTableProps: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// custom fix for specific SharePoint scenario where the RequestStatus property needs to be removed
        /// the idea here is to just find the <RequestStatus></RequestStatus> xml tag and delete it from the file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnDeleteRequestStatus_Click(object sender, EventArgs e)
        {
            try
            {
                List<CustomXmlPart> cxpList;
                lstOutput.Items.Clear();

                if (fType == StringResources.word)
                {
                    foreach (string f in files)
                    {
                        nodeDeleted = false;

                        if (FileUtilities.IsZipArchiveFile(f) == false)
                        {
                            lstOutput.Items.Add(f + " : Not A Valid Office File");
                            return;
                        }

                        using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                        {
                            cxpList = document.MainDocumentPart.CustomXmlParts.ToList();

                            foreach (CustomXmlPart cxp in cxpList)
                            {
                                XmlDocument xDoc = new XmlDocument();
                                xDoc.Load(cxp.GetStream());

                                XPathNavigator navigator = xDoc.CreateNavigator();

                                // we only check the metadata custom xml file for requeststatus xml
                                if (xDoc.DocumentElement.NamespaceURI == StringResources.schemaMetadataProperties)
                                {
                                    // move to the node and delete it
                                    navigator.MoveToChild("properties", StringResources.schemaMetadataProperties);
                                    navigator.MoveToChild("documentManagement", StringResources.emptyString);
                                    navigator.MoveToChild("RequestStatus", StringResources.requestStatusNS);

                                    // check if we actually moved to the RequestStatus node
                                    // if we didn't move there, no changes should happen, it doesn't exist
                                    if (navigator.Name == "RequestStatus")
                                    {
                                        // delete the node
                                        navigator.DeleteSelf();

                                        // re-write the part
                                        using (MemoryStream xmlMS = new MemoryStream())
                                        {
                                            xDoc.Save(xmlMS);
                                            xmlMS.Position = 0;
                                            cxp.FeedData(xmlMS);
                                        }

                                        // flag the part so we can save the file
                                        nodeDeleted = true;
                                    }
                                }
                            }

                            if (nodeDeleted == true)
                            {
                                document.MainDocumentPart.Document.Save();
                                lstOutput.Items.Add(f + " : Request Status Removed");
                            }
                            else
                            {
                                lstOutput.Items.Add(f + " : Request Status Not Found");
                            }
                        }
                    }
                }
                else if (fType == StringResources.excel)
                {
                    foreach (string f in files)
                    {
                        nodeDeleted = false;
                        using (SpreadsheetDocument document = SpreadsheetDocument.Open(f, true))
                        {
                            cxpList = document.WorkbookPart.CustomXmlParts.ToList();

                            foreach (CustomXmlPart cxp in cxpList)
                            {
                                XmlDocument xDoc = new XmlDocument();
                                xDoc.Load(cxp.GetStream());

                                XPathNavigator navigator = xDoc.CreateNavigator();

                                if (xDoc.DocumentElement.NamespaceURI == StringResources.schemaMetadataProperties)
                                {
                                    navigator.MoveToChild("properties", StringResources.schemaMetadataProperties);
                                    navigator.MoveToChild("documentManagement", StringResources.emptyString);
                                    navigator.MoveToChild("RequestStatus", StringResources.requestStatusNS);
                                    
                                    if (navigator.Name == "RequestStatus")
                                    {
                                        navigator.DeleteSelf();

                                        using (MemoryStream xmlMS = new MemoryStream())
                                        {
                                            xDoc.Save(xmlMS);
                                            xmlMS.Position = 0;
                                            cxp.FeedData(xmlMS);
                                        }

                                        nodeDeleted = true;
                                    }
                                }
                            }

                            if (nodeDeleted == true)
                            {
                                document.WorkbookPart.Workbook.Save();
                                lstOutput.Items.Add(f + " : Request Status Removed");
                            }
                            else
                            {
                                lstOutput.Items.Add(f + " : Request Status Not Found");
                            }
                        }
                    }
                }
                else if (fType == StringResources.powerpoint)
                {
                    foreach (string f in files)
                    {
                        nodeDeleted = false;
                        using (PresentationDocument document = PresentationDocument.Open(f, true))
                        {
                            cxpList = document.PresentationPart.CustomXmlParts.ToList();
                            
                            foreach (CustomXmlPart cxp in cxpList)
                            {
                                XmlDocument xDoc = new XmlDocument();
                                xDoc.Load(cxp.GetStream());
                                
                                XPathNavigator navigator = xDoc.CreateNavigator();

                                if (xDoc.DocumentElement.NamespaceURI == StringResources.schemaMetadataProperties)
                                {
                                    navigator.MoveToChild("properties", StringResources.schemaMetadataProperties);
                                    navigator.MoveToChild("documentManagement", StringResources.emptyString);
                                    navigator.MoveToChild("RequestStatus", StringResources.requestStatusNS);

                                    if (navigator.Name == "RequestStatus")
                                    {
                                        navigator.DeleteSelf();

                                        using (MemoryStream xmlMS = new MemoryStream())
                                        {
                                            xDoc.Save(xmlMS);
                                            xmlMS.Position = 0;
                                            cxp.FeedData(xmlMS);
                                        }

                                        nodeDeleted = true;
                                    }
                                }
                            }

                            if (nodeDeleted == true)
                            {
                                document.PresentationPart.Presentation.Save();
                                lstOutput.Items.Add(f + " : Request Status Removed");
                            }
                            else
                            {
                                lstOutput.Items.Add(f + " : Request Status Not Found");
                            }
                        }
                    }
                }
                else
                {
                    return;
                }
            }
            catch (IOException ioe)
            {
                LoggingHelper.Log("BtnDeleteRequestStatus Error: " + ioe.Message);
                lstOutput.Items.Add("Error - " + ioe.Message);
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("BtnDeleteRequestStatus Error: " + ex.Message);
                lstOutput.Items.Add("Error - " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void CkSearchSubfolders_CheckedChanged(object sender, EventArgs e)
        {
            PopulateAndDisplayFiles();
        }
    }
}