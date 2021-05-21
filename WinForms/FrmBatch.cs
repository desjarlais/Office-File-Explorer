// Open Xml SDK refs
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

// app refs
using Office_File_Explorer.App_Helpers;
using Office_File_Explorer.WinForms;
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
        public string fileType = string.Empty;
        public string fType = string.Empty;
        public bool nodeDeleted = false;
        public bool nodeChanged = false;
        public string fromChangeTemplate;

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
                fType = StringResources.wWord;
            }
            else if (rdoExcel.Checked == true)
            {
                fileType = "*.xlsx";
                fType = StringResources.wExcel;
            }
            else if (rdoPowerPoint.Checked == true)
            {
                fileType = "*.pptx";
                fType = StringResources.wPowerpoint;
            }

            return fileType;
        }

        public string DefaultTemplate
        {
            set => fromChangeTemplate = value;
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
            BtnDeleteOpenByDefault.Enabled = false;
            BtnChangeAttachedTemplate.Enabled = false;
            BtnFixExcelHyperlinks.Enabled = false;

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
                BtnDeleteOpenByDefault.Enabled = true;
                BtnChangeAttachedTemplate.Enabled = true;

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
                BtnDeleteOpenByDefault.Enabled = false;
                BtnChangeAttachedTemplate.Enabled = false;
                BtnFixExcelHyperlinks.Enabled = false;
            }

            if (rdoExcel.Checked == true)
            {
                BtnConvertStrict.Enabled = true;
                BtnFixExcelHyperlinks.Enabled = true;

                BtnFixNotesPageSize.Enabled = false;
                BtnFixCorruptBookmarks.Enabled = false;
                BtnFixCorruptRevisions.Enabled = false;
                BtnFixTableProps.Enabled = false;
                BtnRemovePII.Enabled = false;
                BtnPPTResetPII.Enabled = false;
                BtnDeleteOpenByDefault.Enabled = false;
                BtnChangeAttachedTemplate.Enabled = false;
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
                        lstOutput.Items.Add(f + StringResources.wArrow + StringResources.pptNotesSizeReset);
                        LoggingHelper.Log(f + StringResources.wArrow + StringResources.pptNotesSizeReset);
                    }
                }
                catch (NullReferenceException nre)
                {
                    lstOutput.Items.Add(f + StringResources.wArrow + "** Document does not contain Notes Master **");
                    LoggingHelper.Log(f + StringResources.wArrow + StringResources.wErrorText + nre.Message);
                }
                catch (Exception ex)
                {
                    lstOutput.Items.Add(f + StringResources.wArrow + StringResources.wErrorText + ex.Message);
                    LoggingHelper.Log(f + StringResources.wArrow + StringResources.wErrorText + ex.Message);
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
                lstOutput.Items.Add(StringResources.wErrorText + ex.Message);
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
                                    if (oxedr.GetType().ToString() == StringResources.dfowRun)
                                    {
                                        O.Wordprocessing.Run r = (O.Wordprocessing.Run)oxedr;
                                        foreach (OpenXmlElement oxe in oxedr.ChildElements)
                                        {
                                            // you can't have a Text tag inside a DeletedRun
                                            if (oxe.GetType().ToString() == StringResources.dfowText)
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
                lstOutput.Items.Add(StringResources.wErrorText + ex.Message);
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
                lstOutput.Items.Add(StringResources.wErrorText + ex.Message);
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
                        excelcnvPath = string.Empty;
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

                    if (isStrict == true && excelcnvPath != string.Empty)
                    {
                        // setup destination file path
                        string strOriginalFile = f;
                        string strOutputPath = Path.GetDirectoryName(strOriginalFile) + "\\";
                        string strFileExtension = Path.GetExtension(strOriginalFile);
                        string strOutputFileName = strOutputPath + Path.GetFileNameWithoutExtension(strOriginalFile) + StringResources.wFixedFileParentheses + strFileExtension;

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
                lstOutput.Items.Add(StringResources.wErrorText + ex.Message);
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
                string propNameToDelete = string.Empty;
                lstOutput.Items.Clear();

                if (fType == StringResources.wWord)
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
                                            lstOutput.Items.Add(f + StringResources.wColonBuffer + propNameToDelete + " deleted");
                                            customPropFound = true;
                                        }
                                    }

                                    if (customPropFound == false)
                                    {
                                        lstOutput.Items.Add(f + StringResources.noProp);
                                    }
                                }
                                else
                                {
                                    lstOutput.Items.Add(f + StringResources.noProp);
                                }
                            }
                        }
                    }
                }
                else if (fType == StringResources.wExcel)
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
                                            lstOutput.Items.Add(f + StringResources.wColonBuffer + propNameToDelete + " deleted");
                                        }
                                        else
                                        {
                                            lstOutput.Items.Add(f + StringResources.noProp);
                                        }
                                    }
                                }
                                else
                                {
                                    lstOutput.Items.Add(f + StringResources.noProp);
                                }
                            }
                        }
                    }
                }
                else if (fType == StringResources.wPowerpoint)
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
                                            lstOutput.Items.Add(f + StringResources.wColonBuffer + propNameToDelete + " deleted");
                                        }
                                        else
                                        {
                                            lstOutput.Items.Add(f + StringResources.noProp);
                                        }
                                    }
                                }
                                else
                                {
                                    lstOutput.Items.Add(f + StringResources.noProp);
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
                            List<O.Wordprocessing.Table> tbls = document.MainDocumentPart.Document.Descendants<O.Wordprocessing.Table>().ToList();

                            foreach (O.Wordprocessing.Table tbl in tbls)
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
                lstOutput.Items.Add(StringResources.wErrorText + ex.Message);
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

                if (fType == StringResources.wWord)
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
                                    navigator.MoveToChild("documentManagement", string.Empty);
                                    navigator.MoveToChild(StringResources.wCustomXmlRequestStatus, StringResources.wRequestStatusNS);

                                    // check if we actually moved to the RequestStatus node
                                    // if we didn't move there, no changes should happen, it doesn't exist
                                    if (navigator.Name == StringResources.wCustomXmlRequestStatus)
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
                else if (fType == StringResources.wExcel)
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
                                    navigator.MoveToChild("documentManagement", string.Empty);
                                    navigator.MoveToChild(StringResources.wCustomXmlRequestStatus, StringResources.wRequestStatusNS);
                                    
                                    if (navigator.Name == StringResources.wCustomXmlRequestStatus)
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
                else if (fType == StringResources.wPowerpoint)
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
                                    navigator.MoveToChild("documentManagement", string.Empty);
                                    navigator.MoveToChild(StringResources.wCustomXmlRequestStatus, StringResources.wRequestStatusNS);

                                    if (navigator.Name == StringResources.wCustomXmlRequestStatus)
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

        /// <summary>
        /// If the openByDefault custom xml part is true
        /// document will open in Word and display the Info Pane instead of the document
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnDeleteOpenByDefault_Click(object sender, EventArgs e)
        {
            try
            {
                List<CustomXmlPart> cxpList;
                lstOutput.Items.Clear();

                foreach (string f in files)
                {
                    nodeChanged = false;

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

                            if (xDoc.DocumentElement.NamespaceURI == StringResources.schemaCustomXsn)
                            {
                                foreach (XmlNode xNode in xDoc.ChildNodes)
                                {
                                    if (xNode.Name == "customXsn")
                                    {
                                        foreach (XmlNode x in xNode)
                                        {
                                            if (x.Name == "openByDefault")
                                            {
                                                x.FirstChild.Value = "False";
                                                using (MemoryStream xmlMS = new MemoryStream())
                                                {
                                                    xDoc.Save(xmlMS);
                                                    xmlMS.Position = 0;
                                                    cxp.FeedData(xmlMS);
                                                }
                                                nodeChanged = true;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (nodeChanged == true)
                        {
                            document.MainDocumentPart.Document.Save();
                            lstOutput.Items.Add(f + " : openByDefault Changed");
                        }
                        else
                        {
                            lstOutput.Items.Add(f + " : openByDefault Not Found");
                        }
                    }
                }
            }
            catch (IOException ioe)
            {
                LoggingHelper.Log("BtnDeleteOpenByDefault Error: " + ioe.Message);
                lstOutput.Items.Add("Error - " + ioe.Message);
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("BtnDeleteOpenByDefault Error: " + ex.Message);
                lstOutput.Items.Add("Error - " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnFixExcelHyperlinks_Click(object sender, EventArgs e)
        {
            try
            {
                lstOutput.Items.Clear();
                Cursor = Cursors.WaitCursor;

                foreach (string f in files)
                {
                    bool isFileChanged = false;

                    using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(f, true))
                    {
                    // adding a goto since changing the relationship during enumeration causes an error
                    // after making the change, I restart the loops again to look for more corrupt links
                    HLinkStart:
                        foreach (WorksheetPart wsp in excelDoc.WorkbookPart.WorksheetParts)
                        {
                            IEnumerable<O.Spreadsheet.Hyperlink> hLinks = wsp.Worksheet.Descendants<O.Spreadsheet.Hyperlink>();
                            // loop each hyperlink to get the rid
                            foreach (O.Spreadsheet.Hyperlink h in hLinks)
                            {
                                // then check for hyperlinks relationships for the rid
                                if (wsp.HyperlinkRelationships.Count() > 0)
                                {
                                    foreach (HyperlinkRelationship hRel in wsp.HyperlinkRelationships)
                                    {
                                        // if the rid's match, we have the same hyperlink
                                        if (h.Id == hRel.Id)
                                        {
                                            // there is a scenario where files from OpenText appear to be damaged and the url is some temp file path
                                            // not the url path it should be
                                            string badUrl = string.Empty;
                                            string[] separatingStrings = { "livelink" };

                                            // check if the uri contains any of the known bad paths
                                            if (hRel.Uri.ToString().StartsWith("../../../"))
                                            {
                                                badUrl = hRel.Uri.ToString().Replace("../../../", StringResources.wBackslash);
                                            }
                                            else if (hRel.Uri.ToString().Contains("/AppData/Local/Microsoft/Windows/livelink/llsapi.dll/open/"))
                                            {
                                                string[] urlParts = hRel.Uri.ToString().Split(separatingStrings, StringSplitOptions.RemoveEmptyEntries);
                                                badUrl = hRel.Uri.ToString().Replace(urlParts[0], StringResources.wBackslash);
                                            }
                                            else if (hRel.Uri.ToString().Contains("/AppData/Roaming/OpenText/"))
                                            {
                                                string[] urlParts = hRel.Uri.ToString().Split(separatingStrings, StringSplitOptions.RemoveEmptyEntries);
                                                badUrl = hRel.Uri.ToString().Replace(urlParts[0], StringResources.wBackslash);
                                            }

                                            // if a bad path was found, start the work to replace it with the correct path
                                            if (badUrl != string.Empty)
                                            {
                                                // loop the sharedstrings to get the correct replace value
                                                if (excelDoc.WorkbookPart.SharedStringTablePart != null)
                                                {
                                                    SharedStringTable sst = excelDoc.WorkbookPart.SharedStringTablePart.SharedStringTable;
                                                    foreach (SharedStringItem ssi in sst)
                                                    {
                                                        if (ssi.Text != null)
                                                        {
                                                            if (ssi.InnerText.ToString().EndsWith(badUrl))
                                                            {
                                                                // now delete the relationship
                                                                wsp.DeleteReferenceRelationship(h.Id);

                                                                // now add a new relationship with the right address
                                                                wsp.AddHyperlinkRelationship(new Uri(ssi.InnerText, UriKind.Absolute), true, h.Id);
                                                                isFileChanged = true;
                                                                goto HLinkStart;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (isFileChanged == true)
                        {
                            excelDoc.WorkbookPart.Workbook.Save();
                            lstOutput.Items.Add(f + "** Hyperlinks Fixed **");
                        }
                        else
                        {
                            lstOutput.Items.Add(f + "** No Corrupt Hyperlinks Found **");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("BtnFixExcelHyperlinks Error: " + ex.Message);
                lstOutput.Items.Add("Error - " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnChangeAttachedTemplate_Click(object sender, EventArgs e)
        {
            try
            {
                lstOutput.Items.Clear();
                Cursor = Cursors.WaitCursor;

                // get the new template path from the user
                FrmChangeTemplate ctFrm = new FrmChangeTemplate()
                {
                    Owner = this
                };
                ctFrm.ShowDialog();

                foreach (string f in files)
                {
                    bool isFileChanged = false;
                    string attachedTemplateId = "";
                    string filePath = "";

                    using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                    {
                        DocumentSettingsPart dsp = document.MainDocumentPart.DocumentSettingsPart;

                        // if the external rel exists, we need to pull the rid and old uri
                        // we will be deleting this part and re-adding with the new uri
                        if (dsp.ExternalRelationships.Count() > 0)
                        {
                            // just change the attached template
                            foreach (ExternalRelationship er in dsp.ExternalRelationships)
                            {
                                if (er.RelationshipType != null && er.RelationshipType == StringResources.DocumentTemplatePartType)
                                {
                                    // keep track of the existing rId for the template
                                    attachedTemplateId = er.Id;
                                    filePath = er.Uri.ToString();
                                    break;
                                }
                            }
                        }
                        else
                        {
                            // if the part does not exist, this is a Normal.dotm situation
                            // path out to where it should be based on default install settings
                            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                            filePath = userProfile + "\\AppData\\Roaming\\Microsoft\\Templates\\Normal.dotm";

                            if (!File.Exists(filePath))
                            {
                                // Normal.dotm path is not correct?
                                LoggingHelper.Log("BtnChangeDefaultTemplate Error: " + "Invalid Attached Template Path");
                                throw new Exception();
                            }
                        }

                        if (fromChangeTemplate == filePath || fromChangeTemplate == null)
                        {
                            // file path is the same or user closed without wanting changes, do nothing
                            return;
                        }
                        else
                        {
                            filePath = fromChangeTemplate;
                            isFileChanged = true;

                            Uri newFilePath = new Uri(filePath);

                            // delete the old part
                            dsp.DeleteExternalRelationship(attachedTemplateId);

                            // add the new part back in
                            dsp.AddExternalRelationship(StringResources.DocumentTemplatePartType, newFilePath, attachedTemplateId);
                        }

                        if (isFileChanged)
                        {
                            lstOutput.Items.Add(f + "** Attached Template Changed **");
                            document.MainDocumentPart.Document.Save();
                        }
                        else
                        {
                            lstOutput.Items.Add(f + "** No Changed Made To Attached Template **");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("BtnChangeAttachedTemplate Error: " + ex.Message);
                lstOutput.Items.Add("Error - " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }
    }
}