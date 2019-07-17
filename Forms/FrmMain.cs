/****************************** Module Header ******************************\
Module Name:  FrmMain.cs
Project:      Office File Explorer
Copyright (c) Microsoft Corporation.

Main window for OFE.

This source is subject to the Microsoft Public License.
See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Deployment.Application;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using Office_File_Explorer.PowerPoint_Helpers;
using Column = DocumentFormat.OpenXml.Spreadsheet.Column;
using System.Collections.Generic;

namespace Office_File_Explorer
{
    public partial class FrmMain : Form
    {
        // globals
        string _fromAuthor;
        string _FindText;
        string _ReplaceText;
        string fileType;

        // static string variables
        const string _fileDoesNotExist = "** File does not exist **";
        const string _noFootNotes = "** No Footnotes in this document **";
        const string _noEndNotes = "** No Endnotes in this document **";
        const string _themeFileAdded = "Theme File Added.";
        const string _unableToDownloadUpdate = "Unable to download update.";
        const string _word = "Word";
        const string _excel = "Excel";
        const string _powerpoint = "PowerPoint";

        // global numid lists
        ArrayList oNumIdList = new ArrayList();
        ArrayList aNumIdList = new ArrayList();
        ArrayList numIdList = new ArrayList();

        public enum InformationOutput { ClearAndAdd, Append, TextOnly, InvalidFile };

        public FrmMain()
        {
            InitializeComponent();
            Log("App Start");
            _FindText = "";
            _ReplaceText = "";

            // disable all buttons
            DisableButtons();
        }

        #region Class Properties
        public string AuthorProperty
        {
            set => _fromAuthor = value;
        }

        public string FindTextProperty
        {
            set => _FindText = value;
        }

        public string ReplaceTextProperty
        {
            set => _ReplaceText = value;
        }
        #endregion

        public void DisableButtons()
        {
            fileType = "";
            BtnAcceptRevisions.Enabled = false;
            BtnDeleteBreaks.Enabled = false;
            BtnDeleteComments.Enabled = false;
            BtnDeleteEndnotes.Enabled = false;
            BtnDeleteFootnotes.Enabled = false;
            BtnDeleteHdrFtr.Enabled = false;
            BtnDeleteHdrFtr.Enabled = false;
            BtnDeleteHiddenText.Enabled = false;
            BtnDeleteListTemplates.Enabled = false;
            BtnDeleteExternalLinks.Enabled = false;
            BtnListAuthors.Enabled = false;
            BtnListDefinedNames.Enabled = false;
            BtnListComments.Enabled = false;
            BtnListEndnotes.Enabled = false;
            BtnListFonts.Enabled = false;
            BtnListWorksheets.Enabled = false;
            BtnListFootnotes.Enabled = false;
            BtnListFormulas.Enabled = false;
            BtnListHiddenWorksheets.Enabled = false;
            BtnListHiddenRowsColumns.Enabled = false;
            BtnListSharedStrings.Enabled = false;
            BtnListHyperlinks.Enabled = false;
            BtnListLinks.Enabled = false;
            BtnListOle.Enabled = false;
            BtnListRevisions.Enabled = false;
            BtnListStyles.Enabled = false;
            BtnListTemplates.Enabled = false;
            BtnPPTGetAllSlideTitles.Enabled = false;
            BtnPPTListHyperlinks.Enabled = false;
            BtnRemovePII.Enabled = false;
            BtnSearchAndReplace.Enabled = false;
            BtnValidateFile.Enabled = false;
            BtnViewCustomDocProps.Enabled = false;
            BtnComments.Enabled = false;
            BtnDeleteComment.Enabled = false;
            BtnChangeTheme.Enabled = false;
            BtnViewPPTComments.Enabled = false;
        }

        public enum OxmlFileFormat { Xlsx, Xlsm, Docx, Docm, Pptx, Pptm, Invalid };

        public OxmlFileFormat GetFileFormat()
        {
            string fileExt = Path.GetExtension(TxtFileName.Text);
            fileExt = fileExt.ToLower();
            
            if (fileExt == ".docx")
            {
                return OxmlFileFormat.Docx;
            }
            else if (fileExt == ".docm")
            {
                return OxmlFileFormat.Docm;
            }
            else if (fileExt == ".xlsx")
            {
                return OxmlFileFormat.Xlsx;
            }
            else if (fileExt == ".xlsm")
            {
                return OxmlFileFormat.Xlsm;
            }
            else if (fileExt == ".pptx")
            {
                return OxmlFileFormat.Pptx;
            }
            else if (fileExt == ".pptm")
            {
                return OxmlFileFormat.Pptm;
            }
            else
            {
                return OxmlFileFormat.Invalid;
            }
        }

        public void SetUpButtons()
        {
            // disable all buttons first
            DisableButtons();

            if (GetFileFormat() == OxmlFileFormat.Docx || GetFileFormat() == OxmlFileFormat.Docm)
            {
                fileType = _word;

                // WD only files
                BtnAcceptRevisions.Enabled = true;
                BtnDeleteBreaks.Enabled = true;
                BtnDeleteComments.Enabled = true;
                BtnDeleteEndnotes.Enabled = true;
                BtnDeleteFootnotes.Enabled = true;
                BtnDeleteHdrFtr.Enabled = true;
                BtnDeleteHdrFtr.Enabled = true;
                BtnDeleteHiddenText.Enabled = true;
                BtnDeleteListTemplates.Enabled = true;
                BtnListAuthors.Enabled = true;
                BtnListComments.Enabled = true;
                BtnListEndnotes.Enabled = true;
                BtnListFonts.Enabled = true;
                BtnListFootnotes.Enabled = true;
                BtnListHyperlinks.Enabled = true;
                BtnListRevisions.Enabled = true;
                BtnListStyles.Enabled = true;
                BtnListTemplates.Enabled = true;
                BtnSearchAndReplace.Enabled = true;
                BtnListOle.Enabled = true;
                BtnViewCustomDocProps.Enabled = true;
            }
            else if (GetFileFormat() == OxmlFileFormat.Xlsx || GetFileFormat() == OxmlFileFormat.Xlsm)
            {
                fileType = _excel;

                // enable XL only files
                BtnListDefinedNames.Enabled = true;
                BtnListHiddenRowsColumns.Enabled = true;
                BtnDeleteExternalLinks.Enabled = true;
                BtnListLinks.Enabled = true;
                BtnListFormulas.Enabled = true;
                BtnListWorksheets.Enabled = true;
                BtnListHiddenWorksheets.Enabled = true;
                BtnListSharedStrings.Enabled = true;
                BtnComments.Enabled = true;
                BtnDeleteComment.Enabled = true;
            }
            else if (GetFileFormat() == OxmlFileFormat.Pptx || GetFileFormat() == OxmlFileFormat.Pptm)
            {
                fileType = _powerpoint;

                // enable PPT only files
                BtnPPTGetAllSlideTitles.Enabled = true;
                BtnPPTListHyperlinks.Enabled = true;
                BtnViewPPTComments.Enabled = true;
            }
            else if (GetFileFormat() == OxmlFileFormat.Invalid)
            {
                // invalid file format
                MessageBox.Show("Invalid File Format");
            }
            else
            {
                // unknown condition, log details
                Log("GetFileFormat Error: " + TxtFileName.Text);
            }

            // these buttons exists for all file types
            BtnRemovePII.Enabled = true;
            BtnValidateFile.Enabled = true;
            BtnChangeTheme.Enabled = true;
        }

        private void BtnListComments_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                LstDisplay.Items.Clear();

                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    WordprocessingCommentsPart commentsPart = myDoc.MainDocumentPart.WordprocessingCommentsPart;
                    int count = 0;
                    foreach (DocumentFormat.OpenXml.Wordprocessing.Comment cm in commentsPart.Comments)
                    {
                        count++;
                        LstDisplay.Items.Add(count + ". " + cm.InnerText);
                    }
                }
            }
            catch (NullReferenceException)
            {
                DisplayInformation(InformationOutput.ClearAndAdd, "** There are no comments to display **");
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                Log("Word - BtnListComments_Click Error");
                Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// Output text to the listbox
        /// </summary>
        /// <param name="typeOfOutput">This variable specifies the type of output to display</param>
        /// <param name="output">This is the actual data from the document we want to display</param>
        public void DisplayInformation(InformationOutput display, string output)
        {
            switch (display)
            {
                case InformationOutput.ClearAndAdd:
                    LstDisplay.Items.Clear();
                    LstDisplay.Items.Add(output);
                    break;
                case InformationOutput.Append:
                    LstDisplay.Items.Add("");
                    LstDisplay.Items.Add(output);
                    break;
                case InformationOutput.InvalidFile:
                    LstDisplay.Items.Clear();
                    LstDisplay.Items.Add("Invalid File. Please select a valid document.");
                    break;
                default:
                    LstDisplay.Items.Add(output);
                    break;
            }
        }

        private void BtnListStyles_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                LstDisplay.Items.Clear();
                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    MainDocumentPart mainPart = myDoc.MainDocumentPart;
                    StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart;
                    bool containStyle = false;

                    LstDisplay.Items.Clear();
                    try
                    {
                        foreach (OpenXmlElement el in stylePart.Styles.LatentStyles.Elements())
                        {
                            string styleEl = el.GetAttribute("name", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value;
                            int pStyle = Word_Helpers.WordExtensionClass.ParagraphsByStyleName(mainPart, styleEl).Count();
                            int rStyle = Word_Helpers.WordExtensionClass.RunsByStyleName(mainPart, styleEl).Count();
                            int tStyle = Word_Helpers.WordExtensionClass.TablesByStyleName(mainPart, styleEl).Count();

                            if (pStyle > 0)
                            {
                                LstDisplay.Items.Add("Number of paragraphs with " + styleEl + " styles: " + pStyle);
                                containStyle = true;
                            }

                            if (rStyle > 0)
                            {
                                LstDisplay.Items.Add("Number of runs with " + styleEl + " styles: " + rStyle);
                                containStyle = true;
                            }

                            if (tStyle > 0)
                            {
                                LstDisplay.Items.Add("Number of tables with " + styleEl + " styles: " + tStyle);
                                containStyle = true;
                            }
                        }

                        if (containStyle == false)
                        {
                            LstDisplay.Items.Add("** No styles in this document **");
                        }
                    }
                    catch (NullReferenceException)
                    {
                        DisplayInformation(InformationOutput.ClearAndAdd, "Missing StylesWithEffects part.");
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                Log("BtnListStyles_Click Error");
                Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnListHyperlinks_Click(object sender, EventArgs e)
        {
            try
            {
                LstDisplay.Items.Clear();
                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    int hlinkCount = myDoc.MainDocumentPart.HyperlinkRelationships.Count();
                    if (hlinkCount == 0)
                    {
                        LstDisplay.Items.Add("** There are no hyperlinks in this document **");
                    }
                    else
                    {
                        int count = 0;
                        foreach (HyperlinkRelationship hRel in myDoc.MainDocumentPart.HyperlinkRelationships)
                        {
                            count++;
                            LstDisplay.Items.Add(count + ". " + hRel.Uri);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                Log("BtnListHyperlinks_Click Error");
                Log(ex.Message);
            }
        }

        private void BtnListTemplates_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            LstDisplay.Items.Clear();
            numIdList.Clear();
            aNumIdList.Clear();
            oNumIdList.Clear();

            try
            {
                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    MainDocumentPart mainPart = myDoc.MainDocumentPart;
                    NumberingDefinitionsPart numPart = mainPart.NumberingDefinitionsPart;
                    StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart;

                    // Loop each paragraph, get the NumberingId and add it to the array
                    foreach (OpenXmlElement el in mainPart.Document.Descendants<Paragraph>())
                    {
                        if (el.Descendants<NumberingId>().Count() > 0)
                        {
                            foreach (NumberingId pNumId in el.Descendants<NumberingId>())
                            {
                                numIdList.Add(pNumId.Val);
                            }
                        }
                        else
                        {
                            DisplayInformation(InformationOutput.TextOnly, "** There are no List Templates in this document **");
                            return;
                        }
                    }

                    // Loop each header, get the NumId and add it to the array
                    foreach (HeaderPart hdrPart in mainPart.HeaderParts)
                    {
                        foreach (OpenXmlElement el in hdrPart.Header.Elements())
                        {
                            foreach (NumberingId hNumId in el.Descendants<NumberingId>())
                            {
                                numIdList.Add(hNumId.Val);
                            }
                        }
                    }

                    // Loop each footer, get the NumId and add it to the array
                    foreach (FooterPart ftrPart in mainPart.FooterParts)
                    {
                        foreach (OpenXmlElement el in ftrPart.Footer.Elements())
                        {
                            foreach (NumberingId fNumdId in el.Descendants<NumberingId>())
                            {
                                numIdList.Add(fNumdId.Val);
                            }
                        }
                    }

                    // Loop through each style in document and get NumId
                    foreach (OpenXmlElement el in stylePart.Styles.Elements())
                    {
                        try
                        {
                            string styleEl = el.GetAttribute("styleId", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value;
                            int pStyle = Word_Helpers.WordExtensionClass.ParagraphsByStyleName(mainPart, styleEl).Count();

                            if (pStyle > 0)
                            {
                                foreach (NumberingId sEl in el.Descendants<NumberingId>())
                                {
                                    numIdList.Add(sEl.Val);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            // Not all style elements have a styleID, so just skip these scenarios
                            Log("BtnListTemplates_Click : " + ex.Message);
                        }
                    }

                    // try-catch for the scenario where the list is already sorted
                    try
                    {
                        numIdList.Sort();
                        numIdList = RemoveDuplicate(numIdList);
                        LoopArrayList(numIdList);
                    }
                    catch (InvalidOperationException)
                    {
                        // continue on if the list is already sorted
                        numIdList = RemoveDuplicate(numIdList);
                        LoopArrayList(numIdList);
                    }

                    // Loop through each AbstractNumId
                    LstDisplay.Items.Add("");
                    LstDisplay.Items.Add("List Templates in document:");
                    foreach (OpenXmlElement el in numPart.Numbering.Elements())
                    {
                        foreach (AbstractNumId aNumId in el.Descendants<AbstractNumId>())
                        {
                            string strNumId = el.GetAttribute("numId", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value;
                            aNumIdList.Add(strNumId);
                            LstDisplay.Items.Add("numId = " + strNumId + " " + "AbstractNumId = " + aNumId.Val);
                        }
                    }

                    // get the unused list templates
                    oNumIdList = OrphanedListTemplates(numIdList, aNumIdList);
                    LstDisplay.Items.Add("");
                    LstDisplay.Items.Add("Orphaned List Templates:");
                    foreach (object item in oNumIdList)
                    {
                        LstDisplay.Items.Add("numId = " + item);
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                Log("BtnListTemplates_Click Error");
                Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        // method to display the non-duplicate numId used in the document.
        private void LoopArrayList(ArrayList al)
        {
            LstDisplay.Items.Add("numId used in document:");
            foreach (object item in al)
            {
                LstDisplay.Items.Add("numID = " + item);
            }
        }

        // method to return the non-duplicate items in the source arraylist
        private static ArrayList RemoveDuplicate(ArrayList sourceList)
        {
            ArrayList list = new ArrayList();
            foreach (object item in sourceList)
            {
                string val = item.ToString();
                if (!list.Contains(val))
                {
                    list.Add(val);
                }
            }
            return list;
        }

        // loop through both array lists to find which numId is not currently used in the document.
        private static ArrayList OrphanedListTemplates(ArrayList usedNumIdList, ArrayList docNumIdList)
        {
            var copyOfDocNumId = new ArrayList(docNumIdList);

            foreach (var p in usedNumIdList)
            {
                copyOfDocNumId.Remove(p);
            }

            return copyOfDocNumId;
        }

        private void BtnListOle_Click(object sender, EventArgs e)
        {
            LstDisplay.Items.Clear();
            try
            {
                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    int oleObjCount = GetEmbeddedObjectProperties(myDoc.MainDocumentPart);
                    int olePkgPart = GetEmbeddedPackageProperties(myDoc.MainDocumentPart);

                    if (olePkgPart == 0 && oleObjCount == 0)
                    {
                        DisplayInformation(InformationOutput.ClearAndAdd, "** This document does not contain OLE Package objects **");
                    }               
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                Log("BtnListOle_Click Error");
                Log(ex.Message);
            }
        }

        public int GetEmbeddedPackageProperties(MainDocumentPart mPart)
        {
            try
            {
                int x = 0;
                int olePkgCount = mPart.EmbeddedPackageParts.Count();

                String origUri, trimUri;

                do
                {
                    origUri = mPart.EmbeddedPackageParts.ElementAt(x).Uri.ToString();
                    trimUri = origUri.Remove(0, 17);
                    x++;
                    LstDisplay.Items.Add(x + ". " + trimUri);
                }
                while (x < olePkgCount);

                return x;
            }
            catch (ArgumentOutOfRangeException)
            {
                return 0;
            }
        }

        public int GetEmbeddedObjectProperties(MainDocumentPart mPart)
        {
            try
            {
                int x = 0;
                int oleEmbCount = mPart.EmbeddedObjectParts.Count();
                String origUri, trimUri;

                do
                {
                    origUri = mPart.EmbeddedObjectParts.ElementAt(x).Uri.ToString();
                    trimUri = origUri.Remove(0, 17);
                    x++;
                    LstDisplay.Items.Add(x + ". " + trimUri);
                }
                while (x < oleEmbCount);

                return x;
            }
            catch (ArgumentOutOfRangeException)
            {
                return 0;
            }
        }

        private void BtnAcceptRevisions_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                WordprocessingDocument document;

                using (document = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    // get the list of authors
                    _fromAuthor = "";

                    Forms.FrmAuthors aFrm = new Forms.FrmAuthors(TxtFileName.Text, document)
                    {
                        Owner = this
                    };
                    aFrm.ShowDialog();
                }

                if (_fromAuthor == "All Authors")
                {
                    _fromAuthor = "";
                }

                Word_Helpers.WordOpenXml.AcceptAllRevisions(TxtFileName.Text, _fromAuthor);
                DisplayInformation(InformationOutput.ClearAndAdd, "** Revisions Accepted **");
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                Log("BtnAcceptRevisions_Click Error");
                Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnDeleteComments_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                Word_Helpers.WordOpenXml.RemoveComments(TxtFileName.Text);
                DisplayInformation(InformationOutput.ClearAndAdd, "** Comments Removed **");
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                Log("BtnDeleteComments_Click Error");
                Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnDeleteHiddenText_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                Word_Helpers.WordOpenXml.DeleteHiddenText(TxtFileName.Text);
                DisplayInformation(InformationOutput.TextOnly, "** Hidden text deleted **");
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                Log("BtnDeleteHiddenText_Click Error");
                Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnDeleteHdrFtr_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                Word_Helpers.WordOpenXml.RemoveHeadersFooters(TxtFileName.Text);
                DisplayInformation(InformationOutput.TextOnly, "** Headers/Footer removed **");
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                Log("BtnDeleteHdrFtr_Click Error");
                Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnDeleteListTemplates_Click(object sender, EventArgs e)
        {
            try
            {
                BtnListTemplates.PerformClick();
                foreach (object orphanLT in oNumIdList)
                {
                    Word_Helpers.WordOpenXml.RemoveListTemplatesNumId(TxtFileName.Text, orphanLT.ToString());
                }
                DisplayInformation(InformationOutput.TextOnly, "** List Templates removed **");
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
            }
        }

        private void BtnDeleteBreaks_Click(object sender, EventArgs e)
        {
            Word_Helpers.WordOpenXml.RemoveBreaks(TxtFileName.Text);
            DisplayInformation(InformationOutput.ClearAndAdd, "** Page and Section breaks have been removed **");
        }

        private void BtnRemovePII_Click(object sender, EventArgs e)
        {
            if (!File.Exists(TxtFileName.Text))
            {
                DisplayInformation(InformationOutput.InvalidFile, TxtFileName.Text);
                return;
            }

            using (WordprocessingDocument doc = WordprocessingDocument.Open(TxtFileName.Text, true))
            {
                Word_Helpers.WordExtensionClass.RemovePersonalInfo(doc);
                DisplayInformation(InformationOutput.ClearAndAdd, "PII Removed from file.");
            }
        }

        public void DisplayValidationErrorInformation(OpenXmlPackage docPackage)
        {
            OpenXmlValidator validator = new OpenXmlValidator();
            int count = 0;
            LstDisplay.Items.Clear();

            foreach (ValidationErrorInfo error in validator.Validate(docPackage))
            {
                count++;
                LstDisplay.Items.Add("Error " + count);
                LstDisplay.Items.Add("Description: " + error.Description);
                LstDisplay.Items.Add("Path: " + error.Path.XPath);
                LstDisplay.Items.Add("Part: " + error.Part.Uri);
                LstDisplay.Items.Add("-------------------------------------------");
            }

            if (count == 0)
            {
                LstDisplay.Items.Add("** No errors found **");
            }
        }

        private void BtnValidateFile_Click(object sender, EventArgs e)
        {
            try
            {
                if (GetFileFormat() == OxmlFileFormat.Docx)
                {
                    using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, true))
                    {
                        DisplayValidationErrorInformation(myDoc);
                    }
                }
                else if (GetFileFormat() == OxmlFileFormat.Xlsx)
                {
                    using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(TxtFileName.Text, true))
                    {
                        DisplayValidationErrorInformation(myDoc);
                    }
                }
                else if (GetFileFormat() == OxmlFileFormat.Pptx)
                {
                    using (PresentationDocument myDoc = PresentationDocument.Open(TxtFileName.Text, true))
                    {
                        DisplayValidationErrorInformation(myDoc);
                    }
                }
                else
                {
                    Log("BtnValidateFileClick Error");
                    throw new Exception();
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                Log("BtnValidateFile_Click Error");
                Log(ex.Message);
            }

            if (LstDisplay.Items.Count < 0)
            {
                LstDisplay.Items.Add("** No validation errors **");
            }
        }

        private void BtnCopyOutput_Click(object sender, EventArgs e)
        {
            try
            {
                if (LstDisplay.Items.Count <= 0)
                {
                    return;
                }

                StringBuilder buffer = new StringBuilder();
                foreach (object t in LstDisplay.Items)
                {
                    buffer.Append(t);
                    buffer.Append('\n');
                }

                Clipboard.SetText(buffer.ToString());
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.ClearAndAdd, ex.Message);
                Log("BtnCopyOutput Error");
                Log(ex.Message);
            }
        }

        private void BtnListFormulas_Click(object sender, EventArgs e)
        {
            try
            {
                int count = 0;
                LstDisplay.Items.Clear();
                foreach (Sheet sht in Excel_Helpers.ExcelOpenXml.GetWorkSheets(TxtFileName.Text))
                {
                    count++;
                    LstDisplay.Items.Add(count + ". Worksheet = " + sht.Name);
                    SheetData sData = sht.GetFirstChild<SheetData>();
                    foreach (Row row in sht.ChildElements)
                    {
                        foreach (Cell cell in row.Elements<Cell>().ElementAt(2))
                        {
                            LstDisplay.Items.Add(cell.CellValue.ToString() + cell.CellFormula);
                        }
                    }
                }

                if (count == 0)
                {
                    LstDisplay.Items.Add("** No formulas in file **");
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.ClearAndAdd, ex.Message);
                Log("BtnListFormulas_Click Error");
                Log(ex.Message);
            }
        }

        private void BtnListFonts_Click(object sender, EventArgs e)
        {
            try
            {
                LstDisplay.Items.Clear();
                int count = 0;
                using (WordprocessingDocument doc = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    FontTablePart fontPart = doc.MainDocumentPart.FontTablePart;
                    foreach (DocumentFormat.OpenXml.Wordprocessing.Font ft in fontPart.Fonts)
                    {
                        count++;
                        LstDisplay.Items.Add(count + ". " + ft.Name);
                    }
                }

                if (count == 0)
                {
                    LstDisplay.Items.Add("** No Fonts **");
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                Log("BtnListFonts_Click Error");
                Log(ex.Message);
            }
        }

        private void BtnListFootnotes_Click(object sender, EventArgs e)
        {
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    LstDisplay.Items.Clear();
                    FootnotesPart footnotePart = doc.MainDocumentPart.FootnotesPart;
                    if (footnotePart != null)
                    {
                        int count = 0;
                        foreach (Footnote fn in footnotePart.Footnotes)
                        {
                            if (fn.InnerText != "")
                            {
                                count++;
                                DisplayInformation(InformationOutput.TextOnly, count + ". " + fn.InnerText);
                            }
                        }

                        if (count == 0)
                        {
                            DisplayInformation(InformationOutput.TextOnly, _noFootNotes);
                        }
                    }
                    else
                    {
                        DisplayInformation(InformationOutput.TextOnly, _noFootNotes);
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.ClearAndAdd, ex.Message);
                Log("BtnListFootnotes_Click Error");
                Log(ex.Message);
            }
        }

        private void BtnListEndnotes_Click(object sender, EventArgs e)
        {
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    LstDisplay.Items.Clear();
                    EndnotesPart endnotePart = doc.MainDocumentPart.EndnotesPart;
                    if (endnotePart != null)
                    {
                        int count = 0;
                        foreach (Endnote en in endnotePart.Endnotes)
                        {
                            if (en.InnerText != "")
                            {
                                count++;
                                DisplayInformation(InformationOutput.TextOnly, count + ". " + en.InnerText);
                            }
                        }

                        if (count == 0)
                        {
                            DisplayInformation(InformationOutput.TextOnly, _noEndNotes);
                        }
                    }
                    else
                    {
                        DisplayInformation(InformationOutput.TextOnly, _noEndNotes);
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.ClearAndAdd, ex.Message);
                Log("BtnListEndnotes_Click Error");
                Log(ex.Message);
            }
        }

        private void BtnDeleteFootnotes_Click(object sender, EventArgs e)
        {
            try
            {
                LstDisplay.Items.Clear();
                Word_Helpers.WordOpenXml.RemoveFootnotes(TxtFileName.Text);
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.ClearAndAdd, ex.Message);
                Log("BtnDeleteFootnotes_Click Error");
                Log(ex.Message);
            }
        }

        private void BtnDeleteEndnotes_Click(object sender, EventArgs e)
        {
            try
            {
                LstDisplay.Items.Clear();
                Word_Helpers.WordOpenXml.RemoveEndnotes(TxtFileName.Text);
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.ClearAndAdd, ex.Message);
                Log("BtnDeleteEndnotes_Click Error");
                Log(ex.Message);
            }
        }

        private void BtnListRevisions_Click(object sender, EventArgs e)
        {
            int revCount = 0;
            LstDisplay.Items.Clear();
            Cursor = Cursors.WaitCursor;
            try
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    Document doc = document.MainDocumentPart.Document;
                    var paragraphChanged = doc.Descendants<ParagraphPropertiesChange>().ToList();
                    var runChanged = doc.Descendants<RunPropertiesChange>().ToList();
                    var deleted = doc.Descendants<DeletedRun>().ToList();
                    var deletedParagraph = doc.Descendants<Deleted>().ToList();
                    var inserted = doc.Descendants<InsertedRun>().ToList();

                    // get the list of authors
                    _fromAuthor = "";

                    Forms.FrmAuthors aFrm = new Forms.FrmAuthors(TxtFileName.Text, document)
                    {
                        Owner = this
                    };
                    aFrm.ShowDialog();
                    
                    if (!String.IsNullOrEmpty(_fromAuthor))
                    {
                        paragraphChanged = paragraphChanged.Where(item => item.Author == _fromAuthor).ToList();
                        runChanged = runChanged.Where(item => item.Author == _fromAuthor).ToList();
                        deleted = deleted.Where(item => item.Author == _fromAuthor).ToList();
                        inserted = inserted.Where(item => item.Author == _fromAuthor).ToList();
                        deletedParagraph = deletedParagraph.Where(item => item.Author == _fromAuthor).ToList();

                        if ((paragraphChanged.Count + runChanged.Count + deleted.Count + inserted.Count + deletedParagraph.Count) == 0)
                        {
                            DisplayInformation(InformationOutput.ClearAndAdd, "** This author has no changes **");
                            Cursor = Cursors.Default;
                            return;
                        }
                    }
                    else
                    {
                        Cursor = Cursors.Default;
                        DisplayInformation(InformationOutput.ClearAndAdd, "** There are no revisions in this document **");
                        return;
                    }

                    foreach (var item in paragraphChanged)
                    {
                        revCount++;
                        LstDisplay.Items.Add(revCount + ": Paragraph Changed ");
                    }

                    foreach (var item in deletedParagraph)
                    {
                        revCount++;
                        LstDisplay.Items.Add(revCount + ": Paragraph Deleted ");
                    }

                    foreach (var item in runChanged)
                    {
                        revCount++;
                        LstDisplay.Items.Add(revCount + ": Run Changed = " + item.InnerText);
                    }

                    foreach (var item in deleted)
                    {
                        revCount++;
                        LstDisplay.Items.Add(revCount + ": Deletion = " + item.InnerText);
                    }

                    foreach (var item in inserted)
                    {
                        if (item.Parent != null)
                        {
                            var textRuns = item.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>().ToList();
                            var parent = item.Parent;

                            foreach (var textRun in textRuns)
                            {
                                revCount++;
                                LstDisplay.Items.Add(revCount + ": Insertion = " + textRun.InnerText);
                            }
                        }
                    }
                }

                Cursor = Cursors.Default;
            }
            catch(Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                Log("BtnListRevisions_Click Error");
                Log(ex.Message);
                Cursor = Cursors.Default;
            }
        }

        private void BtnListAuthors_Click(object sender, EventArgs e)
        {
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    LstDisplay.Items.Clear();
                    WordprocessingPeoplePart peoplePart = doc.MainDocumentPart.WordprocessingPeoplePart;
                    if (peoplePart != null)
                    {
                        int count = 0;
                        foreach (Person person in peoplePart.People)
                        {
                            count++;
                            PresenceInfo pi = person.PresenceInfo;
                            LstDisplay.Items.Add(count + ". " + person.Author);
                            LstDisplay.Items.Add("   - User Id = " + pi.UserId);
                            LstDisplay.Items.Add("   - Provider Id = " + pi.ProviderId);
                        }
                    }
                    else
                    {
                        DisplayInformation(InformationOutput.TextOnly, "** There are no authors in this document **");
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                Log("BtnListAuthors_Click Error");
                Log(ex.Message);
            }
        }
        
        private void BtnViewCustomDocProps_Click(object sender, EventArgs e)
        {
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    LstDisplay.Items.Clear();
                    DocumentSettingsPart docSettingsPart = doc.MainDocumentPart.DocumentSettingsPart;

                    GetStandardFileProps(doc.PackageProperties);
                    GetExtendedFileProps(doc.ExtendedFilePropertiesPart);

                    try
                    {
                        if (docSettingsPart != null)
                        {
                            DocumentFormat.OpenXml.Wordprocessing.Settings settings = docSettingsPart.Settings;
                            foreach (var setting in settings)
                            {
                                if (setting.LocalName == "compat")
                                {
                                    LstDisplay.Items.Add("");
                                    LstDisplay.Items.Add("---- Compatibility Settings ---- ");
                                    foreach (CompatibilitySetting compat in setting)
                                    {
                                        LstDisplay.Items.Add("   - " + compat.Name + ": " + compat.Val);
                                    }
                                    LstDisplay.Items.Add("");
                                }
                                else
                                {
                                    XmlDocument xDoc = new XmlDocument();
                                    xDoc.LoadXml(setting.OuterXml);

                                    foreach (XmlElement xe in xDoc.ChildNodes)
                                    {
                                        if (xe.Attributes.Count > 1)
                                        {
                                            LstDisplay.Items.Add(xe.LocalName);
                                            foreach (XmlAttribute xa in xe.Attributes)
                                            {
                                                if (!(xa.LocalName == "w" || xa.LocalName == "m" || xa.LocalName == "w14" || xa.LocalName == "w15" || xa.LocalName == "w16"))
                                                {
                                                    if (!xa.Value.StartsWith("http"))
                                                    {
                                                        if (xa.LocalName == "val")
                                                        {
                                                            LstDisplay.Items.Add("-- " + xa.Value);
                                                        }
                                                        else
                                                        {
                                                            LstDisplay.Items.Add("-- " + xa.LocalName + ": " + xa.Value);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            DisplayInformation(InformationOutput.TextOnly, "** There are no custom properties in this document **");
                        }
                    }
                    catch (Exception ex)
                    {
                        DisplayInformation(InformationOutput.TextOnly, ex.Message);
                        Log("BtnViewCustomDocProps (doc settings) Error");
                        Log(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                Log("BtnViewCustomDocProps_Click Error");
                Log(ex.Message);
            }
        }

        public void GetStandardFileProps(System.IO.Packaging.PackageProperties props)
        {
            // Display file package props
            LstDisplay.Items.Add("---- Document Properties ----");
            LstDisplay.Items.Add("Creator : " + props.Creator);
            LstDisplay.Items.Add("Created : " + props.Created);
            LstDisplay.Items.Add("Last Modified By : " + props.LastModifiedBy);
            LstDisplay.Items.Add("Last Printed : " + props.LastPrinted);
            LstDisplay.Items.Add("Modified : " + props.Modified);
            LstDisplay.Items.Add("Subject : " + props.Subject);
            LstDisplay.Items.Add("Revision : " + props.Revision);
            LstDisplay.Items.Add("Title : " + props.Title);
            LstDisplay.Items.Add("Version : " + props.Version);
            LstDisplay.Items.Add("");
        }

        public void GetExtendedFileProps(ExtendedFilePropertiesPart exFilePropPart)
        {
            XmlDocument xmlProps = new XmlDocument();
            xmlProps.Load(exFilePropPart.GetStream());
            XmlNodeList exProps = xmlProps.GetElementsByTagName("Properties");

            LstDisplay.Items.Add("---- Extended File Properties ----");
            try
            {
                foreach (XmlNode xNode in exProps)
                {
                    foreach (XmlElement xElement in xNode)
                    {
                        DisplayElementDetails(xElement);
                    }
                }
            }
            catch (Exception ex)
            {
                // log the error 
                Log("GetExtendedFileProps Error");
                Log(ex.Message);
            }
        }

        public void DisplayElementDetails(XmlElement elem)
        {
            LstDisplay.Items.Add(elem.Name + " : " + elem.InnerText);
        }

        private void MnuAbout_Click(object sender, EventArgs e)
        {
            Forms.FrmAbout frm = new Forms.FrmAbout();
            frm.ShowDialog(this);
            frm.Dispose();
        }

        private void MnuOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog fDialog = new OpenFileDialog
            {
                Title = "Select Office Open Xml File.",
                Filter = "Open XML Files | *.docx; *.dotx; *.docm; *.dotm; *.xlsx; *.xlsm; *.xlst; *.xltm; *.pptx; *.pptm; *.potx; *.potm",
                RestoreDirectory = true,
                InitialDirectory = @"%userprofile%"
            };

            if (fDialog.ShowDialog() == DialogResult.OK)
            {
                TxtFileName.Text = fDialog.FileName.ToString();
                if (!File.Exists(TxtFileName.Text))
                {
                    DisplayInformation(InformationOutput.InvalidFile, _fileDoesNotExist);
                    return;
                }
                else
                {
                    LstDisplay.Items.Clear();
                    SetUpButtons();
                }
            }
            else
            {
                return;
            }
        }

        private void MnuExit_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();
            Application.Exit();
        }

        private void MnuCheckForUpdates_Click(object sender, EventArgs e)
        {
            // Force the application to check for an update
            UpdateCheckInfo info = null;

            if (ApplicationDeployment.IsNetworkDeployed)
            {
                ApplicationDeployment ad = ApplicationDeployment.CurrentDeployment;

                try
                {
                    info = ad.CheckForDetailedUpdate();
                }
                catch (DeploymentDownloadException dde)
                {
                    MessageBox.Show("The new version of the application cannot be downloaded at this time. Please check your network connection, or try again later. Error: " + dde.Message, _unableToDownloadUpdate, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                catch (InvalidDeploymentException ide)
                {
                    MessageBox.Show("Cannot check for a new version of the application. The ClickOnce deployment is corrupt. Please redeploy the application and try again. Error: " + ide.Message, _unableToDownloadUpdate, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                catch (InvalidOperationException ioe)
                {
                    MessageBox.Show("This application cannot be updated. It is likely not a ClickOnce application. Error: " + ioe.Message, _unableToDownloadUpdate, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (info.UpdateAvailable)
                {
                    Boolean doUpdate = true;

                    if (!info.IsUpdateRequired)
                    {
                        DialogResult dr = MessageBox.Show("An update is available. Would you like to update the application now?", "Update Available",
                            MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                        if (!(DialogResult.OK == dr))
                        {
                            doUpdate = false;
                        }
                    }
                    else
                    {
                        // Display a message that the app MUST reboot. Display the minimum required version.
                        MessageBox.Show("This application has detected a mandatory update from your current " +
                            "version to version " + info.MinimumRequiredVersion +
                            ". The application will now install the update and restart.",
                            "Update Available", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }

                    if (doUpdate)
                    {
                        try
                        {
                            ad.Update();
                            MessageBox.Show("The application has been upgraded, and will now restart.", "Upgrade successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            Application.Restart();
                        }
                        catch (DeploymentDownloadException dde)
                        {
                            MessageBox.Show("Cannot install the latest version of the application. Please check your network connection, or try again later. Error: " + dde, _unableToDownloadUpdate, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("You already have the latest version.", "Application Update",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("The new version of the application cannot be downloaded at this time.", _unableToDownloadUpdate, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnPPTListHyperlinks_Click(object sender, EventArgs e)
        {
            try
            {
                // Open the presentation file as read-only.
                using (PresentationDocument document = PresentationDocument.Open(TxtFileName.Text, false))
                {
                    int linkCount = 0;
                    foreach (string s in PowerPointOpenXml.GetAllExternalHyperlinksInPresentation(TxtFileName.Text))
                    {
                        linkCount++;
                        LstDisplay.Items.Add(linkCount + ". " + s);
                    }

                    if (linkCount == 0)
                    {
                        DisplayInformation(InformationOutput.ClearAndAdd, "** No Hyperlinks in file **");
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
            }
        }

        private void BtnPPTGetAllSlideTitles_Click(object sender, EventArgs e)
        {
            try
            {
                // Open the presentation as read-only.
                using (PresentationDocument presentationDocument = PresentationDocument.Open(TxtFileName.Text, false))
                {
                    int slideCount = 0;

                    foreach (string s in PowerPointOpenXml.GetSlideTitles(presentationDocument))
                    {
                        slideCount++;
                        LstDisplay.Items.Add(slideCount + ". " + s);
                    }

                    if (slideCount == 0)
                    {
                        DisplayInformation(InformationOutput.ClearAndAdd, "** No slides in file **");
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
            }
        }

        private void BtnSearchAndReplace_Click(object sender, EventArgs e)
        {
            try
            {
                Forms.FrmSearchAndReplace sFrm = new Forms.FrmSearchAndReplace()
                {
                    Owner = this
                };
                sFrm.ShowDialog();

                if (_FindText == "" && _ReplaceText == "")
                {
                    return;
                }
                else
                {
                    SearchAndReplace(TxtFileName.Text, _FindText, _ReplaceText);
                    LstDisplay.Items.Clear();
                    LstDisplay.Items.Add("** Search and Replace Finished **");
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.TextOnly, ex.Message);
                Log("BtnSearchAndReplace_Click Error");
                Log(ex.Message);
            }
        }

        // To search and replace content in a document part.
        public static void SearchAndReplace(string document, string find, string replace)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexText = new Regex(find);
                docText = regexText.Replace(docText, replace);

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

        private void BtnListLinks_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, true))
                {
                    WorkbookPart wbPart = excelDoc.WorkbookPart;
                    int ExtRelCount = 0;
                    LstDisplay.Items.Clear();

                    if (wbPart.ExternalWorkbookParts.Count() == 0)
                    {
                        LstDisplay.Items.Add("** No External Links **");
                        return;
                    }

                    foreach (ExternalWorkbookPart extWbPart in wbPart.ExternalWorkbookParts)
                    {
                        ExtRelCount++;
                        ExternalRelationship extRel = extWbPart.ExternalRelationships.ElementAt(0);
                        LstDisplay.Items.Add(ExtRelCount + ". " + extWbPart.ExternalRelationships.ElementAt(0).Uri);
                    }
                }
            }
            catch (Exception ex)
            {
                // log the error 
                Log("BtnListLinks_Click Error");
                Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnDeleteExternalLinks_Click(object sender, EventArgs e)
        {
            Excel_Helpers.ExcelOpenXml.RemoveExternalLinks(TxtFileName.Text);
            LstDisplay.Items.Clear();
            LstDisplay.Items.Add("** External References Deleted **");
        }

        public void Log(string logValue)
        {
            Properties.Settings.Default.ErrorLog.Add(DateTime.Now + " : " + logValue);
            Properties.Settings.Default.Save();
        }

        private void BtnErrorLog_Click(object sender, EventArgs e)
        {
            Forms.FrmErrorLog errFrm = new Forms.FrmErrorLog()
            {
                Owner = this
            };
            errFrm.ShowDialog();
        }

        private void BtnListDefinedNames_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                LstDisplay.Items.Clear();
                int nameCount = 0;

                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, true))
                {
                    WorkbookPart wbPart = excelDoc.WorkbookPart;

                    // Retrieve a reference to the defined names collection.
                    DefinedNames definedNames = wbPart.Workbook.DefinedNames;

                    // If there are defined names, add them to the dictionary.
                    if (definedNames != null)
                    {
                        foreach (DefinedName dn in definedNames)
                        {
                            nameCount++;
                            LstDisplay.Items.Add(nameCount + ". " + dn.Name.Value + " = " + dn.Text);
                        }
                    }
                    else
                    {
                        LstDisplay.Items.Add("** No Defined Names **");
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.TextOnly, ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnListHiddenRowsColumns_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                LstDisplay.Items.Clear();

                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, true))
                {
                    WorkbookPart wbPart = excelDoc.WorkbookPart;
                    Sheets theSheets = wbPart.Workbook.Sheets;

                    foreach (Sheet sheet in theSheets)
                    {
                        LstDisplay.Items.Add("Worksheet Name = " + sheet.Name);
                        Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().Where((s) => s.Name == sheet.Name).FirstOrDefault();

                        if (theSheet == null)
                        {
                            Log("BtnListHiddenRowsColumnClickError" + sheet.Name);
                            throw new ArgumentException("sheetName");
                        }
                        else
                        {
                            // The sheet does exist.
                            WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
                            Worksheet ws = wsPart.Worksheet;
                            int rowCount = 0;
                            int colCount = 0;

                            List<uint> rowList = new List<uint>();

                            // Retrieve hidden rows, start by calling the Descendants method of the worksheet, then retrieve a list of all rows. 
                            // The Where method limits the results to only those rows where the Hidden property of the item is not null and the value of the Hidden property is True.
                            // The Select method projects the return value for each row, returning the value of the RowIndex property.
                            // Finally, the ToList < TSource > **method converts the resulting IEnumerable < T > interface into a List<T> object of unsigned integers. 
                            // If there are no hidden rows, the returned list is empty.
                            rowList = ws.Descendants<Row>().Where((r) => r.Hidden != null && r.Hidden.Value).Select(r => r.RowIndex.Value).ToList<uint>();
                            LstDisplay.Items.Add("##    ROWS    ##");
                            foreach (object row in rowList)
                            {
                                rowCount++;
                                LstDisplay.Items.Add(rowCount + ". Row " + row);
                            }

                            if (rowCount == 0)
                            {
                                LstDisplay.Items.Add("    None");
                            }

                            // Retrieve hidden columns is a bit trickier because Excel collapses groups of hidden columns into a single element, 
                            // and provides Min and Max properties that describe the first and last columns in the group. 
                            // Therefore, the code that retrieves the list of hidden columns starts the same as the code that retrieves hidden rows. 
                            // However, it must iterate through the index values (looping each item in the collection, adding each index from the Min to the Max value, inclusively).
                            LstDisplay.Items.Add("##    COLUMNS    ##");
                            var cols = ws.Descendants<Column>().Where((c) => c.Hidden != null && c.Hidden.Value);
                            foreach (Column item in cols)
                            {
                                for (uint i = item.Min.Value; i <= item.Max.Value; i++)
                                {
                                    colCount++;
                                    LstDisplay.Items.Add(colCount + ". Column " + i);
                                }
                            }

                            if (colCount == 0)
                            {
                                LstDisplay.Items.Add("    None");
                            }
                        }
                        LstDisplay.Items.Add("");
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.TextOnly, ex.Message);
                Log("BtnListHiddenRowsColumns_Click Error");
                Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnListWorksheets_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                LstDisplay.Items.Clear();
                int sheetCount = 0;

                foreach (Sheet sht in Excel_Helpers.ExcelOpenXml.GetWorkSheets(TxtFileName.Text))
                {
                    sheetCount++;
                    LstDisplay.Items.Add(sheetCount + ". " + sht.Name);
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.TextOnly, ex.Message);
                Log("BtnListWorksheets_Click Error");
                Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnListHiddenWorksheets_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                LstDisplay.Items.Clear();
                int hiddenCount = 0;

                foreach (Sheet sht in Excel_Helpers.ExcelOpenXml.GetHiddenSheets(TxtFileName.Text))
                {
                    hiddenCount++;
                    LstDisplay.Items.Add(hiddenCount + ". " + sht.Name);
                }

                if (hiddenCount == 0)
                {
                    LstDisplay.Items.Add("** No Hidden Worksheets **");
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.TextOnly, ex.Message);
                Log("BtnListHiddenWorksheets_Click Error");
                Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnListSharedStrings_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                LstDisplay.Items.Clear();
                int sharedStringCount = 0;

                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, true))
                {
                    WorkbookPart wbPart = excelDoc.WorkbookPart;
                    SharedStringTablePart sstp = wbPart.SharedStringTablePart;
                    SharedStringTable sst = sstp.SharedStringTable;
                    foreach (SharedStringItem ssi in sst)
                    {
                        sharedStringCount++;
                        DocumentFormat.OpenXml.Spreadsheet.Text ssValue = ssi.Text;
                        LstDisplay.Items.Add(sharedStringCount + ". " + ssValue.Text);
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.TextOnly, ex.Message);
                Log("BtnListSharedStrings_Click Error");
                Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }
        
        private void BtnComments_Click(object sender, EventArgs e)
        {
            try
            {
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, true))
                {
                    WorkbookPart wbPart = excelDoc.WorkbookPart;
                    int commentCount = 1;
                    LstDisplay.Items.Clear();

                    foreach (WorksheetPart wsp in wbPart.WorksheetParts)
                    {
                        WorksheetCommentsPart wcp = wsp.WorksheetCommentsPart;
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Comment cmt in wcp.Comments.CommentList)
                        {
                            CommentText cText = cmt.CommentText;
                            LstDisplay.Items.Add(commentCount + ". " + cText.InnerText);
                            commentCount++;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("Excel - BtnComments_Click Error:");
                Log(ex.Message);
                DisplayInformation(InformationOutput.TextOnly, "** No Comments **");
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnDeleteComment_Click(object sender, EventArgs e)
        {
            try
            {
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, true))
                {
                    WorkbookPart wbPart = excelDoc.WorkbookPart;
                    LstDisplay.Items.Clear();

                    foreach (WorksheetPart wsp in wbPart.WorksheetParts)
                    {
                        WorksheetCommentsPart wcp = wsp.WorksheetCommentsPart;
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Comment cmt in wcp.Comments.CommentList)
                        {
                            cmt.Remove();
                        }
                    }

                    wbPart.Workbook.Save();
                    LstDisplay.Items.Add("** Comments Deleted **");
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.TextOnly, ex.Message);
                Log("BtnListFormulas_Click Error");
                Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnChangeTheme_Click(object sender, EventArgs e)
        {
            string sThemeFilePath = "";

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

                if (!File.Exists(TxtFileName.Text))
                {
                    DisplayInformation(InformationOutput.InvalidFile, _fileDoesNotExist);
                    return;
                }
                else
                {
                    if (fileType == _word)
                    {
                        // call the replace function using the theme file provided
                        Word_Helpers.WordOpenXml.ReplaceTheme(TxtFileName.Text, sThemeFilePath);
                        DisplayInformation(InformationOutput.ClearAndAdd, _themeFileAdded);
                    }
                    else if (fileType == _excel)
                    {
                        // call the replace function using the theme file provided
                        Excel_Helpers.ExcelOpenXml.ReplaceTheme(TxtFileName.Text, sThemeFilePath);
                        DisplayInformation(InformationOutput.ClearAndAdd, _themeFileAdded);
                    }
                    else if (fileType == _powerpoint)
                    {
                        // call the replace function using the theme file provided
                        PowerPointOpenXml.ReplaceTheme(TxtFileName.Text, sThemeFilePath);
                        DisplayInformation(InformationOutput.ClearAndAdd, _themeFileAdded);
                    }
                    else
                    {
                        DisplayInformation(InformationOutput.ClearAndAdd, "File Not Valid.");
                    }
                }
            }
            else
            {
                return;
            }
        }

        private void BtnViewPPTComments_Click(object sender, EventArgs e)
        {
            try
            {
                // Open the presentation as read-only.
                using (PresentationDocument presentationDocument = PresentationDocument.Open(TxtFileName.Text, false))
                {
                    PresentationPart pPart = presentationDocument.PresentationPart;
                    int commentCount = 0;
                    LstDisplay.Items.Clear();

                    foreach (SlidePart sPart in pPart.SlideParts)
                    {
                        SlideCommentsPart sCPart = sPart.SlideCommentsPart;
                        foreach (DocumentFormat.OpenXml.Presentation.Comment cmt in sCPart.CommentList)
                        {
                            commentCount++;
                            LstDisplay.Items.Add(commentCount + ". " + cmt.InnerText);       
                        }
                    }

                    if (commentCount == 0)
                    {
                        DisplayInformation(InformationOutput.ClearAndAdd, "** File does not have any comments **");
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                Log("PPT - BtnListComments_Click Error");
                Log(ex.Message);
            }
        }
    }
}