﻿/****************************** Module Header ******************************\
Module Name:  FrmMain.cs
Project:      Office File Explorer
Copyright (c) Microsoft Corporation.

Main window for OFE.

This source is subject to the following license.
See https://github.com/desjarlais/Office-File-Explorer/blob/master/LICENSE
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND,
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

// Open Xml SDK refs
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

// shortcut namespace refs
using P = DocumentFormat.OpenXml.Presentation;
using O = DocumentFormat.OpenXml;
using AO = DocumentFormat.OpenXml.Office.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using Column = DocumentFormat.OpenXml.Spreadsheet.Column;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Tag = DocumentFormat.OpenXml.Wordprocessing.Tag;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;

// this app references
using Office_File_Explorer.App_Helpers;
using Office_File_Explorer.Excel_Helpers;
using Office_File_Explorer.Forms;
using Office_File_Explorer.PowerPoint_Helpers;
using Office_File_Explorer.Word_Helpers;

// .Net refs
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Path = System.IO.Path;

namespace Office_File_Explorer
{
    public partial class FrmMain : Form
    {
        // globals
        private string _fromAuthor;
        private string _FindText;
        private string _ReplaceText;
        public static char PrevChar = '<';
        public bool IsRegularXmlTag;
        public bool IsFixed;
        public static string FixedFallback = string.Empty;
        public static string StrOrigFileName = string.Empty;
        public static string StrDestPath = string.Empty;
        public static string StrExtension = string.Empty;
        public static string StrDestFileName = string.Empty;
        private string fileType;

        // global numid lists
        private ArrayList oNumIdList = new ArrayList();
        private ArrayList aNumIdList = new ArrayList();
        private ArrayList numIdList = new ArrayList();

        // fix corrupt doc globals
        private static List<string> _nodes = new List<string>();

        // global packageparts
        private static List<string> _pParts = new List<string>();

        // corrupt doc buffer
        private static StringBuilder _sbNodeBuffer = new StringBuilder();

        public enum InformationOutput { ClearAndAdd, Append, TextOnly, InvalidFile };

        public FrmMain()
        {
            InitializeComponent();
            LoggingHelper.Log("App Start");
            _FindText = StringResources.emptyString;
            _ReplaceText = StringResources.emptyString;

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

        #endregion Class Properties

        /// <summary>
        /// Disable all buttons on the form and reset file type
        /// </summary>
        public void DisableButtons()
        {
            fileType = StringResources.emptyString;
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
            BtnListWSInfo.Enabled = false;
            BtnListCellValuesSAX.Enabled = false;
            BtnListCellValuesDOM.Enabled = false;
            BtnConvertDocmToDocx.Enabled = false;
            BtnListSlideText.Enabled = false;
            BtnFixCorruptDocument.Enabled = false;
            BtnListConnections.Enabled = false;
            BtnListCustomProps.Enabled = false;
            BtnSetCustomProps.Enabled = false;
            BtnSetPrintOrientation.Enabled = false;
            BtnViewParagraphs.Enabled = false;
            BtnConvertPptmToPptx.Enabled = false;
            BtnConvertXlsmToXlsx.Enabled = false;
            BtnListPackageParts.Enabled = false;
            BtnListFieldCodes.Enabled = false;
            BtnListBookmarks.Enabled = false;
            BtnListCC.Enabled = false;
            BtnListShapes.Enabled = false;
            BtnListParagraphStyles.Enabled = false;
            BtnNotesPageSize.Enabled = false;
        }

        public enum OxmlFileFormat { Xlsx, Xlsm, Xlst, Dotx, Docx, Docm, Potx, Pptx, Pptm, Invalid };

        public OxmlFileFormat GetFileFormat()
        {
            string fileExt = System.IO.Path.GetExtension(TxtFileName.Text);
            fileExt = fileExt.ToLower();

            if (fileExt == ".docx")
            {
                return OxmlFileFormat.Docx;
            }
            else if (fileExt == ".docm")
            {
                return OxmlFileFormat.Docm;
            }
            else if (fileExt == ".dotx")
            {
                return OxmlFileFormat.Dotx;
            }
            else if (fileExt == ".xlst")
            {
                return OxmlFileFormat.Xlst;
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
            else if (fileExt == ".potx")
            {
                return OxmlFileFormat.Potx;
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
            OxmlFileFormat ffmt = GetFileFormat();

            if (ffmt == OxmlFileFormat.Docx || ffmt == OxmlFileFormat.Docm || ffmt == OxmlFileFormat.Dotx)
            {
                fileType = StringResources.word;

                // enable WD only files
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
                BtnViewCustomDocProps.Enabled = true;
                BtnSetPrintOrientation.Enabled = true;
                BtnViewParagraphs.Enabled = true;
                BtnRemovePII.Enabled = true;
                BtnListFieldCodes.Enabled = true;
                BtnListBookmarks.Enabled = true;
                BtnListCC.Enabled = true;
                BtnListParagraphStyles.Enabled = true;

                if (ffmt == OxmlFileFormat.Docm)
                {
                    BtnConvertDocmToDocx.Enabled = true;
                }
            }
            else if (ffmt == OxmlFileFormat.Xlsx || ffmt == OxmlFileFormat.Xlsm || ffmt == OxmlFileFormat.Xlst)
            {
                fileType = StringResources.excel;

                // enable XL only files
                BtnListDefinedNames.Enabled = true;
                BtnListHiddenRowsColumns.Enabled = true;
                BtnDeleteExternalLinks.Enabled = true;
                BtnListLinks.Enabled = true;
                BtnListFormulas.Enabled = true;
                BtnListWorksheets.Enabled = true;
                BtnListSharedStrings.Enabled = true;
                BtnComments.Enabled = true;
                BtnDeleteComment.Enabled = true;
                BtnListWSInfo.Enabled = true;
                BtnListCellValuesSAX.Enabled = true;
                BtnListCellValuesDOM.Enabled = true;
                BtnListConnections.Enabled = true;

                if (ffmt == OxmlFileFormat.Xlsm)
                {
                    BtnConvertXlsmToXlsx.Enabled = true;
                }
            }
            else if (ffmt == OxmlFileFormat.Pptx || ffmt == OxmlFileFormat.Pptm || ffmt == OxmlFileFormat.Potx)
            {
                fileType = StringResources.powerpoint;

                // enable PPT only files
                BtnPPTGetAllSlideTitles.Enabled = true;
                BtnPPTListHyperlinks.Enabled = true;
                BtnViewPPTComments.Enabled = true;
                BtnListSlideText.Enabled = true;
                BtnNotesPageSize.Enabled = true;

                if (ffmt == OxmlFileFormat.Pptm)
                {
                    BtnConvertPptmToPptx.Enabled = true;
                }
            }
            else if (ffmt == OxmlFileFormat.Invalid)
            {
                // invalid file format
                MessageBox.Show("Unsupported File Format");
                return;
            }
            else
            {
                // unknown condition, log details
                LoggingHelper.Log("GetFileFormat Error: " + TxtFileName.Text);
                return;
            }

            // these buttons exists for all file types
            BtnValidateFile.Enabled = true;
            BtnChangeTheme.Enabled = true;
            BtnListOle.Enabled = true;
            BtnListCustomProps.Enabled = true;
            BtnSetCustomProps.Enabled = true;
            BtnListPackageParts.Enabled = true;
            BtnListShapes.Enabled = true;
        }

        private void BtnListComments_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                LstDisplay.Items.Clear();

                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    WordprocessingCommentsPart commentsPart = myDoc.MainDocumentPart.WordprocessingCommentsPart;
                    if (commentsPart == null)
                    {
                        DisplayEmptyCount(0, "comments");
                    }
                    else
                    {
                        int count = 0;
                        foreach (DocumentFormat.OpenXml.Wordprocessing.Comment cm in commentsPart.Comments)
                        {
                            count++;
                            LstDisplay.Items.Add(count + StringResources.period + cm.InnerText);
                        }
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
                LoggingHelper.Log("Word - BtnListComments_Click Error");
                LoggingHelper.Log(ex.Message);
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
                    LstDisplay.Items.Add(StringResources.emptyString);
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
                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    MainDocumentPart mainPart = myDoc.MainDocumentPart;
                    StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart;
                    bool containStyle = false;

                    LstDisplay.Items.Clear();
                    try
                    {
                        foreach (OpenXmlElement el in stylePart.Styles.LatentStyles.Elements())
                        {
                            string styleEl = el.GetAttribute("name", StringResources.wordMainAttributeNamespace).Value;
                            int pStyle = WordExtensionClass.ParagraphsByStyleName(mainPart, styleEl).Count();
                            int rStyle = WordExtensionClass.RunsByStyleName(mainPart, styleEl).Count();
                            int tStyle = WordExtensionClass.TablesByStyleName(mainPart, styleEl).Count();

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
                LoggingHelper.Log("BtnListStyles_Click Error");
                LoggingHelper.Log(ex.Message);
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
                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    // first check that hyperlinks exist
                    int count = 0;
                    if (myDoc.MainDocumentPart.HyperlinkRelationships.Count() == 0 && myDoc.MainDocumentPart.RootElement.Descendants<FieldCode>().Count() == 0)
                    {
                        DisplayEmptyCount(0, "hyperlinks");
                    }

                    // then check for regular hyperlinks
                    foreach (HyperlinkRelationship hRel in myDoc.MainDocumentPart.HyperlinkRelationships)
                    {
                        count++;
                        LstDisplay.Items.Add(count + StringResources.period + hRel.Uri);
                    }

                    // now we need to check for field hyperlinks
                    foreach (var field in myDoc.MainDocumentPart.RootElement.Descendants<FieldCode>())
                    {
                        string fldText;
                        if (field.InnerText.StartsWith(" HYPERLINK"))
                        {
                            count++;
                            fldText = field.InnerText.Remove(0, 11);
                            LstDisplay.Items.Add(count + StringResources.period + fldText);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                LoggingHelper.Log("BtnListHyperlinks_Click Error");
                LoggingHelper.Log(ex.Message);
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
                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, false))
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
                            string styleEl = el.GetAttribute("styleId", StringResources.wordMainAttributeNamespace).Value;
                            int pStyle = WordExtensionClass.ParagraphsByStyleName(mainPart, styleEl).Count();

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
                            LoggingHelper.Log("BtnListTemplates_Click : " + ex.Message);
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
                    LstDisplay.Items.Add(StringResources.emptyString);
                    LstDisplay.Items.Add("All List Templates in document:");
                    int aCount = 0;
                    foreach (OpenXmlElement el in numPart.Numbering.Elements())
                    {
                        foreach (AbstractNumId aNumId in el.Descendants<AbstractNumId>())
                        {
                            string strNumId = el.GetAttribute("numId", StringResources.wordMainAttributeNamespace).Value;
                            aNumIdList.Add(strNumId);
                            aCount++;
                            LstDisplay.Items.Add(aCount + ". numId = " + strNumId);
                        }
                    }

                    // get the unused list templates
                    oNumIdList = OrphanedListTemplates(numIdList, aNumIdList);
                    LstDisplay.Items.Add(StringResources.emptyString);
                    LstDisplay.Items.Add("Orphaned List Templates:");
                    int oCount = 0;
                    foreach (object item in oNumIdList)
                    {
                        oCount++;
                        LstDisplay.Items.Add(oCount + ". numId = " + item);
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                LoggingHelper.Log("BtnListTemplates_Click Error");
                LoggingHelper.Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        // method to display the non-duplicate numId used in the document.
        private void LoopArrayList(ArrayList al)
        {
            LstDisplay.Items.Add("Active List Templates in this document:");
            
            // if we don't have any active templates, just continue checking for orphaned
            if (al.Count == 0)
            {
                LstDisplay.Items.Add(" -- none ");
                return;
            }

            int count = 0;
            foreach (object item in al)
            {
                count++;
                LstDisplay.Items.Add(count + ". numID = " + item);
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
                if (fileType == StringResources.word)
                {
                    using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, false))
                    {
                        int wdOleObjCount = GetEmbeddedObjectProperties(myDoc.MainDocumentPart);
                        int wdOlePkgPart = GetEmbeddedPackageProperties(myDoc.MainDocumentPart);

                        if (wdOlePkgPart == 0 && wdOleObjCount == 0)
                        {
                            DisplayInformation(InformationOutput.ClearAndAdd, StringResources.noOle);
                        }
                    }
                }
                else if (fileType == StringResources.excel)
                {
                    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(TxtFileName.Text, false))
                    {
                        int xlOleObjCount = 0;
                        int xlOlePkgPart = 0;

                        foreach (WorksheetPart wPart in doc.WorkbookPart.WorksheetParts)
                        {
                            xlOleObjCount += GetEmbeddedObjectProperties(wPart);
                            xlOlePkgPart += GetEmbeddedPackageProperties(wPart);
                        }

                        if (xlOlePkgPart == 0 && xlOleObjCount == 0)
                        {
                            DisplayInformation(InformationOutput.ClearAndAdd, StringResources.noOle);
                        }
                    }
                }
                else if (fileType == StringResources.powerpoint)
                {
                    using (PresentationDocument doc = PresentationDocument.Open(TxtFileName.Text, false))
                    {
                        int pptOleObjCount = 0;
                        int pptOlePkgPart = 0;

                        foreach (SlidePart sPart in doc.PresentationPart.SlideParts)
                        {
                            pptOleObjCount += GetEmbeddedObjectProperties(sPart);
                            pptOlePkgPart += GetEmbeddedPackageProperties(sPart);
                        }

                        if (pptOlePkgPart == 0 && pptOleObjCount == 0)
                        {
                            DisplayInformation(InformationOutput.ClearAndAdd, StringResources.noOle);
                        }
                    }
                }
                else
                {
                    DisplayInformation(InformationOutput.ClearAndAdd, StringResources.noOle);
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                LoggingHelper.Log("BtnListOle_Click Error");
                LoggingHelper.Log(ex.Message);
            }
        }

        /// <summary>
        /// Return Word Embedded Package count
        /// </summary>
        /// <param name="mPart"></param>
        /// <returns></returns>
        public int GetEmbeddedPackageProperties(MainDocumentPart mPart)
        {
            try
            {
                int x = 0;
                int olePkgCount = mPart.EmbeddedPackageParts.Count();

                do
                {
                    LstDisplay.Items.Add(mPart.EmbeddedPackageParts.ElementAt(x).Uri.ToString());
                    x++;
                }
                while (x < olePkgCount);

                return x;
            }
            catch (ArgumentOutOfRangeException)
            {
                return 0;
            }
        }

        /// <summary>
        /// Return Excel Embedded Package Count
        /// </summary>
        /// <param name="wPart"></param>
        /// <returns></returns>
        public int GetEmbeddedPackageProperties(WorksheetPart wPart)
        {
            try
            {
                int x = 0;
                int olePkgCount = wPart.EmbeddedPackageParts.Count();

                do
                {
                    LstDisplay.Items.Add(wPart.Uri + StringResources.arrow + wPart.EmbeddedPackageParts.ElementAt(x).Uri.ToString());
                    x++;
                }
                while (x < olePkgCount);

                return x;
            }
            catch (ArgumentOutOfRangeException)
            {
                return 0;
            }
        }

        /// <summary>
        /// Return PowerPoint Embedded Package Count
        /// </summary>
        /// <param name="wPart"></param>
        /// <returns></returns>
        public int GetEmbeddedPackageProperties(SlidePart sPart)
        {
            try
            {
                int x = 0;
                int olePkgCount = sPart.EmbeddedPackageParts.Count();

                do
                {
                    LstDisplay.Items.Add(sPart.Uri + StringResources.arrow + sPart.EmbeddedPackageParts.ElementAt(x).Uri.ToString());
                    x++;
                }
                while (x < olePkgCount);

                return x;
            }
            catch (ArgumentOutOfRangeException)
            {
                return 0;
            }
        }

        /// <summary>
        /// Return Word Embedded Object Count
        /// </summary>
        /// <param name="mPart"></param>
        /// <returns></returns>
        public int GetEmbeddedObjectProperties(MainDocumentPart mPart)
        {
            try
            {
                int x = 0;
                int oleEmbCount = mPart.EmbeddedObjectParts.Count();

                do
                {
                    LstDisplay.Items.Add(mPart.EmbeddedObjectParts.ElementAt(x).Uri.ToString());
                    x++;
                }
                while (x < oleEmbCount);

                return x;
            }
            catch (ArgumentOutOfRangeException)
            {
                return 0;
            }
        }

        /// <summary>
        /// Return Excel Embedded Object Count
        /// </summary>
        /// <param name="wPart"></param>
        /// <returns></returns>
        public int GetEmbeddedObjectProperties(WorksheetPart wPart)
        {
            try
            {
                int x = 0;
                int oleEmbCount = wPart.EmbeddedObjectParts.Count();

                do
                {
                    LstDisplay.Items.Add(wPart.Uri + StringResources.arrow + wPart.EmbeddedObjectParts.ElementAt(x).Uri.ToString());
                    x++;
                }
                while (x < oleEmbCount);

                return x;
            }
            catch (ArgumentOutOfRangeException)
            {
                return 0;
            }
        }

        /// <summary>
        /// Return PowerPoint Embedded Object Count
        /// </summary>
        /// <param name="wPart"></param>
        /// <returns></returns>
        public int GetEmbeddedObjectProperties(SlidePart sPart)
        {
            try
            {
                int x = 0;
                int oleEmbCount = sPart.EmbeddedObjectParts.Count();

                do
                {
                    LstDisplay.Items.Add(sPart.Uri + StringResources.arrow + sPart.EmbeddedObjectParts.ElementAt(x).Uri.ToString());
                    x++;
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
                List<string> authors = new List<string>();

                using (document = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    // get the list of authors
                    _fromAuthor = StringResources.emptyString;

                    authors = WordOpenXml.GetAllAuthors(document.MainDocumentPart.Document);

                    FrmAuthors aFrm = new Forms.FrmAuthors(TxtFileName.Text, authors)
                    {
                        Owner = this
                    };
                    aFrm.ShowDialog();
                }

                Cursor = Cursors.WaitCursor;

                if (_fromAuthor == "* No Authors *" || _fromAuthor == "")
                {
                    DisplayInformation(InformationOutput.ClearAndAdd, "** No Revisions To Accept **");
                    return;
                }

                if (_fromAuthor == "* All Authors *")
                {
                    _fromAuthor = "* All Authors *";
                }

                WordOpenXml.AcceptAllRevisions(TxtFileName.Text, _fromAuthor);
                DisplayInformation(InformationOutput.ClearAndAdd, "** Revisions Accepted **");
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                LoggingHelper.Log("BtnAcceptRevisions_Click Error");
                LoggingHelper.Log(ex.Message);
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
                WordOpenXml.RemoveComments(TxtFileName.Text);
                DisplayInformation(InformationOutput.ClearAndAdd, "** Comments Removed **");
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                LoggingHelper.Log("BtnDeleteComments_Click Error");
                LoggingHelper.Log(ex.Message);
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
                if (WordOpenXml.DeleteHiddenText(TxtFileName.Text))
                {
                    DisplayInformation(InformationOutput.ClearAndAdd, "** Hidden text deleted **");
                }
                else
                {
                    DisplayInformation(InformationOutput.ClearAndAdd, "** Document does not contain hiddent text **");
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                LoggingHelper.Log("BtnDeleteHiddenText_Click Error");
                LoggingHelper.Log(ex.Message);
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
                if (WordOpenXml.RemoveHeadersFooters(TxtFileName.Text))
                {
                    DisplayInformation(InformationOutput.ClearAndAdd, "** Headers/Footer removed **");
                }
                else
                {
                    DisplayInformation(InformationOutput.ClearAndAdd, "** Document does not contain a header or footer **");
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                LoggingHelper.Log("BtnDeleteHdrFtr_Click Error");
                LoggingHelper.Log(ex.Message);
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
                    WordOpenXml.RemoveListTemplatesNumId(TxtFileName.Text, orphanLT.ToString());
                }
                DisplayInformation(InformationOutput.ClearAndAdd, "** List Templates removed **");
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
            }
        }

        private void BtnDeleteBreaks_Click(object sender, EventArgs e)
        {
            if (WordOpenXml.RemoveBreaks(TxtFileName.Text))
            {
                DisplayInformation(InformationOutput.ClearAndAdd, "** Page and Section breaks have been removed **");
            }
            else
            {
                DisplayInformation(InformationOutput.ClearAndAdd, "** Document does not contain any page breaks **");
            }
        }

        private void BtnRemovePII_Click(object sender, EventArgs e)
        {
            if (!File.Exists(TxtFileName.Text))
            {
                DisplayInformation(InformationOutput.InvalidFile, TxtFileName.Text);
            }
            else
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    if (WordExtensionClass.HasPersonalInfo(document) == true)
                    {
                        WordExtensionClass.RemovePersonalInfo(document);
                        DisplayInformation(InformationOutput.ClearAndAdd, "PII Removed from file.");
                    }
                    else
                    {
                        DisplayInformation(InformationOutput.ClearAndAdd, "Document does not contain PII.");
                    }
                }
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

            DisplayEmptyCount(count, "errors.");
        }

        private void BtnValidateFile_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

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
                    LoggingHelper.Log("BtnValidateFileClick Error");
                    throw new Exception();
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                LoggingHelper.Log("BtnValidateFile_Click Error");
                LoggingHelper.Log(ex.Message);
            }
            finally
            {
                if (LstDisplay.Items.Count < 0)
                {
                    LstDisplay.Items.Add("** No validation errors **");
                }

                Cursor = Cursors.Default;
            }
        }

        private void BtnListFormulas_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                int count = 0;
                LstDisplay.Items.Clear();

                foreach (Worksheet sht in ExcelOpenXml.GetWorkSheets(TxtFileName.Text, false))
                {
                    foreach (var s in sht)
                    {
                        if (s.LocalName == "sheetData")
                        {
                            IEnumerable<Cell> cells = sht.WorksheetPart.Worksheet.Descendants<Cell>();
                            foreach (Cell c in cells)
                            {
                                if (c.CellFormula != null)
                                {
                                    count++;
                                    LstDisplay.Items.Add(count + StringResources.period + c.CellReference + " = " + c.CellFormula.Text);
                                }
                            }
                        }
                    }
                }

                DisplayEmptyCount(count, "formulas");
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.ClearAndAdd, ex.Message);
                LoggingHelper.Log("BtnListFormulas_Click Error");
                LoggingHelper.Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnListFonts_Click(object sender, EventArgs e)
        {
            try
            {
                LstDisplay.Items.Clear();
                int count = 0;
                using (WordprocessingDocument doc = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    foreach (DocumentFormat.OpenXml.Wordprocessing.Font ft in doc.MainDocumentPart.FontTablePart.Fonts)
                    {
                        count++;
                        LstDisplay.Items.Add(count + StringResources.period + ft.Name);
                    }
                }

                DisplayEmptyCount(count, "fonts");
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                LoggingHelper.Log("BtnListFonts_Click Error");
                LoggingHelper.Log(ex.Message);
            }
        }

        private void BtnListFootnotes_Click(object sender, EventArgs e)
        {
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    LstDisplay.Items.Clear();
                    FootnotesPart footnotePart = doc.MainDocumentPart.FootnotesPart;
                    if (footnotePart != null)
                    {
                        int count = 0;
                        foreach (Footnote fn in footnotePart.Footnotes)
                        {
                            if (fn.InnerText != StringResources.emptyString)
                            {
                                count++;
                                DisplayInformation(InformationOutput.TextOnly, count + StringResources.period + fn.InnerText);
                            }
                        }

                        DisplayEmptyCount(count, "footnotes");
                        
                    }
                    else
                    {
                        DisplayInformation(InformationOutput.TextOnly, StringResources.noFootnotes);
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.ClearAndAdd, ex.Message);
                LoggingHelper.Log("BtnListFootnotes_Click Error");
                LoggingHelper.Log(ex.Message);
            }
        }

        private void BtnListEndnotes_Click(object sender, EventArgs e)
        {
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    LstDisplay.Items.Clear();
                    EndnotesPart endnotePart = doc.MainDocumentPart.EndnotesPart;
                    if (endnotePart != null)
                    {
                        int count = 0;
                        foreach (Endnote en in endnotePart.Endnotes)
                        {
                            if (en.InnerText != StringResources.emptyString)
                            {
                                count++;
                                DisplayInformation(InformationOutput.TextOnly, count + StringResources.period + en.InnerText);
                            }
                        }

                        DisplayEmptyCount(count, "endnotes");
                    }
                    else
                    {
                        DisplayInformation(InformationOutput.TextOnly, StringResources.noEndnotes);
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.ClearAndAdd, ex.Message);
                LoggingHelper.Log("BtnListEndnotes_Click Error");
                LoggingHelper.Log(ex.Message);
            }
        }

        private void BtnDeleteFootnotes_Click(object sender, EventArgs e)
        {
            try
            {
                if (WordOpenXml.RemoveFootnotes(TxtFileName.Text))
                {
                    DisplayInformation(InformationOutput.ClearAndAdd, "** Footnotes removed from document **");
                }
                else
                {
                    DisplayInformation(InformationOutput.ClearAndAdd, "** Document does not contain footnotes **");
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.ClearAndAdd, ex.Message);
                LoggingHelper.Log("BtnDeleteFootnotes_Click Error");
                LoggingHelper.Log(ex.Message);
            }
        }

        private void BtnDeleteEndnotes_Click(object sender, EventArgs e)
        {
            try
            {
                if (WordOpenXml.RemoveEndnotes(TxtFileName.Text))
                {
                    DisplayInformation(InformationOutput.ClearAndAdd, "** Endnotes removed from document **");
                }
                else
                {
                    DisplayInformation(InformationOutput.ClearAndAdd, "** Document does not contain endnotes **");
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.ClearAndAdd, ex.Message);
                LoggingHelper.Log("BtnDeleteEndnotes_Click Error");
                LoggingHelper.Log(ex.Message);
            }
        }

        private void BtnListRevisions_Click(object sender, EventArgs e)
        {
            try
            {
                int revCount = 0;
                LstDisplay.Items.Clear();
                Cursor = Cursors.WaitCursor;

                List<string> authorList = new List<string>();
                
                using (WordprocessingDocument document = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    // if we have an author, go through all the revisions
                    authorList = WordOpenXml.GetAllAuthors(document.MainDocumentPart.Document);

                    Document doc = document.MainDocumentPart.Document;
                    var paragraphChanged = doc.Descendants<ParagraphPropertiesChange>().ToList();
                    var runChanged = doc.Descendants<RunPropertiesChange>().ToList();
                    var deleted = doc.Descendants<DeletedRun>().ToList();
                    var deletedParagraph = doc.Descendants<Deleted>().ToList();
                    var inserted = doc.Descendants<InsertedRun>().ToList();

                    // get the list of authors
                    _fromAuthor = StringResources.emptyString;

                    FrmAuthors aFrm = new Forms.FrmAuthors(TxtFileName.Text, authorList)
                    {
                        Owner = this
                    };
                    aFrm.ShowDialog();

                    if (_fromAuthor == "* All Authors *")
                    {
                        List<string> temp = new List<string>();
                        temp = WordOpenXml.GetAllAuthors(doc);
                        
                        foreach (string s in temp)
                        {
                            var tempParagraphChanged = paragraphChanged.Where(item => item.Author == s).ToList();
                            var tempRunChanged = runChanged.Where(item => item.Author == s).ToList();
                            var tempDeleted = deleted.Where(item => item.Author == s).ToList();
                            var tempInserted = inserted.Where(item => item.Author == s).ToList();
                            var tempDeletedParagraph = deletedParagraph.Where(item => item.Author == s).ToList();

                            if ((tempParagraphChanged.Count + tempRunChanged.Count + tempDeleted.Count + tempInserted.Count + tempDeletedParagraph.Count) == 0)
                            {
                                LstDisplay.Items.Add(s + " has no changes.");
                                continue;
                            }

                            foreach (var item in tempParagraphChanged)
                            {
                                revCount++;
                                LstDisplay.Items.Add(revCount + ". " + s + " : Paragraph Changed ");
                            }

                            foreach (var item in tempDeletedParagraph)
                            {
                                revCount++;
                                LstDisplay.Items.Add(revCount + ". " + s + " : Paragraph Deleted ");
                            }

                            foreach (var item in tempRunChanged)
                            {
                                revCount++;
                                LstDisplay.Items.Add(revCount + ". " + s + " :  Run Changed = " + item.InnerText);
                            }

                            foreach (var item in tempDeleted)
                            {
                                revCount++;
                                LstDisplay.Items.Add(revCount + ". " + s + " :  Deletion = " + item.InnerText);
                            }

                            foreach (var item in tempInserted)
                            {
                                if (item.Parent != null)
                                {
                                    var textRuns = item.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>().ToList();
                                    var parent = item.Parent;

                                    foreach (var textRun in textRuns)
                                    {
                                        revCount++;
                                        LstDisplay.Items.Add(revCount + ". " + s + " :  Insertion = " + textRun.InnerText);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        // list the selected authors revisions
                        if (!String.IsNullOrEmpty(_fromAuthor))
                        {
                            paragraphChanged = paragraphChanged.Where(item => item.Author == _fromAuthor).ToList();
                            runChanged = runChanged.Where(item => item.Author == _fromAuthor).ToList();
                            deleted = deleted.Where(item => item.Author == _fromAuthor).ToList();
                            inserted = inserted.Where(item => item.Author == _fromAuthor).ToList();
                            deletedParagraph = deletedParagraph.Where(item => item.Author == _fromAuthor).ToList();

                            if ((paragraphChanged.Count + runChanged.Count + deleted.Count + inserted.Count + deletedParagraph.Count) == 0)
                            {
                                if (_fromAuthor == "* No Authors *")
                                {
                                    DisplayInformation(InformationOutput.ClearAndAdd, "** There are no revisions in this document **");
                                }
                                else
                                {
                                    DisplayInformation(InformationOutput.ClearAndAdd, "** This author has no changes **");
                                }

                                return;
                            }
                        }
                        else
                        {
                            DisplayInformation(InformationOutput.ClearAndAdd, "** There are no revisions in this document **");
                            return;
                        }

                        foreach (var item in paragraphChanged)
                        {
                            revCount++;
                            LstDisplay.Items.Add(revCount + ". Paragraph Changed ");
                        }

                        foreach (var item in deletedParagraph)
                        {
                            revCount++;
                            LstDisplay.Items.Add(revCount + ". Paragraph Deleted ");
                        }

                        foreach (var item in runChanged)
                        {
                            revCount++;
                            LstDisplay.Items.Add(revCount + ". Run Changed = " + item.InnerText);
                        }

                        foreach (var item in deleted)
                        {
                            revCount++;
                            LstDisplay.Items.Add(revCount + ". Deletion = " + item.InnerText);
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
                                    LstDisplay.Items.Add(revCount + ". Insertion = " + textRun.InnerText);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                LoggingHelper.Log("BtnListRevisions_Click Error");
                LoggingHelper.Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnListAuthors_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                using (WordprocessingDocument doc = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    LstDisplay.Items.Clear();
                    int count = 0;

                    // first check the people part to get authors
                    WordprocessingPeoplePart peoplePart = doc.MainDocumentPart.WordprocessingPeoplePart;
                    if (peoplePart != null)
                    { 
                        foreach (Person person in peoplePart.People)
                        {
                            count++;
                            PresenceInfo pi = person.PresenceInfo;
                            LstDisplay.Items.Add(count + StringResources.period + person.Author);
                            LstDisplay.Items.Add("   - User Id = " + pi.UserId);
                            LstDisplay.Items.Add("   - Provider Id = " + pi.ProviderId);
                        }
                    }

                    // second, sometimes there are authors in a file but they don't exist in people.xml
                    // list those
                    List<string> tempAuthors = WordOpenXml.GetAllAuthors(doc.MainDocumentPart.Document);
                    if (tempAuthors.Count > 0)
                    {
                        foreach (string s in tempAuthors)
                        {
                            count++;
                            LstDisplay.Items.Add(count + ". User Name = " + s);
                        }
                    }

                    // log if we have no authors in the listbox
                    if (count == 0)
                    {
                        DisplayInformation(InformationOutput.TextOnly, "** There are no authors in this document **");
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                LoggingHelper.Log("BtnListAuthors_Click Error");
                LoggingHelper.Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnViewCustomDocProps_Click(object sender, EventArgs e)
        {
            try
            {
                LstDisplay.Items.Clear();
                Cursor = Cursors.WaitCursor;

                using (WordprocessingDocument doc = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    DocumentSettingsPart docSettingsPart = doc.MainDocumentPart.DocumentSettingsPart;
                    GetStandardFileProps(doc.PackageProperties);
                    GetExtendedFileProps(doc.ExtendedFilePropertiesPart);

                    try
                    {
                        if (docSettingsPart != null)
                        {
                            Settings settings = docSettingsPart.Settings;
                            foreach (var setting in settings)
                            {
                                if (setting.LocalName == "compat")
                                {
                                    LstDisplay.Items.Add(StringResources.emptyString);
                                    LstDisplay.Items.Add("---- Compatibility Settings ---- ");

                                    int settingCount = setting.Count();
                                    int settingIndex = 0;

                                    do
                                    {
                                        if (setting.ElementAt(settingIndex).LocalName != "compatSetting")
                                        {
                                            if (setting.ElementAt(0).InnerText != StringResources.emptyString)
                                            {
                                                LstDisplay.Items.Add(setting.ElementAt(0).LocalName + StringResources.colon + setting.ElementAt(0).InnerText);
                                            }
                                            settingIndex++;
                                        }
                                        else
                                        {
                                            CompatibilitySetting cs = (CompatibilitySetting)setting.ElementAt(settingIndex);
                                            if (cs.Name == "compatibilityMode")
                                            {
                                                string compatModeVersion = StringResources.emptyString;

                                                if (cs.Val == "11")
                                                {
                                                    compatModeVersion = " (Word 2003)";
                                                }
                                                else if (cs.Val == "12")
                                                {
                                                    compatModeVersion = " (Word 2007)";
                                                }
                                                else if (cs.Val == "14")
                                                {
                                                    compatModeVersion = " (Word 2010)";
                                                }
                                                else if (cs.Val == "15")
                                                {
                                                    compatModeVersion = " (Word 2013)";
                                                }
                                                else
                                                {
                                                    compatModeVersion = " (Unknown Version)";
                                                }

                                                LstDisplay.Items.Add(cs.Name + StringResources.colon + cs.Val + compatModeVersion);
                                                settingIndex++;
                                            }
                                            else
                                            {
                                                LstDisplay.Items.Add(cs.Name + StringResources.colon + cs.Val);
                                                settingIndex++;
                                            }
                                        }
                                    } while (settingIndex < settingCount);

                                    LstDisplay.Items.Add(StringResources.emptyString);
                                    LstDisplay.Items.Add("---- Settings ---- ");
                                }
                                else
                                {
                                    StringBuilder sb = new StringBuilder();
                                    XmlDocument xDoc = new XmlDocument();
                                    xDoc.LoadXml(setting.OuterXml);

                                    foreach (XmlElement xe in xDoc.ChildNodes)
                                    {
                                        sb.Clear();
                                        if (xe.Attributes.Count > 1)
                                        {
                                            sb.Append(xe.Name + StringResources.colon);
                                            foreach (XmlAttribute xa in xe.Attributes)
                                            {
                                                if (!(xa.LocalName == "w" || xa.LocalName == "m" || xa.LocalName == "w14" || xa.LocalName == "w15" || xa.LocalName == "w16"))
                                                {
                                                    if (!xa.Value.StartsWith("http"))
                                                    {
                                                        if (xa.LocalName == "val")
                                                        {
                                                            sb.Append(xa.Value);
                                                        }
                                                        else
                                                        {
                                                            sb.Append(xa.LocalName + StringResources.colon + xa.Value);
                                                        }
                                                    }
                                                }
                                            }

                                            LstDisplay.Items.Add(sb);
                                        }
                                        else
                                        {
                                            LstDisplay.Items.Add(xe.Name);
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
                        LoggingHelper.Log("BtnViewCustomDocProps (doc settings) Error");
                        LoggingHelper.Log(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                LoggingHelper.Log("BtnViewCustomDocProps_Click Error");
                LoggingHelper.Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public void GetStandardFileProps(PackageProperties props)
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
            LstDisplay.Items.Add("Category : " + props.Category);
            LstDisplay.Items.Add("ContentStatus : " + props.ContentStatus);
            LstDisplay.Items.Add("ContentType : " + props.ContentType);
            LstDisplay.Items.Add("Description : " + props.Description);
            LstDisplay.Items.Add("Language : " + props.Language);
            LstDisplay.Items.Add("Identifier : " + props.Identifier);
            LstDisplay.Items.Add("Keywords : " + props.Keywords);
            LstDisplay.Items.Add(StringResources.emptyString);
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
                LoggingHelper.Log("GetExtendedFileProps Error");
                LoggingHelper.Log(ex.Message);
            }
        }

        public void DisplayElementDetails(XmlElement elem)
        {
            if (elem.Name == StringResources.docSecurity)
            {
                switch (elem.InnerText)
                {
                    case "0":
                        LstDisplay.Items.Add(StringResources.docSecurity + StringResources.colon + "None");
                        break;

                    case "1":
                        LstDisplay.Items.Add(StringResources.docSecurity + StringResources.colon + "Password Protected");
                        break;

                    case "2":
                        LstDisplay.Items.Add(StringResources.docSecurity + StringResources.colon + "Read-Only Recommended");
                        break;

                    case "4":
                        LstDisplay.Items.Add(StringResources.docSecurity + StringResources.colon + "Read-Only Enforced");
                        break;

                    case "8":
                        LstDisplay.Items.Add(StringResources.docSecurity + StringResources.colon + "Locked For Annotation");
                        break;

                    default:
                        break;
                }
            }
            else
            {
                LstDisplay.Items.Add(elem.Name + StringResources.colonBuffer + elem.InnerText);
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmAbout frm = new FrmAbout();
            frm.ShowDialog(this);
            frm.Dispose();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                // disable buttons before each open
                DisableButtons();

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
                        DisplayInformation(InformationOutput.InvalidFile, StringResources.fileDoesNotExist);
                        return;
                    }
                    else
                    {
                        LstDisplay.Items.Clear();
                        // if the file doesn't start with PK, we can stop trying to process it
                        if (!IsZipArchiveFile(TxtFileName.Text))
                        {
                            LstDisplay.Items.Add("Unable to open file, possible causes are:");
                            LstDisplay.Items.Add("  - file corruption");
                            LstDisplay.Items.Add("  - file encrypted");
                            LstDisplay.Items.Add("  - file password protected");
                            LstDisplay.Items.Add("  - not a valid Open Xml file");
                            return;
                        }
                        else
                        {
                            // if the file does start with PK, check if it fails in the SDK
                            OpenWithSdk(TxtFileName.Text, true);
                            PopulatePackageParts();
                        }
                    }
                }
                else
                {
                    // user cancelled dialog, just return
                    return;
                }
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("File Open Error:");
                LoggingHelper.Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public void PopulatePackageParts()
        {
            _pParts.Clear();

            using (Package _package = Package.Open(TxtFileName.Text, FileMode.Open, FileAccess.Read))
            {
                foreach (PackagePart pckg in _package.GetParts())
                {
                    _pParts.Add(pckg.Uri.ToString());
                }
            }
        }

        public bool IsZipArchiveFile(string filePath)
        {
            byte[] buffer = new byte[2];
            try
            {
                // open the file and populate the first 2 bytes into the buffer
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    fs.Read(buffer, 0, buffer.Length);
                }

                // if the buffer starts with PK the file is a zip archive
                if (buffer[0].ToString() == "80" && buffer[1].ToString() == "75")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (UnauthorizedAccessException uae)
            {
                LoggingHelper.Log(uae.Message);
                return false;
            }
            catch (Exception ex)
            {
                LoggingHelper.Log(ex.Message);
                return false;
            }
        }

        /// <summary>
        /// function to open the file in the SDK
        /// if the SDK fails to open the file, it is not a valid docx
        /// </summary>
        /// <param name="file">the path to the initial fix attempt</param>
        public void OpenWithSdk(string file, bool IsFileOpen)
        {
            try
            {
                // if the file is opened by the SDK, we can proceed with opening in tool
                Cursor = Cursors.WaitCursor;

                if (IsFileOpen)
                {
                    SetUpButtons();
                }

                string body = StringResources.emptyString;

                if (fileType == StringResources.word)
                {
                    using (WordprocessingDocument document = WordprocessingDocument.Open(file, false))
                    {
                        // try to get the localname of the document.xml file, if it fails, it is not a Word file
                        body = document.MainDocumentPart.Document.LocalName;
                    }
                }
                else if (fileType == StringResources.excel)
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(file, false))
                    {
                        // try to get the localname of the workbook.xml file if it fails, its not an Excel file
                        body = document.WorkbookPart.Workbook.LocalName;
                    }
                }
                else if (fileType == StringResources.powerpoint)
                {
                    using (PresentationDocument document = PresentationDocument.Open(file, false))
                    {
                        // try to get the presentation.xml local name, if it fails it is not a PPT file
                        body = document.PresentationPart.Presentation.LocalName;
                    }
                }
                else
                {
                    // not a WD, PPT, XL file or file is corrupt
                    LstDisplay.Items.Add("Invalid File: File must be Word, PowerPoint or Excel.");
                    BtnFixCorruptDocument.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                // if the file failed to open in the sdk, it is invalid or corrupt and we need to stop opening
                DisableButtons();
                LstDisplay.Items.Add("Invalid File: Error opening file.");
                LoggingHelper.Log("OpenWithSDK Error: " + ex.Message);
                BtnFixCorruptDocument.Enabled = true;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnPPTListHyperlinks_Click(object sender, EventArgs e)
        {
            try
            {
                LstDisplay.Items.Clear();

                // Open the presentation file as read-only.
                using (PresentationDocument document = PresentationDocument.Open(TxtFileName.Text, false))
                {
                    int linkCount = 0;
                    foreach (string s in PowerPointOpenXml.GetAllExternalHyperlinksInPresentation(TxtFileName.Text))
                    {
                        linkCount++;
                        LstDisplay.Items.Add(linkCount + StringResources.period + s);
                    }

                    DisplayEmptyCount(linkCount, "hyperlinks");
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                LoggingHelper.Log("BtnPPTListHyperlinks_Click Error");
                LoggingHelper.Log(ex.Message);
            }
        }

        private void BtnPPTGetAllSlideTitles_Click(object sender, EventArgs e)
        {
            try
            {
                LstDisplay.Items.Clear();

                // Open the presentation as read-only.
                using (PresentationDocument presentationDocument = PresentationDocument.Open(TxtFileName.Text, false))
                {
                    int slideCount = 0;

                    foreach (string s in PowerPointOpenXml.GetSlideTitles(presentationDocument))
                    {
                        slideCount++;
                        LstDisplay.Items.Add(slideCount + StringResources.period + s);
                    }

                    DisplayEmptyCount(slideCount, "slides");
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                LoggingHelper.Log("BtnGetAllSlideTitles_Click Error");
                LoggingHelper.Log(ex.Message);
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

                if (_FindText == StringResources.emptyString && _ReplaceText == StringResources.emptyString)
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
                LoggingHelper.Log("BtnSearchAndReplace_Click Error");
                LoggingHelper.Log(ex.Message);
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
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, false))
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
                        LstDisplay.Items.Add(ExtRelCount + StringResources.period + extWbPart.ExternalRelationships.ElementAt(0).Uri);
                    }
                }
            }
            catch (Exception ex)
            {
                // log the error
                LstDisplay.Items.Add(StringResources.errorText + ex.Message);
                LoggingHelper.Log("BtnListLinks_Click Error");
                LoggingHelper.Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnDeleteExternalLinks_Click(object sender, EventArgs e)
        {
            ExcelOpenXml.RemoveExternalLinks(TxtFileName.Text);
            LstDisplay.Items.Clear();
            LstDisplay.Items.Add("** External References Deleted **");
        }

        private void BtnListDefinedNames_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                LstDisplay.Items.Clear();
                int nameCount = 0;

                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, false))
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
                            LstDisplay.Items.Add(nameCount + StringResources.period + dn.Name.Value + " = " + dn.Text);
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
                LoggingHelper.Log("BtnListDefinedNames_Click Error");
                LoggingHelper.Log(ex.Message);
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

                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, false))
                {
                    WorkbookPart wbPart = excelDoc.WorkbookPart;
                    Sheets theSheets = wbPart.Workbook.Sheets;

                    foreach (Sheet sheet in theSheets)
                    {
                        LstDisplay.Items.Add("Worksheet Name = " + sheet.Name);
                        Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().Where((s) => s.Name == sheet.Name).FirstOrDefault();

                        if (theSheet == null)
                        {
                            LoggingHelper.Log("BtnListHiddenRowsColumnClickError" + sheet.Name);
                            throw new ArgumentException("sheetName");
                        }
                        else
                        {
                            // The sheet does exist.
                            WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
                            Worksheet ws = wsPart.Worksheet;
                            int rowCount = 0;
                            int colCount = 0;

                            LstDisplay.Items.Add("##    ROWS    ##");
                            IEnumerable<Row> rows = ws.Descendants<Row>().Where((r) => r.Hidden != null && r.Hidden.Value);
                            foreach (Row row in rows)
                            {
                                rowCount++;
                                LstDisplay.Items.Add(rowCount + StringResources.period + row.InnerText);
                            }

                            if (rowCount == 0)
                            {
                                LstDisplay.Items.Add("    None");
                            }

                            LstDisplay.Items.Add("##    COLUMNS    ##");
                            IEnumerable<Column> cols = ws.Descendants<Column>().Where((c) => c.Hidden != null && c.Hidden.Value);
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
                        LstDisplay.Items.Add(StringResources.emptyString);
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.TextOnly, ex.Message);
                LoggingHelper.Log("BtnListHiddenRowsColumns_Click Error");
                LoggingHelper.Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void DisplayEmptyCount(int count, string input)
        {
            if (count == 0)
            {
                LstDisplay.Items.Add("** Document contains no " + input + " **");
            }
        }

        private void BtnListWorksheets_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                LstDisplay.Items.Clear();
                int sheetCount = 0;

                LstDisplay.Items.Add("## VISIBLE WORKSHEETS ##");
                foreach (Sheet sht in ExcelOpenXml.GetSheets(TxtFileName.Text, false))
                {
                    sheetCount++;
                    LstDisplay.Items.Add(sheetCount + StringResources.period + sht.Name);
                }

                DisplayEmptyCount(sheetCount, "worksheets");

                LstDisplay.Items.Add("");
                LstDisplay.Items.Add("## HIDDEN WORKSHEETS ##");
                sheetCount = 0;
                foreach (Sheet sht in ExcelOpenXml.GetHiddenSheets(TxtFileName.Text, false))
                {
                    sheetCount++;
                    LstDisplay.Items.Add(sheetCount + StringResources.period + sht.Name);
                }

                DisplayEmptyCount(sheetCount, "hidden worksheets");

            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.TextOnly, ex.Message);
                LoggingHelper.Log("BtnListWorksheets_Click Error");
                LoggingHelper.Log(ex.Message);
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

                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, false))
                {
                    WorkbookPart wbPart = excelDoc.WorkbookPart;
                    SharedStringTable sst = wbPart.SharedStringTablePart.SharedStringTable;
                    LstDisplay.Items.Add("SharedString Count = " + sst.Count());
                    LstDisplay.Items.Add("Unique Count = " + sst.UniqueCount);
                    LstDisplay.Items.Add(StringResources.emptyString);

                    foreach (SharedStringItem ssi in sst)
                    {
                        sharedStringCount++;
                        DocumentFormat.OpenXml.Spreadsheet.Text ssValue = ssi.Text;
                        LstDisplay.Items.Add(sharedStringCount + StringResources.period + ssValue.Text);
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.TextOnly, "** Document does not contain any shared strings **");
                LoggingHelper.Log("BtnListSharedStrings_Click Error");
                LoggingHelper.Log(ex.Message);
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
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, false))
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
                            LstDisplay.Items.Add(commentCount + StringResources.period + cText.InnerText);
                            commentCount++;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("Excel - BtnComments_Click Error:");
                LoggingHelper.Log(ex.Message);
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
                LoggingHelper.Log("BtnListFormulas_Click Error");
                LoggingHelper.Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnChangeTheme_Click(object sender, EventArgs e)
        {
            string sThemeFilePath = StringResources.emptyString;

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
                    DisplayInformation(InformationOutput.InvalidFile, StringResources.fileDoesNotExist);
                    return;
                }
                else
                {
                    if (fileType == StringResources.word)
                    {
                        // call the replace function using the theme file provided
                        OfficeHelpers.ReplaceTheme(TxtFileName.Text, sThemeFilePath, fileType);
                        DisplayInformation(InformationOutput.ClearAndAdd, StringResources.themeFileAdded);
                    }
                    else if (fileType == StringResources.excel)
                    {
                        OfficeHelpers.ReplaceTheme(TxtFileName.Text, sThemeFilePath, fileType);
                        DisplayInformation(InformationOutput.ClearAndAdd, StringResources.themeFileAdded);
                    }
                    else if (fileType == StringResources.powerpoint)
                    {
                        OfficeHelpers.ReplaceTheme(TxtFileName.Text, sThemeFilePath, fileType);
                        DisplayInformation(InformationOutput.ClearAndAdd, StringResources.themeFileAdded);
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
                        if (sCPart == null)
                        {
                            DisplayInformation(InformationOutput.ClearAndAdd, "** File does not have any comments **");
                            return;
                        }

                        foreach (DocumentFormat.OpenXml.Presentation.Comment cmt in sCPart.CommentList)
                        {
                            commentCount++;
                            LstDisplay.Items.Add(commentCount + StringResources.period + cmt.InnerText);
                        }
                    }

                    DisplayEmptyCount(commentCount, "comments");
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.InvalidFile, ex.Message);
                LoggingHelper.Log("PPT - BtnListComments_Click Error");
                LoggingHelper.Log(ex.Message);
            }
        }

        private void BtnListWSInfo_Click(object sender, EventArgs e)
        {
            GetSheetInfo(TxtFileName.Text);
        }

        public void GetSheetInfo(string fileName)
        {
            // Open file as read-only.
            using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(fileName, false))
            {
                Sheets sheets = mySpreadsheet.WorkbookPart.Workbook.Sheets;
                LstDisplay.Items.Clear();

                // For each sheet, display the sheet information.
                foreach (Sheet sheet in sheets)
                {
                    foreach (OpenXmlAttribute attr in sheet.GetAttributes())
                    {
                        LstDisplay.Items.Add(attr.LocalName + StringResources.colonBuffer + attr.Value);
                    }
                }
            }
        }

        private void BtnListCellValuesDOM_Click(object sender, EventArgs e)
        {
            List<string> list = ExcelOpenXml.ReadExcelFileDOM(TxtFileName.Text);
            LstDisplay.Items.Clear();
            foreach (object o in list)
            {
                LstDisplay.Items.Add(o.ToString());
            }
        }

        private void BtnListCellValuesSAX_Click(object sender, EventArgs e)
        {
            List<string> list = ExcelOpenXml.ReadExcelFileSAX(TxtFileName.Text);
            LstDisplay.Items.Clear();
            foreach (object o in list)
            {
                LstDisplay.Items.Add(o.ToString());
            }
        }

        private void BtnConvertDocmToDocx_Click(object sender, EventArgs e)
        {
            ConvertToNonMacro(StringResources.word);
        }

        public void ConvertToNonMacro(string app)
        {
            try
            {
                DialogResult dr = MessageBox.Show("This will delete the original .docm and replace it with a .docx file!\r\n\r\nAre you sure you would like continue?", "Convert .Docm to .Docx", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if (DialogResult.Yes == dr)
                {
                    LstDisplay.Items.Add("Converted file location = " + OfficeHelpers.ConvertMacroEnabled2NonMacroEnabled(TxtFileName.Text, app));
                }
                else
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                LstDisplay.Items.Add("Unable to convert document.");
                LoggingHelper.Log("BtnConvertDocmToDocx Error:");
                LoggingHelper.Log(ex.Message);
            }
        }

        private void BtnListSlideText_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                LstDisplay.Items.Clear();

                int sCount = PowerPointOpenXml.CountSlides(TxtFileName.Text);
                if (sCount > 0)
                {
                    int count = 0;

                    do
                    {
                        PowerPointOpenXml.GetSlideIdAndText(out string sldText, TxtFileName.Text, count);
                        LstDisplay.Items.Add("Slide " + (count + 1) + StringResources.period + sldText);
                        count++;
                    } while (count < sCount);
                }
                else
                {
                    LstDisplay.Items.Add("Presentation contains no slides.");
                }
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("BtnListSlideText Error:");
                LoggingHelper.Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnFixCorruptDocument_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                StrOrigFileName = TxtFileName.Text;
                StrDestPath = Path.GetDirectoryName(StrOrigFileName) + "\\";
                StrExtension = Path.GetExtension(StrOrigFileName);
                StrDestFileName = Path.GetFileNameWithoutExtension(StrOrigFileName) + "(Fixed)" + StrExtension;

                // check if file we are about to copy exists and append a number so its unique
                if (File.Exists(StrDestFileName))
                {
                    Random rNumber = new Random();
                    StrDestFileName = StrDestPath + Path.GetFileNameWithoutExtension(StrOrigFileName) + "(Fixed)" + rNumber.Next(1, 100) + StrExtension;
                }

                LstDisplay.Items.Clear();

                if (StrExtension == ".docx")
                {
                    if ((File.GetAttributes(StrOrigFileName) & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                    {
                        LstDisplay.Items.Add("ERROR: File is Read-Only.");
                        return;
                    }
                    else
                    {
                        File.Copy(StrOrigFileName, StrDestFileName);
                    }
                }

                using (Package package = Package.Open(StrDestFileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    foreach (PackagePart part in package.GetParts())
                    {
                        if (part.Uri.ToString() == "/word/document.xml")
                        {
                            fileType = StringResources.word;
                            XmlDocument xdoc = new XmlDocument();
                            try
                            {
                                xdoc.Load(part.GetStream(FileMode.Open, FileAccess.Read));
                            }
                            catch (XmlException) // invalid xml found, try to fix the contents
                            {
                                MemoryStream ms = new MemoryStream();
                                InvalidXmlTags invalid = new InvalidXmlTags();

                                using (TextWriter tw = new StreamWriter(ms))
                                {
                                    using (TextReader tr = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read)))
                                    {
                                        string strDocText = tr.ReadToEnd();

                                        foreach (string el in invalid.InvalidTags())
                                        {
                                            foreach (Match m in Regex.Matches(strDocText, el))
                                            {
                                                switch (m.Value)
                                                {
                                                    case ValidXmlTags.StrValidMcChoice1:
                                                        break;

                                                    case ValidXmlTags.StrValidMcChoice2:
                                                        break;

                                                    case ValidXmlTags.StrValidMcChoice3:
                                                        break;

                                                    case InvalidXmlTags.StrInvalidVshape:
                                                        // the original strvalidvshape fixes most corruptions, but there are
                                                        // some that are within a group so I added this for those rare situations
                                                        // where the v:group closing tag needs to be included
                                                        if (Properties.Settings.Default.FixGroupedShapes == "true")
                                                        {
                                                            strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidVshapegroup);
                                                            LstDisplay.Items.Add(StringResources.invalidTag + m.Value);
                                                            LstDisplay.Items.Add(StringResources.replacedWith + ValidXmlTags.StrValidVshapegroup);
                                                        }
                                                        else
                                                        {
                                                            strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidVshape);
                                                            LstDisplay.Items.Add(StringResources.invalidTag + m.Value);
                                                            LstDisplay.Items.Add(StringResources.replacedWith + ValidXmlTags.StrValidVshape);
                                                        }
                                                        break;

                                                    case InvalidXmlTags.StrInvalidOmathWps:
                                                        strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidomathwps);
                                                        LstDisplay.Items.Add(StringResources.invalidTag + m.Value);
                                                        LstDisplay.Items.Add(StringResources.replacedWith + ValidXmlTags.StrValidomathwps);
                                                        break;

                                                    case InvalidXmlTags.StrInvalidOmathWpg:
                                                        strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidomathwpg);
                                                        LstDisplay.Items.Add(StringResources.invalidTag + m.Value);
                                                        LstDisplay.Items.Add(StringResources.replacedWith + ValidXmlTags.StrValidomathwpg);
                                                        break;

                                                    case InvalidXmlTags.StrInvalidOmathWpc:
                                                        strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidomathwpc);
                                                        LstDisplay.Items.Add(StringResources.invalidTag + m.Value);
                                                        LstDisplay.Items.Add(StringResources.replacedWith + ValidXmlTags.StrValidomathwpc);
                                                        break;

                                                    case InvalidXmlTags.StrInvalidOmathWpi:
                                                        strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidomathwpi);
                                                        LstDisplay.Items.Add(StringResources.invalidTag + m.Value);
                                                        LstDisplay.Items.Add(StringResources.replacedWith + ValidXmlTags.StrValidomathwpi);
                                                        break;

                                                    default:
                                                        // default catch for "strInvalidmcChoiceRegEx" and "strInvalidFallbackRegEx"
                                                        // since the exact string will never be the same and always has different trailing tags
                                                        // we need to conditionally check for specific patterns
                                                        // the first if </mc:Choice> is to catch and replace the invalid mc:Choice tags
                                                        if (m.Value.Contains("</mc:Choice>"))
                                                        {
                                                            if (m.Value.Contains("<mc:Fallback id="))
                                                            {
                                                                // secondary check for a fallback that has an attribute.
                                                                // we don't allow attributes in a fallback
                                                                strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidMcChoice4);
                                                                LstDisplay.Items.Add(StringResources.invalidTag + m.Value);
                                                                LstDisplay.Items.Add(StringResources.replacedWith + ValidXmlTags.StrValidMcChoice4);
                                                                break;
                                                            }

                                                            // replace mc:choice and hold onto the tag that follows
                                                            strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidMcChoice3 + m.Groups[2].Value);
                                                            LstDisplay.Items.Add(StringResources.invalidTag + m.Value);
                                                            LstDisplay.Items.Add(StringResources.replacedWith + ValidXmlTags.StrValidMcChoice3 + m.Groups[2].Value);
                                                            break;
                                                        }
                                                        // the second if <w:pict/> is to catch and replace the invalid mc:Fallback tags
                                                        else if (m.Value.Contains("<w:pict/>"))
                                                        {
                                                            if (m.Value.Contains("</mc:Fallback>"))
                                                            {
                                                                // if the match contains the closing fallback we just need to remove the entire fallback
                                                                // this will leave the closing AC and Run tags, which should be correct
                                                                strDocText = strDocText.Replace(m.Value, StringResources.emptyString);
                                                                LstDisplay.Items.Add(StringResources.invalidTag + m.Value);
                                                                LstDisplay.Items.Add(StringResources.replacedWith + "Fallback tag deleted.");
                                                                break;
                                                            }

                                                            // if there is no closing fallback tag, we can replace the match with the omitFallback valid tags
                                                            // then we need to also add the trailing tag, since it's always different but needs to stay in the file
                                                            strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrOmitFallback + m.Groups[2].Value);
                                                            LstDisplay.Items.Add(StringResources.invalidTag + m.Value);
                                                            LstDisplay.Items.Add(StringResources.replacedWith + ValidXmlTags.StrOmitFallback + m.Groups[2].Value);
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            // leaving this open for future checks
                                                            break;
                                                        }
                                                }
                                            }
                                        }

                                        // remove all fallback tags is a 3 step process
                                        // Step 1. start by getting a list of all nodes/values in the document.xml file
                                        // Step 2. call GetAllNodes to add each fallback tag
                                        // Step 3. call ParseOutFallbackTags to remove each fallback
                                        if (Properties.Settings.Default.RemoveFallback == "true")
                                        {
                                            CharEnumerator charEnum = strDocText.GetEnumerator();
                                            while (charEnum.MoveNext())
                                            {
                                                // keep track of previous char
                                                PrevChar = charEnum.Current;

                                                // opening tag
                                                switch (charEnum.Current)
                                                {
                                                    case '<':
                                                        // if we haven't hit a close, but hit another '<' char
                                                        // we are not a true open tag so add it like a regular char
                                                        if (_sbNodeBuffer.Length > 0)
                                                        {
                                                            _nodes.Add(_sbNodeBuffer.ToString());
                                                            _sbNodeBuffer.Clear();
                                                        }
                                                        Node(charEnum.Current);
                                                        break;

                                                    case '>':
                                                        // there are 2 ways to close out a tag
                                                        // 1. self contained tag like <w:sz w:val="28"/>
                                                        // 2. standard xml <w:t>test</w:t>
                                                        // if previous char is '/', then we are an end tag
                                                        if (PrevChar == '/' || IsRegularXmlTag)
                                                        {
                                                            Node(charEnum.Current);
                                                            IsRegularXmlTag = false;
                                                        }
                                                        Node(charEnum.Current);
                                                        _nodes.Add(_sbNodeBuffer.ToString());
                                                        _sbNodeBuffer.Clear();
                                                        break;

                                                    default:
                                                        // this is the second xml closing style, keep track of char
                                                        if (PrevChar == '<' && charEnum.Current == '/')
                                                        {
                                                            IsRegularXmlTag = true;
                                                        }
                                                        Node(charEnum.Current);
                                                        break;
                                                }

                                                // cleanup
                                                charEnum.Dispose();
                                            }

                                            LstDisplay.Items.Add("...removing all fallback tags");
                                            GetAllNodes(strDocText);
                                            strDocText = FixedFallback;
                                        }

                                        tw.Write(strDocText);
                                        tw.Flush();

                                        // rewrite the part
                                        ms.Position = 0;
                                        Stream partStream = part.GetStream(FileMode.Open, FileAccess.Write);
                                        partStream.SetLength(0);
                                        ms.WriteTo(partStream);

                                        LstDisplay.Items.Add("-------------------------------------------------------------");
                                        LstDisplay.Items.Add("Fixed Document Location: " + StrDestFileName);
                                        IsFixed = true;

                                        // open the file in Word
                                        if (Properties.Settings.Default.OpenInWord == "true")
                                        {
                                            Process.Start(StrDestFileName);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (IsFixed == false)
                    {
                        LstDisplay.Items.Add("This document does not contain invalid xml.");
                    }
                }
            }
            catch (IOException ioe)
            {
                LstDisplay.Items.Add(StringResources.errorUnableToFixDocument);
                LoggingHelper.Log("Corrupt Doc IO Exception = " + ioe.Message);
            }
            catch (FileFormatException ffe)
            {
                // list out the possible reasons for this type of exception
                LstDisplay.Items.Add(StringResources.errorUnableToFixDocument);
                LstDisplay.Items.Add("   Possible Causes:");
                LstDisplay.Items.Add("      - File may be password protected");
                LstDisplay.Items.Add("      - File was renamed to the .docx extension, but is not an actual .docx file");
                LstDisplay.Items.Add("      - Error = " + ffe.Message);
            }
            catch (Exception ex)
            {
                LstDisplay.Items.Add(StringResources.errorUnableToFixDocument + ex.Message);
                LoggingHelper.Log("Corrupt Doc Exception = " + ex.Message);
            }
            finally
            {
                // only delete destination file when there is an error
                // need to make sure the file stays when it is fixed
                if (IsFixed == false)
                {
                    // delete the copied file if it exists
                    if (File.Exists(StrDestFileName))
                    {
                        File.Delete(StrDestFileName);
                    }
                }
                else
                {
                    // since we were able to attempt the fixes
                    // check if we can open in the sdk and confirm it was indeed fixed
                    LstDisplay.Items.Add(StringResources.emptyString);
                    OpenWithSdk(StrDestFileName, false);
                }

                // need to reset the globals
                IsFixed = false;
                IsRegularXmlTag = false;
                FixedFallback = string.Empty;
                StrOrigFileName = string.Empty;
                StrDestPath = string.Empty;
                StrExtension = string.Empty;
                StrDestFileName = string.Empty;
                PrevChar = '<';

                Cursor = Cursors.Default;
            }
        }

        public static void Node(char input)
        {
            _sbNodeBuffer.Append(input);
        }

        /// <summary>
        /// this function loops through all nodes parsed out from Step 1
        /// check each node and add fallback tags only to the list
        /// </summary>
        /// <param name="originalText"></param>
        public static void GetAllNodes(string originalText)
        {
            bool isFallback = false;
            var fallback = new List<string>();

            foreach (string o in _nodes)
            {
                if (o == StringResources.txtFallbackStart)
                {
                    isFallback = true;
                }

                if (isFallback)
                {
                    fallback.Add(o);
                }

                if (o == StringResources.txtFallbackEnd)
                {
                    isFallback = false;
                }
            }

            ParseOutFallbackTags(fallback, originalText);
        }

        /// <summary>
        /// we should only have a list of fallback start tags, end tags and each tag in between
        /// the idea is to combine these start/middle/end tags into a long string
        /// then they can be replaced with an empty string
        /// </summary>
        /// <param name="input"></param>
        /// <param name="originalText"></param>
        public static void ParseOutFallbackTags(List<string> input, string originalText)
        {
            var fallbackTagsAppended = new List<string>();
            StringBuilder sbFallback = new StringBuilder();

            foreach (object o in input)
            {
                switch (o.ToString())
                {
                    case StringResources.txtFallbackStart:
                        sbFallback.Append(o);
                        continue;
                    case StringResources.txtFallbackEnd:
                        sbFallback.Append(o);
                        fallbackTagsAppended.Add(sbFallback.ToString());
                        sbFallback.Clear();
                        continue;
                    default:
                        sbFallback.Append(o);
                        continue;
                }
            }

            sbFallback.Clear();

            // loop each item in the list and remove it from the document
            originalText = fallbackTagsAppended.Aggregate(originalText, (current, o) => current.Replace(o.ToString(), StringResources.emptyString));

            // each set of fallback tags should now be removed from the text
            // set it to the global variable so we can add it back into document.xml
            FixedFallback = originalText;
        }

        private void BtnListConnections_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                LstDisplay.Items.Clear();

                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, false))
                {
                    WorkbookPart wbPart = excelDoc.WorkbookPart;
                    ConnectionsPart cPart = wbPart.ConnectionsPart;

                    if (cPart == null)
                    {
                        LstDisplay.Items.Add("** There are no connections in this file **");
                        return;
                    }

                    int cCount = 0;

                    foreach (Connection c in cPart.Connections)
                    {
                        cCount++;
                        if (c.DatabaseProperties.Connection != null)
                        {
                            string cn = c.DatabaseProperties.Connection;
                            string[] cArray = cn.Split(';');

                            LstDisplay.Items.Add(cCount + ". Connection= " + c.Name);
                            foreach (var s in cArray)
                            {
                                LstDisplay.Items.Add("    " + s);
                            }

                            if (c.ConnectionFile != null && c.OlapProperties.RowDrillCount != null)
                            {
                                LstDisplay.Items.Add(StringResources.emptyString);
                                LstDisplay.Items.Add("    Connection File= " + c.ConnectionFile);
                                LstDisplay.Items.Add("    Row Drill Count= " + c.OlapProperties.RowDrillCount);
                            }
                        }
                        else
                        {
                            LstDisplay.Items.Add("Invalid connections.xml");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.TextOnly, ex.Message);
                LoggingHelper.Log("List Connections Failed = " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnListCustomProps_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                LstDisplay.Items.Clear();

                if (fileType == StringResources.word)
                {
                    using (WordprocessingDocument document = WordprocessingDocument.Open(TxtFileName.Text, false))
                    {
                        AddCustomDocPropsToList(document.CustomFilePropertiesPart);
                    }
                }
                else if (fileType == StringResources.excel)
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(TxtFileName.Text, false))
                    {
                        AddCustomDocPropsToList(document.CustomFilePropertiesPart);
                    }
                }
                else if (fileType == StringResources.powerpoint)
                {
                    using (PresentationDocument document = PresentationDocument.Open(TxtFileName.Text, false))
                    {
                        AddCustomDocPropsToList(document.CustomFilePropertiesPart);
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
                LstDisplay.Items.Add(StringResources.noCustomDocProps);
            }
            catch (Exception ex)
            {
                LstDisplay.Items.Add("BtnListCustomProps Error: " + ex.Message);
                LoggingHelper.Log("BtnListCustomProps Error: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public void AddCustomDocPropsToList(CustomFilePropertiesPart cfp)
        {
            if (cfp == null)
            {
                LstDisplay.Items.Add(StringResources.noCustomDocProps);
                return;
            }

            int count = 0;
            
            foreach (var v in CfpList(cfp))
            {
                count++;
                LstDisplay.Items.Add(count + StringResources.period + v);
            }

            DisplayEmptyCount(count, "custom document properties");
        }

        public List<string> CfpList(CustomFilePropertiesPart part)
        {
            List<string> val = new List<string>();
            foreach (CustomDocumentProperty cdp in part.RootElement)
            {
                val.Add(cdp.Name + StringResources.colonBuffer + cdp.InnerText);
            }
            return val;
        }

        private void BtnSetCustomProps_Click(object sender, EventArgs e)
        {
            FrmCustomProperties cFrm = new FrmCustomProperties(TxtFileName.Text, fileType)
            {
                Owner = this
            };
            cFrm.ShowDialog();
        }

        private void BtnSetPrintOrientation_Click(object sender, EventArgs e)
        {
            FrmPrintOrientation pFrm = new FrmPrintOrientation(TxtFileName.Text)
            {
                Owner = this
            };
            pFrm.ShowDialog();
        }

        private void copyOutputToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyAllItems();
        }

        private void BtnViewParagraphs_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            FrmParagraphs paraFrm = new FrmParagraphs(TxtFileName.Text)
            {
                Owner = this
            };
            paraFrm.ShowDialog();
            Cursor = Cursors.Default;
        }

        private void BtnConvertXlsm2Xlsx_Click(object sender, EventArgs e)
        {
            ConvertToNonMacro(StringResources.excel);
        }

        private void BtnConvertPptmToPptx_Click(object sender, EventArgs e)
        {
            ConvertToNonMacro(StringResources.powerpoint);
        }

        private void BtnListPackageParts_Click(object sender, EventArgs e)
        {
            LstDisplay.Items.Clear();

            foreach (var o in _pParts)
            {
                LstDisplay.Items.Add(o);
            }
        }
        
        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmSettings form = new FrmSettings();
            form.Show();
        }

        private void errorLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmErrorLog errFrm = new FrmErrorLog()
            {
                Owner = this
            };
            errFrm.ShowDialog();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();
            Application.Exit();
        }

        private void BtnListFieldCodes_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                using (WordprocessingDocument package = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    IEnumerable<Run> rList = package.MainDocumentPart.Document.Descendants<Run>();
                    LstDisplay.Items.Clear();
                    List<string> fieldCodeList = new List<string>();
                    
                    foreach (Run r in rList)
                    {
                        foreach (OpenXmlElement oxe in r.ChildElements)
                        {
                            if (oxe.LocalName == "fldChar")
                            {
                                FieldChar fc = new FieldChar();
                                fc = (FieldChar)oxe;
                                if (fc.FieldCharType == StringResources.sBegin)
                                {
                                    fieldCodeList.Add(StringResources.sBegin);
                                }
                                else if (fc.FieldCharType == StringResources.sEnd)
                                {
                                    fieldCodeList.Add(StringResources.sEnd);
                                }
                            }
                            else if (oxe.LocalName == "instrText")
                            {
                                fieldCodeList.Add(oxe.InnerText);
                            }
                        }
                    }

                    if (fieldCodeList.Count == 0)
                    {
                        LstDisplay.Items.Add("** Document does not contain any field codes **");
                        return;
                    }
                    else
                    {
                        StringBuilder sb = new StringBuilder();
                        int fCount = 0;

                        foreach (string s in fieldCodeList)
                        {
                            if (s == StringResources.sBegin)
                            {
                                continue;
                            }
                            else if (s == StringResources.sEnd)
                            {
                                // display the field code values
                                fCount++;
                                LstDisplay.Items.Add(fCount + ". " + sb);
                                sb.Clear();
                            }
                            else
                            {
                                sb.Append(s);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LstDisplay.Items.Add(StringResources.errorText + ex.Message);
                LoggingHelper.Log("BtnListFieldCodes: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnListBookmarks_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                using (WordprocessingDocument package = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    IEnumerable<BookmarkStart> bkList = package.MainDocumentPart.Document.Descendants<BookmarkStart>();
                    LstDisplay.Items.Clear();

                    if (bkList.Count() > 0)
                    {
                        int count = 1;

                        foreach (BookmarkStart bk in bkList)
                        {
                            LstDisplay.Items.Add(count + ". " + bk.Name);
                            count++;
                        }
                    }
                    else
                    {
                        LstDisplay.Items.Add("** Document does not contain any bookmarks **");
                    }
                }
            }
            catch (Exception ex)
            {
                LstDisplay.Items.Add(StringResources.errorText + ex.Message);
                LoggingHelper.Log("BtnListBookmarks: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnListCC_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                using (WordprocessingDocument package = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    LstDisplay.Items.Clear();
                    int count = 0;

                    foreach (var cc in package.ContentControls())
                    {
                        string ccType = StringResources.emptyString;
                        string ccVal = StringResources.emptyString;
                        string dropVal = StringResources.emptyString;
                        bool isDropDown = false;

                        SdtProperties props = cc.Elements<SdtProperties>().FirstOrDefault();

                        // loop the properties and get the type
                        foreach (OpenXmlElement oxe in props.ChildElements)
                        {
                            if (oxe.GetType().Name == "Tag")
                            {
                                Tag tag = props.Elements<Tag>().FirstOrDefault();
                                ccType = "Rich / Plain Text";
                                ccVal = tag.Val;
                            }

                            if (oxe.GetType().Name == "SdtContentDropDownList")
                            {
                                Tag tag = props.Elements<Tag>().FirstOrDefault();
                                dropVal = tag.Val;
                                isDropDown = true;
                            }

                            if (oxe.GetType().Name == "SdtContentDocPartList")
                            {
                                SdtContentCheckBox sccb = props.Elements<SdtContentCheckBox>().FirstOrDefault();
                                ccType = "Building Block Gallery";
                            }

                            if (oxe.GetType().Name == "SdtContentCheckBox")
                            {
                                SdtContentCheckBox sccb = props.Elements<SdtContentCheckBox>().FirstOrDefault();
                                ccType = "Check Box";
                            }

                            if (oxe.GetType().Name == "SdtContentPicture")
                            {
                                SdtContentCheckBox sccb = props.Elements<SdtContentCheckBox>().FirstOrDefault();
                                ccType = "Picture";
                            }

                            if (oxe.GetType().Name == "SdtContentComboBox")
                            {
                                SdtContentCheckBox sccb = props.Elements<SdtContentCheckBox>().FirstOrDefault();
                                ccType = "Combo Box";
                            }

                            if (oxe.GetType().Name == "SdtContentDate")
                            {
                                SdtContentCheckBox sccb = props.Elements<SdtContentCheckBox>().FirstOrDefault();
                                ccType = "Date Picker";
                            }
                        }

                        // display the cc type
                        count++;
                        if (isDropDown)
                        {
                            LstDisplay.Items.Add(count + ". " + "Drop Down List: ");
                            continue;
                        }
                        else
                        {
                            LstDisplay.Items.Add(count + ". " + ccType);
                        }
                    }

                    DisplayEmptyCount(count, "content controls.");
                }
            }
            catch (Exception ex)
            {
                LstDisplay.Items.Add(StringResources.errorText + ex.Message);
                LoggingHelper.Log("BtnListCC: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnListShapes_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                LstDisplay.Items.Clear();
                int count = 0;

                if (fileType == StringResources.word)
                {
                    // with Word, we can just run through the entire body and get the shapes
                    using (WordprocessingDocument document = WordprocessingDocument.Open(TxtFileName.Text, false))
                    {
                        foreach (ChartPart c in document.MainDocumentPart.ChartParts)
                        {
                            count++;
                            LstDisplay.Items.Add(count + StringResources.period + c.Uri + StringResources.arrow + StringResources.shpChart);
                        }

                        foreach (AO.Shape shape in document.MainDocumentPart.Document.Body.Descendants<AO.Shape>())
                        {
                            count++;
                            LstDisplay.Items.Add(count + StringResources.shpOfficeDrawing);
                        }

                        foreach (O.Vml.Shape shape in document.MainDocumentPart.Document.Body.Descendants<O.Vml.Shape>())
                        {
                            count++;
                            LstDisplay.Items.Add(count + StringResources.period + shape.Id + StringResources.arrow + StringResources.shpVml);
                        }

                        foreach (O.Math.Shape shape in document.MainDocumentPart.Document.Body.Descendants<O.Math.Shape>())
                        {
                            count++;
                            LstDisplay.Items.Add(count + StringResources.shpMath);
                        }

                        foreach (A.Diagrams.Shape shape in document.MainDocumentPart.Document.Body.Descendants<A.Diagrams.Shape>())
                        {
                            count++;
                            LstDisplay.Items.Add(count + StringResources.shpDrawingDgm);
                        }

                        foreach (A.ChartDrawing.Shape shape in document.MainDocumentPart.Document.Body.Descendants<A.ChartDrawing.Shape>())
                        {
                            count++;
                            LstDisplay.Items.Add(count + StringResources.shpDrawingDgm);
                        }

                        foreach (A.Charts.Shape shape in document.MainDocumentPart.Document.Body.Descendants<A.Charts.Shape>())
                        {
                            count++;
                            LstDisplay.Items.Add(count + StringResources.shpChartShape);
                        }

                        foreach (A.Shape shape in document.MainDocumentPart.Document.Body.Descendants<A.Shape>())
                        {
                            count++;
                            LstDisplay.Items.Add(count + StringResources.shpShape);
                        }

                        foreach (A.Diagrams.Shape3D shape in document.MainDocumentPart.Document.Body.Descendants<A.Diagrams.Shape3D>())
                        {
                            count++;
                            LstDisplay.Items.Add(count + StringResources.shp3D);
                        }

                        DisplayEmptyCount(count, "shapes.");
                    }
                }
                else if (fileType == StringResources.excel)
                {
                    // with XL, we would need to check all sheets
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(TxtFileName.Text, false))
                    {
                        foreach (Sheet sheet in document.WorkbookPart.Workbook.Sheets)
                        {
                            foreach (A.Spreadsheet.Shape shape in sheet.Descendants<A.Spreadsheet.Shape>())
                            {
                                count++;
                                LstDisplay.Items.Add(count + StringResources.shpXlDraw);
                            }

                            foreach (AO.Shape shape in sheet.Descendants<AO.Shape>())
                            {
                                count++;
                                LstDisplay.Items.Add(count + StringResources.shpOfficeDrawing);
                            }

                            foreach (O.Vml.Shape shape in sheet.Descendants<O.Vml.Shape>())
                            {
                                count++;
                                LstDisplay.Items.Add(count + StringResources.period + shape.Id + StringResources.arrow + StringResources.shpVml);
                            }

                            foreach (O.Math.Shape shape in sheet.Descendants<O.Math.Shape>())
                            {
                                count++;
                                LstDisplay.Items.Add(count + StringResources.shpMath);
                            }

                            foreach (A.Diagrams.Shape shape in sheet.Descendants<A.Diagrams.Shape>())
                            {
                                count++;
                                LstDisplay.Items.Add(count + StringResources.shpDrawingDgm);
                            }

                            foreach (A.ChartDrawing.Shape shape in sheet.Descendants<A.ChartDrawing.Shape>())
                            {
                                count++;
                                LstDisplay.Items.Add(count + StringResources.shpChartDraw);
                            }

                            foreach (A.Charts.Shape shape in sheet.Descendants<A.Charts.Shape>())
                            {
                                count++;
                                LstDisplay.Items.Add(count + StringResources.shpChartShape);
                            }

                            foreach (A.Shape shape in sheet.Descendants<A.Shape>())
                            {
                                count++;
                                LstDisplay.Items.Add(count + StringResources.shpShape);
                            }

                            foreach (A.Diagrams.Shape3D shape in sheet.Descendants<A.Diagrams.Shape3D>())
                            {
                                count++;
                                LstDisplay.Items.Add(count + StringResources.shp3D);
                            }

                            DisplayEmptyCount(count, "shapes.");
                        }
                    }
                }
                else if (fileType == StringResources.powerpoint)
                {
                    // with PPT, we need to run through all slides
                    using (PresentationDocument document = PresentationDocument.Open(TxtFileName.Text, false))
                    {
                        foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
                        {
                            foreach (P.Shape shape in slidePart.Slide.Descendants<P.Shape>())
                            {
                                count++;
                                foreach (OpenXmlElement child1 in shape.ChildElements)
                                {
                                    if (child1.GetType().ToString() == "DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties")
                                    {
                                        foreach (OpenXmlElement child2 in child1.ChildElements)
                                        {
                                            if (child2.GetType().ToString() == "DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties")
                                            {
                                                P.NonVisualDrawingProperties nvdp = (P.NonVisualDrawingProperties)child2;
                                                LstDisplay.Items.Add(count + StringResources.period + nvdp.Name);
                                            }
                                        }
                                    }
                                }
                            }

                            foreach (AO.Shape shape in slidePart.Slide.Descendants<AO.Shape>())
                            {
                                count++;
                                LstDisplay.Items.Add(count + StringResources.shpOfficeDrawing);
                            }

                            foreach (O.Vml.Shape shape in slidePart.Slide.Descendants<O.Vml.Shape>())
                            {
                                count++;
                                LstDisplay.Items.Add(count + StringResources.period + shape.Id + StringResources.arrow + StringResources.shpVml);
                            }

                            foreach (O.Math.Shape shape in slidePart.Slide.Descendants<O.Math.Shape>())
                            {
                                count++;
                                LstDisplay.Items.Add(count + StringResources.shpMath);
                            }

                            foreach (A.Diagrams.Shape shape in slidePart.Slide.Descendants<A.Diagrams.Shape>())
                            {
                                count++;
                                LstDisplay.Items.Add(count + StringResources.shpDrawingDgm);
                            }

                            foreach (A.ChartDrawing.Shape shape in slidePart.Slide.Descendants<A.ChartDrawing.Shape>())
                            {
                                count++;
                                LstDisplay.Items.Add(count + StringResources.shpChartDraw);
                            }

                            foreach (A.Charts.Shape shape in slidePart.Slide.Descendants<A.Charts.Shape>())
                            {
                                count++;
                                LstDisplay.Items.Add(count + StringResources.shpChartShape);
                            }

                            foreach (A.Shape shape in slidePart.Slide.Descendants<A.Shape>())
                            {
                                count++;
                                LstDisplay.Items.Add(count + StringResources.shpShape);
                            }

                            foreach (A.Diagrams.Shape3D shape in slidePart.Slide.Descendants<A.Diagrams.Shape3D>())
                            {
                                count++;
                                LstDisplay.Items.Add(count + StringResources.shp3D);
                            }

                            DisplayEmptyCount(count, "shapes.");
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
                LoggingHelper.Log("BtnListShapes Error: " + ioe.Message);
                LstDisplay.Items.Add("Error listing shapes.");
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("BtnListShapes Error: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public static string ParagraphText(XElement e)
        {
            XNamespace w = e.Name.Namespace;
            return e
                   .Elements(w + "r")
                   .Elements(w + "t")
                   .StringConcatenate(element => (string)element);
        }

        private void BtnListParagraphStyles_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                LstDisplay.Items.Clear();

                const string documentRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
                const string stylesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
                const string wordmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                XNamespace w = wordmlNamespace;

                XDocument xDoc = null;
                XDocument styleDoc = null;

                using (Package wdPackage = Package.Open(TxtFileName.Text, FileMode.Open, FileAccess.Read))
                {
                    PackageRelationship docPackageRelationship =
                      wdPackage
                      .GetRelationshipsByType(documentRelationshipType)
                      .FirstOrDefault();
                    if (docPackageRelationship != null)
                    {
                        Uri documentUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), docPackageRelationship.TargetUri);
                        PackagePart documentPart = wdPackage.GetPart(documentUri);

                        //  Load the document XML in the part into an XDocument instance.  
                        xDoc = XDocument.Load(XmlReader.Create(documentPart.GetStream()));

                        //  Find the styles part. There will only be one.  
                        PackageRelationship styleRelation = documentPart.GetRelationshipsByType(stylesRelationshipType).FirstOrDefault();
                        if (styleRelation != null)
                        {
                            Uri styleUri = PackUriHelper.ResolvePartUri(documentUri, styleRelation.TargetUri);
                            PackagePart stylePart = wdPackage.GetPart(styleUri);

                            //  Load the style XML in the part into an XDocument instance.  
                            styleDoc = XDocument.Load(XmlReader.Create(stylePart.GetStream()));
                        }
                    }
                }

                string defaultStyle =
                    (string)(
                        from style in styleDoc.Root.Elements(w + "style")
                        where (string)style.Attribute(w + "type") == "paragraph" &&
                              (string)style.Attribute(w + "default") == "1"
                        select style
                    ).First().Attribute(w + "styleId");

                // Find all paragraphs in the document.  
                var paragraphs =
                    from para in xDoc
                                 .Root
                                 .Element(w + "body")
                                 .Descendants(w + "p")
                    let styleNode = para
                                    .Elements(w + "pPr")
                                    .Elements(w + "pStyle")
                                    .FirstOrDefault()
                    select new
                    {
                        ParagraphNode = para,
                        StyleName = styleNode != null ?
                            (string)styleNode.Attribute(w + "val") :
                            defaultStyle
                    };

                // Retrieve the text of each paragraph.  
                var paraWithText =
                    from para in paragraphs
                    select new
                    {
                        ParagraphNode = para.ParagraphNode,
                        StyleName = para.StyleName,
                        Text = ParagraphText(para.ParagraphNode)
                    };

                int count = 0;

                foreach (var p in paraWithText)
                {
                    count++;
                    LstDisplay.Items.Add(count + ". StyleName: " + p.StyleName + " Text: " + p.Text);
                }
            }
            catch (IOException ioe)
            {
                LoggingHelper.Log("BtnListParagraphStyles Error: " + ioe.Message);
                LstDisplay.Items.Add("Error listing paragraphs.");
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("BtnListParagraphs Error: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void batchProcessingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmBatch bFrm = new FrmBatch()
            {
                Owner = this
            };
            bFrm.ShowDialog();
        }

        private void BtnCopyLine_Click(object sender, EventArgs e)
        {
            try
            {
                if (LstDisplay.Items.Count <= 0)
                {
                    return;
                }

                Clipboard.SetText(LstDisplay.SelectedItem.ToString());
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.ClearAndAdd, ex.Message);
                LoggingHelper.Log("BtnCopyLineOutput Error");
                LoggingHelper.Log(ex.Message);
            }
        }

        private void BtnCopyAll_Click(object sender, EventArgs e)
        {
            CopyAllItems();
        }

        public void CopyAllItems()
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
                LoggingHelper.Log("BtnCopyOutput Error");
                LoggingHelper.Log(ex.Message);
            }
        }

        private void BtnNotesPageSize_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                using (PresentationDocument document = PresentationDocument.Open(TxtFileName.Text, true))
                {
                    PowerPoint_Helpers.PowerPointOpenXml.ChangeNotesPageSize(document);
                    DisplayInformation(InformationOutput.ClearAndAdd, TxtFileName.Text + StringResources.colonBuffer + StringResources.pptNotesSizeReset);
                }
            }
            catch (NullReferenceException nre)
            {
                DisplayInformation(InformationOutput.ClearAndAdd, "** Document does not contain Notes Master **");
                LoggingHelper.Log("BtnNotesPageSize_Click Error");
                LoggingHelper.Log(nre.Message);
            }
            catch (Exception ex)
            {
                DisplayInformation(InformationOutput.ClearAndAdd, ex.Message);
                LoggingHelper.Log("BtnNotesPageSize_Click Error");
                LoggingHelper.Log(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void feedbackToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(StringResources.helpLocation);
        }
    }
}