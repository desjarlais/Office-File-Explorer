/****************************** Module Header ******************************\
Module Name:  FrmMain.cs
Project:      Office File Explorer

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
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Path = System.IO.Path;

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
using System.IO.Compression;

namespace Office_File_Explorer
{
    public partial class FrmMain : Form
    {
        // globals
        private string fromAuthor;
        private string findText;
        private string replaceText;
        public static char PrevChar = '<';
        public bool IsRegularXmlTag;
        public bool IsFixed;
        public static string FixedFallback = string.Empty;
        public static string StrOrigFileName = string.Empty;
        public static string StrDestPath = string.Empty;
        public static string StrExtension = string.Empty;
        public static string StrDestFileName = string.Empty;
        private string fileType;
        public static string StrCopiedFileName = string.Empty;

        // global numid lists
        private ArrayList oNumIdList = new ArrayList();
        private ArrayList aNumIdList = new ArrayList();
        private ArrayList numIdList = new ArrayList();

        // fix corrupt doc globals
        private static List<string> corruptNodes = new List<string>();

        // global lists
        private static List<string> pParts = new List<string>();

        // corrupt doc buffer
        private static StringBuilder sbNodeBuffer = new StringBuilder();

        public enum LogType { ClearAndAdd, Append, TextOnly, InvalidFile, LogException, EmptyCount };

        public FrmMain()
        {
            InitializeComponent();

            // log setup
            LoggingHelper.Clear();
            LoggingHelper.Log(" ## System Information ##");
            LoggingHelper.LogSystemInformation();
            LoggingHelper.Log(string.Empty);
            LoggingHelper.Log(" ## App Logging ##");
            LoggingHelper.Log("App Start");
            
            // init search replace strings
            findText = string.Empty;
            replaceText = string.Empty;

            // disable all buttons
            DisableButtons();
        }

        #region Class Properties

        public string AuthorProperty
        {
            set => fromAuthor = value;
        }

        public string FindTextProperty
        {
            set => findText = value;
        }

        public string ReplaceTextProperty
        {
            set => replaceText = value;
        }

        #endregion Class Properties

        /// <summary>
        /// Disable all buttons on the form and reset file type
        /// </summary>
        public void DisableButtons()
        {
            fileType = string.Empty;
            BtnAcceptRevisions.Enabled = false;
            BtnDeleteBreaks.Enabled = false;
            BtnDeleteComments.Enabled = false;
            BtnDeleteEndnotes.Enabled = false;
            BtnDeleteFootnotes.Enabled = false;
            BtnDeleteHdrFtr.Enabled = false;
            BtnDeleteHdrFtr.Enabled = false;
            BtnDeleteHiddenText.Enabled = false;
            BtnDeleteListTemplates.Enabled = false;
            BtnListAuthors.Enabled = false;
            BtnListDefinedNames.Enabled = false;
            BtnListComments.Enabled = false;
            BtnListEndnotes.Enabled = false;
            BtnListFonts.Enabled = false;
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
            BtnPPTRemovePII.Enabled = false;
            BtnFixDocument.Enabled = false;
            BtnFixPresentation.Enabled = false;
            BtnConvertToNonStrictFormat.Enabled = false;
            BtnListTransitions.Enabled = false;
            BtnMoveSlide.Enabled = false;
            BtnDeleteCustomProps.Enabled = false;
            BtnViewCustomXml.Enabled = false;
            BtnViewImages.Enabled = false;
            BtnListExcelHyperlinks.Enabled = false;
            BtnDeleteUnusedStyles.Enabled = false;
            BtnDeleteEmbeddedLinks.Enabled = false;
            BtnListExcelHyperlinks.Enabled = false;
            BtnListMIPLabels.Enabled = false;
        }

        public enum OxmlFileFormat { Xlsx, Xlsm, Xlst, Dotx, Docx, Docm, Potx, Pptx, Pptm, Invalid };

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
                fileType = StringResources.wWord;

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
                BtnFixDocument.Enabled = true;
                BtnDeleteUnusedStyles.Enabled = true;

                if (ffmt == OxmlFileFormat.Docm)
                {
                    BtnConvertDocmToDocx.Enabled = true;
                }
            }
            else if (ffmt == OxmlFileFormat.Xlsx || ffmt == OxmlFileFormat.Xlsm || ffmt == OxmlFileFormat.Xlst)
            {
                fileType = StringResources.wExcel;

                // enable XL only files
                BtnListDefinedNames.Enabled = true;
                BtnListHiddenRowsColumns.Enabled = true;
                BtnListLinks.Enabled = true;
                BtnListFormulas.Enabled = true;
                BtnListSharedStrings.Enabled = true;
                BtnComments.Enabled = true;
                BtnDeleteComment.Enabled = true;
                BtnListWSInfo.Enabled = true;
                BtnListCellValuesSAX.Enabled = true;
                BtnListConnections.Enabled = true;
                BtnConvertToNonStrictFormat.Enabled = true;
                BtnDeleteEmbeddedLinks.Enabled = true;
                BtnListExcelHyperlinks.Enabled = true;

                if (ffmt == OxmlFileFormat.Xlsm)
                {
                    BtnConvertXlsmToXlsx.Enabled = true;
                }
            }
            else if (ffmt == OxmlFileFormat.Pptx || ffmt == OxmlFileFormat.Pptm || ffmt == OxmlFileFormat.Potx)
            {
                fileType = StringResources.wPowerpoint;

                // enable PPT only files
                BtnPPTGetAllSlideTitles.Enabled = true;
                BtnPPTListHyperlinks.Enabled = true;
                BtnViewPPTComments.Enabled = true;
                BtnListSlideText.Enabled = true;
                BtnPPTRemovePII.Enabled = true;
                BtnFixPresentation.Enabled = true;
                BtnListTransitions.Enabled = true;
                BtnMoveSlide.Enabled = true;

                if (ffmt == OxmlFileFormat.Pptm)
                {
                    BtnConvertPptmToPptx.Enabled = true;
                }
            }
            else if (ffmt == OxmlFileFormat.Invalid)
            {
                // invalid file format
                LoggingHelper.Log("Unknown format" + TxtFileName.Text);
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
            BtnDeleteCustomProps.Enabled = true;
            BtnViewCustomXml.Enabled = true;
            BtnViewImages.Enabled = true;
            BtnListMIPLabels.Enabled = true;
        }

        private void BtnListComments_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                PreButtonClickWork();

                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    int count = 0;

                    // first list all comments in the comment part
                    WordprocessingCommentsPart commentsPart = myDoc.MainDocumentPart.WordprocessingCommentsPart;
                    if (commentsPart == null)
                    {
                        LogInformation(LogType.EmptyCount, StringResources.wComments, string.Empty);
                    }
                    else
                    {
                        foreach (O.Wordprocessing.Comment cm in commentsPart.Comments)
                        {
                            count++;
                            LstDisplay.Items.Add(count + StringResources.wPeriod + cm.Author + StringResources.wColon + cm.InnerText);
                        }
                    }

                    // now we can check how many comment references there are and display that number, they should be the same
                    IEnumerable<CommentReference> crList = myDoc.MainDocumentPart.Document.Descendants<CommentReference>();
                    IEnumerable<OpenXmlUnknownElement> unknownList = myDoc.MainDocumentPart.Document.Descendants<OpenXmlUnknownElement>();
                    int cRefCount = 0;

                    cRefCount = crList.Count();

                    foreach (OpenXmlUnknownElement uk in unknownList)
                    {
                        if (uk.LocalName == "commentReference")
                        {
                            cRefCount++;
                        }
                    }

                    LstDisplay.Items.Add(string.Empty);
                    LstDisplay.Items.Add("Total Comments = " + count);
                    LstDisplay.Items.Add("Comment References = " + cRefCount);
                }
            }
            catch (NullReferenceException nre)
            {
                LogInformation(LogType.LogException, "Word - BtnListComments_Click Error", nre.Message);
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "Word - BtnListComments_Click Error", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// log output based on the type of information
        /// </summary>
        /// <param name="display">scenario to be logged</param>
        /// <param name="output">text to be logged</param>
        /// <param name="ex">exception string if included</param>
        public void LogInformation(LogType display, string output, string ex)
        {
            switch (display)
            {
                case LogType.ClearAndAdd:
                    LstDisplay.Items.Clear();
                    LstDisplay.Items.Add(output);
                    break;
                case LogType.Append:
                    LstDisplay.Items.Add(string.Empty);
                    LstDisplay.Items.Add(output);
                    break;
                case LogType.InvalidFile:
                    LstDisplay.Items.Clear();
                    LstDisplay.Items.Add(StringResources.invalidFile);
                    break;
                case LogType.LogException:
                    LstDisplay.Items.Add(output);
                    LoggingHelper.Log(output);
                    LoggingHelper.Log(ex);
                    break;
                case LogType.EmptyCount:
                    LstDisplay.Items.Add("** Document contains no " + output + " **");
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
                PreButtonClickWork();
                XNamespace w = StringResources.wordMainAttributeNamespace;
                XDocument xDoc = null;
                XDocument styleDoc = null;
                bool containStyle = false;
                bool styleInUse = false;
                int count = 0;

                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    MainDocumentPart mainPart = myDoc.MainDocumentPart;
                    StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart;

                    LstDisplay.Items.Clear();
                    LstDisplay.Items.Add("# Style Summary #");
                    try
                    {
                        // loop the styles in style.xml
                        foreach (OpenXmlElement el in stylePart.Styles.Elements())
                        {
                            string sName = string.Empty;
                            string sType = string.Empty;

                            if (el.LocalName == "style")
                            {
                                Style s = (Style)el;
                                sName = s.StyleId;
                                sType = s.Type;
                                styleInUse = false;

                                int pStyleCount = WordExtensionClass.ParagraphsByStyleName(mainPart, sName).Count();
                                if (sType == "paragraph")
                                {
                                    if (pStyleCount > 0)
                                    {
                                        count += 1;
                                        LstDisplay.Items.Add(count + StringResources.wPeriod + sName + " -> Used in " + pStyleCount + " paragraphs"); 
                                        containStyle = true;
                                        styleInUse = true;
                                        continue;
                                    }
                                }

                                int rStyleCount = WordExtensionClass.RunsByStyleName(mainPart, sName).Count();
                                if (sType == "character")
                                {
                                    if (rStyleCount > 0)
                                    {
                                        count += 1;
                                        LstDisplay.Items.Add(count + StringResources.wPeriod + sName + " -> Used in " + rStyleCount + " runs");
                                        containStyle = true;
                                        styleInUse = true;
                                        continue;
                                    }
                                }

                                int tStyleCount = WordExtensionClass.TablesByStyleName(mainPart, sName).Count();
                                if (sType == "table")
                                {
                                    if (tStyleCount > 0)
                                    {
                                        count += 1;
                                        LstDisplay.Items.Add(count + StringResources.wPeriod + sName + " -> Used in " + tStyleCount + " tables");
                                        containStyle = true;
                                        styleInUse = true;
                                        continue;
                                    }
                                }

                                if (styleInUse == false)
                                {
                                    count += 1;
                                    LstDisplay.Items.Add(count + StringResources.wPeriod + sName + " -> (Not Used)");
                                }

                                if (count == 4079)
                                {
                                    LstDisplay.Items.Add("WARNING: Max Count of Styles for a document is 4079");
                                }
                            }
                        }

                        // add latent style information
                        LstDisplay.Items.Add(string.Empty);
                        LstDisplay.Items.Add("# Latent Style Summary #");
                        foreach (LatentStyleExceptionInfo lex in stylePart.Styles.LatentStyles)
                        {
                            count += 1;
                            if (lex.UnhideWhenUsed != null)
                            {
                                LstDisplay.Items.Add(count + StringResources.wPeriod + lex.Name + " (Hidden)");
                            }
                            else
                            {
                                LstDisplay.Items.Add(count + StringResources.wPeriod + lex.Name);
                            }
                        }

                    }
                    catch (NullReferenceException)
                    {
                        LogInformation(LogType.ClearAndAdd, "** Missing StylesWithEffects part **", string.Empty);
                        return;
                    }
                }

                if (containStyle == false)
                {
                    LogInformation(LogType.EmptyCount, StringResources.wStyles, string.Empty);
                }
                else
                {
                    // list the styles for paragraphs
                    LstDisplay.Items.Add(string.Empty);
                    LstDisplay.Items.Add("# List of paragraph styles #");
                    count = 0;

                    using (Package wdPackage = Package.Open(TxtFileName.Text, FileMode.Open, FileAccess.Read))
                    {
                        PackageRelationship docPackageRelationship = wdPackage.GetRelationshipsByType(StringResources.MainDocumentPartType).FirstOrDefault();
                        if (docPackageRelationship != null)
                        {
                            Uri documentUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), docPackageRelationship.TargetUri);
                            PackagePart documentPart = wdPackage.GetPart(documentUri);

                            //  Load the document XML in the part into an XDocument instance.
                            xDoc = XDocument.Load(XmlReader.Create(documentPart.GetStream()));

                            //  Find the styles part. There will only be one.
                            PackageRelationship styleRelation = documentPart.GetRelationshipsByType(StringResources.StyleDefsPartType).FirstOrDefault();
                            if (styleRelation != null)
                            {
                                Uri styleUri = PackUriHelper.ResolvePartUri(documentUri, styleRelation.TargetUri);
                                PackagePart stylePart = wdPackage.GetPart(styleUri);

                                //  Load the style XML in the part into an XDocument instance.
                                styleDoc = XDocument.Load(XmlReader.Create(stylePart.GetStream()));
                            }
                        }
                    }

                    string defaultStyle = (string)(
                        from style in styleDoc.Root.Elements(w + "style")
                        where (string)style.Attribute(w + "type") == "paragraph" && (string)style.Attribute(w + "default") == "1"
                        select style
                    ).First().Attribute(w + "styleId");

                    // Find all paragraphs in the document.  
                    var paragraphs =
                        from para in xDoc.Root.Element(w + "body").Descendants(w + "p")
                        let styleNode = para.Elements(w + "pPr").Elements(w + "pStyle").FirstOrDefault()
                        select new
                        {
                            ParagraphNode = para,
                            StyleName = styleNode != null ? (string)styleNode.Attribute(w + "val") : defaultStyle
                        };

                    // Retrieve the text of each paragraph.  
                    var paraWithText =
                        from para in paragraphs
                        select new
                        {
                            para.ParagraphNode,
                            para.StyleName,
                            Text = ParagraphText(para.ParagraphNode)
                        };

                    foreach (var p in paraWithText)
                    {
                        count++;
                        LstDisplay.Items.Add(count + ". StyleName: " + p.StyleName + " Text: " + p.Text);
                    }
                }
            }
            catch (IOException ioe)
            {
                LogInformation(LogType.LogException, "BtnListStyles Error: Error listing paragraphs.", ioe.Message);
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnListStyles Error: Error listing paragraphs.", ex.Message);
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
                Cursor = Cursors.WaitCursor;
                PreButtonClickWork();
                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    int count = 0;
                    
                    IEnumerable<O.Wordprocessing.Hyperlink> hLinks = myDoc.MainDocumentPart.Document.Descendants<O.Wordprocessing.Hyperlink>();
                    
                    // handle if no links are found
                    if (myDoc.MainDocumentPart.HyperlinkRelationships.Count() == 0 && myDoc.MainDocumentPart.RootElement.Descendants<FieldCode>().Count() == 0 && hLinks.Count() == 0)
                    {
                        LogInformation(LogType.EmptyCount, StringResources.wHyperlinks, string.Empty);
                    }
                    else
                    {
                        // loop through regular hyperlinks
                        foreach (O.Wordprocessing.Hyperlink h in hLinks)
                        {
                            count++;

                            string hRelUri = null;

                            // then check for hyperlinks relationships
                            foreach (HyperlinkRelationship hRel in myDoc.MainDocumentPart.HyperlinkRelationships)
                            {
                                if (h.Id == hRel.Id)
                                {
                                    hRelUri = hRel.Uri.ToString();
                                }
                            }

                            LstDisplay.Items.Add(count + StringResources.wPeriod + h.InnerText + " Uri = " + hRelUri);
                        }

                        // now we need to check for field hyperlinks
                        foreach (var field in myDoc.MainDocumentPart.RootElement.Descendants<FieldCode>())
                        {
                            string fldText;
                            if (field.InnerText.StartsWith(" HYPERLINK"))
                            {
                                count++;
                                fldText = field.InnerText.Remove(0, 11);
                                LstDisplay.Items.Add(count + StringResources.wPeriod + fldText);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnListHyperlinks_Click Error", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnListTemplates_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            PreButtonClickWork();
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
                    LstDisplay.Items.Add(string.Empty);
                    LstDisplay.Items.Add("All List Templates in document:");
                    int aCount = 0;

                    if (numPart != null)
                    {
                        foreach (OpenXmlElement el in numPart.Numbering.Elements())
                        {
                            foreach (AbstractNumId aNumId in el.Descendants<AbstractNumId>())
                            {
                                string strNumId = el.GetAttribute("numId", StringResources.wordMainAttributeNamespace).Value;
                                aNumIdList.Add(strNumId);
                                aCount++;
                                LstDisplay.Items.Add(aCount + StringResources.wNumId + strNumId);
                            }
                        }
                    }
                    else
                    {
                        LstDisplay.Items.Add(StringResources.wNone);
                    }

                    // get the unused list templates
                    oNumIdList = OrphanedListTemplates(numIdList, aNumIdList);
                    LstDisplay.Items.Add(string.Empty);
                    LstDisplay.Items.Add("Orphaned List Templates:");
                    if (oNumIdList.Count > 0)
                    {
                        int oCount = 0;
                        foreach (object item in oNumIdList)
                        {
                            oCount++;
                            LstDisplay.Items.Add(oCount + StringResources.wNumId + item);
                        }
                    }
                    else
                    {
                        LstDisplay.Items.Add(StringResources.wNone);
                    }
                    
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnListTemplates_Click Error", ex.Message);
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
                LstDisplay.Items.Add(StringResources.wNone);
                return;
            }

            // since we have lists, display them
            int count = 0;
            foreach (object item in al)
            {
                count++;
                LstDisplay.Items.Add(count + StringResources.wNumId + item);
                
                // Word is limited to 2047 total active lists in a document
                if (count == 2047)
                {
                    LogInformation(LogType.LogException, "## You have too many lists in this file. Word will only display up to 2047 lists. ##", "Active List Limit Reached");
                }
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
            PreButtonClickWork();
            try
            {
                if (fileType == StringResources.wWord)
                {
                    using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, false))
                    {
                        int wdOleObjCount = GetEmbeddedObjectProperties(myDoc.MainDocumentPart);
                        int wdOlePkgPart = GetEmbeddedPackageProperties(myDoc.MainDocumentPart);

                        if (wdOlePkgPart == 0 && wdOleObjCount == 0)
                        {
                            LogInformation(LogType.EmptyCount, StringResources.wOle, string.Empty);
                        }
                    }
                }
                else if (fileType == StringResources.wExcel)
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
                            LogInformation(LogType.EmptyCount, StringResources.wOle, string.Empty);
                        }
                    }
                }
                else if (fileType == StringResources.wPowerpoint)
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
                            LogInformation(LogType.EmptyCount, StringResources.wOle, string.Empty);
                        }
                    }
                }
                else
                {
                    LogInformation(LogType.ClearAndAdd, "BtnListOle_Click Error", string.Empty);
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnListOle_Click Error", ex.Message);
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
                    LstDisplay.Items.Add(wPart.Uri + StringResources.wArrow + wPart.EmbeddedPackageParts.ElementAt(x).Uri.ToString());
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
                    LstDisplay.Items.Add(sPart.Uri + StringResources.wArrow + sPart.EmbeddedPackageParts.ElementAt(x).Uri.ToString());
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
                    LstDisplay.Items.Add(wPart.Uri + StringResources.wArrow + wPart.EmbeddedObjectParts.ElementAt(x).Uri.ToString());
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
                    LstDisplay.Items.Add(sPart.Uri + StringResources.wArrow + sPart.EmbeddedObjectParts.ElementAt(x).Uri.ToString());
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
            PreButtonClickWork();

            try
            {
                WordprocessingDocument document;
                List<string> authors = new List<string>();

                using (document = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    // get the list of authors
                    fromAuthor = string.Empty;

                    authors = WordOpenXml.GetAllAuthors(document.MainDocumentPart.Document);

                    FrmAuthors aFrm = new FrmAuthors(authors)
                    {
                        Owner = this
                    };
                    
                    aFrm.ShowDialog();
                }

                Cursor = Cursors.WaitCursor;

                if (fromAuthor == "* No Authors *" || fromAuthor == string.Empty)
                {
                    LogInformation(LogType.ClearAndAdd, "** No Revisions To Accept **", string.Empty);
                    return;
                }

                WordOpenXml.AcceptAllRevisions(TxtFileName.Text, fromAuthor);
                LogInformation(LogType.ClearAndAdd, "** Revisions Accepted **", string.Empty);
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnAcceptRevisions_Click Error", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnDeleteComments_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            PreButtonClickWork();
            try
            {
                WordOpenXml.RemoveComments(TxtFileName.Text);
                LogInformation(LogType.ClearAndAdd, "** Comments Removed **", string.Empty);
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnDeleteComments_Click Error", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnDeleteHiddenText_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            PreButtonClickWork();
            try
            {
                if (WordOpenXml.DeleteHiddenText(TxtFileName.Text))
                {
                    LogInformation(LogType.ClearAndAdd, "** Hidden text deleted **", string.Empty);
                }
                else
                {
                    LogInformation(LogType.ClearAndAdd, "** Document does not contain hiddent text **", string.Empty);
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnDeleteHiddenText_Click Error", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnDeleteHdrFtr_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            PreButtonClickWork();
            try
            {
                if (WordOpenXml.RemoveHeadersFooters(TxtFileName.Text))
                {
                    LogInformation(LogType.ClearAndAdd, "** Headers/Footer removed **", string.Empty);
                }
                else
                {
                    LogInformation(LogType.ClearAndAdd, "** Document does not contain a header or footer **", string.Empty);
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnDeleteHdrFtr_Click Error", ex.Message);
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
                Cursor = Cursors.WaitCursor;
                PreButtonClickWork();
                BtnListTemplates.PerformClick();
                foreach (object orphanLT in oNumIdList)
                {
                    WordOpenXml.RemoveListTemplatesNumId(TxtFileName.Text, orphanLT.ToString());
                }
                LogInformation(LogType.ClearAndAdd, "** List Templates removed **", string.Empty);
            }
            catch (Exception ex)
            {
                LogInformation(LogType.InvalidFile, "BtnDeleteListTemplates_Click Error: ", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnDeleteBreaks_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();
            if (WordOpenXml.RemoveBreaks(TxtFileName.Text))
            {
                LogInformation(LogType.ClearAndAdd, "** Page and Section breaks have been removed **", string.Empty);
            }
            else
            {
                LogInformation(LogType.ClearAndAdd, "** Document does not contain any page breaks **", string.Empty);
            }
        }

        private void BtnRemovePII_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();
            if (!File.Exists(TxtFileName.Text))
            {
                LogInformation(LogType.InvalidFile, TxtFileName.Text, string.Empty);
            }
            else
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    if (WordExtensionClass.HasPersonalInfo(document) == true)
                    {
                        WordExtensionClass.RemovePersonalInfo(document);
                        LogInformation(LogType.ClearAndAdd, "PII Removed from file.", string.Empty);
                    }
                    else
                    {
                        LogInformation(LogType.EmptyCount, StringResources.wPII, string.Empty);
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
                LstDisplay.Items.Add(StringResources.wHeaderLine);
            }
        }

        private void BtnValidateFile_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                PreButtonClickWork();
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
                    LoggingHelper.Log("Unknown file format - BtnValidateFileClick Error");
                    throw new Exception();
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnValidateFile_Click Error", ex.Message);
            }
            finally
            {
                if (LstDisplay.Items.Count < 0)
                {
                    LogInformation(LogType.EmptyCount, StringResources.wValidationErr, string.Empty);
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
                PreButtonClickWork();

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
                                    LstDisplay.Items.Add(count + StringResources.wPeriod + c.CellReference + StringResources.wEqualSign + c.CellFormula.Text);
                                }
                            }
                        }
                    }
                }

                if (count == 0)
                {
                    LogInformation(LogType.EmptyCount, StringResources.wForumulas, string.Empty);
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnListFormulas_Click Error", ex.Message);
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
                PreButtonClickWork();
                int count = 0;

                using (WordprocessingDocument doc = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    foreach (O.Wordprocessing.Font ft in doc.MainDocumentPart.FontTablePart.Fonts)
                    {
                        count++;
                        LstDisplay.Items.Add(count + StringResources.wPeriod + ft.Name);
                    }
                }

                if (count == 0)
                {
                    LogInformation(LogType.EmptyCount, StringResources.wFonts, string.Empty);
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnListFonts_Click Error", ex.Message);
            }
        }

        private void BtnListFootnotes_Click(object sender, EventArgs e)
        {
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    PreButtonClickWork();
                    FootnotesPart footnotePart = doc.MainDocumentPart.FootnotesPart;
                    if (footnotePart != null)
                    {
                        int count = 0;
                        foreach (Footnote fn in footnotePart.Footnotes)
                        {
                            if (fn.InnerText != string.Empty)
                            {
                                count++;
                                LogInformation(LogType.TextOnly, count + StringResources.wPeriod + fn.InnerText, string.Empty);
                            }
                        }

                        if (count == 0)
                        {
                            LogInformation(LogType.EmptyCount, StringResources.wFootnotes, string.Empty);
                        }
                    }
                    else
                    {
                        LogInformation(LogType.TextOnly, StringResources.wFootnotes, string.Empty);
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnListFootnotes_Click Error", ex.Message);
            }
        }

        private void BtnListEndnotes_Click(object sender, EventArgs e)
        {
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    PreButtonClickWork();
                    EndnotesPart endnotePart = doc.MainDocumentPart.EndnotesPart;
                    if (endnotePart != null)
                    {
                        int count = 0;
                        foreach (Endnote en in endnotePart.Endnotes)
                        {
                            if (en.InnerText != string.Empty)
                            {
                                count++;
                                LogInformation(LogType.TextOnly, count + StringResources.wPeriod + en.InnerText, string.Empty);
                            }
                        }

                        if (count == 0)
                        {
                            LogInformation(LogType.EmptyCount, StringResources.wEndnotes, string.Empty);
                        }
                    }
                    else
                    {
                        LogInformation(LogType.TextOnly, StringResources.wEndnotes, string.Empty);
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnListEndnotes_Click Error", ex.Message);
            }
        }

        private void BtnDeleteFootnotes_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();
            try
            {
                if (WordOpenXml.RemoveFootnotes(TxtFileName.Text))
                {
                    LogInformation(LogType.ClearAndAdd, "** Footnotes removed from document **", string.Empty);
                }
                else
                {
                    LogInformation(LogType.EmptyCount, StringResources.wFootnotes, string.Empty);
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnDeleteFootnotes_Click Error", ex.Message);
            }
        }

        private void BtnDeleteEndnotes_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();
            try
            {
                if (WordOpenXml.RemoveEndnotes(TxtFileName.Text))
                {
                    LogInformation(LogType.ClearAndAdd, "** Endnotes removed from document **", string.Empty);
                }
                else
                {
                    LogInformation(LogType.EmptyCount, StringResources.wEndnotes, string.Empty);
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnDeleteEndnotes_Click Error", ex.Message);
            }
        }

        private void BtnListRevisions_Click(object sender, EventArgs e)
        {
            try
            {
                int revCount = 0;
                PreButtonClickWork();
                Cursor = Cursors.WaitCursor;

                List<string> authorList = new List<string>();
                
                using (WordprocessingDocument document = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    // if we have an author, go through all the revisions
                    authorList = WordOpenXml.GetAllAuthors(document.MainDocumentPart.Document);

                    // check people part for authors too
                    WordprocessingPeoplePart peoplePart = document.MainDocumentPart.WordprocessingPeoplePart;
                    if (peoplePart != null)
                    {
                        foreach (Person person in peoplePart.People)
                        {
                            PresenceInfo pi = person.PresenceInfo;
                            authorList.Add(person.Author);
                        }
                    }

                    Document doc = document.MainDocumentPart.Document;
                    var paragraphChanged = doc.Descendants<ParagraphPropertiesChange>().ToList();
                    var runChanged = doc.Descendants<RunPropertiesChange>().ToList();
                    var deleted = doc.Descendants<DeletedRun>().ToList();
                    var deletedParagraph = doc.Descendants<Deleted>().ToList();
                    var inserted = doc.Descendants<InsertedRun>().ToList();

                    // get the list of authors
                    fromAuthor = string.Empty;

                    FrmAuthors aFrm = new FrmAuthors(authorList)
                    {
                        Owner = this
                    };
                    aFrm.ShowDialog();

                    if (fromAuthor == "* All Authors *")
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
                                LstDisplay.Items.Add(revCount + StringResources.wPeriod + s + " : Paragraph Changed ");
                            }

                            foreach (var item in tempDeletedParagraph)
                            {
                                revCount++;
                                LstDisplay.Items.Add(revCount + StringResources.wPeriod + s + " : Paragraph Deleted ");
                            }

                            foreach (var item in tempRunChanged)
                            {
                                revCount++;
                                LstDisplay.Items.Add(revCount + StringResources.wPeriod + s + " :  Run Changed = " + item.InnerText);
                            }

                            foreach (var item in tempDeleted)
                            {
                                revCount++;
                                LstDisplay.Items.Add(revCount + StringResources.wPeriod + s + " :  Deletion = " + item.InnerText);
                            }

                            foreach (var item in tempInserted)
                            {
                                if (item.Parent != null)
                                {
                                    var textRuns = item.Elements<Run>().ToList();
                                    var parent = item.Parent;

                                    foreach (var textRun in textRuns)
                                    {
                                        revCount++;
                                        LstDisplay.Items.Add(revCount + StringResources.wPeriod + s + " :  Insertion = " + textRun.InnerText);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        // list the selected authors revisions
                        if (!string.IsNullOrEmpty(fromAuthor))
                        {
                            paragraphChanged = paragraphChanged.Where(item => item.Author == fromAuthor).ToList();
                            runChanged = runChanged.Where(item => item.Author == fromAuthor).ToList();
                            deleted = deleted.Where(item => item.Author == fromAuthor).ToList();
                            inserted = inserted.Where(item => item.Author == fromAuthor).ToList();
                            deletedParagraph = deletedParagraph.Where(item => item.Author == fromAuthor).ToList();

                            if ((paragraphChanged.Count + runChanged.Count + deleted.Count + inserted.Count + deletedParagraph.Count) == 0)
                            {
                                if (fromAuthor == "* No Authors *")
                                {
                                    LogInformation(LogType.ClearAndAdd, "** There are no revisions in this document **", string.Empty);
                                }
                                else
                                {
                                    LogInformation(LogType.ClearAndAdd, "** This author has no changes **", string.Empty);
                                }

                                return;
                            }
                        }
                        else
                        {
                            LogInformation(LogType.ClearAndAdd, "** There are no revisions in this document **", string.Empty);
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
                                var textRuns = item.Elements<Run>().ToList();
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
                LogInformation(LogType.LogException, "BtnListRevisions_Click Error", ex.Message);
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
                PreButtonClickWork();
                using (WordprocessingDocument doc = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    int count = 0;

                    // check the peoplepart and list those authors
                    WordprocessingPeoplePart peoplePart = doc.MainDocumentPart.WordprocessingPeoplePart;
                    if (peoplePart != null)
                    { 
                        foreach (Person person in peoplePart.People)
                        {
                            count++;
                            PresenceInfo pi = person.PresenceInfo;
                            LstDisplay.Items.Add(count + StringResources.wPeriod + person.Author);
                            LstDisplay.Items.Add("   - User Id = " + pi.UserId);
                            LstDisplay.Items.Add("   - Provider Id = " + pi.ProviderId);
                        }
                    }
                                        
                    List<string> tempAuthors = WordOpenXml.GetAllAuthors(doc.MainDocumentPart.Document);
                    
                    // sometimes there are authors in a file but they don't exist in people.xml
                    if (tempAuthors.Count > 0)
                    {
                        // if the people part count is the same as GetAllAuthors, they must be the same authors
                        if (count == tempAuthors.Count)
                        {
                            return;
                        }

                        // if the count is not the same, display those authors
                        foreach (string s in tempAuthors)
                        {
                            count++;
                            LstDisplay.Items.Add(count + ". User Name = " + s);
                        }
                    }

                    // if the count is 0 at this point, no authors exist
                    if (count == 0)
                    {
                        LogInformation(LogType.EmptyCount, StringResources.wAuthors, string.Empty);
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnListAuthors_Click Error", ex.Message);
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
                Cursor = Cursors.WaitCursor;
                PreButtonClickWork();

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
                                    LstDisplay.Items.Add(string.Empty);
                                    LstDisplay.Items.Add("---- Compatibility Settings ---- ");

                                    int settingCount = setting.Count();
                                    int settingIndex = 0;

                                    do
                                    {
                                        if (setting.ElementAt(settingIndex).LocalName != "compatSetting")
                                        {
                                            if (setting.ElementAt(0).InnerText != string.Empty)
                                            {
                                                LstDisplay.Items.Add(setting.ElementAt(0).LocalName + StringResources.wColon + setting.ElementAt(0).InnerText);
                                            }
                                            settingIndex++;
                                        }
                                        else
                                        {
                                            CompatibilitySetting cs = (CompatibilitySetting)setting.ElementAt(settingIndex);
                                            if (cs.Name == "compatibilityMode")
                                            {
                                                string compatModeVersion = string.Empty;

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

                                                LstDisplay.Items.Add(cs.Name + StringResources.wColon + cs.Val + compatModeVersion);
                                                settingIndex++;
                                            }
                                            else
                                            {
                                                LstDisplay.Items.Add(cs.Name + StringResources.wColon + cs.Val);
                                                settingIndex++;
                                            }
                                        }
                                    } while (settingIndex < settingCount);

                                    LstDisplay.Items.Add(string.Empty);
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
                                            sb.Append(xe.Name + StringResources.wColon);
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
                                                            sb.Append(xa.LocalName + StringResources.wColon + xa.Value);
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
                            LogInformation(LogType.EmptyCount, StringResources.wCustomDocProps, string.Empty);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogInformation(LogType.LogException, "BtnViewCustomDocProps (doc settings) Error", ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnViewCustomDocProps_Click Error", ex.Message);
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
            LstDisplay.Items.Add(string.Empty);
        }

        public void GetExtendedFileProps(ExtendedFilePropertiesPart exFilePropPart)
        {
            LstDisplay.Items.Add("---- Extended File Properties ----");
            try
            {
                if (exFilePropPart != null)
                {
                    XmlDocument xmlProps = new XmlDocument();
                    xmlProps.Load(exFilePropPart.GetStream());
                    XmlNodeList exProps = xmlProps.GetElementsByTagName("Properties");

                    foreach (XmlNode xNode in exProps)
                    {
                        foreach (XmlElement xElement in xNode)
                        {
                            DisplayElementDetails(xElement);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "GetExtendedFileProps Error", ex.Message);
            }
        }

        public void DisplayElementDetails(XmlElement elem)
        {
            if (elem.Name == StringResources.wDocSecurity)
            {
                switch (elem.InnerText)
                {
                    case "0":
                        LstDisplay.Items.Add(StringResources.wDocSecurity + StringResources.wColon + "None");
                        break;
                    case "1":
                        LstDisplay.Items.Add(StringResources.wDocSecurity + StringResources.wColon + "Password Protected");
                        break;
                    case "2":
                        LstDisplay.Items.Add(StringResources.wDocSecurity + StringResources.wColon + "Read-Only Recommended");
                        break;
                    case "4":
                        LstDisplay.Items.Add(StringResources.wDocSecurity + StringResources.wColon + "Read-Only Enforced");
                        break;
                    case "8":
                        LstDisplay.Items.Add(StringResources.wDocSecurity + StringResources.wColon + "Locked For Annotation");
                        break;
                    default:
                        break;
                }
            }
            else
            {
                LstDisplay.Items.Add(elem.Name + StringResources.wColonBuffer + elem.InnerText);
            }
        }

        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmAbout frm = new FrmAbout();
            frm.ShowDialog(this);
            frm.Dispose();
        }

        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                OpenFileDialog fDialog = new OpenFileDialog
                {
                    Title = "Select Office Open Xml File.",
                    Filter = "Open XML Files | *.docx; *.dotx; *.docm; *.dotm; *.xlsx; *.xlsm; *.xlst; *.xltm; *.pptx; *.pptm; *.potx; *.potm",
                    RestoreDirectory = true,
                    InitialDirectory = @"%userprofile%"
                };

                if (fDialog.ShowDialog() == DialogResult.OK)
                {
                    // disable buttons before each open
                    DisableButtons();

                    TxtFileName.Text = fDialog.FileName.ToString();
                    if (!File.Exists(TxtFileName.Text))
                    {
                        LogInformation(LogType.InvalidFile, StringResources.fileDoesNotExist, string.Empty);
                        return;
                    }
                    else
                    {
                        LstDisplay.Items.Clear();
                        // if the file doesn't start with PK, we can stop trying to process it
                        if (!FileUtilities.IsZipArchiveFile(TxtFileName.Text))
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
                LogInformation(LogType.LogException, "File Open Error: ", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public void PopulatePackageParts()
        {
            pParts.Clear();
            using (FileStream zipToOpen = new FileStream(TxtFileName.Text, FileMode.Open, FileAccess.Read))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Read))
                {
                    foreach (ZipArchiveEntry zae in archive.Entries)
                    {
                        pParts.Add(zae.FullName + StringResources.wColonBuffer + FileUtilities.SizeSuffix(zae.Length));
                    }
                }
            }
        }

        /// <summary>
        /// function to open the file in the SDK
        /// if the SDK fails to open the file, it is not a valid docx
        /// </summary>
        /// <param name="file">the path to the initial fix attempt</param>
        /// <param name="isFileOpen">is a file open scenario</param>
        public void OpenWithSdk(string file, bool isFileOpen)
        {
            string body = string.Empty;

            try
            {
                Cursor = Cursors.WaitCursor;

                // if the file is opened by the SDK, we can proceed with opening in tool
                if (isFileOpen)
                {
                    SetUpButtons();
                }

                if (fileType == StringResources.wWord)
                {
                    using (WordprocessingDocument document = WordprocessingDocument.Open(file, false))
                    {
                        // try to get the localname of the document.xml file, if it fails, it is not a Word file
                        body = document.MainDocumentPart.Document.LocalName;
                    }
                }
                else if (fileType == StringResources.wExcel)
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(file, false))
                    {
                        // try to get the localname of the workbook.xml file if it fails, its not an Excel file
                        body = document.WorkbookPart.Workbook.LocalName;
                    }
                }
                else if (fileType == StringResources.wPowerpoint)
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
            catch (OpenXmlPackageException ope)
            {
                // if the exception is related to invalid hyperlinks, use the FixInvalidUri method to change the file
                // once we change the copied file, we can open it in the SDK
                if (ope.ToString().Contains("Invalid Hyperlink"))
                {
                    // known issue in .NET with malformed hyperlinks causing SDK to throw during parse
                    // see UriFixHelper for more details
                    // get the path and make a new file name in the same directory
                    var StrDestPath = Path.GetDirectoryName(TxtFileName.Text) + "\\";
                    var StrExtension = Path.GetExtension(TxtFileName.Text);
                    var StrCopyFileName = StrDestPath + Path.GetFileNameWithoutExtension(TxtFileName.Text) + StringResources.wCopyFileParentheses + StrExtension;

                    // need a copy of the file to change the hyperlinks so we can open the modified version instead of the original
                    if (!File.Exists(StrCopyFileName))
                    {
                        File.Copy(TxtFileName.Text, StrCopyFileName);
                    }
                    else
                    {
                        StrCopyFileName = StrDestPath + Path.GetFileNameWithoutExtension(TxtFileName.Text) + StringResources.wCopyFileParentheses + FileUtilities.GetRandomNumber().ToString() + StrExtension;
                        File.Copy(TxtFileName.Text, StrCopyFileName);
                    }

                    // create the new file with the updated hyperlink
                    using (FileStream fs = new FileStream(StrCopyFileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        UriFixHelper.FixInvalidUri(fs, brokenUri => FixUri(brokenUri));
                    }

                    // now use the new file in the open logic from above
                    if (fileType == StringResources.wWord)
                    {
                        using (WordprocessingDocument document = WordprocessingDocument.Open(StrCopyFileName, false))
                        {
                            // try to get the localname of the document.xml file, if it fails, it is not a Word file
                            body = document.MainDocumentPart.Document.LocalName;
                        }
                    }
                    else if (fileType == StringResources.wExcel)
                    {
                        using (SpreadsheetDocument document = SpreadsheetDocument.Open(StrCopyFileName, false))
                        {
                            // try to get the localname of the workbook.xml file if it fails, its not an Excel file
                            body = document.WorkbookPart.Workbook.LocalName;
                        }
                    }
                    else if (fileType == StringResources.wPowerpoint)
                    {
                        using (PresentationDocument document = PresentationDocument.Open(StrCopyFileName, false))
                        {
                            // try to get the presentation.xml local name, if it fails it is not a PPT file
                            body = document.PresentationPart.Presentation.LocalName;
                        }
                    }

                    // update the main form UI
                    TxtFileName.Text = StrCopyFileName;
                    StrCopiedFileName = StrCopyFileName;
                }
                else
                {
                    // unknown issue opening from .net
                    DisableButtons();
                    LstDisplay.Items.Add("Invalid File: FixUri Failure");
                    LoggingHelper.Log("OpenWithSDK Error: " + ope.Message);
                }
            }
            catch (Exception ex)
            {
                // if the file failed to open in the sdk, it is invalid or corrupt and we need to stop opening
                DisableButtons();
                LstDisplay.Items.Add("Invalid File: Unknown error opening file.");
                LoggingHelper.Log("OpenWithSDK Error: " + ex.Message);
                BtnFixCorruptDocument.Enabled = true;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// given a broken uri this function will return a generic non-broken uri
        /// </summary>
        /// <param name="brokenUri">the uri that is failing in the sdk</param>
        /// <returns></returns>
        private static Uri FixUri(string brokenUri)
        {
            brokenUri = "http://broken-link/";
            return new Uri(brokenUri);
        }

        private void BtnPPTListHyperlinks_Click(object sender, EventArgs e)
        {
            try
            {
                PreButtonClickWork();

                // Open the presentation file as read-only.
                using (PresentationDocument document = PresentationDocument.Open(TxtFileName.Text, false))
                {
                    int linkCount = 0;
                    foreach (string s in PowerPointOpenXml.GetAllExternalHyperlinksInPresentation(TxtFileName.Text))
                    {
                        linkCount++;
                        LstDisplay.Items.Add(linkCount + StringResources.wPeriod + s);
                    }

                    if (linkCount == 0)
                    {
                        LogInformation(LogType.EmptyCount, StringResources.wHyperlinks, string.Empty);
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnPPTListHyperlinks_Click Error", ex.Message);
            }
        }

        private void BtnPPTGetAllSlideTitles_Click(object sender, EventArgs e)
        {
            try
            {
                PreButtonClickWork();

                // Open the presentation as read-only.
                using (PresentationDocument presentationDocument = PresentationDocument.Open(TxtFileName.Text, false))
                {
                    int slideCount = 0;

                    foreach (string s in PowerPointOpenXml.GetSlideTitles(presentationDocument))
                    {
                        slideCount++;
                        LstDisplay.Items.Add(slideCount + StringResources.wPeriod + s);
                    }

                    if (slideCount == 0)
                    {
                        LogInformation(LogType.EmptyCount, StringResources.wSlides, string.Empty);
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnGetAllSlideTitles_Click Error", ex.Message);
            }
        }

        private void BtnSearchAndReplace_Click(object sender, EventArgs e)
        {
            try
            {
                PreButtonClickWork();
                FrmSearchAndReplace sFrm = new FrmSearchAndReplace()
                {
                    Owner = this
                };
                sFrm.ShowDialog();

                if (findText == string.Empty && replaceText == string.Empty)
                {
                    return;
                }
                else
                {
                    SearchAndReplace(TxtFileName.Text, findText, replaceText);
                    LstDisplay.Items.Clear();
                    LstDisplay.Items.Add("** Search and Replace Finished **");
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnSearchAndReplace_Click Error", ex.Message);
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
                PreButtonClickWork();
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
                        LstDisplay.Items.Add(ExtRelCount + StringResources.wPeriod + extWbPart.ExternalRelationships.ElementAt(0).Uri);
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, StringResources.wErrorText, ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnDeleteExternalLinks_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();
            if (ExcelOpenXml.RemoveExternalLinks(TxtFileName.Text))
            {
                LstDisplay.Items.Add("** External References Deleted **");
            }
            else
            {
                LstDisplay.Items.Add("** Document does not contain external references **");
            }
        }

        private void BtnListDefinedNames_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                PreButtonClickWork();
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
                            LstDisplay.Items.Add(nameCount + StringResources.wPeriod + dn.Name.Value + " = " + dn.Text);
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
                LogInformation(LogType.LogException, "BtnListDefinedNames_Click Error", ex.Message);
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
                PreButtonClickWork();

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
                                LstDisplay.Items.Add(rowCount + StringResources.wPeriod + row.InnerText);
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
                        LstDisplay.Items.Add(string.Empty);
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnListHiddenRowsColumns_Click Error", ex.Message);
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
                PreButtonClickWork();
                int sharedStringCount = 0;

                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, false))
                {
                    WorkbookPart wbPart = excelDoc.WorkbookPart;
                    if (wbPart.SharedStringTablePart != null)
                    {
                        SharedStringTable sst = wbPart.SharedStringTablePart.SharedStringTable;
                        LstDisplay.Items.Add("SharedString Count = " + sst.Count());
                        LstDisplay.Items.Add("Unique Count = " + sst.UniqueCount);
                        LstDisplay.Items.Add(string.Empty);

                        foreach (SharedStringItem ssi in sst)
                        {
                            sharedStringCount++;
                            O.Spreadsheet.Text ssValue = ssi.Text;
                            if (ssValue.Text != null)
                            {
                                LstDisplay.Items.Add(sharedStringCount + StringResources.wPeriod + ssValue.Text);
                            }
                        }
                    }
                    else
                    {
                        LogInformation(LogType.EmptyCount, StringResources.wSharedStrings, string.Empty);
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, StringResources.wSharedStrings, ex.Message);
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
                PreButtonClickWork();
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, false))
                {
                    WorkbookPart wbPart = excelDoc.WorkbookPart;
                    int commentCount = 0;
                    LstDisplay.Items.Clear();

                    foreach (WorksheetPart wsp in wbPart.WorksheetParts)
                    {
                        WorksheetCommentsPart wcp = wsp.WorksheetCommentsPart;
                        if (wcp != null)
                        {
                            foreach (O.Spreadsheet.Comment cmt in wcp.Comments.CommentList)
                            {
                                commentCount++;
                                CommentText cText = cmt.CommentText;
                                LstDisplay.Items.Add(commentCount + StringResources.wPeriod + cText.InnerText);
                            }
                        }
                    }

                    if (commentCount == 0)
                    {
                        LogInformation(LogType.EmptyCount, StringResources.wComments, string.Empty);
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "Excel - BtnComments_Click Error:", ex.Message);
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
                PreButtonClickWork();
                bool hasComments = false;

                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, true))
                {
                    WorkbookPart wbPart = excelDoc.WorkbookPart;
                    LstDisplay.Items.Clear();

                    foreach (WorksheetPart wsp in wbPart.WorksheetParts)
                    {
                        if (wsp.WorksheetCommentsPart != null)
                        {
                            if (wsp.WorksheetCommentsPart.Comments.Count() > 0)
                            {
                                WorksheetCommentsPart wcp = wsp.WorksheetCommentsPart;
                                foreach (O.Spreadsheet.Comment cmt in wcp.Comments.CommentList)
                                {
                                    cmt.Remove();
                                }
                            }
                            else
                            {
                                hasComments = true;
                            }
                        }
                    }

                    if (hasComments == true)
                    {
                        wbPart.Workbook.Save();
                        LstDisplay.Items.Add("** Comments Deleted **");
                    }
                    else
                    {
                        LogInformation(LogType.EmptyCount, StringResources.wComments, string.Empty);
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnListFormulas_Click Error", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnChangeTheme_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();
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
                if (!File.Exists(TxtFileName.Text))
                {
                    LogInformation(LogType.InvalidFile, StringResources.fileDoesNotExist, string.Empty);
                    return;
                }
                else
                {
                    if (fileType == StringResources.wWord)
                    {
                        // call the replace function using the theme file provided
                        OfficeHelpers.ReplaceTheme(TxtFileName.Text, sThemeFilePath, fileType);
                        LogInformation(LogType.ClearAndAdd, StringResources.themeFileAdded, string.Empty);
                    }
                    else if (fileType == StringResources.wExcel)
                    {
                        OfficeHelpers.ReplaceTheme(TxtFileName.Text, sThemeFilePath, fileType);
                        LogInformation(LogType.ClearAndAdd, StringResources.themeFileAdded, string.Empty);
                    }
                    else if (fileType == StringResources.wPowerpoint)
                    {
                        OfficeHelpers.ReplaceTheme(TxtFileName.Text, sThemeFilePath, fileType);
                        LogInformation(LogType.ClearAndAdd, StringResources.themeFileAdded, string.Empty);
                    }
                    else
                    {
                        LogInformation(LogType.ClearAndAdd, "ChangeTheme Error:", "File Not Valid.");
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
                PreButtonClickWork();

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
                            LogInformation(LogType.EmptyCount, StringResources.wComments, string.Empty);
                            return;
                        }

                        foreach (P.Comment cmt in sCPart.CommentList)
                        {
                            commentCount++;
                            LstDisplay.Items.Add(commentCount + StringResources.wPeriod + cmt.InnerText);
                        }
                    }

                    if (commentCount == 0)
                    {
                        LogInformation(LogType.EmptyCount, StringResources.wComments, string.Empty);
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "PPT - BtnListComments_Click Error", ex.Message);
            }
        }

        private void BtnListWSInfo_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();
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
                        if (attr.LocalName == "name")
                        {
                            LstDisplay.Items.Add(attr.LocalName + StringResources.wColonBuffer + attr.Value);
                        }
                        else
                        {
                            LstDisplay.Items.Add(StringResources.wMinusSign + attr.LocalName + StringResources.wColonBuffer + attr.Value);
                        }
                    }
                }
            }
        }

        private void BtnListCellValuesSAX_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();
            List<string> list;

            if (Properties.Settings.Default.ListCellValuesSax == true)
            {
                list = ExcelOpenXml.ReadExcelFileSAX(TxtFileName.Text);
            }
            else
            {
                list = ExcelOpenXml.ReadExcelFileDOM(TxtFileName.Text);
            }
            
            LstDisplay.Items.Clear();
            foreach (object o in list)
            {
                LstDisplay.Items.Add(o.ToString());
            }
        }

        private void BtnConvertDocmToDocx_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();
            ConvertToNonMacro(StringResources.wWord);
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
                LogInformation(LogType.LogException, "Unable to convert Docm document.", ex.Message);
            }
        }

        private void BtnListSlideText_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                PreButtonClickWork();

                int sCount = PowerPointOpenXml.CountSlides(TxtFileName.Text);
                if (sCount > 0)
                {
                    int count = 0;

                    do
                    {
                        PowerPointOpenXml.GetSlideIdAndText(out string sldText, TxtFileName.Text, count);
                        LstDisplay.Items.Add("Slide " + (count + 1) + StringResources.wPeriod + sldText);
                        count++;
                    } while (count < sCount);
                }
                else
                {
                    LogInformation(LogType.EmptyCount, StringResources.wSlides, string.Empty);
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnListSlideText Error:", ex.Message);
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
                PreButtonClickWork();

                StrOrigFileName = TxtFileName.Text;
                StrDestPath = Path.GetDirectoryName(StrOrigFileName) + "\\";
                StrExtension = Path.GetExtension(StrOrigFileName);
                StrDestFileName = StrDestPath + Path.GetFileNameWithoutExtension(StrOrigFileName) + StringResources.wFixedFileParentheses + StrExtension;

                // check if file we are about to copy exists and append a number so it is unique
                if (File.Exists(StrDestFileName))
                {
                    StrDestFileName = StrDestPath + Path.GetFileNameWithoutExtension(StrOrigFileName) + StringResources.wFixedFileParentheses + FileUtilities.GetRandomNumber().ToString() + StrExtension;
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
                            fileType = StringResources.wWord;
                            string strDocTextBackup;
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
                                        strDocTextBackup = strDocText;

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
                                                        if (Properties.Settings.Default.FixGroupedShapes == true)
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
                                                        if (m.Value.Contains(StringResources.txtMcChoiceTagEnd))
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
                                                            if (m.Value.Contains(StringResources.txtFallbackEnd))
                                                            {
                                                                // if the match contains the closing fallback we just need to remove the entire fallback
                                                                // this will leave the closing AC and Run tags, which should be correct
                                                                strDocText = strDocText.Replace(m.Value, string.Empty);
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
                                        if (Properties.Settings.Default.RemoveFallback == true)
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
                                                        if (sbNodeBuffer.Length > 0)
                                                        {
                                                            corruptNodes.Add(sbNodeBuffer.ToString());
                                                            sbNodeBuffer.Clear();
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
                                                        corruptNodes.Add(sbNodeBuffer.ToString());
                                                        sbNodeBuffer.Clear();
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

                                            GetAllNodes(strDocText);
                                            strDocText = FixedFallback;
                                        }

                                        // check if any changes were made by comparing the doctext variable
                                        bool result = strDocText.Equals(strDocTextBackup);
                                        if (result == true)
                                        {
                                            LstDisplay.Items.Add(" ## No Corruption Found  ## ");
                                            return;
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
                                        if (Properties.Settings.Default.OpenInWord == true)
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
                        LogInformation(LogType.EmptyCount, StringResources.wInvalidXml, string.Empty);
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
                    LstDisplay.Items.Add(string.Empty);
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
            sbNodeBuffer.Append(input);
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

            foreach (string o in corruptNodes)
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
            originalText = fallbackTagsAppended.Aggregate(originalText, (current, o) => current.Replace(o.ToString(), string.Empty));

            // each set of fallback tags should now be removed from the text
            // set it to the global variable so we can add it back into document.xml
            FixedFallback = originalText;
        }

        private void BtnListConnections_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                PreButtonClickWork();

                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, false))
                {
                    WorkbookPart wbPart = excelDoc.WorkbookPart;
                    ConnectionsPart cPart = wbPart.ConnectionsPart;

                    if (cPart == null)
                    {
                        LogInformation(LogType.EmptyCount, StringResources.wConnections, string.Empty);
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

                            if (c.ConnectionFile != null)
                            {
                                LstDisplay.Items.Add(string.Empty);
                                LstDisplay.Items.Add("    Connection File= " + c.ConnectionFile);
                                
                                if (c.OlapProperties != null)
                                {
                                    LstDisplay.Items.Add("    Row Drill Count= " + c.OlapProperties.RowDrillCount);
                                }
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
                LogInformation(LogType.LogException, "List Connections Failed", ex.Message);
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
                PreButtonClickWork();

                if (fileType == StringResources.wWord)
                {
                    using (WordprocessingDocument document = WordprocessingDocument.Open(TxtFileName.Text, false))
                    {
                        AddCustomDocPropsToList(document.CustomFilePropertiesPart);
                    }
                }
                else if (fileType == StringResources.wExcel)
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(TxtFileName.Text, false))
                    {
                        AddCustomDocPropsToList(document.CustomFilePropertiesPart);
                    }
                }
                else if (fileType == StringResources.wPowerpoint)
                {
                    using (PresentationDocument document = PresentationDocument.Open(TxtFileName.Text, false))
                    {
                        AddCustomDocPropsToList(document.CustomFilePropertiesPart);
                    }
                }
                else
                {
                    LoggingHelper.Log("BtnListCC - unknown app");
                    return;
                }
            }
            catch (IOException ioe)
            {
                LogInformation(LogType.LogException, StringResources.wCustomDocProps, ioe.Message);
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnListCustomProps Error: ", ex.Message);
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
                LogInformation(LogType.EmptyCount, StringResources.wCustomDocProps, string.Empty);
                return;
            }

            int count = 0;
            
            foreach (var v in CfpList(cfp))
            {
                count++;
                LstDisplay.Items.Add(count + StringResources.wPeriod + v);
            }

            if (count == 0)
            {
                LogInformation(LogType.EmptyCount, StringResources.wCustomDocProps, string.Empty);
            }
        }

        public List<string> CfpList(CustomFilePropertiesPart part)
        {
            List<string> val = new List<string>();
            foreach (CustomDocumentProperty cdp in part.RootElement)
            {
                val.Add(cdp.Name + StringResources.wColonBuffer + cdp.InnerText);
            }
            return val;
        }

        private void BtnSetCustomProps_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();
            FrmCustomProperties cFrm = new FrmCustomProperties(TxtFileName.Text, fileType)
            {
                Owner = this
            };
            cFrm.ShowDialog();
        }

        private void BtnSetPrintOrientation_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();
            FrmPrintOrientation pFrm = new FrmPrintOrientation(TxtFileName.Text)
            {
                Owner = this
            };
            pFrm.ShowDialog();

        }

        private void CopyOutputToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyAllItems();
        }

        private void BtnViewParagraphs_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            PreButtonClickWork();
            FrmParagraphs paraFrm = new FrmParagraphs(TxtFileName.Text)
            {
                Owner = this
            };
            paraFrm.ShowDialog();
            Cursor = Cursors.Default;
        }

        private void BtnConvertXlsm2Xlsx_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();
            ConvertToNonMacro(StringResources.wExcel);
        }

        private void BtnConvertPptmToPptx_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();
            ConvertToNonMacro(StringResources.wPowerpoint);
        }

        private void BtnListPackageParts_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();

            foreach (var o in pParts)
            {
                LstDisplay.Items.Add(o);
            }

            // this is the only button that sort makes sense
            // setting to sort here instead of using the check box
            // the prebuttonclickwork function should set the list back to unsorted
            LstDisplay.Sorted = true;
        }
        
        private void SettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmSettings form = new FrmSettings();
            form.Show();
        }

        private void ErrorLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmErrorLog errFrm = new FrmErrorLog()
            {
                Owner = this
            };
            errFrm.ShowDialog();
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AppExitWork();
            Application.Exit();
        }

        private void BtnListFieldCodes_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                PreButtonClickWork();
                using (WordprocessingDocument package = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    IEnumerable<Run> rList = package.MainDocumentPart.Document.Descendants<Run>();
                    IEnumerable<Paragraph> pList = package.MainDocumentPart.Document.Descendants<Paragraph>();

                    List<string> fieldCharList = new List<string>();
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
                                    fieldCharList.Add(StringResources.sBegin);
                                }
                                else if (fc.FieldCharType == StringResources.sEnd)
                                {
                                    fieldCharList.Add(StringResources.sEnd);
                                }
                            }
                            else if (oxe.LocalName == "instrText")
                            {
                                fieldCharList.Add(oxe.InnerText);
                            }
                        }
                    }

                    foreach (Paragraph p in pList)
                    {
                        foreach (OpenXmlElement oxe in p.ChildElements)
                        {
                            if (oxe.LocalName == "fldSimple")
                            {
                                SimpleField sf = new SimpleField();
                                sf = (SimpleField)oxe;
                                fieldCodeList.Add(sf.Instruction);
                            }
                        }
                    }

                    if (fieldCharList.Count == 0 && fieldCodeList.Count == 0)
                    {
                        LogInformation(LogType.EmptyCount, StringResources.wFldCodes, string.Empty);
                        return;
                    }
                    else
                    {
                        StringBuilder sb = new StringBuilder();
                        int fCount = 0;

                        foreach (string s in fieldCharList)
                        {
                            if (s == StringResources.sBegin)
                            {
                                continue;
                            }
                            else if (s == StringResources.sEnd)
                            {
                                // display the field code values
                                fCount++;
                                LstDisplay.Items.Add(fCount + StringResources.wPeriod + sb);
                                sb.Clear();
                            }
                            else
                            {
                                sb.Append(s);
                            }
                        }

                        foreach (string s in fieldCodeList)
                        {
                            fCount++;
                            LstDisplay.Items.Add(fCount + StringResources.wPeriod + s);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LstDisplay.Items.Add(StringResources.wErrorText + ex.Message);
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
                PreButtonClickWork();
                bool hasBookmark = false;

                using (WordprocessingDocument package = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    IEnumerable<BookmarkStart> bkList = package.MainDocumentPart.Document.Descendants<BookmarkStart>();
                    LstDisplay.Items.Add("** Document Bookmarks **");

                    if (bkList.Count() > 0)
                    {
                        int count = 1;
                        hasBookmark = true;

                        foreach (BookmarkStart bk in bkList)
                        {
                            var cElem = bk.Parent;
                            var pElem = bk.Parent;
                            bool endLoop = false;
                            string isCorruptText = string.Empty;

                            do
                            {
                                if (cElem != null && cElem.Parent != null && cElem.Parent.ToString().Contains(StringResources.dfowSdt))
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
                                                    isCorruptText = " <-- ## Warning ## - this bookmark is in a plain text content control which is not allowed";
                                                    endLoop = true;
                                                }
                                            }
                                        }
                                    }
                                    
                                    // set next element
                                    pElem = cElem.Parent;
                                    cElem = pElem;
                                }
                                else
                                {
                                    // if the next element is null, bail
                                    if (cElem == null || cElem.Parent == null)
                                    {
                                        endLoop = true;
                                    }

                                    // set next element
                                    pElem = cElem.Parent;
                                    cElem = pElem;

                                    // if the parent is body, we can stop looping up
                                    // otherwise, we can continue moving up the element chain
                                    if (pElem != null && pElem.ToString() == StringResources.dfowBody)
                                    {
                                        endLoop = true;
                                    }
                                }
                            } while (endLoop == false);

                            LstDisplay.Items.Add(count + StringResources.wPeriod + bk.Name + isCorruptText);
                            count++;
                        }
                    }

                    if (package.MainDocumentPart.WordprocessingCommentsPart != null)
                    {
                        if (package.MainDocumentPart.WordprocessingCommentsPart.Comments != null)
                        {
                            IEnumerable<BookmarkStart> bkCommentList = package.MainDocumentPart.WordprocessingCommentsPart.Comments.Descendants<BookmarkStart>();
                            int bkCommentCount = 0;

                            if (bkCommentList.Count() > 0)
                            {
                                LstDisplay.Items.Add(string.Empty);
                                LstDisplay.Items.Add("** Comment Bookmarks ** ");
                                hasBookmark = true;

                                foreach (BookmarkStart bkc in bkCommentList)
                                {
                                    bkCommentCount++;
                                    LstDisplay.Items.Add(bkCommentCount + StringResources.wPeriod + bkc.Name);
                                }
                            }
                        }
                    }

                    if (hasBookmark == false)
                    {
                        LstDisplay.Items.Add(" None");
                    }
                }
            }
            catch (Exception ex)
            {
                LstDisplay.Items.Add(StringResources.wErrorText + ex.Message);
                LoggingHelper.Log("BtnListBookmarks: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// clear and unsort the list
        /// </summary>
        public void PreButtonClickWork()
        {
            LstDisplay.Items.Clear();
            LstDisplay.Sorted = false;
        }

        private void BtnListCC_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                using (WordprocessingDocument package = WordprocessingDocument.Open(TxtFileName.Text, false))
                {
                    PreButtonClickWork();
                    int count = 0;

                    foreach (var cc in package.ContentControls())
                    {
                        string ccType = string.Empty;
                        bool PropFound = false;
                        SdtProperties props = cc.Elements<SdtProperties>().FirstOrDefault();

                        // loop the properties and get the type
                        foreach (OpenXmlElement oxe in props.ChildElements)
                        {
                            if (oxe.GetType().Name == "SdtContentText")
                            {
                                ccType = "Plain Text";
                                PropFound = true;
                            }

                            if (oxe.GetType().Name == "SdtContentDropDownList")
                            {
                                ccType = "Drop Down List";
                                PropFound = true;
                            }

                            if (oxe.GetType().Name == "SdtContentDocPartList")
                            {
                                ccType = "Building Block Gallery";
                                PropFound = true;
                            }

                            if (oxe.GetType().Name == "SdtContentCheckBox")
                            {
                                ccType = "Check Box";
                                PropFound = true;
                            }

                            if (oxe.GetType().Name == "SdtContentPicture")
                            {
                                ccType = "Picture";
                                PropFound = true;
                            }

                            if (oxe.GetType().Name == "SdtContentComboBox")
                            {
                                ccType = "Combo Box";
                                PropFound = true;
                            }

                            if (oxe.GetType().Name == "SdtContentDate")
                            {
                                ccType = "Date Picker";
                                PropFound = true;
                            }

                            if (oxe.GetType().Name == "SdtRepeatedSection")
                            {
                                ccType = "Repeating Section";
                                PropFound = true;
                            }
                        }

                        // display the cc type
                        count++;
                        if (PropFound == true)
                        {
                            LstDisplay.Items.Add(count + StringResources.wPeriod + ccType);
                        }
                        else
                        {
                            LstDisplay.Items.Add(count + StringResources.wPeriod + "Rich Text");
                        }
                        
                    }

                    if (count == 0)
                    {
                        LogInformation(LogType.EmptyCount, StringResources.wContentControls, string.Empty);
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnListCC Error:", ex.Message);
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
                PreButtonClickWork();
                int count = 0;

                if (fileType == StringResources.wWord)
                {
                    // with Word, we can just run through the entire body and get the shapes
                    using (WordprocessingDocument document = WordprocessingDocument.Open(TxtFileName.Text, false))
                    {
                        foreach (ChartPart c in document.MainDocumentPart.ChartParts)
                        {
                            count++;
                            LstDisplay.Items.Add(count + StringResources.wPeriod + c.Uri + StringResources.wArrow + StringResources.wShpChart);
                        }

                        foreach (AO.Shape shape in document.MainDocumentPart.Document.Body.Descendants<AO.Shape>())
                        {
                            count++;
                            LstDisplay.Items.Add(count + StringResources.shpOfficeDrawing);
                        }

                        foreach (O.Vml.Shape shape in document.MainDocumentPart.Document.Body.Descendants<O.Vml.Shape>())
                        {
                            count++;
                            LstDisplay.Items.Add(count + StringResources.wPeriod + shape.Id + StringResources.wArrow + StringResources.shpVml);
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
                    }
                }
                else if (fileType == StringResources.wExcel)
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
                                LstDisplay.Items.Add(count + StringResources.wPeriod + shape.Id + StringResources.wArrow + StringResources.shpVml);
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
                        }
                    }
                }
                else if (fileType == StringResources.wPowerpoint)
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
                                    if (child1.GetType().ToString() == StringResources.dfopNVSP)
                                    {
                                        foreach (OpenXmlElement child2 in child1.ChildElements)
                                        {
                                            if (child2.GetType().ToString() == StringResources.dfopNVDP)
                                            {
                                                P.NonVisualDrawingProperties nvdp = (P.NonVisualDrawingProperties)child2;
                                                LstDisplay.Items.Add(count + StringResources.wPeriod + nvdp.Name);
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
                                LstDisplay.Items.Add(count + StringResources.wPeriod + shape.Id + StringResources.wArrow + StringResources.shpVml);
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
                        }
                    }
                }
                else
                {
                    return;
                }

                if (count == 0)
                {
                    LogInformation(LogType.EmptyCount, StringResources.wShapes, string.Empty);
                }
            }
            catch (IOException ioe)
            {
                LogInformation(LogType.LogException, "Error listing shapes.", ioe.Message);
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
                LogInformation(LogType.LogException, "BtnCopyLineOutput Error", ex.Message);
            }
        }

        private void BtnCopyAll_Click(object sender, EventArgs e)
        {
            CopyAllItems();
        }

        public void AppExitWork()
        {
            try
            {
                if (Properties.Settings.Default.DeleteCopiesOnExit == true)
                {
                    File.Delete(StrCopiedFileName);
                }

                Properties.Settings.Default.ErrorLog.Clear();
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("App Exit Error: " + ex.Message);
            }
            finally
            {
                Application.Exit();
            }
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
                LogInformation(LogType.LogException, "BtnCopyOutput Error", ex.Message);
            }
        }

        public void FixNotesPageSizeCustom()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                PowerPointOpenXml.UseCustomNotesPageSize(TxtFileName.Text);
                if (Properties.Settings.Default.ResetNotesMaster == false)
                {
                    MessageBox.Show(StringResources.resetNotesMasterRegKey, StringResources.mbWarning, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                LogInformation(LogType.ClearAndAdd, TxtFileName.Text + StringResources.wColonBuffer + StringResources.pptNotesSizeReset, string.Empty);
            }
            catch (NullReferenceException nre)
            {
                LogInformation(LogType.LogException, "FixNotesPageSizeCustom Error", nre.Message);
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "FixNotesPageSizeCustom Error", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public void FixNotesPageSizeDefault()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                using (PresentationDocument document = PresentationDocument.Open(TxtFileName.Text, true))
                {
                    PowerPointOpenXml.ChangeNotesPageSize(document);
                    if (Properties.Settings.Default.ResetNotesMaster == false)
                    {
                        MessageBox.Show(StringResources.resetNotesMasterRegKey, StringResources.mbWarning, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    LogInformation(LogType.ClearAndAdd, TxtFileName.Text + StringResources.wColonBuffer + StringResources.pptNotesSizeReset, string.Empty);
                }
            }
            catch (NullReferenceException nre)
            {
                LogInformation(LogType.EmptyCount, StringResources.wNotesMaster, string.Empty);
                LogInformation(LogType.LogException, "FixNotesPageSizeDefault Error", nre.Message);
            }
            catch (Exception ex)
            {
                LogInformation(LogType.ClearAndAdd, ex.Message, string.Empty);
                LogInformation(LogType.LogException, "FixNotesPageSizeDefault Error", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void FeedbackToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(StringResources.helpLocation);
        }

        private void BtnPPTRemovePII_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                PreButtonClickWork();
                using (PresentationDocument document = PresentationDocument.Open(TxtFileName.Text, true))
                {
                    document.PresentationPart.Presentation.RemovePersonalInfoOnSave = false;
                    document.PresentationPart.Presentation.Save();
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnPPTRemovePII", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public void FixRevisions()
        {
            try
            {
                LstDisplay.Items.Clear();
                bool isFixed = false;
                Cursor = Cursors.WaitCursor;

                using (WordprocessingDocument document = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    Document doc = document.MainDocumentPart.Document;
                    var deleted = doc.Descendants<DeletedRun>().ToList();

                    // loop each DeletedRun
                    foreach (DeletedRun dr in deleted)
                    {
                        foreach (OpenXmlElement oxedr in dr)
                        {
                            // if we have a run, we need to look for Text tags
                            if (oxedr.GetType().ToString() == StringResources.dfowRun)
                            {
                                Run r = (Run)oxedr;
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

                    // now save the file if we have changes
                    if (isFixed == true)
                    {
                        doc.Save();
                        LstDisplay.Items.Add("** Fixed Corrupt Revisions **");
                    }
                    else
                    {
                        LstDisplay.Items.Add("** No Corrupt Revisions Found **");
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnFixCorruptRevisions", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public void FixBookmarks()
        {
            LstDisplay.Items.Clear();

            // there are currently two different bookmark corruptions, check for both
            if (WordOpenXml.RemoveMissingBookmarkTags(TxtFileName.Text) == true || WordOpenXml.RemovePlainTextCcFromBookmark(TxtFileName.Text) == true)
            {
                LstDisplay.Items.Add("** Fixed Corrupt Bookmarks **");
            }
            else
            {
                LstDisplay.Items.Add("** No Corrupt Bookmarks Found **");
            }
        }

        /// <summary>
        /// there are times when a hyperlink is added in between a field code sequence
        /// typically it is begin-separate-end OR begin-end
        /// however, there are times when we see begin-hyperlink-end
        /// that is not valid, this is used to fix those situations
        /// </summary>
        public void FixCorruptHyperlinks()
        {
            try
            {
                LstDisplay.Items.Clear();
                Cursor = Cursors.WaitCursor;

                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    bool fileChanged = false;
                    bool isHyperlinkInBetweenSequence = false;
                    IEnumerable<Paragraph> paras = myDoc.MainDocumentPart.Document.Descendants<Paragraph>();
                    int pCount = paras.Count();
                    int tempCount = 0;

                    // need to keep looping paragraphs until we don't find a bad hlink position
                    do
                    {
                        isHyperlinkInBetweenSequence = false;
                        bool inBeginEndSequence = false;
                        int beginPosition = 0;
                        int endPosition = 0;
                        int elementCount = 0;
                        int beginCount = 0;
                        int prevRunPosition = 0;
                        tempCount = 0;

                        foreach (Paragraph p in paras)
                        {
                            tempCount++;
                            foreach (OpenXmlElement oxe in p.Descendants<OpenXmlElement>())
                            {
                                // keep track of previous run so we can get the right start position
                                // you could just use the first "begin" field code, but fixing it back up later is more challenging
                                // if we grab the begins root run, it makes this much easier
                                elementCount++;
                                if (oxe.GetType().Name == "Run")
                                {
                                    prevRunPosition = elementCount;
                                }

                                // here we are keeping track of begin-end sequences
                                // the beginCount is there for nested begin-end scenarios
                                // you can have begin-begin-separate-end-end
                                if (oxe.GetType().Name == "FieldChar")
                                {
                                    FieldChar fc = (FieldChar)oxe;
                                    if (fc.FieldCharType == FieldCharValues.Begin)
                                    {
                                        beginCount++;
                                        inBeginEndSequence = true;
                                        if (beginPosition == 0)
                                        {
                                            beginPosition = prevRunPosition;
                                        }
                                    }

                                    if (fc.FieldCharType == FieldCharValues.End)
                                    {
                                        // valid sequence, reset values
                                        beginCount--;
                                        if (beginCount == 0)
                                        {
                                            inBeginEndSequence = false;
                                            beginPosition = 0;
                                        }
                                    }
                                }

                                // if we are still in the middle of a begin-end sequence
                                // we can't have a hlink so we know we have a corruption
                                if (oxe.GetType().Name == "Hyperlink" && inBeginEndSequence == true)
                                {
                                    // you can have a hlink in between the begin-end tags or vica versa
                                    // so we are only looking for an hlink that has an end inside it with no begin
                                    if (oxe.InnerXml.Contains(StringResources.txtFieldCodeEnd) && !oxe.InnerXml.Contains(StringResources.txtFieldCodeBegin))
                                    {
                                        isHyperlinkInBetweenSequence = true;
                                        endPosition = elementCount;
                                        break;
                                    }
                                }
                            }

                            if (isHyperlinkInBetweenSequence == true)
                            {
                                break;
                            }
                        }

                        // if isHyperlinkInBetween we need to loop again now that we have the bad position
                        if (isHyperlinkInBetweenSequence == true)
                        {
                            int tCount = 0;
                            bool atEndPosition = false;
                            List<OpenXmlElement> els = new List<OpenXmlElement>();
                            List<Run> runs = new List<Run>();

                            foreach (Paragraph p in paras)
                            {
                                foreach (OpenXmlElement oxe in p.Descendants<OpenXmlElement>())
                                {
                                    tCount++;
                                    if (tCount >= beginPosition)
                                    {
                                        // once we are at the right position where we need to start moving elements around
                                        // create a list of elements
                                        els.Add(oxe);
                                        if (tCount == endPosition)
                                        {
                                            // once we are at the end of the bad sequence
                                            // reverse the list so we can prepend to the hyperlink
                                            atEndPosition = true;
                                            els.Reverse();

                                            foreach (OpenXmlElement e in els)
                                            {
                                                if (e.LocalName != "hyperlink")
                                                {
                                                    // haven't quite figured out why, but I need to create a run for all non-run elements
                                                    // then I can clone it and add it back to the right location in the hyperlink
                                                    // then we can remove the original bad position elements
                                                    if (e.LocalName != "r")
                                                    {
                                                        Run r = new Run();
                                                        r.AppendChild(e.CloneNode(true));
                                                        oxe.PrependChild(r);
                                                        e.Remove();
                                                        fileChanged = true;
                                                    }
                                                    else
                                                    {
                                                        oxe.PrependChild(e.CloneNode(false));
                                                        e.Remove();
                                                        fileChanged = true;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                if (atEndPosition == true)
                                {
                                    break;
                                }
                            }
                        }
                    } while (tempCount < pCount);

                    if (fileChanged)
                    {
                        myDoc.MainDocumentPart.Document.Save();
                        LstDisplay.Items.Add("** Hyperlinks Fixed **");
                    }
                    else
                    {
                        LstDisplay.Items.Add("** No Corrupt Hyperlinks Found **");
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnFixCorruptHyperlinks", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public void AddSeparateToMentionHyperlink(Paragraph p)
        {

        }

        /// <summary>
        /// This is a known scenario that affects Word Co-Auth scenarios
        /// if the field code sequence is missing the separate, it causes problems
        /// </summary>
        public void FixCommentFieldCodeTags()
        {
            try
            {
                LstDisplay.Items.Clear();
                Cursor = Cursors.WaitCursor;

                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    bool isFileChanged = false;
                    Regex emailPattern = new Regex(@"(.*?)<?(\b\S+@\S+\b)>?");
                    WordprocessingCommentsPart commentsPart = myDoc.MainDocumentPart.WordprocessingCommentsPart;
                    foreach (O.Wordprocessing.Comment cmt in commentsPart.Comments)
                    {
                        IEnumerable<Paragraph> pList = cmt.Descendants<Paragraph>();
                        List<Run> rList = new List<Run>();

                        foreach (Paragraph p in pList)
                        {
                            // if the p has the mention style, it passes the first check we need to make
                            if (p.InnerXml.Contains(StringResources.txtAtMentionStyle))
                            {
                                bool beginFound = false;
                                bool separateFound = false;
                                string emailAlias = "";

                                // now we need to loop each run and check the separate is missing
                                foreach (Run r in p.Descendants<Run>())
                                {
                                    if (r.InnerXml.Contains(StringResources.txtFieldCodeBegin))
                                    {
                                        beginFound = true;
                                    }

                                    if (r.InnerXml.Contains(StringResources.txtFieldCodeSeparate))
                                    {
                                        separateFound = true;
                                    }

                                    // hold onto the mailto so we at least have something to use for the mention text
                                    if (beginFound == true && r.InnerXml.Contains("mailto"))
                                    {
                                        var groups = emailPattern.Match(r.InnerXml.ToString()).Groups;

                                        // trim the mailto
                                        emailAlias = groups[2].Value.Remove(0, 7);
                                    }

                                    // once we get to the end, if we haven't found a separate, we need to add it back
                                    if (r.InnerXml.Contains(StringResources.txtFieldCodeEnd))
                                    {
                                        if (r.InnerXml.Contains(StringResources.txtAtMentionStyle) && separateFound == false)
                                        {
                                            // first, remove all children since we are in the area we need to change
                                            // add the new separate run
                                            // add the new text run with the mailto
                                            // add the new end run back
                                            r.RemoveAllChildren();

                                            Run rNewSeparate = new Run();
                                            O.Wordprocessing.RunProperties rPr = new O.Wordprocessing.RunProperties();
                                            RunStyle rs = new RunStyle();
                                            rs.Val = "Mention";
                                            rPr.Append(rs);
                                            rNewSeparate.Append(rPr);
                                            FieldChar fcs = new FieldChar();
                                            fcs.FieldCharType = FieldCharValues.Separate;
                                            rNewSeparate.Append(fcs);
                                            r.Append(rNewSeparate);

                                            Run rNewText = new Run();
                                            O.Wordprocessing.Text t = new O.Wordprocessing.Text();
                                            t.Text = emailAlias;
                                            rNewText.Append(t);
                                            r.Append(rNewText);

                                            Run rNewEnd = new Run();
                                            FieldChar fce = new FieldChar();
                                            fce.FieldCharType = FieldCharValues.End;
                                            rNewEnd.Append(fce);
                                            r.Append(rNewEnd);

                                            isFileChanged = true;
                                        }

                                        // reset logic criteria
                                        beginFound = false;
                                        separateFound = false;
                                        emailAlias = "";
                                    }
                                }
                            }
                        }
                    }

                    if (isFileChanged)
                    {
                        myDoc.MainDocumentPart.Document.Save();
                        LstDisplay.Items.Add("** Corrupt Hyperlinks Fixed **");
                    }
                    else
                    {
                        LstDisplay.Items.Add("** No Corrupt Hyperlinks Found **");
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnFixCommentFieldCodeTags", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// There are times when documents have dangling comment refs where document.xml will reference a comment id
        /// the problem is that comment no longer exists for unknown reasons
        /// this fix should get the id's for those comments and remove them
        /// </summary>
        public void FixMissingCommentRef()
        {
            try
            {
                LstDisplay.Items.Clear();
                Cursor = Cursors.WaitCursor;

                using (WordprocessingDocument document = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    WordprocessingCommentsPart commentsPart = document.MainDocumentPart.WordprocessingCommentsPart;
                    IEnumerable<OpenXmlUnknownElement> unknownList = document.MainDocumentPart.Document.Descendants<OpenXmlUnknownElement>();
                    bool saveFile = false;
                    bool cRefIdExists = false;

                    if (commentsPart == null)
                    {
                        LogInformation(LogType.EmptyCount, StringResources.wComments, string.Empty);
                    }
                    else
                    {
                        foreach (OpenXmlUnknownElement uk in unknownList)
                        {
                            // for some reason these dangling refs are considered unknown types, not commentrefs
                            // convert to an openxmlelement then type it to a commentref to get the id
                            if (uk.LocalName == "commentReference")
                            {
                                // so far I only see the id in the outerxml
                                XmlDocument xDoc = new XmlDocument();
                                xDoc.LoadXml(uk.OuterXml);

                                // traverse the outerxml until we get to the id
                                if (xDoc.ChildNodes.Count > 0)
                                {
                                    foreach (XmlNode xNode in xDoc.ChildNodes)
                                    {
                                        if (xNode.Attributes.Count > 0)
                                        {
                                            foreach (XmlAttribute xa in xNode.Attributes)
                                            {
                                                if (xa.LocalName == "id")
                                                {
                                                    // now that we have the id number, we can use it to compare with the comment part
                                                    // if the id exists in commentref but not the commentpart, it can be deleted
                                                    foreach (O.Wordprocessing.Comment cm in commentsPart.Comments)
                                                    {
                                                        int cId = Convert.ToInt32(cm.Id);
                                                        int cRefId = Convert.ToInt32(xa.Value);
                                                        
                                                        if (cId == cRefId)
                                                        {
                                                            cRefIdExists = true;
                                                        }
                                                    }

                                                    if (cRefIdExists == false)
                                                    {
                                                        uk.Remove();
                                                        saveFile = true;
                                                    }
                                                    else
                                                    {
                                                        cRefIdExists = false;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (saveFile)
                        {
                            document.MainDocumentPart.Document.Save();
                            LstDisplay.Items.Add("** Corrupt Comment Fixed **");
                        }
                        else
                        {
                            LstDisplay.Items.Add("** No Corrupt Comment Found **");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnFixCorruptComments", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// this fix is for a known issue where files contain a table
        /// with a tblGrid element before the first table row, that is not valid per the schema
        /// </summary>
        public void FixTblGrid()
        {
            try
            {
                LstDisplay.Items.Clear();
                Cursor = Cursors.WaitCursor;

                using (WordprocessingDocument document = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    // "global" document variables
                    bool tblModified = false;
                    OpenXmlElement tgClone = null;

                    if (WordOpenXml.IsPartNull(document, "Table") == false)
                    {
                        // get the list of tables in the document
                        List<O.Wordprocessing.Table> tbls = document.MainDocumentPart.Document.Descendants<O.Wordprocessing.Table>().ToList();

                        foreach (O.Wordprocessing.Table tbl in tbls)
                        {
                            // you can have only one tblGrid per table, including nested tables
                            // it needs to be before any row elements so sequence is
                            // 1. check if the tblGrid element is before any table row
                            // 2. check for multiple tblGrid elements
                            bool tRowFound = false;
                            bool tGridBeforeRowFound = false;
                            int tGridCount = 0;

                            foreach (OpenXmlElement oxe in tbl.Elements())
                            {
                                // flag if we found a trow, once we find 1, the rest do not matter
                                if (oxe.GetType().Name == "TableRow")
                                {
                                    tRowFound = true;
                                }

                                // when we get to a tablegrid, we have a few things to check
                                // 1. have we found a table row previously
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

                            // if we had a table grid before a row was found, move it before the first row in the table
                            if (tGridBeforeRowFound == true)
                            {
                                tbl.InsertBefore(tgClone, tbl.GetFirstChild<TableRow>());
                                tblModified = true;
                            }
                        }
                    }

                    // save the file if we modified the table
                    if (tblModified == true)
                    {
                        document.MainDocumentPart.Document.Save();
                        LstDisplay.Items.Add("** Table Fix Completed **");
                    }
                    else
                    {
                        LstDisplay.Items.Add("** No Corrupt Table Found **");
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "FixTbleGrid", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// When the ListTemplate count is too large, Word no longer displays bullets,
        /// this function will try to find 1 single and 1 multi-level bullet,
        /// then go through the document and apply one of those to each bullet used
        /// which should get the document under the count limitation in the document
        /// </summary>
        public void FixListNumbering()
        {
            try
            {
                LstDisplay.Items.Clear();
                Cursor = Cursors.WaitCursor;

                NumberingHelper bulletMultiLevelNumberingValues = new NumberingHelper();
                NumberingHelper bulletSingleLevelNumberingValues = new NumberingHelper();

                List<int> bulletMultiLevelNumIdsInUse = new List<int>();
                List<int> bulletSingleLevelNumIdsInUse = new List<int>();

                using (WordprocessingDocument document = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    if (document.MainDocumentPart.NumberingDefinitionsPart == null)
                    {
                        LstDisplay.Items.Add("** No List Templates Found **");
                        return;
                    }

                    // get the list of numId's and AbstractNum's in numbering.xml
                    var absNumsInUseList = document.MainDocumentPart.NumberingDefinitionsPart.Numbering.Descendants<AbstractNum>().ToList();
                    var numInstancesInUseList = document.MainDocumentPart.NumberingDefinitionsPart.Numbering.Descendants<NumberingInstance>().ToList();

                    bool bulletSingleLevelFound = false;
                    bool bulletMultiLevelFound = false;

                    foreach (AbstractNum an in absNumsInUseList)
                    {
                        foreach (NumberingInstance ni in numInstancesInUseList)
                        {
                            // if the abstractnum and numId match, they are the same listtemplate
                            if (ni.AbstractNumId.Val == an.AbstractNumberId.Value)
                            {
                                // get the level count
                                var lvlNumberingList = an.Descendants<Level>().ToList();

                                // since we have the list template, find out if it is a bullet
                                foreach (OpenXmlElement anChild in an)
                                {
                                    if (anChild.GetType().ToString() == StringResources.dfowLevel)
                                    {
                                        Level lvl = (Level)anChild;
                                        
                                        // try to catch different "types" of numberingformat
                                        // for now, I'm only checking for a single and multi-level bullets
                                        if (lvl.NumberingFormat.Val == "bullet" && lvlNumberingList.Count > 1 && lvl.LevelIndex == 0)
                                        {
                                            // if level is > 1, this is a multi level list
                                            bulletMultiLevelNumIdsInUse.Add(ni.NumberID);

                                            if (bulletMultiLevelFound == false)
                                            {
                                                bulletMultiLevelNumberingValues.AbsNumId = ni.AbstractNumId.Val;
                                                bulletMultiLevelNumberingValues.NumFormat = "bulletMultiLevel";
                                                bulletMultiLevelNumberingValues.NumId = ni.NumberID;
                                                bulletMultiLevelFound = true;
                                            }
                                        }
                                        else if (lvl.NumberingFormat.Val == "bullet" && lvlNumberingList.Count == 1 && lvl.LevelIndex == 0)
                                        {
                                            // if level = 1, this is a single level list
                                            bulletSingleLevelNumIdsInUse.Add(ni.NumberID);

                                            if (bulletSingleLevelFound == false)
                                            {
                                                bulletSingleLevelNumberingValues.AbsNumId = ni.AbstractNumId.Val;
                                                bulletSingleLevelNumberingValues.NumFormat = "bulletSingle";
                                                bulletSingleLevelNumberingValues.NumId = ni.NumberID;
                                                bulletSingleLevelFound = true;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // now that we have bullet numids to use, we can apply it to each paragraph
                    MainDocumentPart mainPart = document.MainDocumentPart;
                    StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart;

                    foreach (OpenXmlElement el in mainPart.Document.Descendants<Paragraph>())
                    {
                        if (el.Descendants<NumberingId>().Count() > 0)
                        {
                            foreach (NumberingId pNumId in el.Descendants<NumberingId>())
                            {
                                foreach (var o in bulletMultiLevelNumIdsInUse)
                                {
                                    if (o == pNumId.Val)
                                    {
                                        pNumId.Val = bulletMultiLevelNumberingValues.NumId;
                                    }
                                }

                                foreach (var o in bulletSingleLevelNumIdsInUse)
                                {
                                    if (o == pNumId.Val)
                                    {
                                        pNumId.Val = bulletSingleLevelNumberingValues.NumId;
                                    }
                                }
                            }
                        }
                    }

                    foreach (HeaderPart hdrPart in mainPart.HeaderParts)
                    {
                        foreach (OpenXmlElement el in hdrPart.Header.Elements())
                        {
                            foreach (NumberingId hNumId in el.Descendants<NumberingId>())
                            {
                                foreach (var o in bulletMultiLevelNumIdsInUse)
                                {
                                    if (o == hNumId.Val)
                                    {
                                        hNumId.Val = bulletMultiLevelNumberingValues.NumId;
                                    }
                                }

                                foreach (var o in bulletSingleLevelNumIdsInUse)
                                {
                                    if (o == hNumId.Val)
                                    {
                                        hNumId.Val = bulletSingleLevelNumberingValues.NumId;
                                    }
                                }
                            }
                        }
                    }

                    foreach (FooterPart ftrPart in mainPart.FooterParts)
                    {
                        foreach (OpenXmlElement el in ftrPart.Footer.Elements())
                        {
                            foreach (NumberingId fNumId in el.Descendants<NumberingId>())
                            {
                                foreach (var o in bulletMultiLevelNumIdsInUse)
                                {
                                    if (o == fNumId.Val)
                                    {
                                        fNumId.Val = bulletMultiLevelNumberingValues.NumId;
                                    }
                                }

                                foreach (var o in bulletSingleLevelNumIdsInUse)
                                {
                                    if (o == fNumId.Val)
                                    {
                                        fNumId.Val = bulletSingleLevelNumberingValues.NumId;
                                    }
                                }
                            }
                        }
                    }

                    foreach (OpenXmlElement el in stylePart.Styles.Elements())
                    {
                        if (el.GetType().ToString() == StringResources.dfowStyle)
                        {
                            string styleEl = el.GetAttribute("styleId", StringResources.wordMainAttributeNamespace).Value;
                            int pStyle = WordExtensionClass.ParagraphsByStyleName(mainPart, styleEl).Count();

                            if (pStyle > 0)
                            {
                                foreach (NumberingId sEl in el.Descendants<NumberingId>())
                                {
                                    foreach (var o in bulletMultiLevelNumIdsInUse)
                                    {
                                        if (o == sEl.Val)
                                        {
                                            sEl.Val = bulletMultiLevelNumberingValues.NumId;
                                        }
                                    }

                                    foreach (var o in bulletSingleLevelNumIdsInUse)
                                    {
                                        if (o == sEl.Val)
                                        {
                                            sEl.Val = bulletSingleLevelNumberingValues.NumId;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    document.MainDocumentPart.Document.Save();
                    LstDisplay.Items.Add("** Numbering Fix Completed **");
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnFixListNumbering", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// there are times when endnotes get inflated with duplicate content
        /// if there are more than 1000 runs of content in a single endnote
        /// this will keep the first endnote paragraph and delete the rest
        /// TODO: would be nice to delete duplicates only as an option
        /// TODO: not sure what could be considered an excessive amount of runs
        /// </summary>
        public void FixEndnotes()
        {
            try
            {
                LstDisplay.Items.Clear();
                Cursor = Cursors.WaitCursor;

                using (WordprocessingDocument document = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    bool corruptEndnotesFound = false;

                    if (document.MainDocumentPart.EndnotesPart != null)
                    {
                        Endnotes ens = document.MainDocumentPart.EndnotesPart.Endnotes;

                        foreach (Endnote en in ens)
                        {
                            // get the paragraph list from the endnote, if it has more than 1000 runs of content
                            // delete it...need to find a way to check for dupes
                            // for now just deleting all but the first paragraph run
                            var paraList = en.Descendants<Paragraph>().ToList();
                            foreach (var p in paraList)
                            {
                                var rList = p.Descendants<Run>().ToList();
                                if (rList.Count > 1000)
                                {
                                    int count = 0;
                                    foreach (var r in rList)
                                    {
                                        if (count > 0)
                                        {
                                            r.Remove();
                                            corruptEndnotesFound = true;
                                        }
                                        count++;
                                    }
                                }
                            }
                        }
                    }

                    if (corruptEndnotesFound == true)
                    {
                        document.MainDocumentPart.Document.Save();
                        LstDisplay.Items.Add("** Endnotes Fix Completed **");
                    }
                    else
                    {
                        LstDisplay.Items.Add("** No Corrupt Endnotes Found **");
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnFixEndnotes", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// Show the fix doc form for Word
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnFixDocument_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();

            using (var f = new FrmFixDocument(StringResources.wWord))
            {
                var result = f.ShowDialog();
                if (result == DialogResult.OK)
                {
                    string val = f.OptionSelected;
                    
                    switch (val)
                    {
                        case "Bookmark":
                            FixBookmarks();
                            break;
                        case "Endnote":
                            FixEndnotes();
                            break;
                        case "ListTemplates":
                            FixListNumbering();
                            break;
                        case "Revision":
                            FixRevisions();
                            break;
                        case "TblGrid":
                            FixTblGrid();
                            break;
                        case "FixComments":
                            FixMissingCommentRef();
                            break;
                        case "FixHyperlinks":
                            FixCorruptHyperlinks();
                            break;
                        case "FixCoAuthHyperlinks":
                            FixCommentFieldCodeTags();
                            break;
                        default:
                            LstDisplay.Items.Add("No Option Selected");
                            LoggingHelper.Log("BtnFixDocument - No Option Selected");
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// show the fix doc form for PPT
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnFixPresentation_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();
            using (var f = new FrmFixDocument(StringResources.wPowerpoint))
            {
                var result = f.ShowDialog();
                if (result == DialogResult.OK)
                {
                    string val = f.OptionSelected;

                    switch (val)
                    {
                        case "Notes":
                            FixNotesPageSizeDefault();
                            break;
                        case "NotesWithFile":
                            FixNotesPageSizeCustom();
                            break;
                        default:
                            LstDisplay.Items.Add("No Option Selected");
                            LoggingHelper.Log("BtnFixPresentation - No Option Selected");
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// this function uses the excelcnv.exe to convert a strict format xlsx to non-strict xlsx
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnConvertToNonStrictFormat_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                PreButtonClickWork();

                // check if the excelcnv.exe exists, without it, no conversion can happen
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
                    // if no path is found, we will be unable to convert
                    excelcnvPath = string.Empty;
                    LstDisplay.Items.Add("** Unable to convert file **");
                    return;
                }

                // check if the file is strict, no changes are made to the file yet
                bool isStrict = false;

                using (Package package = Package.Open(TxtFileName.Text, FileMode.Open, FileAccess.Read))
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
                                LoggingHelper.Log("BtnConvertToNonStrictFormat_Click ReadToEnd Error = " + ex.Message);
                            }
                        }
                    }
                }

                // if the file is strict format
                // run the command to convert it to non-strict
                if (isStrict == true)
                {
                    // setup destination file path
                    string strOriginalFile = TxtFileName.Text;
                    string strOutputPath = Path.GetDirectoryName(strOriginalFile) + "\\";
                    string strFileExtension = Path.GetExtension(strOriginalFile);
                    string strOutputFileName = strOutputPath + Path.GetFileNameWithoutExtension(strOriginalFile) + StringResources.wFixedFileParentheses + strFileExtension;

                    // run the command to convert the file "excelcnv.exe -nme -oice "strict-file-path" "converted-file-path""
                    string cParams = " -nme -oice " + '"' + TxtFileName.Text + '"' + StringResources.wSpaceChar + '"' + strOutputFileName + '"';
                    var proc = Process.Start(excelcnvPath, cParams);
                    proc.Close();
                    LstDisplay.Items.Add(StringResources.fileConvertSuccessful);
                    LstDisplay.Items.Add("File Location: " + strOutputFileName);
                }
                else
                {
                    LstDisplay.Items.Add("** File Is Not Strict Open Xml Format **");
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnConvertToNonStrictFormat_Click Error = ", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void FontViewerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmFontViewer fFrm = new FrmFontViewer(StringResources.sampleSentence)
            {
                Owner = this
            };
            fFrm.ShowDialog();
        }

        private void ClipboardViewerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmClipboardViewer cFrm = new FrmClipboardViewer()
            {
                Owner = this
            };
            cFrm.ShowDialog();
        }

        private void PrinterSettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmPrinterSettings pFrm = new FrmPrinterSettings()
            {
                Owner = this
            };
            pFrm.ShowDialog();
        }

        private void BtnListTransitions_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            PreButtonClickWork();

            try
            {
                using (PresentationDocument ppt = PresentationDocument.Open(TxtFileName.Text, false))
                {
                    int transitionCount = 0;

                    foreach (string s in PowerPointOpenXml.GetSlideTransitions(ppt))
                    {
                        transitionCount++;
                        LstDisplay.Items.Add(transitionCount + StringResources.wPeriod + s);
                    }

                    if (transitionCount == 0)
                    {
                        LogInformation(LogType.EmptyCount, StringResources.wTransitions, string.Empty);
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("BtnListTransitions Error: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnMoveSlide_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            PreButtonClickWork();
            try
            {
                using (PresentationDocument ppt = PresentationDocument.Open(TxtFileName.Text, true))
                {
                    FrmMoveSlide mvFrm = new FrmMoveSlide(ppt)
                    {
                        Owner = this
                    };
                    mvFrm.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("BtnMoveSlide Error: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// Form to allow deleting custom doc props
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnDeleteCustomProps_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                PreButtonClickWork();

                if (fileType == StringResources.wWord)
                {
                    using (WordprocessingDocument document = WordprocessingDocument.Open(TxtFileName.Text, true))
                    {
                        AddCustomDocPropsToList(document.CustomFilePropertiesPart);
                        LstDisplay.Items.Clear();

                        using (var f = new FrmDeleteCustomProps(document.CustomFilePropertiesPart))
                        {
                            var result = f.ShowDialog();
                            if (f.PartModified)
                            {
                                document.MainDocumentPart.Document.Save();
                                AddCustomDocPropsToList(document.CustomFilePropertiesPart);
                                
                            }
                        }
                    }
                }
                else if (fileType == StringResources.wExcel)
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(TxtFileName.Text, true))
                    {
                        AddCustomDocPropsToList(document.CustomFilePropertiesPart);
                        LstDisplay.Items.Clear();

                        using (var f = new FrmDeleteCustomProps(document.CustomFilePropertiesPart))
                        {
                            var result = f.ShowDialog();
                            if (f.PartModified)
                            {
                                document.WorkbookPart.Workbook.Save();
                                AddCustomDocPropsToList(document.CustomFilePropertiesPart);
                            }
                        }
                    }
                }
                else if (fileType == StringResources.wPowerpoint)
                {
                    using (PresentationDocument document = PresentationDocument.Open(TxtFileName.Text, true))
                    {
                        AddCustomDocPropsToList(document.CustomFilePropertiesPart);
                        LstDisplay.Items.Clear();

                        using (var f = new FrmDeleteCustomProps(document.CustomFilePropertiesPart))
                        {
                            var result = f.ShowDialog();
                            if (f.PartModified)
                            {
                                document.PresentationPart.Presentation.Save();
                                AddCustomDocPropsToList(document.CustomFilePropertiesPart);
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
                LogInformation(LogType.LogException, "BtnListCustomProps Error: ", ioe.Message);
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnListCustomProps Error: ", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnViewCustomXml_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                PreButtonClickWork();

                using (var f = new FrmCustomXmlViewer(TxtFileName.Text, fileType))
                {
                    var result = f.ShowDialog();
                }
            }
            catch (IOException ioe)
            {
                LogInformation(LogType.LogException, "No Customer Xml - BtnViewCustomXml Error: ", ioe.Message);
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "No Customer Xml - BtnViewCustomXml Error: ", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BatchFileProcessingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmBatch bFrm = new FrmBatch()
            {
                Owner = this
            };
            bFrm.ShowDialog();
        }

        private void BtnViewImages_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();
            FrmViewImages imgFrm = new FrmViewImages(TxtFileName.Text, fileType)
            {
                Owner = this
            };
            imgFrm.ShowDialog();
        }

        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            AppExitWork();
        }

        private void BtnListExcelHyperlinks_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();
            Cursor = Cursors.WaitCursor;
            try
            {
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, false))
                {
                    int count = 0;

                    foreach (WorksheetPart wsp in excelDoc.WorkbookPart.WorksheetParts)
                    {
                        IEnumerable<O.Spreadsheet.Hyperlink> hLinks = wsp.Worksheet.Descendants<O.Spreadsheet.Hyperlink>();
                        foreach (O.Spreadsheet.Hyperlink h in hLinks)
                        {
                            count++;

                            string hRelUri = null;

                            // then check for hyperlinks relationships
                            if (wsp.HyperlinkRelationships.Count() > 0)
                            {
                                foreach (HyperlinkRelationship hRel in wsp.HyperlinkRelationships)
                                {
                                    if (h.Id == hRel.Id)
                                    {
                                        hRelUri = hRel.Uri.ToString();
                                        LstDisplay.Items.Add(count + StringResources.wPeriod + h.InnerText + " Uri = " + hRelUri);
                                    }
                                }
                            }
                            else
                            {
                                LogInformation(LogType.EmptyCount, "hyperlinks", string.Empty);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnListExcelHyperlinks", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnDeleteUnusedStyles_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                PreButtonClickWork();
                int count = 0;

                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(TxtFileName.Text, true))
                {
                    MainDocumentPart mainPart = myDoc.MainDocumentPart;
                    StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart;
                    bool styleDeleted = false;

                    LstDisplay.Items.Clear();

                    try
                    {
                        List<string> baseStyleChains = new List<string>();
                        List<string> baseStyles = new List<string>();
                        string[] words = null;

                        // first, get all the base styles
                        foreach (OpenXmlElement tempEl in stylePart.Styles.Elements())
                        {
                            if (tempEl.LocalName == "style")
                            {
                                Style tempStyle = (Style)tempEl;
                                if (tempStyle.BasedOn == null)
                                {
                                    baseStyles.Add(tempStyle.StyleId);
                                }
                            }
                        }

                        // loop base styles and recursively get basedon chain
                        // this should create a string of the linked list sequence of styles
                        foreach (string sBase in baseStyles)
                        {
                            StringBuilder tempBaseStyleChain = new StringBuilder();
                            tempBaseStyleChain.Append(sBase);

                            StringBuilder baseStyleChain = WordOpenXml.GetBasedOnStyleChain(stylePart, sBase, tempBaseStyleChain);

                            if (baseStyleChain.ToString().Contains(StringResources.wArrowOnly))
                            {
                                baseStyleChains.Add(baseStyleChain.ToString());
                            }
                        }

                        // now we need to parse out the style chains for each individual styleid in reverse order
                        // if the style is applied to any p, r, or t, don't delete
                        // if the style is default, nextParagraphStyle or LinkedStyle, don't delete
                        // if neither of these is true, we can delete the style
                        if (baseStyleChains.Count > 0)
                        {
                            foreach (string b in baseStyleChains)
                            {
                                bool doNotDeleteAnyInChain = false;
                                string[] separatingStrings = { StringResources.wArrowOnly };
                                words = b.Split(separatingStrings, StringSplitOptions.None);

                                if (words.Count() > 0)
                                {
                                    foreach (string w in words.Reverse())
                                    {
                                        int pWStyleCount = WordExtensionClass.ParagraphsByStyleId(mainPart, w).Count();
                                        int rWStyleCount = WordExtensionClass.RunsByStyleId(mainPart, w).Count();
                                        int tWStyleCount = WordExtensionClass.TablesByStyleId(mainPart, w).Count();
                                        count += 1;

                                        // if the style is used in a para, run or table, don't delete
                                        if (pWStyleCount > 0 || rWStyleCount > 0 || tWStyleCount > 0)
                                        {
                                            LstDisplay.Items.Add(count + StringResources.wPeriod + StringResources.doNotDeleteStyle + w);
                                            doNotDeleteAnyInChain = true;
                                        }

                                        // if the style is not used, candidate for delete
                                        if (pWStyleCount == 0 && rWStyleCount == 0 && tWStyleCount == 0)
                                        {
                                            // if the previous style in the chain was true, we need to leave these alone
                                            if (doNotDeleteAnyInChain == true)
                                            {
                                                LstDisplay.Items.Add(count + StringResources.wPeriod + StringResources.doNotDeleteStyle + w);
                                            }
                                            else
                                            {
                                                // if we get here, the style is not applied and we can check the rest of the requirements for deletion
                                                foreach (OpenXmlElement tempEl in stylePart.Styles.Elements())
                                                {
                                                    if (tempEl.LocalName == "style")
                                                    {
                                                        Style tempStyle = (Style)tempEl;
                                                        if (tempStyle.StyleId == w)
                                                        {
                                                            // this is the last leg of the style still in use checks
                                                            // if default, nextpara and linked are all null, this style can be deleted
                                                            if (tempStyle.Default == null && tempStyle.NextParagraphStyle == null && tempStyle.LinkedStyle == null)
                                                            {
                                                                LstDisplay.Items.Add(count + StringResources.wPeriod + StringResources.deleteStyle + w);
                                                                tempEl.Remove();
                                                                styleDeleted = true;
                                                            }
                                                            else
                                                            {
                                                                LstDisplay.Items.Add(count + StringResources.wPeriod + StringResources.doNotDeleteStyle + w);
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
                    }
                    catch (NullReferenceException)
                    {
                        LogInformation(LogType.LogException, "BtnDeleteUnusedStyles - Missing StylesWithEffects part.", string.Empty);
                        return;
                    }

                    // if we deleted any style, save the file
                    if (styleDeleted == true)
                    {
                        myDoc.MainDocumentPart.Document.Save();
                    }
                    else
                    {
                        LstDisplay.Items.Add("** Document does not contain unused styles **");
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnDeleteUnusedStyles Error: ", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnDeleteEmbeddedLinks_Click(object sender, EventArgs e)
        {
            PreButtonClickWork();
            Cursor = Cursors.WaitCursor;
            try
            {
                bool linkDeleted = false;

                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(TxtFileName.Text, true))
                {
                    foreach (WorksheetPart wsp in excelDoc.WorkbookPart.WorksheetParts)
                    {
                        var oleObjects = wsp.Worksheet.Descendants<OleObject>().ToList();

                        foreach (OleObject oo in oleObjects)
                        {
                            oo.RemoveAttribute("link", string.Empty);
                            oo.RemoveAttribute("oleUpdate", string.Empty);
                            
                            if (oo.EmbeddedObjectProperties != null)
                            {
                                oo.EmbeddedObjectProperties.RemoveAttribute("dde", string.Empty);
                            }

                            // set the change flag
                            linkDeleted = true;
                        }
                    }

                    if (linkDeleted == true)
                    {
                        excelDoc.WorkbookPart.Workbook.Save();
                        LstDisplay.Items.Add("** Embedded Links Removed **");
                    }
                    else
                    {
                        LogInformation(LogType.EmptyCount, "Embedded Links", "");
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogType.LogException, "BtnListExcelHyperlinks", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnListMIPLabels_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            PreButtonClickWork();
            try
            {
                using (Package wdPackage = Package.Open(TxtFileName.Text, FileMode.Open, FileAccess.Read))
                {
                    PackageRelationship mipPackageRelationship = wdPackage.GetRelationshipsByType(StringResources.ClpRelationship).FirstOrDefault();
                    if (mipPackageRelationship != null)
                    {
                        Uri documentUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), mipPackageRelationship.TargetUri);
                        PackagePart mipPart = wdPackage.GetPart(documentUri);

                        XmlDocument doc = new XmlDocument();
                        doc.Load(XmlReader.Create(mipPart.GetStream()));
                        XmlElement root = doc.DocumentElement;

                        int count = 0;
                        foreach (XmlNode xn in root .ChildNodes)
                        {
                            count++;
                            LstDisplay.Items.Add("Label " + count + StringResources.wColon);
                            foreach (XmlAttribute xa in xn.Attributes)
                            {
                                LstDisplay.Items.Add(StringResources.wMinusSign + xa.Name + StringResources.wColonBuffer + xa.Value);
                            }
                        }

                        if (count == 0)
                        {
                            LstDisplay.Items.Add("** Document does not contain any labels **");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("BtnListMIPLabels Error: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }
    }
}