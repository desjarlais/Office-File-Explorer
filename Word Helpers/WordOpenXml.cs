/****************************** Module Header ******************************\
Module Name:  WordOpenXml.cs
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

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.Collections;
using System.IO;
using Office_File_Explorer.App_Helpers;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.VariantTypes;

namespace Office_File_Explorer.Word_Helpers
{
    class WordOpenXml
    {
        // This method can be used to replace the theme part in a package.
        public static void ReplaceTheme(string document, string themeFile)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                // Delete the old document part.
                mainPart.DeletePart(mainPart.ThemePart);

                // Add a new document part and then add content.
                ThemePart themePart = mainPart.AddNewPart<ThemePart>();

                using (StreamReader streamReader = new StreamReader(themeFile))
                using (StreamWriter streamWriter = new StreamWriter(themePart.GetStream(FileMode.Create)))
                {
                    streamWriter.Write(streamReader.ReadToEnd());
                }
            }
        }

        public static void RemoveBreaks(string filename)
        {
            // this function will remove both page and section breaks in a document
            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filename, true))
            {
                MainDocumentPart mainPart = myDoc.MainDocumentPart;

                List<Break> breaks = mainPart.Document.Descendants<Break>().ToList();

                foreach (Break b in breaks)
                {
                    b.Remove();
                }

                List<ParagraphProperties> paraProps = mainPart.Document.Descendants<ParagraphProperties>()
                .Where(pPr => IsSectionProps(pPr)).ToList();

                foreach (ParagraphProperties pPr in paraProps)
                {
                    pPr.RemoveChild<SectionProperties>(pPr.GetFirstChild<SectionProperties>());
                }

                mainPart.Document.Save();
            }
        }

        public static void RemoveListTemplatesNumId(string filename, string numId)
        {
            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filename, true))
            {
                MainDocumentPart mainPart = myDoc.MainDocumentPart;
                NumberingDefinitionsPart numPart = mainPart.NumberingDefinitionsPart;
                ArrayList absNumId = new ArrayList();

                foreach (OpenXmlElement el in numPart.Numbering.Elements())
                {
                    foreach (AbstractNumId aNumId in el.Descendants<AbstractNumId>())
                    {
                        string strNumId = el.GetAttribute("numId", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value;
                        if (strNumId.Equals(numId))
                        {
                            absNumId.Add(aNumId.Val);
                            el.Remove();
                        }
                    }
                }

                foreach (object obj in absNumId)
                {
                    foreach (OpenXmlElement el in numPart.Numbering.Elements())
                    {
                        try
                        {
                            string x = el.GetAttribute("abstractNumId", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value.ToString();
                            string y = obj.ToString();

                            if (x == y)
                            {
                                el.Remove();
                            }
                        }
                        catch (Exception ex)
                        {
                            LoggingHelper.Log("RemoveListTemplatesNumId" + ex.Message);
                        }
                    }
                }

                mainPart.Document.Save();
            }
        }

        static bool IsSectionProps(ParagraphProperties pPr)
        {
            SectionProperties sectPr = pPr.GetFirstChild<SectionProperties>();

            if (sectPr == null)
                return false;
            else
                return true;
        }

        public static void RemoveComments(string filename)
        {
            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filename, true))
            {
                MainDocumentPart mainPart = myDoc.MainDocumentPart;

                //Delete the comment part, plus any other part referenced, like image parts 
                mainPart.DeletePart(mainPart.WordprocessingCommentsPart);

                //Find all elements that are associated with comments 
                IEnumerable<OpenXmlElement> elementList = mainPart.Document.Descendants()
                .Where(el => el is CommentRangeStart || el is CommentRangeEnd || el is CommentReference);

                //Delete every found element 
                foreach (OpenXmlElement e in elementList)
                {
                    e.Remove();
                }

                //Save changes 
                mainPart.Document.Save();
            }
        }

        // Given a .docm file (with macro storage), remove the VBA 
        // project, reset the document type, and save the document with a new name.
        public static string ConvertDOCMtoDOCX(string fileName)
        {
            bool fileChanged = false;
            string newFileName = "";

            using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, true))
            {
                // Access the main document part.
                var docPart = document.MainDocumentPart;

                // Look for the vbaProject part. If it is there, delete it.
                var vbaPart = docPart.VbaProjectPart;
                if (vbaPart != null)
                {
                    // Delete the vbaProject part and then save the document.
                    docPart.DeletePart(vbaPart);
                    docPart.Document.Save();

                    // Change the document type to not macro-enabled.
                    document.ChangeDocumentType(WordprocessingDocumentType.Document);

                    // Track that the document has been changed.
                    fileChanged = true;
                }
            }

            // If anything goes wrong in this file handling,
            // the code will raise an exception back to the caller.
            if (fileChanged)
            {
                // Create the new .docx filename.
                newFileName = Path.ChangeExtension(fileName, ".docx");

                // If it already exists, it will be deleted!
                if (File.Exists(newFileName))
                {
                    File.Delete(newFileName);
                }

                // Rename the file.
                File.Move(fileName, newFileName);
            }

            return newFileName;
        }

        // Given a document name and an author name, accept all revisions by the specified author. 
        // Pass an empty string for the author to accept all revisions.
        public static void AcceptAllRevisions(string docName, string authorName)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(docName, true))
            {
                Document doc = document.MainDocumentPart.Document;
                var paragraphChanged = doc.Descendants<ParagraphPropertiesChange>().ToList();
                var runChanged = doc.Descendants<RunPropertiesChange>().ToList();
                var deleted = doc.Descendants<DeletedRun>().ToList();
                var deletedParagraph = doc.Descendants<Deleted>().ToList();
                var inserted = doc.Descendants<InsertedRun>().ToList();

                if (!String.IsNullOrEmpty(authorName))
                {
                    paragraphChanged = paragraphChanged.Where(item => item.Author == authorName).ToList();
                    runChanged = runChanged.Where(item => item.Author == authorName).ToList();
                    deleted = deleted.Where(item => item.Author == authorName).ToList();
                    inserted = inserted.Where(item => item.Author == authorName).ToList();
                    deletedParagraph = deletedParagraph.Where(item => item.Author == authorName).ToList();
                }

                foreach (var item in paragraphChanged)
                    item.Remove();

                foreach (var item in deletedParagraph)
                    item.Remove();

                foreach (var item in runChanged)
                    item.Remove();

                foreach (var item in deleted)
                    item.Remove();

                foreach (var item in inserted)
                {
                    if (item.Parent != null)
                    {
                        var textRuns = item.Elements<Run>().ToList();
                        var parent = item.Parent;
                        foreach (var textRun in textRuns)
                        {
                            item.RemoveAttribute("rsidR", parent.NamespaceUri);
                            item.RemoveAttribute("sidRPr", parent.NamespaceUri);
                            parent.InsertBefore(textRun.CloneNode(true), item);
                        }
                        item.Remove();
                    }
                }
                doc.Save();
            }
        }

        // Delete headers and footers from a document.
        public static void RemoveHeadersFooters(string docName)
        {
            // Given a document name, remove all headers and footers.
            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(docName, true))
            {
                if (wdDoc.MainDocumentPart.GetPartsCountOfType<HeaderPart>() > 0 ||
                  wdDoc.MainDocumentPart.GetPartsCountOfType<FooterPart>() > 0)
                {
                    // Remove header and footer parts.
                    wdDoc.MainDocumentPart.DeleteParts(wdDoc.MainDocumentPart.HeaderParts);
                    wdDoc.MainDocumentPart.DeleteParts(wdDoc.MainDocumentPart.FooterParts);

                    Document doc = wdDoc.MainDocumentPart.Document;

                    // Remove references to the headers and footers.
                    var headers =
                      doc.Descendants<HeaderReference>().ToList();
                    foreach (var header in headers)
                    {
                        header.Parent.RemoveChild(header);
                    }

                    var footers =
                      doc.Descendants<FooterReference>().ToList();
                    foreach (var footer in footers)
                    {
                        footer.Parent.RemoveChild(footer);
                    }
                    doc.Save();
                }
            }
        }

        // Given a document, remove all hidden text.
        public static void DeleteHiddenText(string docName)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(docName, true))
            {
                Document doc = document.MainDocumentPart.Document;
                var hiddenItems = doc.Descendants<Vanish>().ToList();
                foreach (var item in hiddenItems)
                {
                    // Need to go up at least two levels to get to the run.
                    if ((item.Parent != null) &&
                      (item.Parent.Parent != null) &&
                      (item.Parent.Parent.Parent != null))
                    {
                        var topNode = item.Parent.Parent;
                        var topParentNode = item.Parent.Parent.Parent;
                        if (topParentNode != null)
                        {
                            topNode.Remove();
                            // No more children? Remove the parent node, as well.
                            if (!topParentNode.HasChildren)
                            {
                                topParentNode.Remove();
                            }
                        }
                    }
                }
                doc.Save();
            }
        }

        // Delete headers and footers from a document.
        public static void RemoveFootnotes(string docName)
        {
            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(docName, true))
            {
                FootnotesPart fnp = wdDoc.MainDocumentPart.FootnotesPart;
                if (fnp != null)
                {
                    var footnotes = fnp.Footnotes.Elements<Footnote>();
                    var references = wdDoc.MainDocumentPart.Document.Body.Descendants<FootnoteReference>().ToArray();

                    foreach (var reference in references)
                    {
                        reference.Remove();
                    }

                    foreach (var footnote in footnotes)
                    {
                        footnote.Remove();
                    }
                }

                wdDoc.MainDocumentPart.Document.Save();
                wdDoc.Close();
            }
        }

        // Delete headers and footers from a document.
        public static void RemoveEndnotes(string docName)
        {
            // Given a document name, remove all headers and footers.
            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(docName, true))
            {
                if (wdDoc.MainDocumentPart.GetPartsCountOfType<EndnotesPart>() > 0)
                {
                    MainDocumentPart mainPart = wdDoc.MainDocumentPart;
                    mainPart.DeletePart(mainPart.EndnotesPart);

                    var enr = mainPart.Document.Descendants<EndnoteReference>().ToList();
                    foreach (var e in enr)
                    {
                        e.Parent.RemoveChild(e);
                    }

                    var en = mainPart.Document.Descendants<Endnote>().ToList();
                    foreach (var e in en)
                    {
                        e.Parent.RemoveChild(e);
                    }

                    mainPart.Document.Save();
                    wdDoc.Close();
                }
            }
        }

        /// <summary>
        /// Set the font for a text run.
        /// </summary>
        /// <param name="fileName"></param>        
        public static void SetRunFont(string fileName)
        {
            // Open a Wordprocessing document for editing.
            using (WordprocessingDocument package = WordprocessingDocument.Open(fileName, true))
            {
                // Set the font to Arial to the first Run.
                // Use an object initializer for RunProperties and rPr.
                RunProperties rPr = new RunProperties(
                    new RunFonts()
                    {
                        Ascii = "Arial"
                    });

                Run r = package.MainDocumentPart.Document.Descendants<Run>().First();
                r.PrependChild<RunProperties>(rPr);

                // Save changes to the MainDocumentPart part.
                package.MainDocumentPart.Document.Save();
            }
        }

        public enum PropertyTypes : int
        {
            YesNo,
            Text,
            DateTime,
            NumberInteger,
            NumberDouble
        }

        public static string SetCustomProperty(string fileName, string propertyName, object propertyValue, PropertyTypes propertyType)
        {
            // Given a document name, a property name/value, and the property type, 
            // add a custom property to a document. The method returns the original
            // value, if it existed.

            string returnValue = null;

            var newProp = new CustomDocumentProperty();
            bool propSet = false;

            // Calculate the correct type.
            switch (propertyType)
            {
                case PropertyTypes.DateTime:

                    // Be sure you were passed a real date, 
                    // and if so, format in the correct way. 
                    // The date/time value passed in should 
                    // represent a UTC date/time.
                    if ((propertyValue) is DateTime)
                    {
                        newProp.VTFileTime = new VTFileTime(string.Format("{0:s}Z", Convert.ToDateTime(propertyValue)));
                        propSet = true;
                    }

                    break;

                case PropertyTypes.NumberInteger:
                    if ((propertyValue) is int)
                    {
                        newProp.VTInt32 = new VTInt32(propertyValue.ToString());
                        propSet = true;
                    }

                    break;

                case PropertyTypes.NumberDouble:
                    if (propertyValue is double)
                    {
                        newProp.VTFloat = new VTFloat(propertyValue.ToString());
                        propSet = true;
                    }

                    break;

                case PropertyTypes.Text:
                    newProp.VTLPWSTR = new VTLPWSTR(propertyValue.ToString());
                    propSet = true;

                    break;

                case PropertyTypes.YesNo:
                    if (propertyValue is bool)
                    {
                        // Must be lowercase.
                        newProp.VTBool = new VTBool(Convert.ToBoolean(propertyValue).ToString().ToLower());
                        propSet = true;
                    }
                    break;
            }

            if (!propSet)
            {
                // If the code was not able to convert the 
                // property to a valid value, throw an exception.
                throw new InvalidDataException("propertyValue");
            }

            // Now that you have handled the parameters, start
            // working on the document.
            newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
            newProp.Name = propertyName;

            using (var document = WordprocessingDocument.Open(fileName, true))
            {
                var customProps = document.CustomFilePropertiesPart;
                if (customProps == null)
                {
                    // No custom properties? Add the part, and the
                    // collection of properties now.
                    customProps = document.AddCustomFilePropertiesPart();
                    customProps.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties();
                }

                var props = customProps.Properties;
                if (props != null)
                {
                    // This will trigger an exception if the property's Name 
                    // property is null, but if that happens, the property is damaged, 
                    // and probably should raise an exception.
                    var prop = props.Where(p => ((CustomDocumentProperty)p).Name.Value == propertyName).FirstOrDefault();

                    // Does the property exist? If so, get the return value, 
                    // and then delete the property.
                    if (prop != null)
                    {
                        returnValue = prop.InnerText;
                        prop.Remove();
                    }

                    // Append the new property, and 
                    // fix up all the property ID values. 
                    // The PropertyId value must start at 2.
                    props.AppendChild(newProp);
                    int pid = 2;
                    foreach (CustomDocumentProperty item in props)
                    {
                        item.PropertyId = pid++;
                    }
                    props.Save();
                }
            }
            return returnValue;
        }

        /// <summary>
        /// Given a document name, set the print orientation for all the sections of the document.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="newOrientation"></param>
        public static void SetPrintOrientation(string fileName, PageOrientationValues newOrientation)
        {
            using (var document = WordprocessingDocument.Open(fileName, true))
            {
                bool documentChanged = false;

                var docPart = document.MainDocumentPart;
                var sections = docPart.Document.Descendants<SectionProperties>();

                foreach (SectionProperties sectPr in sections)
                {
                    bool pageOrientationChanged = false;

                    PageSize pgSz = sectPr.Descendants<PageSize>().FirstOrDefault();
                    if (pgSz != null)
                    {
                        // No Orient property? Create it now. Otherwise, just 
                        // set its value. Assume that the default orientation 
                        // is Portrait.
                        if (pgSz.Orient == null)
                        {
                            // Need to create the attribute. You do not need to 
                            // create the Orient property if the property does not 
                            // already exist, and you are setting it to Portrait. 
                            // That is the default value.
                            if (newOrientation != PageOrientationValues.Portrait)
                            {
                                pageOrientationChanged = true;
                                documentChanged = true;
                                pgSz.Orient = new EnumValue<PageOrientationValues>(newOrientation);
                            }
                        }
                        else
                        {
                            // The Orient property exists, but its value
                            // is different than the new value.
                            if (pgSz.Orient.Value != newOrientation)
                            {
                                pgSz.Orient.Value = newOrientation;
                                pageOrientationChanged = true;
                                documentChanged = true;
                            }
                        }

                        if (pageOrientationChanged)
                        {
                            // Changing the orientation is not enough. You must also 
                            // change the page size.
                            var width = pgSz.Width;
                            var height = pgSz.Height;
                            pgSz.Width = height;
                            pgSz.Height = width;

                            PageMargin pgMar = sectPr.Descendants<PageMargin>().FirstOrDefault();
                            if (pgMar != null)
                            {
                                // Rotate margins. Printer settings control how far you 
                                // rotate when switching to landscape mode. Not having those
                                // settings, this code rotates 90 degrees. You could easily
                                // modify this behavior, or make it a parameter for the 
                                // procedure.
                                var top = pgMar.Top.Value;
                                var bottom = pgMar.Bottom.Value;
                                var left = pgMar.Left.Value;
                                var right = pgMar.Right.Value;

                                pgMar.Top = new Int32Value((int)left);
                                pgMar.Bottom = new Int32Value((int)right);
                                pgMar.Left = new UInt32Value((uint)System.Math.Max(0, bottom));
                                pgMar.Right = new UInt32Value((uint)System.Math.Max(0, top));
                            }
                        }
                    }
                }
                if (documentChanged)
                {
                    docPart.Document.Save();
                }
            }
        }
    }
}
