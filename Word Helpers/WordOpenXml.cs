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

namespace Office_File_Explorer.Word_Helpers
{
    class WordOpenXml
    {
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
                        catch (Exception)
                        {
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
    }
}
