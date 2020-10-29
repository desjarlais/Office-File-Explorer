/****************************** Module Header ******************************\
Module Name:  WordOpenXml.cs
Project:      Office File Explorer

Word Open Xml Helper Functions

This source is subject to the following license.
See https://github.com/desjarlais/Office-File-Explorer/blob/master/LICENSE
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

using Office_File_Explorer.App_Helpers;

namespace Office_File_Explorer.Word_Helpers
{
    class WordOpenXml
    {
        public static bool fWorked;

        public static bool RemoveBreaks(string filename)
        {
            fWorked = false;

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
                fWorked = true;
            }

            return fWorked;
        }

        /// <summary>
        /// given a numId, find the element in the Numbering.xml file with that same numId
        /// get the abstractnumid and delete the numId element
        /// then you can delete the abstractnumid
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="numId">orphaned ListTemplate to delete</param>
        public static void RemoveListTemplatesNumId(string filename, string numId)
        {
            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filename, true))
            {
                string absNumIdToRemove = "";

                var absNumsInUseList = myDoc.MainDocumentPart.NumberingDefinitionsPart.Numbering.Descendants<AbstractNum>().ToList();
                var numInstancesInUseList = myDoc.MainDocumentPart.NumberingDefinitionsPart.Numbering.Descendants<NumberingInstance>().ToList();

                // loop each numberinginstance and if the orphaned id matches, delete it and keep track of the abstractnumid
                foreach (NumberingInstance ni in numInstancesInUseList)
                {
                    if (ni.NumberID == numId)
                    {
                        absNumIdToRemove = ni.AbstractNumId.Val.ToString();
                        ni.Remove();
                        break;
                    }
                }

                // now that we have the abstractnum, loop that list and delete that one
                foreach (AbstractNum an in absNumsInUseList)
                {
                    if (an.AbstractNumberId == Int32.Parse(absNumIdToRemove))
                    {
                        an.Remove();
                        break;
                    }
                }

                myDoc.MainDocumentPart.Document.Save();
            }
        }

        static bool IsSectionProps(ParagraphProperties pPr)
        {
            SectionProperties sectPr = pPr.GetFirstChild<SectionProperties>();

            if (sectPr == null)
            {
                return false;
            }
            else
            {
                return true;
            }
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

        /// <summary>
        /// get a list of all authors in each story of the document
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static List<string> GetAllAuthors(Document doc)
        {
            bool nullAuthor = false;
            List<string> allAuthorsInDocument = new List<string>();

            var paragraphChanged = doc.Descendants<ParagraphPropertiesChange>().ToList();
            var runChanged = doc.Descendants<RunPropertiesChange>().ToList();
            var deleted = doc.Descendants<DeletedRun>().ToList();
            var deletedParagraph = doc.Descendants<Deleted>().ToList();
            var inserted = doc.Descendants<InsertedRun>().ToList();

            // loop through each revision and catalog the authors
            // some authors show up as null, check and ignore
            foreach (ParagraphPropertiesChange ppc in paragraphChanged)
            {
                if (ppc.Author != null)
                {
                    allAuthorsInDocument.Add(ppc.Author);
                }
                else
                {
                    nullAuthor = true;
                }
            }

            foreach (RunPropertiesChange rpc in runChanged)
            {
                if (rpc.Author != null)
                {
                    allAuthorsInDocument.Add(rpc.Author);
                }
                else
                {
                    nullAuthor = true;
                }
            }

            foreach (DeletedRun dr in deleted)
            {
                if (dr.Author != null)
                {
                    allAuthorsInDocument.Add(dr.Author);
                }
                else
                {
                    nullAuthor = true;
                }
            }

            foreach (Deleted d in deletedParagraph)
            {
                if (d.Author != null)
                {
                    allAuthorsInDocument.Add(d.Author);
                }
                else
                {
                    nullAuthor = true;
                }
            }

            foreach (InsertedRun ir in inserted)
            {
                if (ir.Author != null)
                {
                    allAuthorsInDocument.Add(ir.Author);
                }
                else
                {
                    nullAuthor = true;
                }
            }

            // log if we have a null author, not sure how this happens yet
            if (nullAuthor)
            {
                LoggingHelper.Log("Null Author Found");
            }
            
            List<string> distinctAuthors = allAuthorsInDocument.Distinct().ToList();

            return distinctAuthors;
        }

        /// <summary>
        /// Given a document name and an author name, accept all revisions by the specified author. 
        /// Pass an empty string for the author to accept all revisions.
        /// </summary>
        /// <param name="docName"></param>
        /// <param name="authorName"></param>
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

                if (authorName == "* All Authors *")
                {
                    List<string> temp = new List<string>();
                    temp = GetAllAuthors(document.MainDocumentPart.Document);

                    // create a temp list for each author so we can loop the changes individually and list them
                    foreach (string s in temp)
                    {
                        var tempParagraphChanged = paragraphChanged.Where(item => item.Author == s).ToList();
                        var tempRunChanged = runChanged.Where(item => item.Author == s).ToList();
                        var tempDeleted = deleted.Where(item => item.Author == s).ToList();
                        var tempInserted = inserted.Where(item => item.Author == s).ToList();
                        var tempDeletedParagraph = deletedParagraph.Where(item => item.Author == s).ToList();

                        foreach (var item in tempParagraphChanged)
                            item.Remove();

                        foreach (var item in tempDeletedParagraph)
                            item.Remove();

                        foreach (var item in tempRunChanged)
                            item.Remove();

                        foreach (var item in tempDeleted)
                            item.Remove();

                        foreach (var item in tempInserted)
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
                    }
                    doc.Save();
                }
                else
                {
                    // for single author, just loop that authors from the original list
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
        }

        /// <summary>
        /// Delete headers and footers from a document.
        /// </summary>
        /// <param name="docName"></param>
        /// <returns></returns>
        public static bool RemoveHeadersFooters(string docName)
        {
            fWorked = false;

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

                    var footers = doc.Descendants<FooterReference>().ToList();
                    foreach (var footer in footers)
                    {
                        footer.Parent.RemoveChild(footer);
                    }
                    doc.Save();
                    fWorked = true;
                }
            }

            return fWorked;
        }

        // Given a document, remove all hidden text.
        public static bool DeleteHiddenText(string docName)
        {
            fWorked = false;

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
                fWorked = true;
            }

            return fWorked;
        }

        /// <summary>
        /// delete the footnotes in a file
        /// </summary>
        /// <param name="docName"></param>
        /// <returns></returns>
        public static bool RemoveFootnotes(string docName)
        {
            fWorked = false;

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
                fWorked = true;
            }

            return fWorked;
        }

        /// <summary>
        /// Delete endnotes from the document
        /// </summary>
        /// <param name="docName"></param>
        /// <returns></returns>
        public static bool RemoveEndnotes(string docName)
        {
            fWorked = false;

            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(docName, true))
            {
                if (wdDoc.MainDocumentPart.GetPartsCountOfType<EndnotesPart>() > 0)
                {
                    MainDocumentPart mainPart = wdDoc.MainDocumentPart;
                    
                    var enr = mainPart.Document.Descendants<EndnoteReference>().ToArray();
                    foreach (var e in enr)
                    {
                        e.Remove();
                    }

                    EndnotesPart ep = mainPart.EndnotesPart;
                    ep.Endnotes = WordOpenXml.CreateDefaultEndnotes();
                    mainPart.Document.Save();
                    fWorked = true;
                }
            }

            return fWorked;
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

        /// <summary>
        /// Sometimes bookmarks are added and the start/end tag is missing
        /// This function will try to find those orphan tags and remove them
        /// </summary>
        /// <param name="filename">file to be scanned</param>
        /// <returns>true for successful removal and false if none are found</returns>
        public static bool RemoveMissingBookmarkTags(string filename)
        {
            bool isFixed = false;
            
            try
            {
                using (WordprocessingDocument package = WordprocessingDocument.Open(filename, true))
                {
                    if (package.MainDocumentPart.WordprocessingCommentsPart == null)
                    {
                        return false;
                    }

                    if (package.MainDocumentPart.WordprocessingCommentsPart.Comments == null)
                    {
                        return false;
                    }

                    IEnumerable<BookmarkStart> bkStartList = package.MainDocumentPart.WordprocessingCommentsPart.Comments.Descendants<BookmarkStart>();
                    IEnumerable<BookmarkEnd> bkEndList = package.MainDocumentPart.WordprocessingCommentsPart.Comments.Descendants<BookmarkEnd>();

                    // create temp lists so we can loop and remove any that exist in both lists
                    // if we have a start and end, the bookmark is valid and we can remove the rest
                    List<string> bkStartTagIds = new List<string>();
                    List<string> bkEndTagIds = new List<string>();

                    // check each start and find if there is a matching end tag id
                    foreach (BookmarkStart bks in bkStartList)
                    {
                        foreach (BookmarkEnd bke in bkEndList)
                        {
                            if (bke.Id.ToString() == bks.Id.ToString())
                            {
                                bkStartTagIds.Add(bke.Id);
                            }
                        }
                    }

                    // now we can check if there is a end tag with a matching start tag id
                    foreach (BookmarkEnd bke in bkEndList)
                    {
                        foreach (BookmarkStart bks in bkStartList)
                        {
                            if (bks.Id.ToString() == bke.Id.ToString())
                            {
                                bkEndTagIds.Add(bks.Id);
                            }
                        }
                    }

                    // now that we know all the id's that match, we can loop again and remove id's that are not in the lists
                    bool startTagFound = false;

                    foreach (BookmarkStart bks in bkStartList)
                    {
                        foreach (object o in bkEndTagIds)
                        {
                            // if the end tag matches and we can ignore doing anything
                            if (o.ToString() == bks.Id.ToString())
                            {
                                startTagFound = true;
                            }
                        }

                        // if we get here and no match was found, it is orphaned and we can delete
                        if (startTagFound == false)
                        {
                            bks.Remove();
                            isFixed = true;
                        }
                        else
                        {
                            // reset the value for the next start tag check
                            startTagFound = false;
                        }
                    }

                    bool endTagFound = false;

                    foreach (BookmarkEnd bke in bkEndList)
                    {
                        foreach (object o in bkStartTagIds)
                        {
                            if (o.ToString() == bke.Id.ToString())
                            {
                                endTagFound = true;
                            }
                        }

                        if (endTagFound == false)
                        {
                            bke.Remove();
                            isFixed = true;
                        }
                        else
                        {
                            endTagFound = false;
                        }
                    }

                    if (isFixed)
                    {
                        package.MainDocumentPart.Document.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("RemoveMissingBookmarkTags: " + ex.Message);
                return false;
            }

            return isFixed;
        }

        public static bool RemovePlainTextCcFromBookmark(string filename)
        {
            bool isFixed = false;

            try
            {
                using (WordprocessingDocument package = WordprocessingDocument.Open(filename, true))
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
                                // first check if we are a content control
                                if (cElem.Parent != null && cElem.Parent.ToString().Contains("DocumentFormat.OpenXml.Wordprocessing.Sdt"))
                                {
                                    foreach (OpenXmlElement oxe in cElem.Parent.ChildElements)
                                    {
                                        // get the properties
                                        if (oxe.GetType().Name == "SdtProperties")
                                        {
                                            foreach (OpenXmlElement oxeSdtAlias in oxe)
                                            {
                                                // check for plain text
                                                if (oxeSdtAlias.GetType().Name == "SdtContentText")
                                                {
                                                    // if the parent is a plain text content control, bookmark is not allowed
                                                    // add the id to the list of bookmarks that need to be deleted
                                                    removedBookmarkIds.Add(bk.Id);
                                                    endLoop = true;
                                                }
                                            }
                                        }
                                    }

                                    // set the next element to the parent and continue moving up the element chain
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
                                    else
                                    {
                                        // set pElem to the parent so we can check for the end of the loop
                                        // set cElem to the parent also so we can continue moving up the element chain
                                        pElem = cElem.Parent;
                                        cElem = pElem;

                                        // loop should continue until we get to the body element, then we can stop looping
                                        if (pElem.ToString() == "DocumentFormat.OpenXml.Wordprocessing.Body")
                                        {
                                            endLoop = true;
                                        }
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

                        // check if there were any fixes made and update the output display
                        if (removedBookmarkIds.Count > 0)
                        {
                            isFixed = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("RemovePlainTextCcFromBookmark: " + ex.Message);
                return false;
            }

            return isFixed;
        }

        // Creates an Endnotes instance and adds its children.
        public static Endnotes CreateDefaultEndnotes()
        {
            Endnotes endnotes1 = new Endnotes() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex wp14" } };
            endnotes1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            endnotes1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            endnotes1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            endnotes1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            endnotes1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            endnotes1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            endnotes1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            endnotes1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            endnotes1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            endnotes1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            endnotes1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            endnotes1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            endnotes1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            endnotes1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            endnotes1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            endnotes1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            endnotes1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            endnotes1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            endnotes1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            endnotes1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            endnotes1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            endnotes1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            endnotes1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            endnotes1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            endnotes1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            endnotes1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            endnotes1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            endnotes1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            endnotes1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            endnotes1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            endnotes1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Endnote endnote1 = new Endnote() { Type = FootnoteEndnoteValues.Separator, Id = -1 };

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00822C7D", RsidParagraphProperties = "00A30061", RsidRunAdditionDefault = "00822C7D", ParagraphId = "34CB4C4C", TextId = "77777777" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            paragraphProperties1.Append(spacingBetweenLines1);

            Run run1 = new Run();
            SeparatorMark separatorMark1 = new SeparatorMark();

            run1.Append(separatorMark1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            endnote1.Append(paragraph1);

            Endnote endnote2 = new Endnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00822C7D", RsidParagraphProperties = "00A30061", RsidRunAdditionDefault = "00822C7D", ParagraphId = "4BBA76AE", TextId = "77777777" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            paragraphProperties2.Append(spacingBetweenLines2);

            Run run2 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark1 = new ContinuationSeparatorMark();

            run2.Append(continuationSeparatorMark1);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);

            endnote2.Append(paragraph2);

            endnotes1.Append(endnote1);
            endnotes1.Append(endnote2);
            return endnotes1;
        }

        /// <summary>
        /// Function to check if a part is null
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="partType"></param>
        /// <returns></returns>
        public static bool IsPartNull(WordprocessingDocument doc, string partType)
        {
            try
            {
                if (partType == "DeletedRun")
                {
                    var deleted = doc.MainDocumentPart.Document.Descendants<DeletedRun>().ToList();
                }
                else if (partType == "Table")
                {
                    var tbls = doc.MainDocumentPart.Document.Descendants<Table>().ToList();
                }

                return false;
            }
            catch (Exception)
            {
                return true;
            }
        }
    }
}
