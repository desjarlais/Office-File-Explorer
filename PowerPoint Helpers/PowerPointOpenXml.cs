using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using PShape = DocumentFormat.OpenXml.Presentation.Shape;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using Drawing = DocumentFormat.OpenXml.Drawing;

using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using System.Windows.Forms;
using ShapeStyle = DocumentFormat.OpenXml.Presentation.ShapeStyle;
using ShapeProperties = DocumentFormat.OpenXml.Presentation.ShapeProperties;
using NonVisualShapeProperties = DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties;
using NonVisualDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties;
using NonVisualGroupShapeProperties = DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties;
using ColorMap = DocumentFormat.OpenXml.Presentation.ColorMap;
using TextBody = DocumentFormat.OpenXml.Presentation.TextBody;
using NonVisualShapeDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties;
using NonVisualGroupShapeDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties;

namespace Office_File_Explorer.PowerPoint_Helpers
{
    public static class PowerPointOpenXml
    {
        // Returns all the external hyperlinks in the slides of a presentation.
        public static IEnumerable<String> GetAllExternalHyperlinksInPresentation(string fileName)
        {
            // Declare a list of strings.
            List<string> ret = new List<string>();

            // Open the presentation file as read-only.
            using (PresentationDocument document = PresentationDocument.Open(fileName, false))
            {
                // Iterate through all the slide parts in the presentation part.
                foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
                {
                    IEnumerable<Drawing.HyperlinkType> links = slidePart.Slide.Descendants<Drawing.HyperlinkType>();

                    // Iterate through all the links in the slide part.
                    foreach (Drawing.HyperlinkType link in links)
                    {
                        // Iterate through all the external relationships in the slide part. 
                        foreach (HyperlinkRelationship relation in slidePart.HyperlinkRelationships)
                        {
                            // If the relationship ID matches the link ID
                            if (relation.Id.Equals(link.Id))
                            {
                                // Add the URI of the external relationship to the list of strings.
                                ret.Add(relation.Uri.AbsoluteUri);
                            }
                        }
                    }
                }
            }

            // Return the list of strings.
            return ret;
        }

        /// <summary>
        /// Check the notes page size and reset
        /// generate a new notes master if setting is enabled
        /// </summary>
        /// <param name="pDoc"></param>
        public static void ChangeNotesPageSize(PresentationDocument pDoc)
        {
            if (pDoc == null)
            {
                throw new ArgumentNullException("pDoc = null");
            }

            // Get the presentation part of document
            PresentationPart presentationPart = pDoc.PresentationPart;

            if (presentationPart != null)
            {
                Presentation p = presentationPart.Presentation;

                // Step 1 : Resize the presentation notesz prop
                // if the notes size is already the default, no need to make any changes
                if (p.NotesSize.Cx != 6858000 || p.NotesSize.Cy != 9144000)
                {
                    // setup default size
                    NotesSize defaultNotesSize = new NotesSize() { Cx = 6858000L, Cy = 9144000L };

                    // first reset the notes size values        
                    p.NotesSize = defaultNotesSize;

                    // now save up the part
                    p.Save();
                }

                // Step 2 : loop the shapes in the notes master and reset their sizes
                // need to find a way to flag a file if the notes master and/or notes slides become corrupt
                // hiding behind a setting checkbox for now
                if (Properties.Settings.Default.ResetNotesMaster == "true")
                {
                    // we need to reset sizes in the notes master for each shape
                    ShapeTree mSt = presentationPart.NotesMasterPart.NotesMaster.CommonSlideData.ShapeTree;
                    foreach (var mShp in mSt)
                    {
                        if (mShp.ToString() == "DocumentFormat.OpenXml.Presentation.Shape")
                        {
                            PShape ps = (PShape)mShp;
                            NonVisualShapeProperties nvsp = ps.NonVisualShapeProperties;
                            NonVisualDrawingProperties nvdpr = nvsp.NonVisualDrawingProperties;
                            ShapeProperties sp = ps.ShapeProperties;
                            Transform2D t2d = ps.ShapeProperties.Transform2D;

                            if (nvdpr.Name == "Header Placeholder 1")
                            {
                                t2d.Offset.X = 0L;
                                t2d.Offset.Y = 0L;
                                t2d.Extents.Cx = 2971800L;
                                t2d.Extents.Cy = 458788L;
                            }

                            if (nvdpr.Name == "Date Placeholder 2")
                            {
                                t2d.Offset.X = 3884613L;
                                t2d.Offset.Y = 0L;
                                t2d.Extents.Cx = 2971800L;
                                t2d.Extents.Cy = 458788L;
                            }

                            if (nvdpr.Name == "Slide Image Placeholder 3")
                            {
                                t2d.Offset.X = 685800L;
                                t2d.Offset.Y = 1143000L;
                                t2d.Extents.Cx = 5486400L;
                                t2d.Extents.Cy = 3086100L;
                            }

                            if (nvdpr.Name == "Notes Placeholder 4")
                            {
                                t2d.Offset.X = 685800L;
                                t2d.Offset.Y = 4400550L;
                                t2d.Extents.Cx = 5486400L;
                                t2d.Extents.Cy = 3600450L;
                            }

                            if (nvdpr.Name == "Footer Placeholder 5")
                            {
                                t2d.Offset.X = 0L;
                                t2d.Offset.Y = 8685213L;
                                t2d.Extents.Cx = 2971800L;
                                t2d.Extents.Cy = 458787L;
                            }

                            if (nvdpr.Name == "Slide Number Placeholder 6")
                            {
                                t2d.Offset.X = 3884613L;
                                t2d.Offset.Y = 8685213L;
                                t2d.Extents.Cx = 2971800L;
                                t2d.Extents.Cy = 458787L;
                            }
                        }
                    }

                    // Step 3 : we need to delete the transform2d data so that each slide will pull the size from the master
                    foreach (var slideId in p.SlideIdList.Elements<SlideId>())
                    {
                        SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                        NotesSlidePart nsp = slidePart.NotesSlidePart;
                        NotesSlide ns = nsp.NotesSlide;
                        CommonSlideData csd = ns.CommonSlideData;
                        ShapeTree st = csd.ShapeTree;
                        
                        foreach (var s in st)
                        {
                            // we only want to make changes to the shapes
                            if (s.ToString() == "DocumentFormat.OpenXml.Presentation.Shape")
                            {
                                PShape ps = (PShape)s;
                                NonVisualShapeProperties nvsp = ps.NonVisualShapeProperties;
                                NonVisualDrawingProperties nvdpr = nvsp.NonVisualDrawingProperties;
                                ShapeProperties sp = ps.ShapeProperties;
                                Transform2D t2d = ps.ShapeProperties.Transform2D;

                                // if the transform exists, delete it for each slide
                                if (t2d != null)
                                {
                                    t2d.Remove();
                                }

                                // if there are drawing paragraph props, reset the margin and indent to 0
                                if (ps.TextBody != null)
                                {
                                    TextBody tb = ps.TextBody;

                                    foreach (var x in tb.ChildElements)
                                    {
                                        if (x.ToString() == "DocumentFormat.OpenXml.Drawing.Paragraph")
                                        {
                                            DocumentFormat.OpenXml.Drawing.Paragraph para = (DocumentFormat.OpenXml.Drawing.Paragraph)x;
                                            if (para.ParagraphProperties != null)
                                            {
                                                if (para.ParagraphProperties.LeftMargin != null)
                                                {
                                                    para.ParagraphProperties.LeftMargin = 0;
                                                }

                                                if (para.ParagraphProperties.Indent != null)
                                                {
                                                    para.ParagraphProperties.Indent = 0;
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

        // Get a list of the titles of all the slides in the presentation.
        public static IList<string> GetSlideTitles(PresentationDocument presentationDocument)
        {
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            // Get a PresentationPart object from the PresentationDocument object.
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            if (presentationPart != null && presentationPart.Presentation != null)
            {
                // Get a Presentation object from the PresentationPart object.
                Presentation presentation = presentationPart.Presentation;

                if (presentation.SlideIdList != null)
                {
                    List<string> titlesList = new List<string>();

                    // Get the title of each slide in the slide order.
                    foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
                    {
                        SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;

                        // Get the slide title.
                        string title = GetSlideTitle(slidePart);

                        // An empty title can also be added.
                        titlesList.Add(title);
                    }

                    return titlesList;
                }

            }

            return null;
        }

        // Get the title string of the slide.
        public static string GetSlideTitle(SlidePart slidePart)
        {
            if (slidePart == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            // Declare a paragraph separator.
            string paragraphSeparator = null;

            if (slidePart.Slide != null)
            {
                // Find all the title shapes.
                var shapes = from shape in slidePart.Slide.Descendants<PShape>()
                             where IsTitleShape(shape)
                             select shape;

                StringBuilder paragraphText = new StringBuilder();

                foreach (var shape in shapes)
                {
                    // Get the text in each paragraph in this shape.
                    foreach (var paragraph in shape.TextBody.Descendants<Drawing.Paragraph>())
                    {
                        // Add a line break.
                        paragraphText.Append(paragraphSeparator);

                        foreach (var text in paragraph.Descendants<Drawing.Text>())
                        {
                            paragraphText.Append(text.Text);
                        }

                        paragraphSeparator = "\n";
                    }
                }

                return paragraphText.ToString();
            }

            return string.Empty;
        }

        /// <summary>
        /// Determines whether the shape is a title shape.
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        private static bool IsTitleShape(PShape shape)
        {
            var placeholderShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();
            if (placeholderShape != null && placeholderShape.Type != null && placeholderShape.Type.HasValue)
            {
                switch ((PlaceholderValues)placeholderShape.Type)
                {
                    // Any title shape.
                    case PlaceholderValues.Title:

                    // A centered title.
                    case PlaceholderValues.CenteredTitle:
                        return true;

                    default:
                        return false;
                }
            }
            return false;
        }

        public static int CountSlides(string presentationFile)
        {
            // Open the presentation as read-only.
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
            {
                // Pass the presentation to the next CountSlides method
                // and return the slide count.
                return CountSlides(presentationDocument);
            }
        }

        /// <summary>
        /// Get the slideId and text for that slide
        /// </summary>
        /// <param name="sldText">string returned to caller</param>
        /// <param name="docName">path to powerpoint file</param>
        /// <param name="index">slide number</param>
        public static void GetSlideIdAndText(out string sldText, string docName, int index)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(docName, false))
            {
                // Get the relationship ID of the first slide.
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

                string relId = (slideIds[index] as SlideId).RelationshipId;

                // Get the slide part from the relationship ID.
                SlidePart slide = (SlidePart)part.GetPartById(relId);
                
                // Build a StringBuilder object.
                StringBuilder paragraphText = new StringBuilder();

                // Get the inner text of the slide:
                IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
                foreach (A.Text text in texts)
                {
                    paragraphText.Append(text.Text);
                }
                sldText = paragraphText.ToString();
            }
        }

        /// <summary>
        /// Count the slides in the presentation.
        /// </summary>
        /// <param name="presentationDocument"></param>
        /// <returns></returns>
        public static int CountSlides(PresentationDocument presentationDocument)
        {
            // Check for a null document object.
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            int slidesCount = 0;

            // Get the presentation part of document.
            PresentationPart presentationPart = presentationDocument.PresentationPart;
            // Get the slide count from the SlideParts.
            if (presentationPart != null)
            {
                slidesCount = presentationPart.SlideParts.Count();
            }
            // Return the slide count to the previous method.
            return slidesCount;
        }

        // Move a slide to a different position in the slide order in the presentation.
        public static void MoveSlide(string presentationFile, int from, int to)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
            {
                MoveSlide(presentationDocument, from, to);
            }
        }

        /// <summary>
        /// Move a slide to a different position in the slide order in the presentation.
        /// </summary>
        /// <param name="presentationDocument"></param>
        /// <param name="from"></param>
        /// <param name="to"></param>
        public static void MoveSlide(PresentationDocument presentationDocument, int from, int to)
        {
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            // Call the CountSlides method to get the number of slides in the presentation.
            int slidesCount = CountSlides(presentationDocument);

            // Verify that both from and to positions are within range and different from one another.
            if (from < 0 || from >= slidesCount)
            {
                throw new ArgumentOutOfRangeException("from");
            }

            if (to < 0 || from >= slidesCount || to == from)
            {
                throw new ArgumentOutOfRangeException("to");
            }

            // Get the presentation part from the presentation document.
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // The slide count is not zero, so the presentation must contain slides.            
            Presentation presentation = presentationPart.Presentation;
            SlideIdList slideIdList = presentation.SlideIdList;

            // Get the slide ID of the source slide.
            SlideId sourceSlide = slideIdList.ChildElements[from] as SlideId;

            SlideId targetSlide = null;

            // Identify the position of the target slide after which to move the source slide.
            if (to == 0)
            {
                targetSlide = null;
            }
            if (from < to)
            {
                targetSlide = slideIdList.ChildElements[to] as SlideId;
            }
            else
            {
                targetSlide = slideIdList.ChildElements[to - 1] as SlideId;
            }

            // Remove the source slide from its current position.
            sourceSlide.Remove();

            // Insert the source slide at its new position after the target slide.
            slideIdList.InsertAfter(sourceSlide, targetSlide);

            // Save the modified presentation.
            presentation.Save();
        }

        /// <summary>
        /// Change the fill color of a shape, docName must have a filled shape as the first shape on the first slide.
        /// </summary>
        /// <param name="docName">path to the file</param>
        public static void SetPPTShapeColor(string docName)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(docName, true))
            {
                // Get the relationship ID of the first slide.
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
                string relId = (slideIds[0] as SlideId).RelationshipId;

                // Get the slide part from the relationship ID.
                SlidePart slide = (SlidePart)part.GetPartById(relId);

                if (slide != null)
                {
                    // Get the shape tree that contains the shape to change.
                    ShapeTree tree = slide.Slide.CommonSlideData.ShapeTree;

                    // Get the first shape in the shape tree.
                    PShape shape = tree.GetFirstChild<PShape>();

                    if (shape != null)
                    {
                        // Get the style of the shape.
                        ShapeStyle style = shape.ShapeStyle;

                        // Get the fill reference.
                        Drawing.FillReference fillRef = style.FillReference;

                        // Set the fill color to SchemeColor Accent 6;
                        fillRef.SchemeColor = new Drawing.SchemeColor();
                        fillRef.SchemeColor.Val = Drawing.SchemeColorValues.Accent6;

                        // Save the modified slide.
                        slide.Slide.Save();
                    }
                }
            }
        }

        /// <summary>
        /// Function to retrieve the number of slides
        /// </summary>
        /// <param name="fileName">path to the file</param>
        /// <param name="includeHidden">default is true, pass false if you don't want hidden slides counted</param>
        /// <returns></returns>
        public static int RetrieveNumberOfSlides(string fileName, bool includeHidden = true)
        {
            int slidesCount = 0;

            using (PresentationDocument doc =
                PresentationDocument.Open(fileName, false))
            {
                // Get the presentation part of the document.
                PresentationPart presentationPart = doc.PresentationPart;
                if (presentationPart != null)
                {
                    if (includeHidden)
                    {
                        slidesCount = presentationPart.SlideParts.Count();
                    }
                    else
                    {
                        // Each slide can include a Show property, which if hidden 
                        // will contain the value "0". The Show property may not 
                        // exist, and most likely will not, for non-hidden slides.
                        var slides = presentationPart.SlideParts.Where(
                            (s) => (s.Slide != null) &&
                              ((s.Slide.Show == null) || (s.Slide.Show.HasValue &&
                              s.Slide.Show.Value)));
                        slidesCount = slides.Count();
                    }
                }
            }
            return slidesCount;
        }
    }
}
