using Drawing = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using System.Text;
using System.Linq;
using System.IO;

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
                            // If the relationship ID matches the link ID…
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

        // Get a list of the titles of all the slides in the presentation.
        public static IList<string> GetSlideTitles(PresentationDocument presentationDocument)
        {
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            // Get a PresentationPart object from the PresentationDocument object.
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            if (presentationPart != null &&
                presentationPart.Presentation != null)
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
                var shapes = from shape in slidePart.Slide.Descendants<Shape>()
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

        public static void ReplaceTheme(string document, string themeFile)
        {
            using (PresentationDocument presDoc = PresentationDocument.Open(document, true))
            {
                PresentationPart mainPart = presDoc.PresentationPart;

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

        // Determines whether the shape is a title shape.
        private static bool IsTitleShape(Shape shape)
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
    }
}
