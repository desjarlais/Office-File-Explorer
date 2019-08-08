﻿/****************************** Module Header ******************************\
Module Name:  WordExtensionClass.cs
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

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Office_File_Explorer.Word_Helpers
{
    public static class WordExtensionClass
    {
        public static XDocument GetXDocument(this OpenXmlPart part)
        {
            XDocument xdoc = part.Annotation<XDocument>();
            if (xdoc != null)
            {
                return xdoc;
            }

            using (StreamReader sr = new StreamReader(part.GetStream()))
            using (XmlReader xr = XmlReader.Create(sr))
            {
                xdoc = XDocument.Load(xr);
            }

            part.AddAnnotation(xdoc);
            return xdoc;
        }

        public static bool HasPersonalInfo(WordprocessingDocument document)
        {
            // check for company name from /docProps/app.xml
            XNamespace x = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
            OpenXmlPart extendedFilePropertiesPart = document.ExtendedFilePropertiesPart;
            XDocument extendedFilePropertiesXDoc = extendedFilePropertiesPart.GetXDocument();
            string company = extendedFilePropertiesXDoc.Elements(x + "Properties").Elements(x + "Company").Select(e => (string)e)
                .Aggregate("", (s, i) => s + i);
            if (company.Length > 0)
            {
                return true;
            }

            // check for dc:creator, cp:lastModifiedBy from /docProps/core.xml
            XNamespace dc = "http://purl.org/dc/elements/1.1/";
            XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
            OpenXmlPart coreFilePropertiesPart = document.CoreFilePropertiesPart;
            XDocument coreFilePropertiesXDoc = coreFilePropertiesPart.GetXDocument();
            string creator = coreFilePropertiesXDoc.Elements(cp + "coreProperties").Elements(dc + "creator").Select(e => (string)e)
                .Aggregate("", (s, i) => s + i);
            if (creator.Length > 0)
            {
                return true;
            }

            string lastModifiedBy = coreFilePropertiesXDoc.Elements(cp + "coreProperties").Elements(cp + "lastModifiedBy").Select(e => (string)e)
                .Aggregate("", (s, i) => s + i);
            if (lastModifiedBy.Length > 0)
            {
                return true;
            }

            // check for nonexistence of removePersonalInformation and removeDateAndTime
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            OpenXmlPart documentSettingsPart = document.MainDocumentPart.DocumentSettingsPart;
            XDocument documentSettingsXDoc = documentSettingsPart.GetXDocument();
            XElement settings = documentSettingsXDoc.Root;
            if (settings.Element(w + "removePersonalInformation") == null)
            {
                return true;
            }

            if (settings.Element(w + "removeDateAndTime") == null)
            {
                return true;
            }

            return false;
        }

        public static void RemovePersonalInfo(WordprocessingDocument document)
        {
            // remove the company name from /docProps/app.xml
            // set TotalTime to "0"
            XNamespace x = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
            OpenXmlPart extendedFilePropertiesPart = document.ExtendedFilePropertiesPart;
            XDocument extendedFilePropertiesXDoc = extendedFilePropertiesPart.GetXDocument();
            extendedFilePropertiesXDoc.Elements(x + "Properties").Elements(x + "Company").Remove();
            XElement totalTime = extendedFilePropertiesXDoc.Elements(x + "Properties").Elements(x + "TotalTime").FirstOrDefault();
            if (totalTime != null)
            {
                totalTime.Value = "0";
            }

            using (XmlWriter xw = XmlWriter.Create(extendedFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write)))
            {
                extendedFilePropertiesXDoc.Save(xw);
            }

            // remove the values of dc:creator, cp:lastModifiedBy from /docProps/core.xml
            // set cp:revision to "1"
            XNamespace dc = "http://purl.org/dc/elements/1.1/";
            XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
            OpenXmlPart coreFilePropertiesPart = document.CoreFilePropertiesPart;
            XDocument coreFilePropertiesXDoc = coreFilePropertiesPart.GetXDocument();
            foreach (var textNode in coreFilePropertiesXDoc.Elements(cp + "coreProperties")
                                                           .Elements(dc + "creator")
                                                           .Nodes()
                                                           .OfType<XText>())
            {
                textNode.Value = "";
            }

            foreach (var textNode in coreFilePropertiesXDoc.Elements(cp + "coreProperties")
                                                           .Elements(cp + "lastModifiedBy")
                                                           .Nodes()
                                                           .OfType<XText>())
            {
                textNode.Value = "";
            }

            XElement revision = coreFilePropertiesXDoc.Elements(cp + "coreProperties").Elements(cp + "revision").FirstOrDefault();
            if (revision != null)
            {
                revision.Value = "1";
            }

            using (XmlWriter xw = XmlWriter.Create(coreFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write)))
            {
                coreFilePropertiesXDoc.Save(xw);
            }

            // add w:removePersonalInformation, w:removeDateAndTime to /word/settings.xml
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            OpenXmlPart documentSettingsPart = document.MainDocumentPart.DocumentSettingsPart;
            XDocument documentSettingsXDoc = documentSettingsPart.GetXDocument();
            // add the new elements in the right position.  Add them after the following three elements
            // (which may or may not exist in the xml document).
            XElement settings = documentSettingsXDoc.Root;
            XElement lastOfTop3 = settings.Elements()
                .Where(e => e.Name == w + "writeProtection" ||
                    e.Name == w + "view" ||
                    e.Name == w + "zoom")
                .InDocumentOrder()
                .LastOrDefault();
            if (lastOfTop3 == null)
            {
                // none of those three exist, so add as first children of the root element
                settings.AddFirst(
                    settings.Elements(w + "removePersonalInformation").Any() ?
                        null :
                        new XElement(w + "removePersonalInformation"),
                    settings.Elements(w + "removeDateAndTime").Any() ?
                        null :
                        new XElement(w + "removeDateAndTime")
                );
            }
            else
            {
                // one of those three exist, so add after the last one
                lastOfTop3.AddAfterSelf(
                    settings.Elements(w + "removePersonalInformation").Any() ?
                        null :
                        new XElement(w + "removePersonalInformation"),
                    settings.Elements(w + "removeDateAndTime").Any() ?
                        null :
                        new XElement(w + "removeDateAndTime")
                );
            }
            using (XmlWriter xw = XmlWriter.Create(documentSettingsPart.GetStream(FileMode.Create, FileAccess.Write)))
            {
                documentSettingsXDoc.Save(xw);
            }
        }

        private static string GetStyleIdFromStyleName(MainDocumentPart mainPart, string styleName)
        {
            StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart;
            string styleId = stylePart.Styles.Descendants<StyleName>()
                .Where(s => s.Val.Value.Equals(styleName))
                .Select(n => ((Style)n.Parent).StyleId).FirstOrDefault();
            return styleId ?? styleName;
        }

        public static IEnumerable<Paragraph> ParagraphsByStyleName(this MainDocumentPart mainPart, string styleName)
        {
            string styleId = GetStyleIdFromStyleName(mainPart, styleName);
            IEnumerable<Paragraph> paraList = mainPart.Document.Descendants<Paragraph>()
                .Where(p => IsParagraphInStyle(p, styleId));
            return paraList;
        }

        private static bool IsParagraphInStyle(Paragraph p, string styleId)
        {
            ParagraphProperties pPr = p.GetFirstChild<ParagraphProperties>();
            if (pPr != null)
            {
                ParagraphStyleId paraStyle = pPr.ParagraphStyleId;

                if (paraStyle != null)
                {
                    return paraStyle.Val.Value.Equals(styleId);
                }
            }
            return false;
        }

        public static IEnumerable<Run> RunsByStyleName(this MainDocumentPart mainPart, string styleName)
        {
            string styleId = GetStyleIdFromStyleName(mainPart, styleName);

            IEnumerable<Run> runList = mainPart.Document.Descendants<Run>()
                .Where(r => IsRunInStyle(r, styleId));
            return runList;
        }

        private static bool IsRunInStyle(Run r, string styleId)
        {
            RunProperties rPr = r.GetFirstChild<RunProperties>();

            if (rPr != null)
            {
                RunStyle runStyle = rPr.RunStyle;
                if (runStyle != null)
                {
                    return runStyle.Val.Value.Equals(styleId);
                }
            }
            return false;
        }

        public static IEnumerable<Table> TablesByStyleName(this MainDocumentPart mainPart, string styleName)
        {
            string styleId = GetStyleIdFromStyleName(mainPart, styleName);

            IEnumerable<Table> tableList = mainPart.Document.Descendants<Table>()
                .Where(t => IsTableInStyle(t, styleId));
            return tableList;
        }

        private static bool IsTableInStyle(Table tbl, string styleId)
        {
            TableProperties tblPr = tbl.GetFirstChild<TableProperties>();

            if (tblPr != null)
            {
                TableStyle tblStyle = tblPr.TableStyle;

                if (tblStyle != null)
                {
                    return tblStyle.Val.Value.Equals(styleId);
                }
            }
            return false;
        }
    }
}
