/****************************** Module Header ******************************\
Module Name:  ExcelOpenXml.cs
Project:      Office File Explorer

Excel Open Xml Helper Functions

This source is subject to the following license.
See https://github.com/desjarlais/Office-File-Explorer/blob/master/LICENSE
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using Office_File_Explorer.App_Helpers;
using System;

namespace Office_File_Explorer.Excel_Helpers
{
    class ExcelOpenXml
    {
        public static bool RemoveExternalLinks(string docName)
        {
            // step 1 remove the externalworkbookpart
            // step 2 remove the oleobject attributes in the worksheet
            bool linksRemoved = false;

            using (SpreadsheetDocument xlDoc = SpreadsheetDocument.Open(docName, true))
            {
                var ewParts = xlDoc.WorkbookPart.ExternalWorkbookParts.ToList();
                var wksParts = xlDoc.WorkbookPart.WorksheetParts.ToList();

                if (ewParts.Count > 0)
                {
                    foreach (ExternalWorkbookPart ewp in ewParts)
                    {
                        xlDoc.WorkbookPart.DeletePart(ewp);
                        linksRemoved = true;
                    }

                    foreach (WorksheetPart wp in wksParts)
                    {
                        foreach (OpenXmlElement child in wp.Worksheet.ChildElements)
                        {
                            if (child.LocalName == "oleObjects")
                            {
                                foreach (OpenXmlElement childOleObjects in child.ChildElements)
                                {
                                    if (childOleObjects.LocalName == "AlternateContent")
                                    {
                                        foreach (OpenXmlElement ac in childOleObjects.ChildElements)
                                        {
                                            if (ac.LocalName == "Choice")
                                            {
                                                foreach (OpenXmlElement oleObj in ac.ChildElements)
                                                {
                                                    if (oleObj.LocalName == "oleObject")
                                                    {
                                                        oleObj.ClearAllAttributes();
                                                        linksRemoved = true;
                                                    }
                                                }
                                            }

                                            if (ac.LocalName == "Fallback")
                                            {
                                                ac.Remove();
                                                linksRemoved = true;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (linksRemoved)
                    {
                        // save the document
                        xlDoc.WorkbookPart.Workbook.Save();
                    }

                    xlDoc.Close();
                }
            }

            return linksRemoved;
        }

        public static void RemoveExternalLink(string fileName, StringValue linkToDelete)
        {
            bool linkRemoved = false;

            foreach (ExternalReference eLink in GetExternalLinks(fileName))
            {
                if (linkToDelete == eLink.Id)
                {
                    ExternalReference er = (ExternalReference)eLink;
                    er.Remove();
                    linkRemoved = true;
                }
            }

            if (linkRemoved == true)
            {

            }
        }

        public static List<ExternalReference> GetExternalLinks(string fileName)
        {
            List<ExternalReference> returnVal = new List<ExternalReference>();

            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(fileName, false))
            {
                var wbLinks = excelDoc.WorkbookPart.Workbook.Descendants<ExternalReference>().ToList();

                foreach (ExternalReference eLink in wbLinks)
                {
                    returnVal.Add(eLink);
                }
            }

            return returnVal;
        }

        public static List<Sheet> GetSheets(string fileName, bool fileIsEditable)
        {
            List<Sheet> returnVal = new List<Sheet>();

            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(fileName, fileIsEditable))
            {
                foreach (Sheet sheet in excelDoc.WorkbookPart.Workbook.Sheets)
                {
                    returnVal.Add(sheet);
                }
            }

            return returnVal;
        }

        public static List<Worksheet> GetWorkSheets(string fileName, bool fileIsEditable)
        {
            List<Worksheet> returnVal = new List<Worksheet>();

            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(fileName, fileIsEditable))
            {
                foreach (WorksheetPart wsPart in excelDoc.WorkbookPart.WorksheetParts)
                {
                    returnVal.Add(wsPart.Worksheet);
                }
            }

            return returnVal;
        }

        public static List<Sheet> GetHiddenSheets(string fileName, bool fileIsEditable)
        {
            List<Sheet> returnVal = new List<Sheet>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, fileIsEditable))
            {
                var sheets = document.WorkbookPart.Workbook.Descendants<Sheet>();

                // Look for sheets where there is a State attribute defined, 
                // where the State has a value,
                // and where the value is either Hidden or VeryHidden.
                var hiddenSheets = sheets.Where((item) => item.State != null && item.State.HasValue &&
                (item.State.Value == SheetStateValues.Hidden || item.State.Value == SheetStateValues.VeryHidden));

                returnVal = hiddenSheets.ToList();
            }
            return returnVal;
        }
        
        // The DOM approach.
        // Note that the code below works only for cells that contain numeric values.
        // 
        public static List<string> ReadExcelFileDOM(string fileName)
        {
            List<string> values = new List<string>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                string text;
                
                foreach (Row r in sheetData.Elements<Row>())
                {
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        if (c.CellValue != null)
                        {
                            text = c.CellValue.Text;
                            values.Add(text + StringResources.wSpaceChar);
                        }
                    }
                }

                return values;
            }
        }

        // The SAX approach.
        public static List<string> ReadExcelFileSAX(string fileName)
        {
            List<string> values = new List<string>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                
                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                string text;

                while (reader.Read())
                {
                    if (reader.ElementType == typeof(CellValue))
                    {
                        text = reader.GetText();
                        values.Add(text + StringResources.wSpaceChar);
                    }
                }

                return values;
            }
        }
    }
}
