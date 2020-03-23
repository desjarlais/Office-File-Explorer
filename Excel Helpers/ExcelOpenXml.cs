using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace Office_File_Explorer.Excel_Helpers
{
    class ExcelOpenXml
    {
        public static bool RemoveExternalLinks(string docName)
        {
            bool linkRemoved = false;

            // Given a document name, remove all headers and footers.
            using (SpreadsheetDocument xlDoc = SpreadsheetDocument.Open(docName, true))
            {
                if (xlDoc.WorkbookPart.GetPartsCountOfType<ExternalWorkbookPart>() > 0)
                {
                    // Remove header and footer parts.
                    xlDoc.WorkbookPart.DeleteParts(xlDoc.WorkbookPart.ExternalWorkbookParts);

                    // Remove references to the headers and footers.
                    var wbLinks = xlDoc.WorkbookPart.Workbook.Descendants<ExternalReference>().ToList();
                    foreach (var link in wbLinks)
                    {
                        link.Parent.RemoveChild(link);
                        linkRemoved = true;
                    }

                    xlDoc.Close();
                }
            }

            return linkRemoved;
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
                            values.Add(text + " ");
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
                        values.Add(text + " ");
                    }
                }

                return values;
            }
        }
    }
}
