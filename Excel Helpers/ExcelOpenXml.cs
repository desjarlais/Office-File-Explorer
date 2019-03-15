using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.Collections;

namespace Office_File_Explorer.Excel_Helpers
{
    class ExcelOpenXml
    {
        public static void RemoveExternalLinks(string docName)
        {
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
                    }

                    xlDoc.Close();
                }
            }
        }

        public static List<Sheet> GetWorkSheets(string fileName)
        {
            List<Sheet> returnVal = new List<Sheet>();

            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(fileName, true))
            {
                WorkbookPart wbPart = excelDoc.WorkbookPart;
                Sheets theSheets = wbPart.Workbook.Sheets;
                

                foreach (Sheet sheet in theSheets)
                {
                    returnVal.Add(sheet);
                }
            }

            return returnVal;
        }

        public static List<Sheet> GetHiddenSheets(string fileName)
        {
            List<Sheet> returnVal = new List<Sheet>();

            using (SpreadsheetDocument document =
                SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                var sheets = wbPart.Workbook.Descendants<Sheet>();

                // Look for sheets where there is a State attribute defined, 
                // where the State has a value,
                // and where the value is either Hidden or VeryHidden.
                var hiddenSheets = sheets.Where((item) => item.State != null &&
                    item.State.HasValue &&
                    (item.State.Value == SheetStateValues.Hidden ||
                    item.State.Value == SheetStateValues.VeryHidden));

                returnVal = hiddenSheets.ToList();
            }
            return returnVal;
        }
    }
}
