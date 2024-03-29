﻿using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;

namespace Office_File_Explorer.App_Helpers
{
    class OfficeHelpers
    {
        public enum PropertyTypes : int
        {
            YesNo,
            Text,
            DateTime,
            NumberInteger,
            NumberDouble
        }

        /// <summary>
        /// Given a document name, a property name/value, and the property type add a custom property to a document. 
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="propertyName"></param>
        /// <param name="propertyValue"></param>
        /// <param name="propertyType"></param>
        /// <param name="fileType"></param>
        /// <returns>returns the original value, if it existed</returns>
        public static string SetCustomProperty(string fileName, string propertyName, object propertyValue, PropertyTypes propertyType, string fileType)
        {
            string returnValue = null;

            var newProp = new CustomDocumentProperty();
            bool propSet = false;

            // Calculate the correct type.
            switch (propertyType)
            {
                case PropertyTypes.DateTime:

                    // Be sure you were passed a real date and if so, format in the correct way. 
                    // The date/time value passed in should represent a UTC date/time.
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
                // If the code was not able to convert the property to a valid value, throw an exception.
                MessageBox.Show("The value entered does not match the specific type.  The value will be stored as text.", "Invalid Type", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                newProp.VTLPWSTR = new VTLPWSTR(propertyValue.ToString());
                propSet = true;
            }

            // Now that you have handled the parameters, start working on the document.
            newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
            newProp.Name = propertyName;

            if (fileType == "Word")
            {
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
            }
            else if (fileType == "Excel")
            {
                using (var document = SpreadsheetDocument.Open(fileName, true))
                {
                    var customProps = document.CustomFilePropertiesPart;
                    if (customProps == null)
                    {
                        customProps = document.AddCustomFilePropertiesPart();
                        customProps.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties();
                    }

                    var props = customProps.Properties;
                    if (props != null)
                    {
                        var prop = props.Where(p => ((CustomDocumentProperty)p).Name.Value == propertyName).FirstOrDefault();

                        if (prop != null)
                        {
                            returnValue = prop.InnerText;
                            prop.Remove();
                        }

                        props.AppendChild(newProp);
                        int pid = 2;
                        foreach (CustomDocumentProperty item in props)
                        {
                            item.PropertyId = pid++;
                        }
                        props.Save();
                    }
                }
            }
            else
            {
                using (var document = PresentationDocument.Open(fileName, true))
                {
                    var customProps = document.CustomFilePropertiesPart;
                    if (customProps == null)
                    {
                        customProps = document.AddCustomFilePropertiesPart();
                        customProps.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties();
                    }

                    var props = customProps.Properties;
                    if (props != null)
                    {
                        var prop = props.Where(p => ((CustomDocumentProperty)p).Name.Value == propertyName).FirstOrDefault();

                        if (prop != null)
                        {
                            returnValue = prop.InnerText;
                            prop.Remove();
                        }

                        props.AppendChild(newProp);
                        int pid = 2;
                        foreach (CustomDocumentProperty item in props)
                        {
                            item.PropertyId = pid++;
                        }

                        props.Save();
                    }
                }
            }

            return returnValue;
        }

        /// <summary>
        /// replace the current theme with a user specified theme
        /// </summary>
        /// <param name="document">document file lcoation</param>
        /// <param name="themeFile">theme xml file location</param>
        /// <param name="app">which app is the document</param>
        public static void ReplaceTheme(string document, string themeFile, string app)
        {
            if (app == "Word")
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
            else if (app == "PowerPoint")
            {
                using (PresentationDocument presDoc = PresentationDocument.Open(document, true))
                {
                    PresentationPart mainPart = presDoc.PresentationPart;                   
                    mainPart.DeletePart(mainPart.ThemePart);
                    ThemePart themePart = mainPart.AddNewPart<ThemePart>();

                    using (StreamReader streamReader = new StreamReader(themeFile))
                    using (StreamWriter streamWriter = new StreamWriter(themePart.GetStream(FileMode.Create)))
                    {
                        streamWriter.Write(streamReader.ReadToEnd());
                    }
                }
            }
            else
            {
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(document, true))
                {
                    WorkbookPart mainPart = excelDoc.WorkbookPart;
                    mainPart.DeletePart(mainPart.ThemePart);
                    ThemePart themePart = mainPart.AddNewPart<ThemePart>();

                    using (StreamReader streamReader = new StreamReader(themeFile))
                    using (StreamWriter streamWriter = new StreamWriter(themePart.GetStream(FileMode.Create)))
                    {
                        streamWriter.Write(streamReader.ReadToEnd());
                    }
                }
            }
        }

        /// <summary>
        /// Function to convert a macro enabled file to a non-macro enabled file
        /// </summary>
        /// <param name="fileName">file location</param>
        /// <param name="app">app type</param>
        /// <returns></returns>
        public static string ConvertMacroEnabled2NonMacroEnabled(string fileName, string app)
        {
            bool fileChanged = false;
            string newFileName = string.Empty;
            string fileExtension;

            if (app == "Word")
            {
                fileExtension = ".docx";
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
            }
            else if (app == "PowerPoint")
            {
                fileExtension = ".pptx";
                using (PresentationDocument document = PresentationDocument.Open(fileName, true))
                {
                    var docPart = document.PresentationPart;
                    var vbaPart = docPart.VbaProjectPart;
                    if (vbaPart != null)
                    {
                        docPart.DeletePart(vbaPart);
                        docPart.Presentation.Save();
                        document.ChangeDocumentType(PresentationDocumentType.Presentation);
                        fileChanged = true;
                    }
                }
            }
            else
            {
                fileExtension = ".xlsx";
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
                {
                    var docPart = document.WorkbookPart;
                    var vbaPart = docPart.VbaProjectPart;
                    if (vbaPart != null)
                    {
                        docPart.DeletePart(vbaPart);
                        docPart.Workbook.Save();
                        document.ChangeDocumentType(SpreadsheetDocumentType.Workbook);
                        fileChanged = true;
                    }
                }
            }

            // If anything goes wrong in this file handling,
            // the code will raise an exception back to the caller.
            if (fileChanged)
            {
                // Create the new filename.
                newFileName = Path.ChangeExtension(fileName, fileExtension);

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

        /// <summary>
        /// create a random guid for rsid values
        /// </summary>
        /// <returns></returns>
        public static string CreateNewRsid()
        {
            Guid g = Guid.NewGuid();
            return g.ToString();
        }

        /// <summary>
        /// add a new part that needs a relationship id
        /// </summary>
        /// <param name="document"></param>
        public static void AddNewPart(string document)
        {
            // Create a new word processing document.
            WordprocessingDocument wordDoc = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document);

            // Add the MainDocumentPart part in the new word processing document.
            var mainDocPart = wordDoc.AddNewPart<MainDocumentPart>("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "rId1");
            mainDocPart.Document = new Document();

            // Add the CustomFilePropertiesPart part in the new word processing document.
            var customFilePropPart = wordDoc.AddCustomFilePropertiesPart();
            customFilePropPart.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties();

            // Add the CoreFilePropertiesPart part in the new word processing document.
            var coreFilePropPart = wordDoc.AddCoreFilePropertiesPart();
            using (XmlTextWriter writer = new XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8))
            {
                writer.WriteRaw("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"></cp:coreProperties>");
                writer.Flush();
            }

            // Add the DigitalSignatureOriginPart part in the new word processing document.
            wordDoc.AddNewPart<DigitalSignatureOriginPart>("rId4");

            // Add the ExtendedFilePropertiesPart part in the new word processing document.
            var extendedFilePropPart = wordDoc.AddNewPart<ExtendedFilePropertiesPart>("rId5");
            extendedFilePropPart.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties();

            // Add the ThumbnailPart part in the new word processing document.
            wordDoc.AddNewPart<ThumbnailPart>("image/jpeg", "rId6");

            wordDoc.Close();
        }

        // To add a new document part to a package.
        public static void AddNewPart(string document, string fileName)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                CustomXmlPart myXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);

                using (FileStream stream = new FileStream(fileName, FileMode.Open))
                {
                    myXmlPart.FeedData(stream);
                }
            }
        }
    }
}
