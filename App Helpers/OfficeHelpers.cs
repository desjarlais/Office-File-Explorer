﻿using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using Office_File_Explorer.Word_Helpers;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

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

        public static string SetCustomProperty(string fileName, string propertyName, object propertyValue, PropertyTypes propertyType, string fileType)
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
                // If the code was not able to convert the property to a valid value, throw an exception.
                MessageBox.Show("The value entered does not match the specific type.  The value will be stored as text.", "Invalid Type", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                newProp.VTLPWSTR = new VTLPWSTR(propertyValue.ToString());
                propSet = true;
            }

            // Now that you have handled the parameters, start
            // working on the document.
            Guid id = Guid.NewGuid();
            newProp.FormatId = id.ToString();
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
            else
            {
                using (var document = PresentationDocument.Open(fileName, true))
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
            return returnValue;
        }


        /// <summary>
        /// replace the current theme with a user specified theme
        /// </summary>
        /// <param name="document"></param>
        /// <param name="themeFile"></param>
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
            else if (app == "Powerpoint")
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
            else if (app == "Excel")
            {
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(document, true))
                {
                    WorkbookPart mainPart = excelDoc.WorkbookPart;

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
            else
            {
                return;
            }
        }
    }
}
