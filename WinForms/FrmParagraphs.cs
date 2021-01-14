using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Office_File_Explorer.App_Helpers;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Color = System.Drawing.Color;

using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Xml.Linq;
using Office_File_Explorer.Word_Helpers;
using System.Linq;
using System.Globalization;

namespace Office_File_Explorer.Forms
{
    public partial class FrmParagraphs : Form
    {
        string filePath;
        string styleName, fontColor;
        int fontSize;

        public FrmParagraphs(string file)
        {
            InitializeComponent();
            filePath = file;
            PopulateParagraphComboBox();

            if (cbParagraphs.Items.Count > 0)
            {
                cbParagraphs.SelectedIndex = 0;
            }
        }

        public void PopulateParagraphComboBox()
        {
            try
            {
                int count = 0;

                using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true))
                {
                    MainDocumentPart mPart = package.MainDocumentPart;
                    IEnumerable<Paragraph> pList = mPart.Document.Descendants<Paragraph>();
                    
                    foreach (var v in pList)
                    {
                        count++;
                    }
                }

                if (count == 0)
                {
                    cbParagraphs.Items.Add("None");
                }
                else
                {
                    int n = 0;
                    do
                    {
                        n++;
                        cbParagraphs.Items.Add("Paragraph #" + n);
                    } while (n < count);
                }

                lblParaCount.Text = "Paragraph Count = " + count;
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("PopulateParagraphComboBox Error: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public void ListParagraphs()
        {
            try
            {
                string sNum = cbParagraphs.SelectedItem.ToString();
                char last = sNum[sNum.Length - 1];
                double pNum = Char.GetNumericValue(last);

                using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true))
                {
                    MainDocumentPart mPart = package.MainDocumentPart;
                    IEnumerable<Paragraph> pList = mPart.Document.Descendants<Paragraph>();
                    int count = 0;

                    richTextBox1.Clear();
                    
                    foreach (Paragraph p in pList)
                    {
                        if (p.InnerText == string.Empty)
                        {
                            continue;
                        }
                        else
                        {
                            count++;
                            if (count == pNum)
                            {
                                GetRunDetails(p);

                                ParagraphProperties pPr = p.Elements<ParagraphProperties>().First();
                                StyleDefinitionsPart stPart = package.MainDocumentPart.StyleDefinitionsPart;
                                
                                foreach (var obj in stPart.Styles)
                                {
                                    if (obj.GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Style" && pPr.ParagraphStyleId != null)
                                    {
                                        Style style = (Style)obj;
                                        if (style.StyleId.ToString() == pPr.ParagraphStyleId.Val)
                                        {
                                            StyleRunProperties srPr = style.StyleRunProperties;
                                            if (srPr != null)
                                            {
                                                if (srPr.FontSize != null)
                                                {
                                                    fontSize = Convert.ToInt32(srPr.FontSize.Val);
                                                    LblFontSize.Text = srPr.FontSize.Val;
                                                }

                                                styleName = style.StyleId.ToString();
                                                LblStyleName.Text = styleName;

                                                if (srPr.Color != null)
                                                {
                                                    fontColor = "#" + srPr.Color.Val;
                                                    LblFontColor.Text = fontColor;
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
            catch (NullReferenceException nre)
            {
                LoggingHelper.Log("ListParagraphs Error: " + nre.Message);
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("ListParagraphs Error: " + ex.Message);
            }
        }

        public void GetRunDetails(Paragraph p)
        {
            RunProperties rPr = new RunProperties();
            foreach (Run r in p.Descendants<Run>())
            {
                rPr = r.RunProperties;
                richTextBox1.Text += r.InnerText;
            }
        }
        
        private void CbParagraphs_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            ListParagraphs();
            Cursor = Cursors.Default;
        }

        public static string ParagraphText(XElement e)
        {
            XNamespace w = e.Name.Namespace;
            return e
                   .Elements(w + "r")
                   .Elements(w + "t")
                   .StringConcatenate(element => (string)element);
        }

        private void BtnShowViewer_Click(object sender, EventArgs e)
        {
            FrmFontViewer fFrm = new FrmFontViewer(richTextBox1.Text)
            {
                Owner = this
            };
            fFrm.ShowDialog();
        }

        /// <summary>
        /// Gets the System.Drawing.Color object from hex string.
        /// </summary>
        /// <param name="hexString">The hex string.</param>
        /// <returns></returns>
        private Color GetSystemDrawingColorFromHexString(string hexString)
        {
            try
            {
                if (!System.Text.RegularExpressions.Regex.IsMatch(hexString, @"\B#(?:[a-fA-F0–9]{6}|[a-fA-F0–9]{3})\b"))
                {
                    throw new ArgumentException();
                }

                int red = int.Parse(hexString.Substring(1, 2), NumberStyles.HexNumber);
                int green = int.Parse(hexString.Substring(3, 2), NumberStyles.HexNumber);
                int blue = int.Parse(hexString.Substring(5, 2), NumberStyles.HexNumber);
                return Color.FromArgb(red, green, blue);
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("GetSystemDrawingColorFromHexString Error: " + ex.Message);
                return Color.Black;
            }
        }
    }
}
