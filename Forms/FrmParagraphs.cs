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
                    // TODO: lot of work to getting this dialog to render the content from the paragraph
                    foreach (Paragraph p in pList)
                    {
                        if (p.InnerText == "")
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
                                    if (obj.GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Style")
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
            PBoxFont.Invalidate();
            Cursor = Cursors.Default;
        }

        private static void RenderText(IDeviceContext hdc, string text, string fontFamily, System.Drawing.Color color, Rectangle region, int size)
        {
            // create the handle of DC
            HandleRef h = new HandleRef(null, hdc.GetHdc());
            // create the font
            HandleRef p = new HandleRef(null, NativeMethods.CreateFont
                (size, 0, 0, 0, 0, 0, 0, 0, 1/*Ansi_encoding*/, 0, 0, 4, 0, fontFamily));
            try
            {
                // use the font in the DC
                NativeMethods.SelectObject((IntPtr)h, p.Handle);
                // set the background to transparent
                NativeMethods.SetBkMode((IntPtr)h, 1);
                // set the color of the text
                NativeMethods.SetTextColor((IntPtr)h, ColorTranslator.ToWin32(color));
                // draw the text to the region
                NativeMethods.DrawText(h, text, region, 0x0100);
            }
            finally
            {
                // release the resources
                NativeMethods.DeleteObject((IntPtr)p);
                hdc.ReleaseHdc();
            }
        }

        public static string ParagraphText(XElement e)
        {
            XNamespace w = e.Name.Namespace;
            return e
                   .Elements(w + "r")
                   .Elements(w + "t")
                   .StringConcatenate(element => (string)element);
        }

        /// <summary>
        /// Gets the System.Drawing.Color object from hex string.
        /// </summary>
        /// <param name="hexString">The hex string.</param>
        /// <returns></returns>
        private System.Drawing.Color GetSystemDrawingColorFromHexString(string hexString)
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

        private void PBoxFont_Paint_1(object sender, PaintEventArgs e)
        {
            try
            {
                if (fontColor != null)
                {
                    Color fColor = GetSystemDrawingColorFromHexString(fontColor);
                    RenderText(e.Graphics, richTextBox1.Text, "Calibri", fColor, PBoxFont.ClientRectangle, fontSize);
                }
                else
                {
                    RenderText(e.Graphics, richTextBox1.Text, "Calibri", Color.Black, PBoxFont.ClientRectangle, fontSize);
                }
            }
            catch (NullReferenceException)
            {
                return;
            }
        }
    }

    internal class NativeMethods
    {
        private NativeMethods()
        {
        }

        struct Rect
        {
            public long Left, Top, Right, Bottom;
            public Rect(Rectangle rect)
            {
                Left = rect.Left;
                Top = rect.Top;
                Right = rect.Right;
                Bottom = rect.Bottom;
            }
        }
        [DllImport("gdi32.dll")]
        internal static extern IntPtr CreateFont(
            int nHeight,
            int nWidth,
            int nEscapement,
            int nOrientation,
            int fnWeight,
            uint fdwItalic,
            uint fdwUnderline,
            uint fdwStrikeOut,
            uint fdwCharSet,
            uint fdwOutputPrecision,
            uint fdwClipPrecision,
            uint fdwQuality,
            uint fdwPitchAndFamily,
            string lpszFace
            );

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int DrawText(HandleRef hDC, string lpchText, int nCount, ref Rect lpRect, uint uFormat);
        internal static int DrawText(HandleRef hDC, string text, Rectangle rect, uint format)
        {
            var r = new Rect(rect);
            return DrawText(hDC, text, text.Length, ref r, format);
        }

        /// <summary>Selects an object into the specified device context (DC). The new object replaces the previous object of the same type.</summary>
        /// <param name="hdc">A handle to the DC.</param>
        /// <param name="hgdiobj">A handle to the object to be selected.</param>
        /// <returns>
        ///   <para>If the selected object is not a region and the function succeeds, the return value is a handle to the object being replaced. If the selected object is a region and the function succeeds, the return value is one of the following values.</para>
        ///   <para>SIMPLEREGION - Region consists of a single rectangle.</para>
        ///   <para>COMPLEXREGION - Region consists of more than one rectangle.</para>
        ///   <para>NULLREGION - Region is empty.</para>
        ///   <para>If an error occurs and the selected object is not a region, the return value is <c>NULL</c>. Otherwise, it is <c>HGDI_ERROR</c>.</para>
        /// </returns>
        /// <remarks>
        ///   <para>This function returns the previously selected object of the specified type. An application should always replace a new object with the original, default object after it has finished drawing with the new object.</para>
        ///   <para>An application cannot select a single bitmap into more than one DC at a time.</para>
        ///   <para>ICM: If the object being selected is a brush or a pen, color management is performed.</para>
        /// </remarks>
        [DllImport("gdi32.dll", EntryPoint = "SelectObject")]
        public static extern IntPtr SelectObject([In] IntPtr hdc, [In] IntPtr hgdiobj);

        [DllImport("gdi32.dll")]
        public static extern int SetBkMode(IntPtr hdc, int iBkMode);

        [DllImport("gdi32.dll")]
        public static extern uint SetTextColor(IntPtr hdc, int crColor);

        /// <summary>Deletes a logical pen, brush, font, bitmap, region, or palette, freeing all system resources associated with the object. After the object is deleted, the specified handle is no longer valid.</summary>
        /// <param name="hObject">A handle to a logical pen, brush, font, bitmap, region, or palette.</param>
        /// <returns>
        ///   <para>If the function succeeds, the return value is nonzero.</para>
        ///   <para>If the specified handle is not valid or is currently selected into a DC, the return value is zero.</para>
        /// </returns>
        /// <remarks>
        ///   <para>Do not delete a drawing object (pen or brush) while it is still selected into a DC.</para>
        ///   <para>When a pattern brush is deleted, the bitmap associated with the brush is not deleted. The bitmap must be deleted independently.</para>
        /// </remarks>
        [DllImport("gdi32.dll", EntryPoint = "DeleteObject")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DeleteObject([In] IntPtr hObject);
    }
}
