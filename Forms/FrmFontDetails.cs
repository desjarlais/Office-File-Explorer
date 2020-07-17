using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Office_File_Explorer.Forms
{
    public partial class FrmFontDetails : Form
    {
        string fName;

        public FrmFontDetails(string fontName, int fSize)
        {
            InitializeComponent();

            fName = fontName;
            LstFontInfo.Items.Clear();

            // setup font variables
            int ascent;
            float ascentPixel;
            int descent;
            float descentPixel;
            int lineSpacing;
            float lineSpacingPixel;

            FontFamily fontFamily = new FontFamily(fontName);
            Font font = new Font(fontFamily, fSize, FontStyle.Regular, GraphicsUnit.Pixel);
            int fEmHeight = fontFamily.GetEmHeight(FontStyle.Regular);

            // update list box
            LstFontInfo.Items.Add("Font Name: " + font.Name);
            LstFontInfo.Items.Add("System Font Name: " + font.SystemFontName);
            LstFontInfo.Items.Add("Font size: " + font.Size);
            LstFontInfo.Items.Add("Font size in points: " + font.SizeInPoints);
            LstFontInfo.Items.Add("Font height: " + font.Height);
            LstFontInfo.Items.Add("Font EmHeight: " + fEmHeight);
            LstFontInfo.Items.Add("Font unit: " + font.Unit);
            LstFontInfo.Items.Add("GdiCharSet: " + GdiCharacterSet(font.GdiCharSet));
            LstFontInfo.Items.Add("Is System Font: " + font.IsSystemFont);
            LstFontInfo.Items.Add("GdiVerticalFont: " + font.GdiVerticalFont);
            LstFontInfo.Items.Add("");

            // Display the ascent in design units and pixels.
            ascent = fontFamily.GetCellAscent(FontStyle.Regular);

            // 14.484375 = 16.0 * 1854 / 2048
            ascentPixel = font.Size * ascent / fEmHeight;
            LstFontInfo.Items.Add("The ascent is " + ascent + " design units, " + ascentPixel + " pixels.");

            // Display the descent in design units and pixels.
            descent = fontFamily.GetCellDescent(FontStyle.Regular);

            // 3.390625 = 16.0 * 434 / 2048
            descentPixel = font.Size * descent / fEmHeight;
            LstFontInfo.Items.Add("The descent is " + descent + " design units, " + descentPixel + " pixels.");

            // Display the line spacing in design units and pixels.
            lineSpacing = fontFamily.GetLineSpacing(FontStyle.Regular);

            // 18.398438 = 16.0 * 2355 / 2048
            lineSpacingPixel = font.Size * lineSpacing / fEmHeight;
            LstFontInfo.Items.Add("The line spacing is " + lineSpacing + " design units, " + lineSpacingPixel + " pixels.");
        }

        public string GdiCharacterSet(byte fGdiCharSet)
        {
            string output;

            switch (fGdiCharSet)
            {
                case 0:
                    output = "Ansi";
                    break;
                case 1:
                    output = "Default";
                    break;
                case 2:
                    output = "Symbol";
                    break;
                case 77:
                    output = "Mac";
                    break;
                case 128:
                    output = "ShiftJis";
                    break;
                case 129:
                    output = "Hangul";
                    break;
                case 130:
                    output = "Johab";
                    break;
                case 134:
                    output = "GB2312";
                    break;
                case 136:
                    output = "ChineseBig5";
                    break;
                case 161:
                    output = "Greek";
                    break;
                case 162:
                    output = "Turkish";
                    break;
                case 163:
                    output = "Vietnamese";
                    break;
                case 177:
                    output = "Hebrew";
                    break;
                case 178:
                    output = "Arabic";
                    break;
                case 186:
                    output = "Baltic";
                    break;
                case 204:
                    output = "Russian";
                    break;
                case 222:
                    output = "Thai";
                    break;
                case 238:
                    output = "EastEurope";
                    break;
                case 255:
                    output = "OEM";
                    break;
                default:
                    output = "";
                    break;
            }
            return output;
        }

        private void PBoxAlias_Paint(object sender, PaintEventArgs e)
        {
            // setup font display
            FontFamily fontFamily = new FontFamily(fName);
            Font font = new Font(
               fontFamily,
               16, FontStyle.Regular,
               GraphicsUnit.Pixel);
            PointF pointF = new PointF(10, 10);
            SolidBrush solidBrush = new SolidBrush(Color.Black);

            // display the antialias comparison text
            pointF.Y += font.Height * 2;
            string string1 = "SingleBitPerPixel";
            e.Graphics.TextRenderingHint = TextRenderingHint.SingleBitPerPixel;
            e.Graphics.DrawString(string1, font, solidBrush, pointF);

            string string2 = "AntiAlias";
            pointF.Y += font.Height;
            e.Graphics.TextRenderingHint = TextRenderingHint.AntiAlias;
            e.Graphics.DrawString(string2, font, solidBrush, pointF);
        }
    }
}
