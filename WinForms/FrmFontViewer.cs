using System;
using System.Drawing;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Office_File_Explorer.Forms
{
    public partial class FrmFontViewer : Form
    {
        // enum variables
        FontFamily[] fontFamilies;
        InstalledFontCollection installedFontCollection = new InstalledFontCollection();
        Color textColor;

        public FrmFontViewer(string inputText)
        {
            InitializeComponent();

            // populate the font list
            fontFamilies = installedFontCollection.Families;

            int count = fontFamilies.Length;
            for (int j = 1; j < count; j++)
            {
                CboFonts.Items.Add(fontFamilies[j].Name);
            }

            // select the first item
            CboFonts.SelectedIndex = 0;
            textColor = Color.Black;
            txbInput.Text = inputText;
        }

        private void UpdateDisplay()
        {
            if (CboFonts.SelectedItem.ToString() == "")
            {
                return;
            }

            // update picture box
            pBoxFont.Invalidate();
        }

        private void PBoxFont_Paint(object sender, PaintEventArgs e)
        {
            try
            {
                FontFamily fontFamily = new FontFamily(CboFonts.SelectedItem.ToString());
                Font font = new Font(fontFamily, (int)NudFontSize.Value, FontStyle.Regular, GraphicsUnit.Pixel);

                if (RdoRenderText.Checked == true)
                {
                    RenderText(e.Graphics, txbInput.Text, CboFonts.SelectedItem.ToString(), textColor, pBoxFont.ClientRectangle, (int)NudFontSize.Value);
                }
                else if (RdoDrawString.Checked == true)
                {
                    PointF pointF = new PointF(10, 10);
                    SolidBrush solidBrush = new SolidBrush(textColor);
                    e.Graphics.DrawString(txbInput.Text, font, solidBrush, pBoxFont.ClientRectangle);
                }
                else
                {
                    TextRenderer.DrawText(e.Graphics, txbInput.Text, font, new Point(10, 10), textColor);
                }
            }
            catch (NullReferenceException)
            {
                // if there is no selected item, just return
                return;
            }
        }

        private static void RenderText(IDeviceContext hdc, string text, string fontFamily, Color color, Rectangle region, int size)
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

        private void CboFonts_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDisplay();
        }

        private void BtnFontDetails_Click(object sender, EventArgs e)
        {
            FrmFontDetails ffd = new FrmFontDetails(CboFonts.SelectedItem.ToString(), (int)NudFontSize.Value);
            ffd.ShowDialog(this);
            ffd.Dispose();
        }

        private void BtnColorDlg_Click_1(object sender, EventArgs e)
        {
            DialogResult result = colorDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                textColor = colorDialog1.Color;
                if (textColor.IsKnownColor)
                {
                    LblColor.Text = textColor.Name;
                }
                else
                {
                    LblColor.Text = "Unknown: Hex = " + textColor.Name.ToUpper();
                }
                UpdateDisplay();
            }
        }

        private void CboFonts_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            UpdateDisplay();
        }

        private void TxbInput_TextChanged(object sender, EventArgs e)
        {
            UpdateDisplay();
        }

        private void RdoDrawString_CheckedChanged(object sender, EventArgs e)
        {
            UpdateDisplay();
        }

        private void RdoRenderText_CheckedChanged(object sender, EventArgs e)
        {
            UpdateDisplay();
        }

        private void RdoDrawText_CheckedChanged(object sender, EventArgs e)
        {
            UpdateDisplay();
        }

        private void NudFontSize_ValueChanged(object sender, EventArgs e)
        {
            UpdateDisplay();
        }

        private void BtnFontDetails_Click_1(object sender, EventArgs e)
        {
            FrmFontDetails ffd = new FrmFontDetails(CboFonts.SelectedItem.ToString(), (int)NudFontSize.Value);
            ffd.ShowDialog(this);
            ffd.Dispose();
        }
    }

    internal class NativeMethods
    {
        private NativeMethods() { }

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
