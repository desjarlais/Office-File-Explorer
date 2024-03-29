﻿using Office_File_Explorer.App_Helpers;
using System;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Office_File_Explorer.Forms
{
    public partial class FrmClipboardViewer : Form
    {
        // global var
        private IDataObject data;
        private bool DisplayMemoryInHex;
        private bool DisplayRichText;
        private bool DisplayPictures;
        private bool AutoRefresh;
        private IntPtr _chainedWnd = (IntPtr)0;

        // win32 imports
        [DllImport(StringResources.user32)]
        static extern IntPtr SetClipboardViewer(IntPtr hWndNewViewer);

        [DllImport(StringResources.user32)]
        static extern IntPtr ChangeClipboardChain(IntPtr hWndRemove, IntPtr hWndNewNext);

        [DllImport(StringResources.user32)]
        public static extern void SendMessage(IntPtr hwnd, uint wMsg, IntPtr wParam, IntPtr lParam);

        [DllImport(StringResources.user32, SetLastError = true)]
        static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport(StringResources.user32)]
        static extern IntPtr GetClipboardData(uint uFormat);

        [DllImport(StringResources.user32, SetLastError = true)]
        static extern bool CloseClipboard();

        [DllImport(StringResources.user32)]
        private static extern int IsClipboardFormatAvailable(int wFormat);

        [DllImport(StringResources.user32, SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport(StringResources.user32)]
        static extern IntPtr GetClipboardOwner();

        [DllImport(StringResources.gdi32)]
        static extern IntPtr CopyEnhMetaFile(IntPtr hemfSrc, string lpszFile);

        [DllImport(StringResources.gdi32)]
        static extern bool DeleteEnhMetaFile(IntPtr hemf);

        const int WM_DRAWCLIPBOARD = 0x0308;
        const int WM_CHANGECBCHAIN = 0x030D;

        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case WM_DRAWCLIPBOARD:
                    if (AutoRefresh)
                    {
                        RefreshClipboard();
                    }
                    break;
                case WM_CHANGECBCHAIN:
                    if (m.WParam == _chainedWnd)
                    {
                        _chainedWnd = m.LParam;
                    }
                    else
                    {
                        SendMessage(_chainedWnd, (uint)m.Msg, m.WParam, m.LParam);
                    }
                    break;
            }
            base.WndProc(ref m);
        }

        private void SetDataText(string text)
        {
            rtbClipData.Text = text;
            rtbClipData.SelectAll();
            rtbClipData.SelectionFont = new Font("Courier New", 10, (FontStyle)0 /* no style */ );
            rtbClipData.SelectionColor = Color.FromKnownColor(KnownColor.Black);
            rtbClipData.SelectionBackColor = Color.FromKnownColor(KnownColor.White);
            rtbClipData.Select(0, 0);
            rtbClipData.Visible = true;
            pbClipData.Visible = false;
        }

        private void UpdateDisplay()
        {
            // check for empty clipboard
            if (lbClipFormats.SelectedItem.ToString() == "The clipboard is emtpy.")
            {
                return;
            }

            string Format = (string)lbClipFormats.SelectedItem;
            if (data.GetDataPresent(Format))
            {
                object dataobject;

                try
                {
                    Cursor = Cursors.WaitCursor;

                    // check for image formats
                    if (Format == "EnhancedMetafile")
                    {
                        Image imgFormat = GetEnhMetaImageFromClipboard();
                        pbClipData.Image = imgFormat;
                        pbClipData.Visible = true;
                        rtbClipData.Visible = false;
                        return;
                    }

                    dataobject = data.GetData(Format);

                    // no data in format
                    if (dataobject == null)
                    {
                        SetDataText(string.Format("No data returned for the {0} format.", Format));
                        return;
                    }

                    // rtf
                    if (Format == "Rich Text Format" && DisplayRichText)
                    {
                        rtbClipData.Rtf = dataobject.ToString();
                    }
                    else
                    {
                        // string
                        if (dataobject is string)
                        {
                            if (Format == "System.String")
                            {
                                SetDataText(Clipboard.GetText(TextDataFormat.Text));
                            }
                            else if (Format == "UnicodeText")
                            {
                                SetDataText(Clipboard.GetText(TextDataFormat.UnicodeText));
                            }
                            else
                            {
                                // display other type of string text
                                SetDataText((string)dataobject);
                            }
                        }
                        else
                        {
                            // bitmap
                            if (dataobject is Bitmap && DisplayPictures)
                            {
                                pbClipData.Image = (Bitmap)dataobject;
                                pbClipData.Visible = true;
                            }
                            else
                            {
                                // memory stream
                                if (dataobject is Stream)
                                {
                                    Stream dataStream = dataobject as Stream;
                                    byte[] buffer = new byte[dataStream.Length];
                                    int bytesread = dataStream.Read(buffer, 0, buffer.Length);
                                    StringBuilder s = new StringBuilder(buffer.Length + 100);

                                    s.AppendFormat("{0}: {1:N0} bytes", dataobject, dataStream.Length);
                                    s.Append("\r\n\r\n");
                                    for (int i = 0; i < bytesread; i++)
                                    {
                                        byte b = buffer[i];
                                        if (DisplayMemoryInHex)
                                        {
                                            s.Append(b.ToString("X2"));
                                            if ((i & 0x7) == 0x7)
                                            {
                                                s.Append(' ');
                                            }
                                        }
                                        else
                                        {
                                            if (b >= 32)
                                            {
                                                char c = (char)b;
                                                s.Append(c);
                                            }
                                            else
                                            {
                                                s.Append('.');
                                            }
                                        }
                                    }
                                    SetDataText(s.ToString());
                                }
                                else
                                {
                                    if (Clipboard.ContainsFileDropList())
                                    {
                                        StringCollection returnList = Clipboard.GetFileDropList();
                                        StringBuilder s = new StringBuilder();
                                        foreach (object o in returnList)
                                        {
                                            s.Append(o.ToString());
                                        }
                                        SetDataText(s.ToString());
                                    }
                                    else
                                    {
                                        SetDataText(dataobject.ToString());
                                    }
                                }
                            }
                        }
                    }
                }
                catch
                {
                    SetDataText(string.Format("No viewer for the {0} format.", Format));
                    return;
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
            else
            {
                SetDataText(string.Format("No data matches the {0} format.", Format));
            }
        }

        public Image GetEnhMetaImageFromClipboard()
        {
            string fileName = Environment.GetEnvironmentVariable("TEMP") + "\\" + Guid.NewGuid().ToString() + ".emf";

            OpenClipboard(IntPtr.Zero);
            IntPtr pointer = GetClipboardData(14);
            IntPtr handle = CopyEnhMetaFile(pointer, fileName);

            Image image;
            using (Metafile metafile = new Metafile(fileName))
            {
                image = new Bitmap(metafile.Width, metafile.Height);
                Graphics g = Graphics.FromImage(image);
                g.DrawImage(metafile, 0, 0, image.Width, image.Height);
                CloseClipboard();
            }

            DeleteEnhMetaFile(handle);
            File.Delete(fileName);
            return image;
        }

        public FrmClipboardViewer()
        {
            InitializeComponent();

            // enable auto refresh by default
            autoRefreshToolStripMenuItem.Checked = true;
            AutoRefresh = true;
        }

        private void RefreshClipboard()
        {
            try
            {
                lbClipFormats.Items.Clear();
                data = Clipboard.GetDataObject();
                string[] Formats = data.GetFormats();

                // check for empty list
                if (Formats.Length == 0)
                {
                    lbClipFormats.Items.Add("The clipboard is emtpy.");
                    rtbClipData.Text = string.Empty;
                    pbClipData.Visible = false;
                    return;
                }

                lbClipFormats.Items.AddRange(Formats);
                lbClipFormats.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Refresh Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FrmClipboardViewer_Shown(object sender, EventArgs e)
        {
            RefreshClipboard();
            _chainedWnd = SetClipboardViewer(this.Handle);
        }

        private void FrmClipboardViewer_FormClosed(object sender, FormClosedEventArgs e)
        {
            ChangeClipboardChain(this.Handle, _chainedWnd);
        }

        private void LbClipFormats_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDisplay();
        }

        private void RefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RefreshClipboard();
        }

        private void OwnerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                uint threadID = GetWindowThreadProcessId(GetClipboardOwner(), out uint processID);
                string processOwnerName = Process.GetProcessById((int)processID).ProcessName;
                MessageBox.Show("Process = " + processOwnerName, "Clipboard Owner", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ClearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.Clear();
            RefreshClipboard();
        }

        private void SaveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string DlgFilterText;

                // get the clipboard contents
                bool isImage = Clipboard.ContainsImage();

                // set dialog filter text
                if (isImage)
                {
                    DlgFilterText = "Bitmap (*.bmp)|*.bmp|PNG (*.png)|*.png|JPEG (*.jpg;*.jpeg)|*.jpg;*.jpeg|All files (*.*)|*.*";
                }
                else
                {
                    DlgFilterText = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                }

                // launch the dialog
                SaveFileDialog fDialog = new SaveFileDialog
                {
                    Title = "Save clipboard format to file.",
                    Filter = DlgFilterText,
                    FilterIndex = 1,
                    RestoreDirectory = true,
                    InitialDirectory = @"%userprofile%"
                };

                // handle the dialog return
                if (fDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get file name.
                    string name = fDialog.FileName;

                    // save out image based on file extension
                    if (name.EndsWith(".jpg") || name.EndsWith(".jpeg"))
                    {
                        Clipboard.GetImage().Save(name, ImageFormat.Jpeg);
                    }
                    else if (name.EndsWith(".bmp"))
                    {
                        Clipboard.GetImage().Save(name, ImageFormat.Bmp);
                    }
                    else if (name.EndsWith(".png"))
                    {
                        Clipboard.GetImage().Save(name, ImageFormat.Png);
                    }
                    else if (name.EndsWith(".emf"))
                    {
                        Clipboard.GetImage().Save(name, ImageFormat.Emf);
                    }
                    else if (name.EndsWith(".gif"))
                    {
                        Clipboard.GetImage().Save(name, ImageFormat.Gif);
                    }
                    else if (name.EndsWith(".wmf"))
                    {
                        Clipboard.GetImage().Save(name, ImageFormat.Wmf);
                    }
                    else if (name.EndsWith(".ico"))
                    {
                        Clipboard.GetImage().Save(name, ImageFormat.Icon);
                    }
                    else
                    {
                        // Write to the file name selected, you can write the text from a TextBox instead of a string literal.
                        rtbClipData.SelectAll();
                        File.WriteAllText(name, rtbClipData.Text);
                        rtbClipData.Select(0, 0);
                    }
                }
                else
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AutoRefreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AutoRefresh = !autoRefreshToolStripMenuItem.Checked;
            autoRefreshToolStripMenuItem.Checked = AutoRefresh;
        }

        private void ShowRichTextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DisplayRichText = !showRichTextToolStripMenuItem.Checked;
            showRichTextToolStripMenuItem.Checked = DisplayRichText;
            UpdateDisplay();
        }

        private void ShowMemoryInHexToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DisplayMemoryInHex = !showMemoryInHexToolStripMenuItem.Checked;
            showMemoryInHexToolStripMenuItem.Checked = DisplayMemoryInHex;
            UpdateDisplay();
        }

        private void ShowPicturesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DisplayPictures = !showPicturesToolStripMenuItem.Checked;
            showPicturesToolStripMenuItem.Checked = DisplayPictures;
            UpdateDisplay();
        }
    }
}