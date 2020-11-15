using DocumentFormat.OpenXml.Packaging;
using Office_File_Explorer.App_Helpers;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace Office_File_Explorer.Forms
{
    public partial class FrmViewImages : Form
    {
        string appName, fileName;

        public FrmViewImages(string fName, string fType)
        {
            InitializeComponent();

            appName = fType;
            fileName = fName;

            if (appName == StringResources.word)
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, false))
                {
                    foreach (ImagePart ip in document.MainDocumentPart.ImageParts)
                    {
                        LstImages.Items.Add(ip.Uri);
                    }
                }
            }
            else if (appName == StringResources.excel)
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                {
                    foreach (WorksheetPart wp in document.WorkbookPart.WorksheetParts)
                    {
                        foreach (ImagePart ip in wp.DrawingsPart.ImageParts)
                        {
                            LstImages.Items.Add(ip.Uri);
                        }
                    }
                }
            }
            else if (appName == StringResources.powerpoint)
            {
                using (PresentationDocument document = PresentationDocument.Open(fileName, false))
                {
                    foreach (SlidePart sp in document.PresentationPart.SlideParts)
                    {
                        foreach (ImagePart ip in sp.ImageParts)
                        {
                            LstImages.Items.Add(ip.Uri);
                        }
                    }
                }
            }

            if (LstImages.Items.Count > 0)
            {
                LstImages.SelectedIndex = 0;
            }
            else
            {
                MessageBox.Show("Document does not contain any images.", "Images", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Load += (s, e) => Close();
                return;
            }
        }

        private void LstImages_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            if (appName == StringResources.word)
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, false))
                {
                    foreach (ImagePart ip in document.MainDocumentPart.ImageParts)
                    {
                        if (ip.Uri.ToString() == LstImages.SelectedItem.ToString())
                        {
                            DisplayImage(ip);
                        }
                    }
                }
            }
            else if (appName == StringResources.excel)
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                {
                    foreach (WorksheetPart wp in document.WorkbookPart.WorksheetParts)
                    {
                        foreach (ImagePart ip in wp.DrawingsPart.ImageParts)
                        {
                            if (ip.Uri.ToString() == LstImages.SelectedItem.ToString())
                            {
                                DisplayImage(ip);
                            }
                        }
                    }
                }
            }
            else if (appName == StringResources.powerpoint)
            {
                using (PresentationDocument document = PresentationDocument.Open(fileName, false))
                {
                    foreach (SlidePart sp in document.PresentationPart.SlideParts)
                    {
                        foreach (ImagePart ip in sp.ImageParts)
                        {
                            if (ip.Uri.ToString() == LstImages.SelectedItem.ToString())
                            {
                                DisplayImage(ip);
                            }
                        }
                    }
                }
            }
            else
            {
                return;
            }
        }

        public void DisplayImage(ImagePart ip)
        {
            try
            {
                Stream stream = ip.GetStream();

                if (ip.Uri.ToString().EndsWith(".svg"))
                {
                    MessageBox.Show("Format not currently supported.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    pbImage.Image = pbImage.ErrorImage;
                    pbImage.SizeMode = PictureBoxSizeMode.CenterImage;
                }
                else
                {
                    Image imgStream = Image.FromStream(stream);
                    pbImage.Image = imgStream;
                    
                    if (imgStream.Height > pbImage.Size.Height || imgStream.Width > pbImage.Size.Width)
                    {
                        pbImage.SizeMode = PictureBoxSizeMode.Zoom;
                    }
                    else
                    {
                        pbImage.SizeMode = PictureBoxSizeMode.CenterImage;
                    }
                }

                pbImage.Visible = true;
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("ViewImages::UnableToDisplayImage : " + ex.Message);
            }
        }
    }
}
