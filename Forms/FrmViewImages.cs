using DocumentFormat.OpenXml.Packaging;
using Office_File_Explorer.App_Helpers;
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
            else
            {
                return;
            }

            if (LstImages.Items.Count > 0)
            {
                LstImages.SelectedIndex = 0;
            }
            else
            {
                MessageBox.Show("Document does not contain any images.", "Images", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
            Stream stream = ip.GetStream();
            pbImage.Image = Image.FromStream(stream);
            pbImage.SizeMode = PictureBoxSizeMode.StretchImage;
            pbImage.Visible = true;
        }
    }
}
