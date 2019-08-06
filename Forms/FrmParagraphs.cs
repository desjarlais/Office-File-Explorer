using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;

namespace Office_File_Explorer.Forms
{
    public partial class FrmParagraphs : Form
    {
        string filePath;

        public FrmParagraphs(string file)
        {
            InitializeComponent();
            filePath = file;
            PopulateParagraphComboBox();
        }

        public void PopulateParagraphComboBox()
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
        }

        public void ListParagraphs()
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
                        }
                    }
                }
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

        private void FrmParagraphs_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void CbParagraphs_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListParagraphs();
        }
    }
}
