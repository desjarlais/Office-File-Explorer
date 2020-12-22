using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Windows.Forms;

namespace Office_File_Explorer.Forms
{
    public partial class FrmPrintOrientation : Form
    {
        static string fName;

        public FrmPrintOrientation(string fileName)
        {
            InitializeComponent();
            fName = fileName;
            rdoPortrait.Checked = true;
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (rdoLandscape.Checked)
            {
                Word_Helpers.WordOpenXml.SetPrintOrientation(fName, PageOrientationValues.Landscape);
            }
            else
            {
                Word_Helpers.WordOpenXml.SetPrintOrientation(fName, PageOrientationValues.Portrait);
            }

            Close();
        }
    }
}
