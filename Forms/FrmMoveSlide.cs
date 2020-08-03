using DocumentFormat.OpenXml.Packaging;
using Office_File_Explorer.App_Helpers;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Office_File_Explorer.Forms
{
    public partial class FrmMoveSlide : Form
    {
        public int _SlideCount;
        public PresentationDocument _pDoc;

        public FrmMoveSlide(PresentationDocument pDoc)
        {
            InitializeComponent();
            _pDoc = pDoc;
            _SlideCount = PowerPoint_Helpers.PowerPointOpenXml.CountSlides(_pDoc);

            if (_SlideCount > 1)
            {
                // for ui-only, use non zero based values
                for (int i = 0; i < _SlideCount; i++)
                {
                    cboFrom.Items.Add(i + 1);
                    cboTo.Items.Add(i + 1);
                }

                // since we have items to list, pre-select the first one
                cboFrom.SelectedIndex = 0;
                cboTo.SelectedIndex = 0;
            }
            else
            {
                MessageBox.Show("Not enough slides to move.", "Slide Warning", MessageBoxButtons.OK);
                Close();
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            try
            {
                PowerPoint_Helpers.PowerPointOpenXml.MoveSlide(_pDoc, cboFrom.SelectedIndex, cboTo.SelectedIndex);
                Close();
            }
            catch(Exception ex)
            {
                LoggingHelper.Log(ex.Message);
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
