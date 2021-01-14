using System;
using System.Windows.Forms;

namespace Office_File_Explorer.Forms
{
    public partial class FrmSearchAndReplace : Form
    {
        public FrmSearchAndReplace()
        {
            InitializeComponent();
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (Owner is FrmMain f)
            {
                f.FindTextProperty = TxtBxFind.Text;
                f.ReplaceTextProperty = TxtBxReplace.Text;
            }
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            if (Owner is FrmMain f)
            {
                f.FindTextProperty = string.Empty;
                f.ReplaceTextProperty = string.Empty;
            }
            Close();
        }
    }
}
