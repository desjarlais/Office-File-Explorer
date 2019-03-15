using System;
using System.Text;
using System.Windows.Forms;

namespace Office_File_Explorer.Forms
{
    public partial class FrmErrorLog : Form
    {
        public FrmErrorLog()
        {
            InitializeComponent();
        }

        private void FrmErrorLog_Load(object sender, EventArgs e)
        {
            foreach (var obj in Properties.Settings.Default.ErrorLog)
            {
                LstErrorLog.Items.Add(obj);
            }
        }

        private void BtnClearLog_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.ErrorLog.Clear();
            Properties.Settings.Default.Save();
            LstErrorLog.Items.Clear();
        }

        private void BtnCopyResults_Click(object sender, EventArgs e)
        {
            try
            {
                if (LstErrorLog.Items.Count <= 0)
                {
                    return;
                }

                StringBuilder buffer = new StringBuilder();
                foreach (object t in LstErrorLog.Items)
                {
                    buffer.Append(t);
                    buffer.Append('\n');
                }

                Clipboard.SetText(buffer.ToString());
            }
            catch (Exception ex)
            {
                //
                MessageBox.Show(ex.Message);
            }
        }
    }
}
