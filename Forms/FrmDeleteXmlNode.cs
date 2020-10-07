using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Office_File_Explorer.Forms
{
    public partial class FrmDeleteXmlNode : Form
    {
        public FrmDeleteXmlNode(List<string> nodeList)
        {
            InitializeComponent();
            foreach (string s in nodeList)
            {
                cboNodes.Items.Add(s);
            }
        }

        private void BtnDeleteNode_Click(object sender, EventArgs e)
        {

        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
