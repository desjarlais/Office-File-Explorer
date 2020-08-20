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
    public partial class FrmBatchDeleteCustomProps : Form
    {
        public string PropName { get; set; }

        public FrmBatchDeleteCustomProps()
        {
            InitializeComponent();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            PropName = "Cancel";
            Close();
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            PropName = txbPropName.Text;
            Close();
        }
    }
}
