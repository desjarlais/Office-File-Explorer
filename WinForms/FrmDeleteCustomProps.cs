using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using Office_File_Explorer.App_Helpers;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Office_File_Explorer.Forms
{
    public partial class FrmDeleteCustomProps : Form
    {
        CustomFilePropertiesPart part;
        public bool PartModified { get; set; }

        public FrmDeleteCustomProps(CustomFilePropertiesPart cfp)
        {
            InitializeComponent();
            PartModified = false;
            part = cfp;
            UpdateList();
        }

        public void UpdateList()
        {
            if (part == null)
            {
                lbProps.Items.Add(StringResources.noCustomDocProps);
                return;
            }

            int count = 0;

            foreach (var v in CfpList(part))
            {
                count++;
                lbProps.Items.Add(count + StringResources.period + v);
            }

            lbProps.SelectedIndex = 0;
        }

        public List<string> CfpList(CustomFilePropertiesPart part)
        {
            List<string> val = new List<string>();
            foreach (CustomDocumentProperty cdp in part.RootElement)
            {
                val.Add(cdp.Name);
            }
            return val;
        }

        private void BtnDeleteProp_Click(object sender, System.EventArgs e)
        {
            string[] valToDelete = lbProps.SelectedItem.ToString().Split('.');
            string val = valToDelete[1].TrimStart();
            foreach (CustomDocumentProperty cdp in part.RootElement)
            {
                if (val == cdp.Name)
                {
                    cdp.Remove();
                    PartModified = true;
                    lbProps.Items.RemoveAt(lbProps.SelectedIndex);
                    lbProps.Items.Clear();
                    UpdateList();
                    lbProps.SelectedIndex = 0;
                }
            }
        }

        private void BtnOK_Click(object sender, System.EventArgs e)
        {
            Close();
        }
    }
}
