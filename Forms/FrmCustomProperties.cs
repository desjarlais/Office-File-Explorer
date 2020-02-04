using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

using Office_File_Explorer.App_Helpers;

namespace Office_File_Explorer.Forms
{
    public partial class FrmCustomProperties : Form
    {
        string fName, fType;
        List<string> bFiles;
        bool isBatch;

        // single file constructor
        public FrmCustomProperties(string filePath, string fileType)
        {
            InitializeComponent();
            fName = filePath;
            fType = fileType;
            rdoNo.Enabled = false;
            rdoYes.Enabled = false;
            TxtBoxNumber.Enabled = false;
            TxtBoxText.Enabled = false;
            DtDateTime.Enabled = false;
            isBatch = false;
        }

        // multiple file constructor
        public FrmCustomProperties(List<string> files, string fileType)
        {
            InitializeComponent();
            fType = fileType;
            bFiles = files;
            rdoNo.Enabled = false;
            rdoYes.Enabled = false;
            TxtBoxNumber.Enabled = false;
            TxtBoxText.Enabled = false;
            DtDateTime.Enabled = false;
            isBatch = true;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                bool value;
                int num;
                double number;

                switch (CbType.SelectedItem)
                {
                    case "YesNo":       
                        if (rdoNo.Checked)
                        {
                            value = false;
                        }
                        else
                        {
                            value = true;
                        }
                        
                        if (isBatch == true)
                        {
                            foreach (string f in bFiles)
                            {
                                OfficeHelpers.SetCustomProperty(f, TxtName.Text, value, OfficeHelpers.PropertyTypes.YesNo, fType);
                            }
                        }
                        else
                        {
                            OfficeHelpers.SetCustomProperty(fName, TxtName.Text, value, OfficeHelpers.PropertyTypes.YesNo, fType);
                        }
                        
                        break;
                    case "Date":
                        if (isBatch == true)
                        {
                            foreach (string f in bFiles)
                            {
                                OfficeHelpers.SetCustomProperty(f, TxtName.Text, DtDateTime.Value, OfficeHelpers.PropertyTypes.DateTime, fType);
                            }
                        }
                        else
                        {
                            OfficeHelpers.SetCustomProperty(fName, TxtName.Text, DtDateTime.Value, OfficeHelpers.PropertyTypes.DateTime, fType);
                        }
                        
                        break;
                    case "Number":                       
                        if (Int32.TryParse(TxtBoxNumber.Text, out num))
                        {
                            if (isBatch == true)
                            {
                                foreach (string f in bFiles)
                                {
                                    OfficeHelpers.SetCustomProperty(f, TxtName.Text, num, OfficeHelpers.PropertyTypes.NumberInteger, fType);
                                }
                            }
                            else
                            {
                                OfficeHelpers.SetCustomProperty(fName, TxtName.Text, num, OfficeHelpers.PropertyTypes.NumberInteger, fType);
                            }
                            
                        }
                        else if (Double.TryParse(TxtBoxNumber.Text, out number))
                        {
                            if (isBatch == true)
                            {
                                foreach (string f in bFiles)
                                {
                                    OfficeHelpers.SetCustomProperty(f, TxtName.Text, number, OfficeHelpers.PropertyTypes.NumberDouble, fType);
                                }
                            }
                            else
                            {
                                OfficeHelpers.SetCustomProperty(fName, TxtName.Text, number, OfficeHelpers.PropertyTypes.NumberDouble, fType);
                            }
                        }
                        else
                        {
                            // if the value isn't an int or double, just use text format
                            MessageBox.Show("The value entered is not a valid number and will be stored as text.", "Invalid Number", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                            if (isBatch == true)
                            {
                                foreach (string f in bFiles)
                                {
                                    OfficeHelpers.SetCustomProperty(f, TxtName.Text, TxtBoxNumber.Text, OfficeHelpers.PropertyTypes.Text, fType);
                                }
                            }
                            else
                            {
                                OfficeHelpers.SetCustomProperty(fName, TxtName.Text, TxtBoxNumber.Text, OfficeHelpers.PropertyTypes.Text, fType);
                            }
                        }
                        break;
                    default:
                        // Text is default
                        if (isBatch == true)
                        {
                            foreach (string f in bFiles)
                            {
                                OfficeHelpers.SetCustomProperty(f, TxtName.Text, TxtBoxText.Text, OfficeHelpers.PropertyTypes.Text, fType);
                            }
                        }
                        else
                        {
                            OfficeHelpers.SetCustomProperty(fName, TxtName.Text, TxtBoxText.Text, OfficeHelpers.PropertyTypes.Text, fType);
                        }
                        
                        break;
                }

                Close();
            }
            catch (InvalidDataException ide)
            {
                LoggingHelper.Log("SetCustomProperty: Invalid Property Value");
                LoggingHelper.Log(ide.Message);
            }
            catch (Exception ex)
            {
                LoggingHelper.Log("BtnOKCustomProps Error: " + ex.Message);
                Close();
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void CbType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CbType.SelectedItem.ToString() == "Text")
            {
                rdoNo.Enabled = false;
                rdoYes.Enabled = false;
                TxtBoxNumber.Enabled = false;
                TxtBoxText.Enabled = true;
                DtDateTime.Enabled = false;
            }
            else if (CbType.SelectedItem.ToString() == "YesNo")
            {
                rdoNo.Enabled = true;
                rdoYes.Enabled = true;
                TxtBoxNumber.Enabled = false;
                TxtBoxText.Enabled = false;
                DtDateTime.Enabled = false;
            }
            else if (CbType.SelectedItem.ToString() == "Number")
            {
                rdoNo.Enabled = false;
                rdoYes.Enabled = false;
                TxtBoxNumber.Enabled = true;
                TxtBoxText.Enabled = false;
                DtDateTime.Enabled = false;
            }
            else
            {
                rdoNo.Enabled = false;
                rdoYes.Enabled = false;
                TxtBoxNumber.Enabled = false;
                TxtBoxText.Enabled = false;
                DtDateTime.Enabled = true;
            }
        }
    }
}
