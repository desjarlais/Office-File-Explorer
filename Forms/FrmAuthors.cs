/****************************** Module Header ******************************\
Module Name:  FrmAuthors.cs
Project:      Office File Explorer
Copyright (c) Microsoft Corporation.

Main window for OFE.

This source is subject to the Microsoft Public License.
See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Windows.Forms;

namespace Office_File_Explorer.Forms
{
    public partial class FrmAuthors : Form
    {
        string author = "";

        public FrmAuthors(string filename, WordprocessingDocument doc)
        {
            InitializeComponent();

            WordprocessingPeoplePart peoplePart = doc.MainDocumentPart.WordprocessingPeoplePart;
            if (peoplePart != null)
            {
                foreach (Person person in peoplePart.People)
                {
                    cmbAuthors.Items.Add(person.Author);
                }

                cmbAuthors.SelectedIndex = 0;
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            try
            {
                author = cmbAuthors.SelectedItem.ToString();

                if (Owner is FrmMain f)
                {
                    f.AuthorProperty = author;
                }

                Close();
            }
            catch (NullReferenceException)
            {
                MessageBox.Show("Please choose author from the dropdown list.", "No Author Selected.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
