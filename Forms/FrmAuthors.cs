/****************************** Module Header ******************************\
Module Name:  FrmAuthors.cs
Project:      Office File Explorer
Copyright (c) Microsoft Corporation.

List of Authors Form

This source is subject to the following license.
See https://github.com/desjarlais/Office-File-Explorer/blob/master/LICENSE
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;

using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Office_File_Explorer.Forms
{
    public partial class FrmAuthors : Form
    {
        string author = "";

        public FrmAuthors(string filename, List<string> authors)
        {
            InitializeComponent();

            foreach (string s in authors)
            {
                cmbAuthors.Items.Add(s);
            }

            // handle documents with no authors
            if (cmbAuthors.Items.Count == 0)
            {
                cmbAuthors.Items.Add("* No Authors *");
                cmbAuthors.SelectedIndex = 0;
            }

            cmbAuthors.SelectedIndex = 0;
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
