﻿using DocumentFormat.OpenXml.Packaging;
using Office_File_Explorer.App_Helpers;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Xml;

namespace Office_File_Explorer.Forms
{
    public partial class FrmCustomXmlViewer : Form
    {
        static List<CustomXmlPart> cxpList;
        static List<string> nodeNames;
        static string fType, fName;

        public FrmCustomXmlViewer(string fileName, string fileType)
        {
            InitializeComponent();
            fType = fileType;
            fName = fileName;

            nodeNames = new List<string>();

            if (fType == StringResources.word)
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(fName, true))
                {
                    cxpList = document.MainDocumentPart.CustomXmlParts.ToList();
                    
                    foreach (CustomXmlPart cxp in cxpList)
                    {
                        lstCustomXmlFiles.Items.Add(cxp.Uri);
                    }
                }
            }
            else if (fType == StringResources.excel)
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fName, true))
                {
                    cxpList = document.WorkbookPart.CustomXmlParts.ToList();

                    foreach (CustomXmlPart cxp in cxpList)
                    {
                        lstCustomXmlFiles.Items.Add(cxp.Uri);
                    }
                }
            }
            else if (fType == StringResources.powerpoint)
            {
                using (PresentationDocument document = PresentationDocument.Open(fName, true))
                {
                    cxpList = document.PresentationPart.CustomXmlParts.ToList();

                    foreach (CustomXmlPart cxp in cxpList)
                    {
                        lstCustomXmlFiles.Items.Add(cxp.Uri);
                    }
                }
            }
            else
            {
                return;
            }
        }

        private void LstCustomXmlFiles_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            if (fType == StringResources.word)
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(fName, true))
                {
                    cxpList = document.MainDocumentPart.CustomXmlParts.ToList();

                    foreach (CustomXmlPart c in cxpList)
                    {
                        if (c.Uri.ToString() == lstCustomXmlFiles.SelectedItem.ToString())
                        {
                            treeView1.Nodes.Clear();
                            nodeNames.Clear();
                            XmlDocument xDoc = new XmlDocument();
                            xDoc.Load(c.GetStream());
                            PopulateBaseNodes(xDoc);
                        }
                    }
                }
            }
            else if (fType == StringResources.excel)
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fName, true))
                {
                    cxpList = document.WorkbookPart.CustomXmlParts.ToList();

                    foreach (CustomXmlPart c in cxpList)
                    {
                        if (c.Uri.ToString() == lstCustomXmlFiles.SelectedItem.ToString())
                        {
                            treeView1.Nodes.Clear();
                            nodeNames.Clear();
                            XmlDocument xDoc = new XmlDocument();
                            xDoc.Load(c.GetStream());
                            PopulateBaseNodes(xDoc);
                        }
                    }
                }
            }
            else if (fType == StringResources.powerpoint)
            {
                using (PresentationDocument document = PresentationDocument.Open(fName, true))
                {
                    cxpList = document.PresentationPart.CustomXmlParts.ToList();

                    foreach (CustomXmlPart c in cxpList)
                    {
                        if (c.Uri.ToString() == lstCustomXmlFiles.SelectedItem.ToString())
                        {
                            treeView1.Nodes.Clear();
                            nodeNames.Clear();
                            XmlDocument xDoc = new XmlDocument();
                            xDoc.Load(c.GetStream());
                            PopulateBaseNodes(xDoc);
                        }
                    }
                }
            }
            else
            {
                return;
            }
        }

        /// <summary>
        /// populate the treeview
        /// </summary>
        /// <param name="docXml"></param>
        private void PopulateBaseNodes(XmlDocument docXml)
        {
            treeView1.Nodes.Clear();
            treeView1.BeginUpdate();

            treeView1.Nodes.Add(new TreeNode(docXml.DocumentElement.Name));
            TreeNode tNode = (TreeNode)treeView1.Nodes[0];
            AddTreeNode(docXml.DocumentElement, tNode);

            treeView1.EndUpdate();
            treeView1.Refresh();
            treeView1.ExpandAll();
            treeView1.Nodes[0].EnsureVisible();
        }

        private void TreeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e != null)
            {
                toolStripStatusNodePath.Text = "Node Path = " + e.Node.FullPath;
            }
            else
            {
                toolStripStatusNodePath.Text = "Node Path = ";
            }
        }

        /// <summary>
        /// This function is called recursively until all nodes are loaded
        /// </summary>
        /// <param name="xmlNode"></param>
        /// <param name="treeNode"></param>
        private void AddTreeNode(XmlNode xmlNode, TreeNode treeNode)
        {
            XmlNode xNode;
            TreeNode tNode;
            XmlNodeList xNodeList;

            // The current node has children
            if (xmlNode.HasChildNodes)
            {
                // Loop through the child nodes
                xNodeList = xmlNode.ChildNodes;
                for (int x = 0; x <= xNodeList.Count - 1; x++)
                {
                    xNode = xmlNode.ChildNodes[x];
                    
                    if (xNode.Name != null)
                    {
                        nodeNames.Add(xmlNode.Name);
                    }
                    
                    treeNode.Nodes.Add(new TreeNode(xNode.Name));
                    tNode = treeNode.Nodes[x];
                    AddTreeNode(xNode, tNode);
                }
            }
            else
            {
                // No children, so add the outer xml (trimming off whitespace)
                treeNode.Text = xmlNode.OuterXml.Trim();
            }
        }
    }
}
