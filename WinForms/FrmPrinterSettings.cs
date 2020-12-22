using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing.Printing;
using static System.Drawing.Printing.PrinterSettings;
using System;
using System.Text;
using System.Printing;

namespace Office_File_Explorer.Forms
{
    public partial class FrmPrinterSettings : Form
    {
        public List<string> Printers = new List<string>();

        public FrmPrinterSettings()
        {
            InitializeComponent();
            GetListOfPrinters();
            PopulatePrinterList();
        }

        public void PopulatePrinterList()
        {
            // clear the ui list and add each printer name
            CboPrinters.Items.Clear();

            foreach (object o in Printers)
            {
                CboPrinters.Items.Add(o);
            }

            // set the first printer as the combo box default
            CboPrinters.SelectedIndex = 0;
        }

        public void GetListOfPrinters()
        {
            // clear the existing list and re-populate it
            Printers.Clear();

            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                Printers.Add(printer);
            }
        }

        private void BtnCopy_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (LstDisplay.Items.Count <= 0)
                {
                    return;
                }

                StringBuilder buffer = new StringBuilder();
                foreach (object t in LstDisplay.Items)
                {
                    buffer.Append(t);
                    buffer.Append('\n');
                }

                Clipboard.SetText(buffer.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CboPrinters_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            try
            {
                LstDisplay.Items.Clear();
                PrintServer psrv;
                PrintQueue pq;
                PrinterResolutionCollection prc;
                PaperSizeCollection pszc;
                PaperSourceCollection psc;
                PageSettings pgs;

                PrinterSettings ps = new PrinterSettings
                {
                    PrinterName = CboPrinters.SelectedItem.ToString()
                };

                prc = ps.PrinterResolutions;
                psc = ps.PaperSources;
                pszc = ps.PaperSizes;
                pgs = ps.DefaultPageSettings;

                // display single value properties
                LstDisplay.Items.Add("Printer Name: " + ps.PrinterName);
                LstDisplay.Items.Add("Is Default Printer: " + ps.IsDefaultPrinter);
                LstDisplay.Items.Add("Print Range: " + ps.PrintRange);
                LstDisplay.Items.Add("PrintToFile: " + ps.PrintToFile);
                LstDisplay.Items.Add("Supports Color: " + ps.SupportsColor);
                LstDisplay.Items.Add("Duplex: " + ps.Duplex);
                LstDisplay.Items.Add("Can Duplex: " + ps.CanDuplex);
                LstDisplay.Items.Add("Collate: " + ps.Collate);
                LstDisplay.Items.Add("Copies: " + ps.Copies);
                LstDisplay.Items.Add("Maximum Copies: " + ps.MaximumCopies);
                LstDisplay.Items.Add("Minimum Page: " + ps.MinimumPage);
                LstDisplay.Items.Add("Maximum Page: " + ps.MaximumPage);
                LstDisplay.Items.Add("");

                // display setting collections
                LstDisplay.Items.Add("Printer Resolutions:");
                if (prc.Count > 0)
                {
                    foreach (object o in prc)
                    {
                        LstDisplay.Items.Add("   " + o);
                    }
                }
                else
                {
                    LstDisplay.Items.Add("   No Printer Resolutions.");
                }

                LstDisplay.Items.Add("");
                LstDisplay.Items.Add("Paper Sources: ");
                if (psc.Count > 0)
                {
                    foreach (object o in psc)
                    {
                        LstDisplay.Items.Add("   " + o);
                    }
                }
                else
                {
                    LstDisplay.Items.Add("   No Paper Sources.");
                }

                LstDisplay.Items.Add("");
                LstDisplay.Items.Add("Paper Sizes: ");
                if (pszc.Count > 0)
                {
                    foreach (object o in pszc)
                    {
                        LstDisplay.Items.Add("   " + o);
                    }
                }
                else
                {
                    LstDisplay.Items.Add("   No Paper Sizes.");
                }

                LstDisplay.Items.Add("");
                LstDisplay.Items.Add("Default Page Settings: ");
                if (pgs.Landscape == true)
                {
                    LstDisplay.Items.Add("Orientation = Landscape");
                }
                else
                {
                    LstDisplay.Items.Add("Orientation = Portrait");
                }
                LstDisplay.Items.Add("Color = " + pgs.Color);
                LstDisplay.Items.Add("Printable Area = " + pgs.PrintableArea);
                LstDisplay.Items.Add("Margins = " + pgs.Margins);
                LstDisplay.Items.Add("Paper Source = " + pgs.PaperSource);
                LstDisplay.Items.Add("Page Bounds = " + pgs.Bounds);

                // display  print queues
                LstDisplay.Items.Add("");
                LstDisplay.Items.Add("Print Queue:");

                // check if network printer and adjust print server
                if (CboPrinters.Text.StartsWith("\\\\"))
                {
                    string[] path;
                    path = CboPrinters.Text.Split('\\');
                    psrv = new PrintServer("\\\\" + path[2]);
                    pq = psrv.GetPrintQueue(path[3]);
                }
                else
                {
                    psrv = new PrintServer();
                    pq = psrv.GetPrintQueue(CboPrinters.Text);
                }

                int pjCount = 1;

                if (pq == null || pq.NumberOfJobs == 0)
                {
                    LstDisplay.Items.Add("   Empty");
                }
                else
                {
                    // loop the print job collection
                    foreach (PrintSystemJobInfo pjsi in pq.GetPrintJobInfoCollection())
                    {
                        LstDisplay.Items.Add(pjCount + ". Print Job Name: " + pjsi.Name);
                        LstDisplay.Items.Add("   Job Size: " + pjsi.JobSize);
                        LstDisplay.Items.Add("   Job Status: " + pjsi.JobStatus);
                        LstDisplay.Items.Add("   Position In PrintQueue: " + pjsi.PositionInPrintQueue);
                        LstDisplay.Items.Add("   Number of pages: " + pjsi.NumberOfPages);
                        LstDisplay.Items.Add("   Number of pages printed: " + pjsi.NumberOfPagesPrinted);
                        LstDisplay.Items.Add("   Submitter: " + pjsi.Submitter);
                        LstDisplay.Items.Add("   StartTimeOfDay: " + pjsi.StartTimeOfDay);
                        LstDisplay.Items.Add("   TimeJobSubmitted: " + pjsi.TimeJobSubmitted);
                        LstDisplay.Items.Add("   TimeSinceStartedPrinting: " + pjsi.TimeSinceStartedPrinting);
                        LstDisplay.Items.Add("   UntilTimeOfDay: " + pjsi.UntilTimeOfDay);
                        LstDisplay.Items.Add("   IsBlocked: " + pjsi.IsBlocked);
                        LstDisplay.Items.Add("   IsCompleted: " + pjsi.IsCompleted);
                        LstDisplay.Items.Add("   IsDeleted: " + pjsi.IsDeleted);
                        LstDisplay.Items.Add("   IsDeleting: " + pjsi.IsDeleting);
                        LstDisplay.Items.Add("   IsInError: " + pjsi.IsInError);
                        LstDisplay.Items.Add("   IsOffline: " + pjsi.IsOffline);
                        LstDisplay.Items.Add("   IsPaperOut: " + pjsi.IsPaperOut);
                        LstDisplay.Items.Add("   IsPaused: " + pjsi.IsPaused);
                        LstDisplay.Items.Add("   IsPrinted: " + pjsi.IsPrinted);
                        LstDisplay.Items.Add("   IsPrinting: " + pjsi.IsPrinting);
                        LstDisplay.Items.Add("   IsRestarted: " + pjsi.IsRestarted);
                        LstDisplay.Items.Add("   IsRetained: " + pjsi.IsRetained);
                        LstDisplay.Items.Add("   IsSpooling: " + pjsi.IsSpooling);
                        LstDisplay.Items.Add("   IsSpooling: " + pjsi.IsUserInterventionRequired);
                        pjCount++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                LstDisplay.Items.Add("Error retrieving print settings");
            }
        }
    }
}
