using System;
using System.Threading;
using System.Windows.Forms;

namespace Office_File_Explorer
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            // create a named mutex
            Mutex mutex = new Mutex(false, "brandesoft office file explorer", out bool noInstanceCurrently);

            // let the user know if we already exist.
            if (noInstanceCurrently == false)
            {
                MessageBox.Show("The application is already running.", "Application Launch Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FrmMain());
        }
    }
}
