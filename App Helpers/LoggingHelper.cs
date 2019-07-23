using System;

namespace Office_File_Explorer.App_Helpers
{
    class LoggingHelper
    {
        public static void Log(string logValue)
        {
            Properties.Settings.Default.ErrorLog.Add(DateTime.Now + " : " + logValue);
            Properties.Settings.Default.Save();
        }
    }
}
