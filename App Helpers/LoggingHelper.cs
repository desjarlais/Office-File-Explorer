using System;

namespace Office_File_Explorer.App_Helpers
{
    class LoggingHelper
    {
        /// <summary>
        /// Generic log function, add data to the app property and save
        /// </summary>
        /// <param name="logValue">string to put in the app property for logging purposes</param>
        public static void Log(string logValue)
        {
            Properties.Settings.Default.ErrorLog.Add(DateTime.Now + " : " + logValue);
            Properties.Settings.Default.Save();
        }
    }
}
