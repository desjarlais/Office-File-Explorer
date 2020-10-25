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

        public static void Clear()
        {
            Properties.Settings.Default.ErrorLog.Clear();
            Properties.Settings.Default.Save();
        }

        public static void LogSystemInformation()
        {
            Properties.Settings.Default.ErrorLog.Add("");
            Properties.Settings.Default.ErrorLog.Add("Operation System: " + Environment.OSVersion);
            Properties.Settings.Default.ErrorLog.Add("Processor Architecture: " + Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE"));
            Properties.Settings.Default.ErrorLog.Add("Processor Model: " + Environment.GetEnvironmentVariable("PROCESSOR_IDENTIFIER"));
            Properties.Settings.Default.ErrorLog.Add("Processor Level: " + Environment.GetEnvironmentVariable("PROCESSOR_LEVEL"));
            Properties.Settings.Default.ErrorLog.Add("ProcessorCount: " + Environment.ProcessorCount);
            Properties.Settings.Default.ErrorLog.Add("Version: " + Environment.Version);
            Properties.Settings.Default.ErrorLog.Add("");
        }
    }
}
