using System;
using System.Diagnostics;

namespace VoiceBrowser
{
    class DependencyInstaller
    {
        /// <summary>
        /// Installs a given .MSI file, and waits for the installation to complete
        /// </summary>
        /// <param name="msiPath"></param>
        public static void Install(string msiPath)
        {
            Process process = new Process();
            process.StartInfo.FileName = "msiexec.exe";
            process.StartInfo.Arguments = String.Format(" /qb /i \"{0}\" ALLUSERS=1", msiPath); //(almost) UI-less install
            process.Start();
            process.WaitForExit();
        }
    }
}
