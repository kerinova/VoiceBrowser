using System;
using System.Diagnostics;

namespace VoiceScroll
{
    class DependencyInstaller
    {
        public static void Install(string msiPath)
        {
            Process process = new Process();
            process.StartInfo.FileName = "msiexec.exe";
            process.StartInfo.Arguments = String.Format(" /qb /i \"{0}\" ALLUSERS=1", msiPath);
            process.Start();
            process.WaitForExit();
        }
    }
}
