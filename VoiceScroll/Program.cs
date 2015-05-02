#region License
/*
    Copyright (C) 2015 Kerinova Studios

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
#endregion
using System;
using System.Reflection;
using System.Threading;

namespace VoiceScroll
{
    class Program
    {
        static bool done = false;
        static void Main(string[] args)
        {
            Console.WriteLine("Voice Listener [Version 0.0.3.25]");
            Console.WriteLine("Copyright \u00a9 2015-{0} Kerinova Studios. All rights reserved.", DateTime.Now.Year);
            Console.WriteLine("GPL v3 License");
            try
            {
                Recognize();
            }
            catch (System.IO.FileNotFoundException e) //if the speech runtime is not installed, install it and restart.
            {
                InstallDependencies();
            }
        }
        private static void InstallDependencies()
        {
            DependencyInstaller.Install("SpeechPlatformRuntime.msi");
            DependencyInstaller.Install("MSSpeech_SR_en-US_TELE.msi");
        }
        private static void Recognize()
        {
            VoiceListener voiceListener = new VoiceListener();
            voiceListener.Recognize();
            while (done == false)
            {
                ;
            }
            Console.WriteLine("Say 'Exit' to close");
        }
    }
}
