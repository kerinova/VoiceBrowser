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
using Microsoft.Speech.Recognition;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Linq;
using System.IO;

namespace VoiceBrowser
{
    class VoiceListener
    {
        private string[] speechCommands; //create an array with 8 indices for the voice commands.

        SpeechRecognitionEngine speechRecogEngine;
        public VoiceListener()
        {
            //initialize the speech recognition engine
            speechRecogEngine = new SpeechRecognitionEngine(new CultureInfo("en-us"));
            speechRecogEngine.SetInputToDefaultAudioDevice();
            speechRecogEngine.SpeechRecognized += speechRecogEngine_SpeechRecognized;
            Grammar grammar = InitializeGrammar();
            speechRecogEngine.LoadGrammarAsync(grammar);
        }

        /// <summary>
        /// Initialize the list of commands pertaining to browsing.
        /// </summary>
        /// <returns></returns>
        private Grammar InitializeGrammar()
        {
            speechCommands = LoadGrammar();
            Choices commands = new Choices();
            commands.Add(speechCommands.Take(7).ToArray()); //adds all commands except for 'resume'
            GrammarBuilder grammarBuilder = new GrammarBuilder();
            grammarBuilder.Append(commands);
            Grammar grammar = new Grammar(grammarBuilder);
            return grammar;
        }

        private string[] LoadGrammar()
        {
            try
            {
                return File.ReadAllLines("./commands.txt");
            }
            catch(IOException e)
            {
                Console.WriteLine("Failed to load custom commands: {0}", e.Message);
            }
            return new string[] { "scroll", "scroll up", "next", "back", "close tab", "pause", "exit", "resume" }; //if loading fails, load default commands
        }
        /// <summary>
        /// Initialize a grammar containing resume and exit, for when recognition is paused.
        /// </summary>
        /// <returns></returns>
        private Grammar InitializeMenuGrammar()
        {
            Choices commands = new Choices();
            commands.Add(speechCommands.Skip(6).ToArray()); //adds only 'exit' and 'resume' for the menu commands
            GrammarBuilder grammarBuilder = new GrammarBuilder();
            grammarBuilder.Append(commands);
            Grammar grammar = new Grammar(grammarBuilder);
            return grammar;
        }

        /// <summary>
        /// Begin recognition
        /// </summary>
        public void Recognize()
        {
            speechRecogEngine.RecognizeAsync(RecognizeMode.Multiple);
        }

        /// <summary>
        /// Parse the recognized voice command.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void speechRecogEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string txt = e.Result.Text;
            float confidence = e.Result.Confidence;
            Console.WriteLine("\nRecognized: {0}, Confidence: {1}", txt, confidence.ToString());
            if(confidence < 0.70 )
            {
                return;
            }
            if(txt.IndexOf(speechCommands[1]) >= 0) //scroll up
            {
                ScrollUp();
            }
            else if (txt.IndexOf(speechCommands[0]) >= 0) //scroll
            {
                ScrollDown();
            }
            else if (txt.IndexOf(speechCommands[4]) >= 0) //close tab
            {
                CloseTab();
            }
            else if (txt.IndexOf(speechCommands[2]) >= 0) //next tab
            {
                NextTab();
            }
            else if (txt.IndexOf(speechCommands[3]) >= 0) //previous tab
            {
                PrevTab();
            }
            else if (txt.IndexOf(speechCommands[5]) >= 0) //switch to the menu grammar set
            {
                speechRecogEngine.UnloadAllGrammars();
                speechRecogEngine.LoadGrammarAsync(InitializeMenuGrammar());
            }
            else if (txt.IndexOf(speechCommands[6]) >= 0) //switch to the normal grammar set
            {
                speechRecogEngine.UnloadAllGrammars();
                speechRecogEngine.LoadGrammarAsync(InitializeGrammar());
            }
            else if (txt.IndexOf(speechCommands[7]) >= 0) //exit
            {
                System.Environment.Exit(0); //clean exit the application
            }
        }

        #region SendKeys
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint numberOfInputs, INPUT[] inputs, int sizeOfInputStructure);

        /// <summary>
        /// http://msdn.microsoft.com/en-us/library/windows/desktop/ms646270(v=vs.85).aspx
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct INPUT
        {
            public uint Type;
            public MOUSEKEYBDHARDWAREINPUT Data;
        }

        /// <summary>
        /// http://social.msdn.microsoft.com/Forums/en/csharplanguage/thread/f0e82d6e-4999-4d22-b3d3-32b25f61fb2a
        /// </summary>
        [StructLayout(LayoutKind.Explicit)]
        internal struct MOUSEKEYBDHARDWAREINPUT
        {
            [FieldOffset(0)]
            public HARDWAREINPUT Hardware;
            [FieldOffset(0)]
            public KEYBDINPUT Keyboard;
            [FieldOffset(0)]
            public MOUSEINPUT Mouse;
        }

        /// <summary>
        /// http://msdn.microsoft.com/en-us/library/windows/desktop/ms646310(v=vs.85).aspx
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct HARDWAREINPUT
        {
            public uint Msg;
            public ushort ParamL;
            public ushort ParamH;
        }

        /// <summary>
        /// http://msdn.microsoft.com/en-us/library/windows/desktop/ms646310(v=vs.85).aspx
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct KEYBDINPUT
        {
            public ushort Vk;
            public ushort Scan;
            public uint Flags;
            public uint Time;
            public IntPtr ExtraInfo;
        }

        /// <summary>
        /// http://social.msdn.microsoft.com/forums/en-US/netfxbcl/thread/2abc6be8-c593-4686-93d2-89785232dacd
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct MOUSEINPUT
        {
            public int X;
            public int Y;
            public uint MouseData;
            public uint Flags;
            public uint Time;
            public IntPtr ExtraInfo;
        }
        #endregion
        private void ScrollUp()
        {
            SendKey(0x21); //page up key
        }

        private void NextTab()
        {
            SendKeyDown(0x11); //down ctrl
            SendKey(0x09); //press tab
            SendKeyUp(0x11); //up ctrl
        }
        private void PrevTab()
        {
            SendKeyDown(0x11); //down ctrl
            SendKeyDown(0x10); //down shift
            SendKey(0x09); //press tab
            SendKeyUp(0x11); //up ctrl
            SendKeyUp(0x10); //up hift
        }
        private void CloseTab()
        {
            SendKeyDown(0x11); //down ctrl
            SendKey(0x57); //press w
            SendKeyUp(0x11); //up ctrl
        }
        private void ScrollDown()
        {
            SendKey(0x22); //page down key
        }
        private void SendKey(int keyCode)
        {
            //Process[] procIE = Process.GetProcessesByName("iexplorer");
            //foreach (Process ie in procIE)
            //{
                //if (ie.MainWindowHandle != IntPtr.Zero)
                //{
                    // Set focus on the window so that the key input can be received.
                    //SetForegroundWindow(ie.MainWindowHandle);
            INPUT[] inputs;
            INPUT ip = new INPUT { Type = 1 };
            ip.Data.Keyboard = new KEYBDINPUT();
            ip.Data.Keyboard.Vk = (ushort)keyCode;
            ip.Data.Keyboard.Scan = 0;
            ip.Data.Keyboard.Flags = 0;
            ip.Data.Keyboard.Time = 0;
            ip.Data.Keyboard.ExtraInfo = IntPtr.Zero;

            inputs = new INPUT[] { ip };

            SendInput(1, inputs, Marshal.SizeOf(typeof(INPUT)));
        }
        /// <summary>
        /// Send a key down and hold it down until sendkeyup method is called
        /// </summary>
        /// <param name="keyCode"></param>
        public static void SendKeyDown(int keyCode)
        {
            INPUT input = new INPUT
            {
                Type = 1
            };
            input.Data.Keyboard = new KEYBDINPUT();
            input.Data.Keyboard.Vk = (ushort)keyCode;
            input.Data.Keyboard.Scan = 0;
            input.Data.Keyboard.Flags = 0;
            input.Data.Keyboard.Time = 0;
            input.Data.Keyboard.ExtraInfo = IntPtr.Zero;
            INPUT[] inputs = new INPUT[] { input };
            SendInput(1, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        /// <summary>
        /// Release a key that is being hold down
        /// </summary>
        /// <param name="keyCode"></param>
        public static void SendKeyUp(int keyCode)
        {
            INPUT input = new INPUT
            {
                Type = 1
            };
            input.Data.Keyboard = new KEYBDINPUT();
            input.Data.Keyboard.Vk = (ushort)keyCode;
            input.Data.Keyboard.Scan = 0;
            input.Data.Keyboard.Flags = 2;
            input.Data.Keyboard.Time = 0;
            input.Data.Keyboard.ExtraInfo = IntPtr.Zero;
            INPUT[] inputs = new INPUT[] { input };
            SendInput(1, inputs, Marshal.SizeOf(typeof(INPUT)));

        }
    }
}
