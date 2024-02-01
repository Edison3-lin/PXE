/*
* CaptainWin.Common - Common API for test items
* DoSleep.cs - Common test operations for test items
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Edison Lin  <Edison.Lin@quantatw.com>
*/

using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace CaptainWin.CommonAPI {

    /// <summary>
    /// This method sleeps for a specified duration based on the sleep type and count.
    /// </summary>
    public class DoSleep {
        /// <summary>
        /// This method sleeps for a specified duration based on the sleep type and count.
        /// </summary>
        /// <param name="type">The sleep type.</param>
        /// <param name="count">The duration to sleep for, depending on the sleep type.</param>    /// 
        public static void Sleep (int type, int count) {
            string executablePath = @"c:\TestManager\ItemDownload\pwrtest.exe";
            string arguments = string.Format("/sleep /c:{1} /s:{0} /d:90 /p:60", type, count);

            ProcessStartInfo startInfo = new ProcessStartInfo(executablePath) {
                Arguments = arguments,
                WorkingDirectory = @"c:\TestManager\ItemDownload",
                Verb = "runas"
            };

            try {
                Process process = new Process
                {
                    StartInfo = startInfo
                };
                process.Start();
                process.WaitForExit();
            }
            catch (Exception ex) {
                Console.WriteLine("Error: " + ex.Message);
            }
        } //Sleep
    } //DoSleep
}