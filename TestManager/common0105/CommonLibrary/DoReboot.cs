/*
* CaptainWin.Common - Common API for test items
* DoReboot.cs - Common test operations for test items
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
    /// Reboot the system
    /// </summary>
    public class DoReboot {
        /// <summary>
        /// Reboot the system
        /// </summary>
        /// <param name="sec">Seconds to wait before reboot</param>
        public static void Reboot (int sec) {
            string exeFilePath = "shutdown";
            ProcessStartInfo startInfo = new ProcessStartInfo(exeFilePath);
            startInfo.WorkingDirectory = ".\\";
            startInfo.Arguments = $"/r /t {sec}";
            Process process = new Process();
            process.StartInfo = startInfo;
            process.Start();
            process.WaitForExit();
            while (true){}
        }
    }
}
