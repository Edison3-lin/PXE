/*
* CaptainWin.Common - Common API for test items
* Smode.cs - Common test operations for test items
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Edison Lin  <Edison.Lin@quantatw.com>
*/

using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Management;
using Microsoft.Win32;

namespace CaptainWin.CommonAPI {

    /// <summary>
    /// Get WMI function
    /// </summary>
    public class Smode {
        /// <summary>
        /// TitleLog
        /// </summary>
        public static void TitleLog(string content) {
           using (StreamWriter writer = new StreamWriter("c:\\TestManager\\ItemDownload\\GetSmode.log", true))
           {
               writer.Write("\n[[ "+DateTime.Now.ToString()+" ]] -- "+content+" --\n");
           }
        }
        /// <summary>
        /// Log
        /// </summary>
        public static void ProcessLog(string content) {
            try {
                // appand content
                using (StreamWriter writer = new StreamWriter("c:\\TestManager\\ItemDownload\\GetSmode.log", true))
                {
                    writer.Write(content+'\n');
                }
            }
            catch (Exception ex) {
                Console.WriteLine("Error!!! " + ex.Message);
            }
        }      

        /// <summary>
        /// GetSetSmode
        /// </summary>
        public static void GetSmode() {
            TitleLog("GetSmode");
            // *** The first method
            // using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem")) {
            //     foreach (ManagementObject os in searcher.Get()) {
            //         string edition = os["OperatingSystemSKU"].ToString();
            //         ProcessLog(String.Format("OperatingSystemSKU: {0}", edition));
            //         if (edition == "125" || edition == "126" || edition == "27") {
            //             ProcessLog("This OS is in S-mode");
            //         }
            //         else {
            //             ProcessLog("This OS is not in S-mode");
            //         }
            //     }
            // }

            // *** Sencond method
            // try {
            //     using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Store", false))
            //     {
            //         if (key != null)
            //         {
            //             var value = key.GetValue("SystemPaneSuggestionsEnabled");
            //             if (value != null && value is int sModeValue)
            //             {
            //                 ProcessLog("This OS is in S-mode");
            //                 return;
            //             }
            //         }
            //     }
            // }
            // catch (Exception ex) {
            //     Console.WriteLine($"An error occurred: {ex.Message}");
            // }
            // ProcessLog("This OS is not in S-mode");

            // *** Third method
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.Arguments = $"/c systeminfo";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.Start();
            string output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();
            // ProcessLog(output);            
            if(output.Contains("S Mode")) {
                ProcessLog("\nThis system is in S-mode\n");            
            }
            else {
                ProcessLog("\nThis system is not in S-mode\n");            
            };
        }
    }        
}
