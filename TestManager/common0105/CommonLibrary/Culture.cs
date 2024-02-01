/*
* CaptainWin.Common - Common API for test items
* Culture.cs - Common test operations for test items
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
using System.Globalization;

namespace CaptainWin.CommonAPI {
    public class Culture {
        /// <summary>
        /// TitleLog
        /// </summary>
        public static void TitleLog(string content) {
           using (StreamWriter writer = new StreamWriter("c:\\TestManager\\ItemDownload\\GetCulture.log", true))
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
                using (StreamWriter writer = new StreamWriter("c:\\TestManager\\ItemDownload\\GetCulture.log", true))
                {
                    writer.Write(content+'\n');
                }

            }
            catch (Exception ex) {
                Console.WriteLine("Error!!! " + ex.Message);
            }
        }      
        /// <summary>
        /// GetCulture UI
        /// </summary>
        public static void GetCulture() {
            TitleLog("GetCulture");

            CultureInfo uiCulture = CultureInfo.InstalledUICulture;
            if (uiCulture.Equals(CultureInfo.InvariantCulture)) {
                ProcessLog("OS support single load language");
            }
            else {
                ProcessLog($"OS support Multi-load language, UI is {uiCulture.Name}");
            }            
            ProcessLog("\nInstalled UI Culture:");
            ProcessLog("Name: " + uiCulture.Name);
            ProcessLog("DisplayName: " + uiCulture.DisplayName);
            ProcessLog("EnglishName: " + uiCulture.EnglishName);
            ProcessLog("TwoLetterISOLanguageName: " + uiCulture.TwoLetterISOLanguageName);
            ProcessLog("ThreeLetterISOLanguageName: " + uiCulture.ThreeLetterISOLanguageName);
        }
    }
}
