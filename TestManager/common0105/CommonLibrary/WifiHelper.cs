﻿/*
* WifiHelper.cs - Use netsh to manipulate wifi settings
* AddProfile - Add a profile into wifi interface
* Connect - Connect to a AP by ssid
* Enable - Enable wifi interface
* Disable - Disable wifi interface
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Chris Huang   <Chris_Huang@quantatw.com>
*/

using System;

namespace CaptainWin.CommonAPI {
    /// <summary>
    /// Thie class uitilize netsh to enable or disable wifi, connect to SSID and load wifi profile
    /// </summary>
    public class WifiHelper {
        private static string runSync(object command) {
            string result;
            try {

                System.Diagnostics.ProcessStartInfo procStartInfo = new System.Diagnostics.ProcessStartInfo("cmd", "/c " + command);

                procStartInfo.RedirectStandardOutput = true;
                procStartInfo.UseShellExecute = false;
                // Do not create the black window.
                procStartInfo.CreateNoWindow = true;

                // Now we create a process, assign its ProcessStartInfo and start it
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo = procStartInfo;
                proc.Start();

                // Get the output into a string
                result = proc.StandardOutput.ReadToEnd();

            }
            catch (Exception objException) {
                result = "ExecuteCommandSync failed" + objException.Message;
            }
            return result;
        }
        /// <summary>
        /// Add the SSID profile of the AP which DUT wants to connet to.
        /// The SSID profile can be generated by netsh.
        /// </summary>
        /// <param name="fileName">The XML file name of the wifi profile </param>
        /// <returns>output from netsh</returns>
        public static string AddProfile(string fileName) {
            string output;

            output = WifiHelper.runSync("netsh wlan add profile filename=" + fileName);
            return output;

        }
        /// <summary>
        /// Connect to AP by ssid
        /// </summary>
        /// <param name="ssid">The ssid of the AP</param>
        /// <returns>output from netsh</returns>
        public static string Connect(string ssid) {
            string output;

            output = WifiHelper.runSync("netsh wlan connect name=" + ssid + " ssid=" + ssid);
            return output;

        }
        /// <summary>
        /// Enable wifi interface
        /// </summary>
        /// <returns>output from netsh</returns>
        public static string Enable() {
            string output;

            output = WifiHelper.runSync("netsh interface set interface \"Wi-Fi\" enable");
            return output;

        }
        /// <summary>
        /// Disable wifi interface
        /// </summary>
        /// <returns>output from netsh</returns>
        public static string Disable() {
            string output;

            output = WifiHelper.runSync("netsh interface set interface \"Wi-Fi\" disable");
            return output;

        }
    }
}