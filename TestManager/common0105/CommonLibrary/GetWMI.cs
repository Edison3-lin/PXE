/*
* CaptainWin.Common - Common API for test items
* GetWMI.cs - Common test operations for test items
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

namespace CaptainWin.CommonAPI {

    /// <summary>
    /// Get WMI function
    /// </summary>
    public class SysInfo {
        /// <summary>
        /// Log
        /// </summary>
        public static void ProcessLog(string content) {
            try {
                // appand content
                using (StreamWriter writer = new StreamWriter("c:\\TestManager\\ItemDownload\\GetWMI.log", true))
                {
                    writer.Write(content+'\n');
                }

            }
            catch (Exception ex) {
                Console.WriteLine("Error!!! " + ex.Message);
            }
        }      

        /// <summary>
        /// Get WMI function
        /// </summary>
        public static void GetWMI(string className) {
           using (StreamWriter writer = new StreamWriter("c:\\TestManager\\ItemDownload\\GetWMI.log", true))
           {
               writer.Write("\n[[ "+DateTime.Now.ToString()+" ]] WMI Class - "+className+'\n');
           }

            ManagementClass mgmtClass = new ManagementClass(className);

            try {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher($"SELECT * FROM {className}");            
                ManagementObjectCollection collection = searcher.Get();

                // Console.WriteLine($"Class Name: {mgmtClass["__CLASS"]}");
                // Console.WriteLine($"Description: {mgmtClass["__CLASS"]}");
                // Console.WriteLine("Properties:");
                // ProcessLog($"Class Name: {mgmtClass["__CLASS"]}");
                // ProcessLog($"Description: {mgmtClass["__CLASS"]}");
                // ProcessLog("Properties:");

                foreach (ManagementObject obj in collection) {
                    foreach (PropertyData prop in mgmtClass.Properties) {
                        if (obj[prop.Name] != null) {
                            if (obj[prop.Name].GetType() == typeof(string[])) {
                                foreach (string item in (string[])obj[prop.Name]) {
                                    // Console.WriteLine($"{prop.Name}: {item}");
                                    ProcessLog($"{prop.Name}: {item}");
                                }
                            }
                            else if (obj[prop.Name].GetType() == typeof(UInt16[])) {
                                foreach (UInt16 item in (UInt16[])obj[prop.Name]) {
                                    // Console.WriteLine($"{prop.Name}: {item.ToString()}");
                                    ProcessLog($"{prop.Name}: {item.ToString()}");
                                }
                            } 
                            else {   
                                string propertyValue = obj[prop.Name].ToString();
                                // Console.WriteLine($"{prop.Name}: {propertyValue}");
                                ProcessLog($"{prop.Name}: {propertyValue}");
                            }    
                        }
                        else {
                            // Console.WriteLine($"{prop.Name}: Not available");
                            ProcessLog($"{prop.Name}: Not available");
                        }
                    }      
                    ProcessLog("-- Next Device --\n");
                }
            }
            catch (Exception ex) {
                Console.WriteLine($"Error: {ex.Message}");
                ProcessLog($"Error: {ex.Message}");
            }
        }
        /// <summary>
        /// Get WMI function
        /// </summary>
        public static string GetWMI(string className, string propName) {

            ManagementClass mgmtClass = new ManagementClass(className);

            try {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher($"SELECT * FROM {className}");            
                ManagementObjectCollection collection = searcher.Get();
                foreach (ManagementObject obj in collection) {
                    foreach (PropertyData prop in mgmtClass.Properties) {
                        if (obj[prop.Name] != null) {
                            if((prop.Name).ToString() == propName) {
                                if (obj[prop.Name].GetType() == typeof(string[])) {
                                    foreach (string item in (string[])obj[prop.Name]) {
                                        return item;
                                    }
                                }
                                else if (obj[prop.Name].GetType() == typeof(UInt16[])) {
                                    foreach (UInt16 item in (UInt16[])obj[prop.Name]) {
                                        return item.ToString();
                                    }
                                } 
                                else {   
                                    string propertyValue = obj[prop.Name].ToString();
                                    return propertyValue;
                                }    
                            }      
                        }
                    }      
                }
            }
            catch (Exception ex) {
                Console.WriteLine($"Error: {ex.Message}");
            }
            return false.ToString();
        }

        public static void GetDeviceManager() {
            var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity");
            
            int i = 0;
            foreach (var device in searcher.Get())
            {
                i++;
                ProcessLog($"Name: {device["Name"]}");
                ProcessLog($"Description: {device["Description"]}");
                ProcessLog($"Status: {device["Status"]}");
                ProcessLog($"DeviceID: {device["DeviceID"]}");
                ProcessLog($"PNPDeviceID: {device["PNPDeviceID"]}");
                ProcessLog("-----------------------------------------------------");
            }

            ProcessLog($"===> {i} devices in total");     
        }

    }        
}
