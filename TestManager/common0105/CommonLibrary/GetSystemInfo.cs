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
    public class GetSystemInfo {
        /// <summary>
        /// TitleLog
        /// </summary>
        public static void TitleLog(string content) {
           using (StreamWriter writer = new StreamWriter("c:\\TestManager\\ItemDownload\\GetSystemInfo.log", true))
           {
               writer.Write("\n[[ "+DateTime.Now.ToString()+" ]] -- "+content+" --\n");
           }
        }
        /// <summary>
        /// ProcessLog
        /// </summary>
        public static void ProcessLog(string content) {
            try {
                // appand content
                using (StreamWriter writer = new StreamWriter("c:\\TestManager\\ItemDownload\\GetSystemInfo.log", true))
                {
                    writer.Write(content+'\n');
                }

            }
            catch (Exception ex) {
                Console.WriteLine("Error!!! " + ex.Message);
            }
        }      
        /// <summary>
        /// GetOSVersion
        /// </summary>
		public static void GetOSVersion() {
            TitleLog("GetOSVersion");
            try {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("Select * From Win32_OperatingSystem");
                foreach (ManagementObject mo in searcher.Get()) {
                    ProcessLog(mo.Properties["Caption"].Value.ToString());
                    break;
                }
            }
            catch (Exception ex) {
                ProcessLog(new StringBuilder("Error getting operating system version").Append(ex.Message).ToString());
            }
        }
        /// <summary>
        /// GetSystemType
        /// </summary>
		public static void GetSystemType()
        {
            TitleLog("GetSystemType");
            if (Environment.Is64BitOperatingSystem)
            {
                ProcessLog("64bit operating system");
            }
            else
            {
                ProcessLog("32bit operating system");
            }
        }
        /// <summary>
        /// GetCPUCount
        /// </summary>
		public static void GetCPUCount()
        {
            TitleLog("GetCPUCount");
            ProcessLog(Environment.ProcessorCount.ToString());
        }
        /// <summary>
        /// CPUID
        /// </summary>
		public static void GetCpuId()
        {
            TitleLog("GetCpuId");
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("Select * From Win32_Processor");
                foreach (ManagementObject mo in searcher.Get())
                {
                    ProcessLog(mo.Properties["ProcessorId"].Value.ToString());
                    break;
                }
            }
            catch (Exception ex) {
                ProcessLog(new StringBuilder("Failed to get disk name").Append(ex.Message).ToString());
            }
        }
        /// <summary>
        /// GetDiskDevice
        /// </summary>
		public static void GetDiskDevice() {
            TitleLog("GetDiskDevice");
            int index;
            try {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("Select * From Win32_DiskDrive");
                index = 0;
                foreach (ManagementObject mo in searcher.Get()) {
                    index++;
                    // disk = mo.Properties["Caption"].Value.ToString();
                    ProcessLog(index.ToString()+". "+mo.Properties["Caption"].Value.ToString());
                }
            }
            catch (Exception ex) {
                ProcessLog(new StringBuilder("Failed to get disk name").Append(ex.Message).ToString());
            }
        }
        /// <summary>
        /// GetDiskSpace
        /// </summary>
		public static void GetDiskSpace() {
            TitleLog("GetDiskSpace");
            try {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("Select * From Win32_LogicalDisk");
                foreach (ManagementObject mo in searcher.Get()) {
                    ulong totalSpace = Convert.ToUInt64(mo.Properties["Size"].Value) / (1024*1024*1024);
                    ulong freeSpace = Convert.ToUInt64(mo.Properties["FreeSpace"].Value) / (1024*1024*1024);     
                    string Caption = mo.Properties["Caption"].Value.ToString();
                    string VolumeName = mo.Properties["VolumeName"].Value.ToString();
                    string myString = string.Format("{0,-2} <VolumeName>: {1,-10} <Size>: {2,6} GB <Free Space>: {3,6} GB", Caption, VolumeName, totalSpace, freeSpace);
                    ProcessLog(myString);
                }
            }
            catch (Exception ex) {
                ProcessLog(new StringBuilder("Failed to get disk name").Append(ex.Message).ToString());
            }
        }
        /// <summary>
        /// GetProcessorName
        /// </summary>
		public static void GetProcessorName() {
            TitleLog("GetProcessorName");
            try {
                ManagementClass mos = new ManagementClass("Win32_Processor");
                foreach (ManagementObject mo in mos.GetInstances()) {
                    if (mo["Name"] != null) {
                        ProcessLog(mo["Name"].ToString());
                    }
                }
            }
            catch (Exception ex) {
                ProcessLog(new StringBuilder("Failed to get processor name").Append(ex.Message).ToString());
            }
        }

        /// <summary>
        /// GetRunningTime
        /// </summary>
		public static void GetRunningTime() {
            TitleLog("GetRunningTime");
            string result = string.Empty;
            try {
                int uptime = Environment.TickCount & Int32.MaxValue;
                TimeSpan ts = new TimeSpan(Convert.ToInt64(Convert.ToInt64(uptime) * 10000));
                result = new StringBuilder(ts.Days.ToString()).Append(" day ").Append(ts.Hours).Append(":").
                    Append(ts.Minutes).Append(":").Append(ts.Seconds).ToString();
            }
            catch (Exception ex) {
                ProcessLog(new StringBuilder("Failed to obtain boot time: ").Append(ex.Message).ToString());
            }
            ProcessLog(result);
        }
        /// <summary>
        /// GetDiskMediaType
        /// </summary>
		public static string GetDiskMediaType() {
            TitleLog("GetDiskPartition");
            string result="";
            try {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive");
                ManagementObjectCollection queryCollection = searcher.Get();

                foreach (ManagementObject disk in queryCollection)
                {
                    string d = disk["Name"].ToString();
                    if ( d.Contains("PHYSICALDRIVE0") ) {
                        result = disk["MediaType"].ToString();
                    }
                }
            }
            catch (Exception e) {
                Console.WriteLine("Error: " + e.Message);
            }
            return result;
        }
        /// <summary>
        /// GetDiskFormat
        /// </summary>
		public static bool GetDiskFormat() {
            TitleLog("GetDiskFormat");

            DriveInfo[] drives = DriveInfo.GetDrives();
            bool result = true;

            foreach (DriveInfo drive in drives)
            {

                if (drive.IsReady)
                {
                    switch (drive.Name)
                    {

                        case "C:\\":

                            if ( (drive.DriveType.ToString() != "Fixed") || (drive.VolumeLabel.ToString() != "Acer") ||(drive.DriveFormat.ToString() != "NTFS") ) {
                                ProcessLog(drive.Name + " " +  drive.VolumeLabel + " "  + drive.DriveType + " " + drive.DriveFormat);
                                result = false;
                            }
                            break;
                        case "D:\\":
                            if ( (drive.DriveType.ToString() != "Fixed") || (drive.VolumeLabel.ToString() != "Data") || (drive.DriveFormat.ToString() != "NTFS") ) {
                                ProcessLog(drive.Name + " " +  drive.VolumeLabel + " "  + drive.DriveType + " " + drive.DriveFormat);
                                result = false;
                            }
                            break;
                        case "E:\\":
                            if ( (drive.DriveType.ToString() != "Fixed") || (drive.VolumeLabel.ToString() != "Data2") || (drive.DriveFormat.ToString() != "NTFS") ) {
                                ProcessLog(drive.Name + " " +  drive.VolumeLabel + " "  + drive.DriveType + " " + drive.DriveFormat);
                                result = false;
                            }
                            break;
                    }
                }
            }
            return result;
        }
        /// <summary>
        /// GetPhysicalMemory
        /// </summary>
		public static void GetPhysicalMemory() {
            TitleLog("GetPhysicalMemory");
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.Arguments = $"/c systeminfo";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.Start();
            string output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();
            // ProcessLog(output); 
            string[] substrings = output.Split('\n');           
            foreach (string line in substrings)
            {
                if( line.Contains("Total Physical Memory:") | line.Contains("Available Physical Memory:") ) {
                    ProcessLog(line);            
                }
            }
        }
    }
}
