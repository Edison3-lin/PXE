﻿using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;
using System.Xml.Linq;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading;
using CaptainWin.CommonAPI;

namespace Template {
    public class MyTemplate {
        private const string TR = "C:\\TestManager\\TR_Result.json";
        public static void Run()
        {
            string ftpJson = System.IO.File.ReadAllText(TR);
            JObject fjson = JObject.Parse(ftpJson);
            int index = (int)fjson["Reboot"];

            int DllIndex = 0;
            //********* SIT 依序填寫執行的DLL的項目 /Start/
            // DllIndex++; // 1
            // if( DllIndex > index ) {
            //     RecordDllIndex(DllIndex);
            //     CaptainWin.CommonAPI.DoSleep.Sleep(3, 1);
            // }    

            DllIndex++; // 2
            if( DllIndex > index ) {
                RecordDllIndex(DllIndex);
                CaptainWin.CommonAPI.Smode.GetSmode();
            }    

            DllIndex++; // 2
            if( DllIndex > index ) {
                RecordDllIndex(DllIndex);
               CaptainWin.CommonAPI.HDMI.HdmiConnectionStatus();
            }    

            DllIndex++; // 2
            if( DllIndex > index ) {
                RecordDllIndex(DllIndex);
                CaptainWin.CommonAPI.SysInfo.GetWMI("Win32_BIOS");
            }    

            DllIndex++; // 3
            if( DllIndex > index ) {
                RecordDllIndex(DllIndex);
                CaptainWin.CommonAPI.DoReboot.Reboot(5);
            }    

            DllIndex++; // 4
            if( DllIndex > index ) {
                RecordDllIndex(DllIndex);
                CaptainWin.CommonAPI.SysInfo.GetWMI("Win32_SMBIOSMemory");
            }    

            DllIndex++; // 5
            if( DllIndex > index ) {
                RecordDllIndex(DllIndex);
                CaptainWin.CommonAPI.SysInfo.GetWMI("Win32_OperatingSystem");
            }    


            DllIndex++; // 7
            if( DllIndex > index ) {
                RecordDllIndex(DllIndex);
                CaptainWin.CommonAPI.Culture.GetCulture();
            }    

            DllIndex++; // 4
            if( DllIndex > index ) {
                RecordDllIndex(DllIndex);
                CaptainWin.CommonAPI.GetSystemInfo.GetOSVersion();
                CaptainWin.CommonAPI.GetSystemInfo.GetSystemType();
                CaptainWin.CommonAPI.GetSystemInfo.GetProcessorName();
                CaptainWin.CommonAPI.GetSystemInfo.GetRunningTime();
                CaptainWin.CommonAPI.GetSystemInfo.GetPhysicalMemory();
                CaptainWin.CommonAPI.GetSystemInfo.GetCpuId();
                CaptainWin.CommonAPI.GetSystemInfo.GetCPUCount();
                CaptainWin.CommonAPI.GetSystemInfo.GetDiskDevice();
                CaptainWin.CommonAPI.GetSystemInfo.GetDiskSpace();
            }    
            //********* SIT 依序填寫執行的DLL的項目 /End/

            RecordDllIndex(0);
        }

        public static void UpdateResults() {
                CaptainWin.CommonAPI.GetSystemInfo.GetOSVersion();
        }

        public static void Setup() {
                CaptainWin.CommonAPI.GetSystemInfo.GetPhysicalMemory();
        }

        public static void TearDown() {
                CaptainWin.CommonAPI.GetSystemInfo.GetSystemType();
        }

        public static void RecordDllIndex(int DllIndex) {
            try {
                string ftpJson = System.IO.File.ReadAllText(TR);
                JObject fjson = JObject.Parse(ftpJson);
                fjson["Reboot"] = DllIndex;
                string updatedJson = fjson.ToString();
                System.IO.File.WriteAllText(TR, updatedJson);
            }
            catch (Exception ex) {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }   //RecordDllIndex    

    } //Class1
} //namespace
