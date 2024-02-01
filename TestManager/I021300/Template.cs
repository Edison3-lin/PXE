using System;
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
using Microsoft.Win32;

using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Template {
    public class MyTemplate {
        private const string TR = "C:\\TestManager\\TR_Result.json";

        public static void TestResult(string TestResult) {
            try {
                string ftpJson = System.IO.File.ReadAllText(TR);
                JObject fjson = JObject.Parse(ftpJson);
                fjson["TestResult"] = TestResult;
                string updatedJson = fjson.ToString();
                System.IO.File.WriteAllText(TR, updatedJson);
            }
            catch (Exception ex) {
                Console.WriteLine($"sdfsadfsdfsd An error occurred: {ex.Message}");
            }
        }

        public static void readExcel(string inName, string Sheet) {
            if (!File.Exists(inName)) {
                Console.WriteLine($"Can't find {inName}");
                return;
            }    

            var app = new Excel.Application();
            var wbk = app.Workbooks.Add(inName);
            //app.Visible = true;
            int index;    
            for (index = 1; index <= (wbk.Sheets.Count); index++) {
                // Console.WriteLine((wbk.Sheets[i]).Name);
                if( (wbk.Sheets[index]).Name == Sheet ) {
                    break;
                }
                if( index == (wbk.Sheets.Count) ) {
                    Console.WriteLine($"Can't find sheet name: {Sheet}");
                    return;
                }
            }    

            var sh = wbk.Sheets[index];
            sh.Activate();
            var usedRange = sh.UsedRange.CurrentRegion;

            // Console.WriteLine("Open -> " + sh.Name);
            // Console.WriteLine($"Row: {usedRange.Rows.Count}, Column: {usedRange.Columns.Count}");

            for (int i = 0; i < usedRange.Rows.Count; i++)
            {
                for (int j = 0; j < usedRange.Columns.Count; j++)
                    Console.Write($"{sh.Cells[i + 1, j + 1].Text} ");
                Console.Write("\n");
            }
            wbk.Close();
            app.Quit();
        }
        public static string readExcel(string inName, string Sheet, int[] Cell) {

            if (!File.Exists(inName)) {
                Console.WriteLine($"Can't find {inName}");
                return "Can't find";
            }    

            var app = new Excel.Application();
            var wbk = app.Workbooks.Add(inName);
            //app.Visible = true;
            int index;    
            for (index = 1; index <= (wbk.Sheets.Count); index++) {
                // Console.WriteLine((wbk.Sheets[i]).Name);
                if( (wbk.Sheets[index]).Name == Sheet ) {
                    break;
                }
                if( index == (wbk.Sheets.Count) ) {
                    Console.WriteLine($"Can't find sheet name: {Sheet}");
                    return "Can't find sheet name";
                }
            }    
            var sh = wbk.Sheets[index];
            sh.Activate();
            string result = sh.Cells[Cell[0], Cell[1]].Text;
            wbk.Close();
            app.Quit();
            return result;
        }

        public static void Run()
        {

            int index = 0;
            int DllIndex = 0;
            //********* SIT 依序填寫執行的DLL的項目 /Start/


            string path = @"c:\TestManager\ItemDownload\SCD_RV07RC.xls";

            // readExcel(path, "Support Language");
            string Enduro = readExcel(path, "SCL Content", new int[] { 4, 2 });
            string productType = CaptainWin.CommonAPI.SysInfo.GetWMI("Win32_SystemEnclosure", "ChassisTypes");
            // if( (Enduro == "Enduro") && (productType == "10")) {
            if( true ) {
                // Read Touch keyboard setting
                RegistryHive hive = RegistryHive.CurrentUser;
                string keyPath = "Software\\Microsoft\\TabletTip\\1.7";
                string itemName = "TipbandDesiredVisibility";

                var result = CaptainWin.CommonAPI.RegistryHelper.ReadRegistryValue(hive, keyPath, itemName);
                CaptainWin.CommonAPI.RegistryHelper.ReadRegistryValue(hive, keyPath, itemName);

                if (result.isFind)
                {
                    string value = result.getValue;
                    if( value == "1") {
                        Console.WriteLine("xxxxxxxxxxx");
                        TestResult("Pass");
                    }
                    else {
                        Console.WriteLine("ttttttttttttt");
                        TestResult("Fail");
                    }    
                } else {
                    Console.WriteLine("Can't find out registry!");
                }
            }
            else {

            }

            // Console.WriteLine(productType);

/* (EdisonLin-20240122-1)
            RegistryHive hive = RegistryHive.LocalMachine;
            string keyPath = "SOFTWARE\\Policies\\Microsoft\\Windows Defender";
            string itemName = "DisableAntiSpyware";

            var result = CaptainWin.CommonAPI.RegistryHelper.ReadRegistryValue(hive, keyPath, itemName);
            CaptainWin.CommonAPI.RegistryHelper.ReadRegistryValue(hive, keyPath, itemName);

            if (result.isFind)
            {
                string value = result.getValue;
                Console.WriteLine($"A value: {value}");
            } else
            {
                Console.WriteLine("xxxxxxxxxxxxxxxxxxxxxxxxx");
            }
(EdisonLin-20240122-1) */


/* (EdisonLin-20240123-1)
(EdisonLin-20240123-1) */
// Read Touch keyboard setting
            // RegistryHive hive = RegistryHive.CurrentUser;
            // string keyPath = "Software\\Microsoft\\TabletTip\\1.7";
            // string itemName = "TipbandDesiredVisibility";

            // var result = CaptainWin.CommonAPI.RegistryHelper.ReadRegistryValue(hive, keyPath, itemName);
            // CaptainWin.CommonAPI.RegistryHelper.ReadRegistryValue(hive, keyPath, itemName);

            // if (result.isFind)
            // {
            //     string value = result.getValue;
            //     Console.WriteLine($"A value: {value}");
            // } else
            // {
            //     Console.WriteLine("Can't find out!");
            // }



/* (EdisonLin-20240122-2)

            DateTime fromDate = new DateTime(2024, 1, 22, 08, 30, 0);
            DateTime toDate = new DateTime(2024, 1, 22, 15, 40, 0);
            List<EventLogEntryDetails> eventLogEntries = new List<EventLogEntryDetails>();
            eventLogEntries = CaptainWin.CommonAPI.EventLogHelper.QueryEventLog(fromDate ,toDate, "Application", "c:\\TestManager\\");

            foreach (EventLogEntryDetails entry in eventLogEntries)
            {
                Console.WriteLine(entry.TimeGenerated);
                Console.WriteLine(entry.Source);
                Console.WriteLine(entry.EntryType);
                Console.WriteLine(entry.Id);
            }
(EdisonLin-20240122-2) */


/* (EdisonLin-20240122-3)

            string timeZone = CaptainWin.CommonAPI.CommonReadOOBESpecTable.GetTimeZone();
            string UUID = CaptainWin.CommonAPI.RegistryHelper.GetWindowsUUID();
            Console.WriteLine(timeZone);
            Console.WriteLine(UUID);
            Console.ReadKey();
(EdisonLin-20240122-3) */

            // DllIndex++; // 1
            // if( DllIndex > index ) {
            //     RecordDllIndex(DllIndex);
            //     CaptainWin.CommonAPI.DoSleep.Sleep(3, 1);
            // }    

            // DllIndex++; // 2
            // if( DllIndex > index ) {
            //     RecordDllIndex(DllIndex);
            //     CaptainWin.CommonAPI.Smode.GetSmode();
            // }    

            // DllIndex++; // 2
            // if( DllIndex > index ) {
            //     RecordDllIndex(DllIndex);
            //    CaptainWin.CommonAPI.HDMI.HdmiConnectionStatus();
            // }    

            // DllIndex++; // 2
            // if( DllIndex > index ) {
            //     RecordDllIndex(DllIndex);
            //     CaptainWin.CommonAPI.SysInfo.GetWMI("Win32_BIOS");
            // }    

            // DllIndex++; // 2
            // if( DllIndex > index ) {
            //     RecordDllIndex(DllIndex);
            //     CaptainWin.CommonAPI.SysInfo.GetWMI("Win32_BaseBoard");
            // }    

            DllIndex++; // 2
            if( DllIndex > index ) {
                RecordDllIndex(DllIndex);
                // CaptainWin.CommonAPI.SysInfo.GetWMI("Win32_SystemEnclosure");
                // CaptainWin.CommonAPI.SysInfo.GetWMI("Win32_SystemEnclosure");
                // CaptainWin.CommonAPI.SysInfo.GetWMI("Win32_SystemEnclosure", "ChassisTypes");
                // CaptainWin.CommonAPI.SysInfo.GetWMI("Win32_SystemEnclosure", "InstallDate");
            }    

            // DllIndex++; // 3
            // if( DllIndex > index ) {
            //     RecordDllIndex(DllIndex);
            //     CaptainWin.CommonAPI.DoReboot.Reboot(60);
            // }    

            // DllIndex++; // 4
            // if( DllIndex > index ) {
            //     RecordDllIndex(DllIndex);
            //     CaptainWin.CommonAPI.SysInfo.GetWMI("Win32_SMBIOSMemory");
            // }    

            // DllIndex++; // 5
            // if( DllIndex > index ) {
            //     RecordDllIndex(DllIndex);
            //     CaptainWin.CommonAPI.SysInfo.GetWMI("Win32_OperatingSystem");
            // }    


            // DllIndex++; // 7
            // if( DllIndex > index ) {
            //     RecordDllIndex(DllIndex);
            //     CaptainWin.CommonAPI.Culture.GetCulture();
            // }    

            DllIndex++; // 4
            if( DllIndex > index ) {
                RecordDllIndex(DllIndex);
                // CaptainWin.CommonAPI.GetSystemInfo.GetOSVersion();
                // CaptainWin.CommonAPI.GetSystemInfo.GetSystemType();
                // CaptainWin.CommonAPI.GetSystemInfo.GetProcessorName();
                // CaptainWin.CommonAPI.GetSystemInfo.GetRunningTime();
                // CaptainWin.CommonAPI.GetSystemInfo.GetPhysicalMemory();
                // CaptainWin.CommonAPI.GetSystemInfo.GetCpuId();
                // CaptainWin.CommonAPI.GetSystemInfo.GetCPUCount();
                // CaptainWin.CommonAPI.GetSystemInfo.GetDiskDevice();
                // CaptainWin.CommonAPI.GetSystemInfo.GetDiskSpace();
                // bool r = CaptainWin.CommonAPI.GetSystemInfo.GetDiskFormat();
            }    
            //********* SIT 依序填寫執行的DLL的項目 /End/

            RecordDllIndex(0);
        }

        public static void UpdateResults() {

            // RegistryHive hive = RegistryHive.CurrentUser;
            // string keyPath = "SOFTWARE\\Microsoft\\Microsoft\\Windows\\CurrentVersion\\Run";
            // string itemName = "OneDrive";

            // var result = CaptainWin.CommonAPI.RegistryHelper.ReadRegistryValue(hive, keyPath, itemName);

            // if (result.isFind)
            // {
            //     string value = result.getValue;
            //     Console.WriteLine($"A value: {value}");
            // } else
            // {
            //     Console.WriteLine("4444444444444444444444444444444");
            // }

                // CaptainWin.CommonAPI.GetSystemInfo.GetOSVersion();
        }

        public static void Setup() {
                // CaptainWin.CommonAPI.GetSystemInfo.GetPhysicalMemory();
        }

        public static void TearDown() {
                // CaptainWin.CommonAPI.GetSystemInfo.GetSystemType();
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
                Console.WriteLine($"sdfsadfsdfsd An error occurred: {ex.Message}");
            }
        }   //RecordDllIndex    

    } //Class1
} //namespace
