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

namespace MyTestAll2 {
    public class Class1 {
        private const string TR = "C:\\TestManager\\TR_Result.json";
        public static void Setup() {
        }

        public static void Run()
        {
            string ftpJson = System.IO.File.ReadAllText(TR);
            JObject fjson = JObject.Parse(ftpJson);
            int index = (int)fjson["Reboot"];

            int DllIndex = 0;
            //********* SIT 依序填寫執行的DLL的項目 /Start/
            DllIndex++; // 1
            if( DllIndex > index ) {
                RecordDllIndex(DllIndex);
                CaptainWin.CommonAPI.DoSleep.Sleep(3, 1);
            }    

            DllIndex++; // 2
            if( DllIndex > index ) {
                RecordDllIndex(DllIndex);
                CaptainWin.CommonAPI.DoReboot.Reboot(10);
            }    

            // DllIndex++; // 3
            // if( DllIndex > index ) {
            //     RecordDllIndex(DllIndex);
            //     CaptainWin.CommonAPI.DoReboot.Reboot(10);
            // }    

            // DllIndex++; // 4
            // if( DllIndex > index ) {
            //     RecordDllIndex(DllIndex);
            //     CaptainWin.CommonAPI.DoReboot.Reboot(10);
            // }    
            //********* SIT 依序填寫執行的DLL的項目 /End/

            RecordDllIndex(0);
        }

        public static void UpdateResults() {
        }

        public static void TearDown() {
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
