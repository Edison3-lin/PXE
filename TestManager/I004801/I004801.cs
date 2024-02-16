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

namespace I004801 {
    public class MyI004801 {
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
                Console.WriteLine($"Write TR.json error occurred: {ex.Message}");
            }
        }
        public static void Run()
        {
            string a = CaptainWin.CommonAPI.GetSystemInfo.GetDiskMediaType();
            bool result = false;
            if (a == "Fixed hard disk media") {
                result = true;
            }
            if (result) {
                TestResult("Pass");
            } else {
                TestResult("Fail");
            }
        }

        public static void UpdateResults() {
        }

        public static void Setup() {
        }

        public static void TearDown() {
        }
    } //Class1
} //namespace
