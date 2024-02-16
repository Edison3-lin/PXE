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
using System.Globalization;

using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;

namespace I005000 {
    public class MyI005000 {
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

        public static bool readExcel(string inName, string Sheet, string region, string timeZone) {

            if (!File.Exists(inName)) {
                Console.WriteLine($"Can't find {inName}");
                return false;
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
                    return false;
                }
            }    
            var sh = wbk.Sheets[index];
            sh.Activate();
            var usedRange = sh.UsedRange.CurrentRegion;
            int rows = usedRange.Rows.Count;
            int columns = usedRange.Columns.Count;
            bool result = false;
            for (int i = 3; i < rows; i++) {
                // Console.WriteLine( sh.Cells[i, 6].Text );                
                // Console.WriteLine( sh.Cells[i, 8].Text );                
                // Console.WriteLine( '\n' );                
                if( region == sh.Cells[i, 6].Text ){
                    if( timeZone == sh.Cells[i, 8].Text ) {
                        result = true;
                        break;
                    }
                }
            }
            wbk.Close();
            app.Quit();
            return result;
        }

        public static void Run()
        {
            CultureInfo currentCulture = CultureInfo.CurrentCulture;
            RegionInfo currentRegion = new RegionInfo(currentCulture.Name);
            string region = currentRegion.DisplayName;
            TimeZoneInfo tZone = TimeZoneInfo.Local;
            string timeZone = tZone.DisplayName;

            if( timeZone == "(UTC+08:00) 台北" ) {
                timeZone = "(UTC+08:00) Taipei";
            }

            string path = @"c:\TestManager\ItemDownload\Win11_SV2_OOBE_SPEC_20231108.xlsx";
            bool result = readExcel(path, "Lang_Region_Keyboard_Timezone", region, timeZone);

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
