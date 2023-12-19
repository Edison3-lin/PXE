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

namespace MyTestAll2
{
    public class Class1
    {

        private const string ThisFileName = "MyTestAll2.dll";
        private static int DllIndex;
        private static string ItemDownload = "C:\\TestManager\\ItemDownload\\";
        // private const string TR = "C:\\TestManager\\TR_Result.json";
 
       public static void Setup()
        {
        }

        public static void Run()
        {
Console.WriteLine("RebbotDllIndex dsfdsf" );
string jsonString = "{\"name\":\"John\",\"age\":30,\"city\":\"New York\"}";
try{
            string TR = @"c:\\TestManager\\TR_Result.json"; // 將路徑替換為你的JSON文件的實際路徑

//                     string jsonString = File.ReadAllText(TR);
// Console.WriteLine(jsonString);

                     JObject json = JObject.Parse(jsonString);
                    // int timeout = (int)json["Test_TimeOut"];
// Console.WriteLine(timeout);
}
catch (FileNotFoundException)
        {
            Console.WriteLine("找不到指定的文件。");
        }
        catch (IOException)
        {
            Console.WriteLine("讀取文件時出現了錯誤。");
        }
        catch (Exception ex)
        {
            Console.WriteLine("發生了未處理的錯誤：" + ex.Message);
        }

            //********* SIT 依序填寫執行的DLL的項目 /Start/
            // RecordDllIndex(0);
            // // CaptainWin.CommonAPI.TestOperation.Sleep(3, 1);
            // RecordDllIndex(1);
            // CaptainWin.CommonAPI.TestOperation.Reboot(10);
            // RecordDllIndex(2);
            //********* SIT 依序填寫執行的DLL的項目 /End/


            // HadRun("_kIll_");
        }

        public static void UpdateResults()
        {
        }

        public static void TearDown()
        {
        }

        public static void RecordDllIndex(int DllIndex)
        {

try{
                            //string ftpJson = System.IO.File.ReadAllText(TR);
                            // JObject fjson = JObject.Parse(ftpJson);
                            // fjson["TestStatus"] = "DONE";
                            // string updatedJson = fjson.ToString();
                            // System.IO.File.WriteAllText(TR, updatedJson);
}
                catch (Exception ex)
                {
                    Console.WriteLine($"sdfsadfsdfsd An error occurred: {ex.Message}");
                }
// jsonString = System.IO.File.ReadAllText(TR);     
// json = JObject.Parse(jsonString);
// Console.WriteLine("RebbotDllIndex {0}" ,json["Reboot"]);
// Console.ReadKey();
           
        }    

        public static bool HadRun(string DllFileName)
        {
            // DoneDll.txt 紀錄已執行到第幾個DLL
            string log_path = ItemDownload + "DoneDll.txt";
            DllIndex++;

            // 檢查檔案是否存在，如果不存在則建立
            if (!File.Exists(log_path))
            {
                using (FileStream fs = File.Create(log_path));
            }
            else
            {
                    
                if(DllFileName == "_kIll_")
                {
                    // 使用File.Delete删除文件
                    File.Delete(log_path);
                    return true;
                }

                try
                {
                    // 跳過已經執行的DLL
                    using (StreamReader reader = new StreamReader(log_path))
                    {
                        string strNumber = reader.ReadToEnd();
                        int number = int.Parse(strNumber);
                        // 如果Dll index小於紀錄的值，表示已執行過
                        if(number >= DllIndex)
                            return true;
                    }
                }
                catch (FileNotFoundException)
                {
                    Console.WriteLine($"File not found: {DllFileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                }
            }

            try
            {
                // 紀錄執行到第DllIndex個DLL
                using (StreamWriter writer = new StreamWriter(log_path))
                {
                    writer.Write(DllIndex.ToString());
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("WriteLog Error!!! " + ex.Message);
            }

            return false;
        }

    }
}
