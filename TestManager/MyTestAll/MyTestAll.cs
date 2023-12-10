using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Common;

namespace MyTestAll
{
    public class MyTestAll
    {
        private const string ThisFileName = "MyTestAll.dll";
        private static int DllIndex;
        private static string ItemDownload = "C:\\TestManager\\ItemDownload\\";
 
       public int Setup()
        {
            Testflow.General.WriteLog("MyTestAll", "Setup MyTestAll.dll...");
            return 0;
        }

        public int Run()
        {
            Testflow.General.WriteLog("MyTestAll", "Start MyTestAll.dll...");
            DllIndex = 0;

            //********* SIT 依序填寫執行的DLL的項目 /Start/
            //Execute_dll("Dll檔案名稱", Setup()參數, Run()參數, UpdateResults()參數, TearDown()參數);
            Execute_dll("MyPowerShell.dll", new object[]{}, new object[]{}, new object[]{}, new object[]{});
            Execute_dll("MyParameter.dll",  new object[] { 1122, "Quanta" }, new object[] { 13, 61, "Edison" }, new object[]{"斌", 77}, new object[]{'a', "==== tear down 結束 ===="});
            // Execute_dll("MyReboot.dll", new object[]{}, new object[]{}, new object[]{}, new object[]{});   //reboot
            Execute_dll("MyParameter.dll",  new object[] { 13332, "林" }, new object[] { 13, 61, "Sam" }, new object[]{"仙劍奇俠", 77}, new object[]{'a', "==== tear down 結束 ===="});
            // Execute_dll("MyReboot.dll", new object[]{}, new object[]{}, new object[]{}, new object[]{});   //reboot
            Execute_dll("MyParameter.dll",  new object[] { 22, "G2" }, new object[] { 20, 20, "Edison" }, new object[]{"Hello! ", 77}, new object[]{'c', "==== tear down 結束 ===="});
            //********* SIT 依序填寫執行的DLL的項目 /End/

            Testflow.General.WriteLog("MyTestAll", "Finish MyTestAll.dll...");

            HadRun("_kIll_");
            return 0;
        }

        public int UpdateResults()
        {
            Testflow.General.WriteLog("MyTestAll", "UpdateResults MyTestAll.dll...");
            return 0;
        }

        public int TearDown()
        {
            Testflow.General.WriteLog("MyTestAll", "TearDown MyTestAll.dll...");
            return 0;
        }

        public static void Execute_dll(string DllFileName, object[] S, object[] R, object[] U, object[] T)
        {
            try
            {
                if(!HadRun(DllFileName))
                {
                    Common.Runnner.RunTestItem(ItemDownload+DllFileName, S, R, U, T);
                }    
            }
            catch (Exception ex)
            {
                Console.WriteLine(DllFileName + ex.Message);
            }
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
