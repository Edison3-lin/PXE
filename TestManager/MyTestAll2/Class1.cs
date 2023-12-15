using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LoadDll;

namespace MyTestAll2
{
    public class Class1
    {

        private const string ThisFileName = "MyTestAll2.dll";
        private static int DllIndex;
        private static string ItemDownload = "C:\\TestManager\\ItemDownload\\";
 
       public int Setup()
        {
            return 0;
        }

        public int Run()
        {
            DllIndex = 0;

            //********* SIT 依序填寫執行的DLL的項目 /Start/
            Execute_dll("MyPowerShell.dll");
            Execute_dll("MyParameter.dll");
            // Execute_dll("MyReboot.dll");   //reboot
            Execute_dll("MyParameter.dll");
            // Execute_dll("MyReboot.dll");   //reboot
            Execute_dll("MyParameter.dll");
            //********* SIT 依序填寫執行的DLL的項目 /End/


            HadRun("_kIll_");
            return 0;
        }

        public int UpdateResults()
        {
            return 0;
        }

        public int TearDown()
        {
            return 0;
        }

        public static void Execute_dll(string DllFileName)
        {
            try
            {
                if(!HadRun(DllFileName))
                {
                    LoadDll.Runnner.RunTestItem(ItemDownload+DllFileName);
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
