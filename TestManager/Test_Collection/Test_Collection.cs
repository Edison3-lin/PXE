using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Common;

namespace Test_Collection
{
    public class Test_Collection
    {
        private const string ThisFileName = "Test_Collection.dll";
        private static int DllIndex;
        private static string currentDirectory = Directory.GetCurrentDirectory() + "\\ItemDownload\\";
 
       public int Setup()
        {
            // common.Setup
            return 81;
        }

        public int Run()
        {
            Testflow.Run("TEST");
            DllIndex = 0;

            //********* SIT 依序填寫執行的DLL的項目 /Start/
            Execute_dll("Test3.dll", new object[]{}, new object[]{}, new object[]{}, new object[]{});
            Execute_dll("TestItem2.dll", new object[]{}, new object[]{}, new object[]{}, new object[]{});
            // Execute_dll("TestItem1.dll", new object[]{}, new object[]{}, new object[]{}, new object[]{});   //reboot
            // Execute_dll("T2.dll", new object[]{}, new object[]{}, new object[]{}, new object[]{});
            // Execute_dll("TestItem1.dll", new object[]{}, new object[]{}, new object[]{}, new object[]{});   //reboot
            // Execute_dll("T3.dll", new object[]{}, new object[]{}, new object[]{}, new object[]{});
            // Execute_dll("TestItem1.dll", new object[]{}, new object[]{}, new object[]{}, new object[]{});   //reboot
            Execute_dll("C1.dll",  new object[] { 22, "Grace" }, new object[] { 20, 30, "Edison" }, new object[]{"林淑芳", 77}, new object[]{'a', "林宏斌"});
            //********* SIT 依序填寫執行的DLL的項目 /End/

            HadRun("");
            return 82;
        }

        public int UpdateResults()
        {

            return 83;
        }

        public int TearDown()
        {
            return 84;
        }

        public static void Execute_dll(string DllFileName, object[] S, object[] R, object[] U, object[] T)
        {
            try
            {
                if(!HadRun(DllFileName))
                {
                    Common.Runnner.RunTestItem(currentDirectory+DllFileName, S, R, U, T);
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
            string log_path = currentDirectory + "DoneDll.txt";
            DllIndex++;

            // 檢查檔案是否存在，如果不存在則建立
            if (!File.Exists(log_path))
            {
                using (FileStream fs = File.Create(log_path));
            }
            else
            {
                    
                if(DllFileName == "")
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
