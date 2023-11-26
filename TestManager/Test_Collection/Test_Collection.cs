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
        private static string currentDirectory = Directory.GetCurrentDirectory() + "\\ItemDownload\\";
 
       public int Setup()
        {
            // common.Setup
            // Console.WriteLine("collect setup~~~~~~~~~~~~~~~~~~~");

            return 81;
        }

        public int Run()
        {
            // Testflow.Run(DllName);

            //********* SIT 依序填寫執行的DLL的項目 /Start/
            Execute_dll("Test3.dll");
            Execute_dll("TestItem2.dll");
            Execute_dll("TestItem1.dll");
            Execute_dll("T2.dll");
            Execute_dll("T3.dll");
            //********* SIT 依序填寫執行的DLL的項目 /End/

            HadRun("");
            return 82;
        }

        public int UpdateResults()
        {

            // Return the test results from 'Run'
            return 83;
        }

        public int TearDown()
        {
            return 84;
        }

        public static void Execute_dll(string DllFileName)
        {

            try
            {
                if(!HadRun(DllFileName))
                {
                    // Console.WriteLine(DllFileName+" XXXXXX");
                    // Console.ReadKey();
                    Common.Runnner.RunTestItem(currentDirectory+DllFileName);
                }    
            }
            catch (Exception ex)
            {
                Console.WriteLine(DllFileName + ex.Message);
            }
        }    

        public static bool HadRun(string DllFileName)
        {

            string log_path = currentDirectory + "DoneDll.txt";

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
                    // 使用File.ReadAllLines按行讀取文件的內容
                    string[] lines = File.ReadAllLines(log_path);

                    // 列印每一行的內容
                    // Console.WriteLine("File Content:");
                    foreach (string line in lines)
                    {
                        if(line == DllFileName)
                        {
                            return true;
                        }
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
                // 使用 StreamWriter 打開檔案並appand內容
                using (StreamWriter writer = new StreamWriter(log_path, true))
                {
                    writer.Write(DllFileName+'\n');
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
