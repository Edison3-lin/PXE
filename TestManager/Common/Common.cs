using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Common
{
    public class Runnner
    {
        //private static string test_dll_folder;

        public void SetTestDllFolder(string folder)
        {
            // Runnner.test_dll_folder = folder;
        }

        /*
         * DllName: the unique file name (without '.dll') of the test dll. For example: MyTestItem_1
         * 
         */
        public static bool RunTestItem(string dllPath)
        {

            Testflow.General.WriteLog("RunTestItem", dllPath);
            try
            {
                Assembly myDll = Assembly.LoadFile(dllPath);
                var myTest=myDll.GetTypes().First(m=>!m.IsAbstract && m.IsClass);
                object myObj = myDll.CreateInstance(myTest.FullName);
                myTest.GetMethod("Setup").Invoke(myObj, new object[]{});            
                object myResult = myTest.GetMethod("Run").Invoke(myObj, new object[]{});            
                Testflow.General.WriteLog("RunTestItem", myResult.ToString());
                myTest.GetMethod("UpdateResults").Invoke(myObj, new object[]{});            
                myTest.GetMethod("TearDown").Invoke(myObj, new object[]{});   
                if(myResult.ToString() == "True") return true;
                else return false;
            }
            catch (Exception ex)
            {
                Testflow.General.WriteLog("RunTestItem", "Common Error!!! " + ex.Message);
            }   
            return false;
        }
    }
    public class Testflow
    {
        public static int Setup(string DllName)
        {
            General.WriteLog(DllName, "Testflow::Setup");
            return 90;
        }
        public static int Run(string DllName)
        {
            General.WriteLog(DllName, "Testflow::Run");
            return 90;
        }

        public static int UpdateResults(string DllName, bool passFail)
        {
            General.WriteLog(DllName, "Testflow::UpdateResults");
            return 90;
        }

        public static int TearDown(string DllName)
        {
            General.WriteLog(DllName, "Testflow::TearDown");
            return 90;
        }

        public class General
        {
            public static int WriteLog(string DllName, string content)
            {
                string log_path = "C:\\TestManager\\ResultUpload\\" + DllName+".log";
                // 檢查目錄是否存在，如果不存在則建立
                if (!Directory.Exists("C:\\TestManager\\ResultUpload\\"))
                {
                    Directory.CreateDirectory("C:\\TestManager\\ResultUpload\\");
                }                

                try
                {
                    // 使用 StreamWriter 打開檔案並appand內容
                    using (StreamWriter writer = new StreamWriter(log_path, true))
                    {
                        writer.Write("["+DateTime.Now.ToString()+"] "+content+'\n');
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("WriteLog Error!!! " + ex.Message);
                }
            
                return 0;
            }
        }
    }
}
