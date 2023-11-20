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
            Assembly myDll = Assembly.LoadFile(dllPath);
            var myTest=myDll.GetTypes().First(m=>!m.IsAbstract && m.IsClass);
            object myObj = myDll.CreateInstance(myTest.FullName);
            object myResult = null;

            try
            {
                Testflow.General.WriteLog("Common", "Invoke "+dllPath+".Setup()" );
                myTest.GetMethod("Setup").Invoke(myObj, new object[]{});            
            }
            catch (Exception ex)
            {
                Testflow.General.WriteLog("Common", "Setup() Error!!! " + ex.Message);
            }  

            try
            {
                Testflow.General.WriteLog("Common", "Invoke "+dllPath+".Run()" );
                myResult = myTest.GetMethod("Run").Invoke(myObj, new object[]{});            
            }
            catch (Exception ex)
            {
                Testflow.General.WriteLog("Common", "Run() Error!!! " + ex.Message);
            }   

            try
            {
                Testflow.General.WriteLog("Common", "Invoke "+dllPath+".UpdateResults()" );
                myTest.GetMethod("UpdateResults").Invoke(myObj, new object[]{});            
            }
            catch (Exception ex)
            {
                Testflow.General.WriteLog("Common", "UpdateResults() Error!!! " + ex.Message);
            }   

            try
            {
               Testflow.General.WriteLog("Common", "Invoke "+dllPath+".TearDown()" );
               myTest.GetMethod("TearDown").Invoke(myObj, new object[]{});   
            }   
            catch (Exception ex)
            {
                Testflow.General.WriteLog("Common", "TearDown() Error!!! " + ex.Message);
            }   

            if(myResult.ToString() == "True") return true;
            else return false;
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
            private static string currentDirectory = Directory.GetCurrentDirectory() + '\\';
            // private static string MyLog = currentDirectory+"MyLog\\";
            private static string TestLog = currentDirectory+"TestLog\\";

            public static int WriteLog(string DllName, string content)
            {
                string log_file = TestLog+DllName+".log";
                try
                {
                    // 使用 StreamWriter 打開檔案並appand內容
                    using (StreamWriter writer = new StreamWriter(log_file, true))
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
