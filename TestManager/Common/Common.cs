using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Common
{
    public class Runnner {
        public static bool RunTestItem(string dllPath, object[] S, object[] R, object[] U, object[] T ) {
            Assembly myDll = Assembly.LoadFile(dllPath);
            var myTest=myDll.GetTypes().First(m=>!m.IsAbstract && m.IsClass);
            object myObj = myDll.CreateInstance(myTest.FullName);
            object myResult = null;

            try {
                Testflow.General.WriteLog("Common", dllPath+".Setup()" );
                myTest.GetMethod("Setup").Invoke(myObj, S);            
            }
            catch (Exception ex) {
                Testflow.General.WriteLog("Common", "Setup() Error!!! " + ex.Message);
            }  

            try {
                Testflow.General.WriteLog("Common", dllPath+".Run()" );
                myResult = myTest.GetMethod("Run").Invoke(myObj, R);            
            }
            catch (Exception ex) {
                Testflow.General.WriteLog("Common", "Run() Error!!! " + ex.Message);
            }   

            try {
                Testflow.General.WriteLog("Common", dllPath+".UpdateResults()" );
                myTest.GetMethod("UpdateResults").Invoke(myObj, U);            
            }
            catch (Exception ex) {
                Testflow.General.WriteLog("Common", "UpdateResults() Error!!! " + ex.Message);
            }   

            try {
               Testflow.General.WriteLog("Common", "Invoke "+dllPath+".TearDown()" );
               myTest.GetMethod("TearDown").Invoke(myObj, T);   
            }   
            catch (Exception ex) {
                Testflow.General.WriteLog("Common", "TearDown() Error!!! " + ex.Message);
            }   

            if(myResult.ToString() == "True") 
                return true;
            else 
                return false;
        }
    }
    public class Testflow
    {
        private const string TMDIRECTORY = "C:\\TestManager\\";
        private const string TESTLOGDIRECTORY = "C:\\TestManager\\TestLog\\";

        public static void Setup(string logFileName)
        {
            General.WriteLog(logFileName, $"General.WriteLog({logFileName}, \"Setup\");");
        }
        public static void Run(string logFileName)
        {
            General.WriteLog(logFileName, $"General.WriteLog({logFileName}, \"Run\");");
        }

        public static void UpdateResults(string logFileName, bool passFail)
        {
            General.WriteLog(logFileName, $"General.WriteLog({logFileName}, \"UpdateResults\");");
        }

        public static void TearDown(string logFileName)
        {
            General.WriteLog(logFileName, $"General.WriteLog({logFileName}, \"TearDown\");");
        }

        public class General
        {
            public static void WriteLog(string logFileName, string content)
            {
                string LogFile = TESTLOGDIRECTORY+logFileName+".log";
                try
                {
                    // appand content
                    using (StreamWriter writer = new StreamWriter(LogFile, true))
                    {
                        writer.Write("["+DateTime.Now.ToString()+"] "+content+'\n');
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("WriteLog Error!!! " + ex.Message);
                }
            
            }
        }
    }
}
