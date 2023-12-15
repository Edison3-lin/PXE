using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace LoadDll {
    public class Runnner {
        public static void WriteLog(string content)
        {
            string LogFile = "C:\\TestManager\\MyLog\\LoadDll.log";
            if (!File.Exists(LogFile))
            {
                using (FileStream fs = File.Create(LogFile));
            }

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

        public static bool RunTestItem(string dllPath) {
            Assembly myDll = Assembly.LoadFile(dllPath);
            var myTest=myDll.GetTypes().First(m=>!m.IsAbstract && m.IsClass);
            object myObj = myDll.CreateInstance(myTest.FullName);
            object myResult = null;

            try {
                myTest.GetMethod("Setup").Invoke(myObj, new object[] {});
            }
            catch (Exception ex) {
                WriteLog( "Setup() Error!!! " + ex.Message);
            }  

            try {
                myResult = myTest.GetMethod("Run").Invoke(myObj, new object[] {});            
            }
            catch (Exception ex) {
                WriteLog( "Run() Error!!! " + ex.Message);
            }   

            try {
                myTest.GetMethod("UpdateResults").Invoke(myObj, new object[] {});            
            }
            catch (Exception ex) {
                WriteLog( "UpdateResults() Error!!! " + ex.Message);
            }   

            try {
               myTest.GetMethod("TearDown").Invoke(myObj, new object[] {});   
            }   
            catch (Exception ex) {
                WriteLog( "TearDown() Error!!! " + ex.Message);
            }   

            if(myResult.ToString() == "True") 
                return true;
            else 
                return false;
        }
    }
}
