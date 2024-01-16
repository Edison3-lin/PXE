using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Threading;
using System.Reflection;
using System.Diagnostics;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Security.Cryptography;

namespace TM1007b1 {

    public class TestManager {
        private const string TMDIRECTORY = "C:\\TestManager\\";
        private const string ITEMDOWNLOAD = "C:\\TestManager\\ItemDownload\\";
        private const string SIGNKEY = "C:\\TestManager\\Key\\";
        private const string TMLOG = "C:\\TestManager\\MyLog\\TestManager.log";
        private const string TR = "C:\\TestManager\\TR_Result.json";
        static Stopwatch ItemWatch = new Stopwatch();
        private static int timeout = int.MaxValue;

        private static string arg0 = null;
        private static string arg1 = null;

        // **** CheckTR ****
/*
    "TCM_ID":  0,
    "TR_ID":  0,
    "TestResult":  "",
    "TestStatus":  "",
    "Test_TimeOut":  0,
    "TCM_Done":  false,
    "Text_Log_File_Path":  [
*/
        static bool CheckTR() {
            // Read Reboot status    
            string ftpJson = System.IO.File.ReadAllText(TR);
            JObject fjson = JObject.Parse(ftpJson);
            if( (fjson["TCM_ID"] == null) || (fjson["TR_ID"] == null) ) {
                Console.WriteLine("TR_Result.json lost \"TCM_ID\"");
                return false;
            }    

            if( (fjson["TestResult"] == null) || (fjson["TestStatus"] == null) ) {
                Console.WriteLine("TR_Result.json lost \"TestResult\"");
                return false;
            }    

            if( fjson["Test_TimeOut"] == null ) {
                Console.WriteLine("TR_Result.json lost \"Test_TimeOut\"");
                return false;
            }    

            if( fjson["TCM_Done"] == null ) {
                Console.WriteLine("TR_Result.json lost \"TCM_Done\"");
                return false;
            }    

            if( fjson["Text_Log_File_Path"] == null ) {
                Console.WriteLine("TR_Result.json lost \"Text_Log_File_Path\"");
                return false;
            }    

            return true;
        }

        // **** Check Need Update ****
        static bool UpgradeCheck() {
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();
            try {
                pipeline.Commands.AddScript(TMDIRECTORY + "UpgradeCheck.ps1");
                var result = pipeline.Invoke();
                runspace.Close();  
                foreach (var psObject in result)
                {
                    if(psObject != null)
                    {
                        if(psObject.ToString() == "True") 
                            return true;
                        else 
                            return false;
                    }    
                    else
                        return false;
                }
            }
            catch {
                runspace.Close();
                ProcessLog("Waiting 2 sec for ready");
                Thread.Sleep(2000);
                return false;
            }
            return false;
        }    

        // **** Update TestManager ****
        static void UpgradTestManager() {
            int currentProcessId = Process.GetCurrentProcess().Id;
            File.Copy("C:\\TestManager\\UpgradTestManager.ps1", "C:\\TestManager\\UT.ps1", true);
            string scriptCommand = $"Start-Process powershell -ArgumentList '-NoExit -File C:\\TestManager\\UT.ps1' -WindowStyle Hidden; Stop-Process -Id {currentProcessId}";

            // Build ProcessStartInfo object，setting process information
             ProcessStartInfo psi = new ProcessStartInfo
             {
                 FileName = "powershell.exe",
                 Arguments = $"-Command \"{scriptCommand}\"",
                 RedirectStandardOutput = true,
                 UseShellExecute = false,
                 CreateNoWindow = true
             };

             // Create Process object
             Process process = new Process
             {
                 StartInfo = psi
             };
             
             process.Start();       // Start process
             process.WaitForExit();
             string output = process.StandardOutput.ReadToEnd();
             Console.WriteLine(output);
             process.Close();       // Close process
             Environment.Exit(0);   // Close TestManger
        }

        // **** 創建log file ****
        static void CreateDirectoryAndFile() {
            try {
                if (!Directory.Exists(ITEMDOWNLOAD))
                {
                    Directory.CreateDirectory(ITEMDOWNLOAD);
                }                
                if (!Directory.Exists(TMDIRECTORY+"MyLog\\"))
                {
                    Directory.CreateDirectory(TMDIRECTORY+"MyLog\\");
                }                
                if (!Directory.Exists(TMDIRECTORY+"TestLog\\"))
                {
                    Directory.CreateDirectory(TMDIRECTORY+"TestLog\\");
                }                
                // if (!Directory.Exists(TMDIRECTORY+"Key\\"))
                // {
                //     Directory.CreateDirectory(TMDIRECTORY+"Key\\");
                // }                

                if (!File.Exists(TMLOG))
                {
                    using (FileStream fs = File.Create(TMLOG));
                }
                // else
                // {
                //     // Clear TMLOG content
                //     using (FileStream fs = new FileStream(TMLOG, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                //     {
                //         fs.SetLength(0);
                //     }                    
                // }
            }
            catch (Exception ex) {
                Console.WriteLine("Error!!! " + ex.Message);
            }
        }

        // **** TestManager.log ****
        public static void ProcessLog(string content) {
            try {
                // appand content
                using (StreamWriter writer = new StreamWriter(TMLOG, true))
                {
                    writer.Write("["+DateTime.Now.ToString()+"] "+content+'\n');
                }

            }
            catch (Exception ex) {
                Console.WriteLine("Error!!! " + ex.Message);
            }
        }      
        
        // ***** get jobs from DB *****
        static string DBimage() {
            string jobList = null;
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();
            try {
                pipeline.Commands.AddScript(TMDIRECTORY + "DBimage.ps1");
                var result = pipeline.Invoke();
                foreach (var psObject in result)
                {
                    if(psObject != null)
                        jobList = psObject.ToString();
                    else
                        jobList = null;
                }
                runspace.Close();                
            }
            catch {
                runspace.Close();
                ProcessLog("Waiting 2 sec for DBimage ready");
                Thread.Sleep(2000);
                return null;
            }
            return jobList;
        }

        // ***** get jobs from DB *****
        static string DBjob() {
            string jobList = "";
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();
            try {
                pipeline.Commands.AddScript(TMDIRECTORY+"DBjob.ps1");
                var result = pipeline.Invoke();
                runspace.Close();
                foreach (var psObject in result)
                {
                    if(psObject != null)
                        jobList = psObject.ToString();
                    else
                        jobList = null;
                }
            }    
            catch {
                runspace.Close();
                ProcessLog("Waiting 2 sec for DBjob ready");
                Thread.Sleep(2000);
                return null;
            }
            return jobList;
        }

        // ***** FTPdownload *****
        static bool FTPdownload(string jobList) {
            ProcessLog("   Log => C:\\TestManager\\MyLog\\FTPdownload_process.log");
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();

            try {
                pipeline.Commands.AddScript("$remoteFile = \""+jobList+"\"");
                pipeline.Commands.AddScript(TMDIRECTORY+"FTPdownload.ps1");
                var result = pipeline.Invoke();
                runspace.Close();
                if(result[0].ToString() == "True") return true;
                else return false;
            }    
            catch (Exception ex) {
                ProcessLog("Error!!! Downloading "+ex.Message);
                runspace.Close();
                return false;
            }
        }
        // ***** upload a program from FTP *****
        static void FTPupload() {
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();

            try {
                pipeline.Commands.AddScript(TMDIRECTORY+"FTPupload.ps1 ");
                var result = pipeline.Invoke();
            }    
            catch (Exception ex) {
                ProcessLog("Upload "+ex.Message);
                runspace.Close();
                return;
            }

            runspace.Close();

            // Clear LOG content
            // using (FileStream fs = new FileStream(TMLOG, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            // {
            //     fs.SetLength(0);
            // }                    
            // using (FileStream fs = new FileStream("C:\\TestManager\\MyLog\\LoadDll.log", FileMode.OpenOrCreate, FileAccess.ReadWrite))
            // {
            //     fs.SetLength(0);
            // }                    

            return;
        }

        // ***** DB test *****
        static bool DBtest() {
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();
            try {
                pipeline.Commands.AddScript(TMDIRECTORY+"DBtest.ps1");
                var result = pipeline.Invoke();
                if(result[0].ToString() != "Open")
                {
                    runspace.Close();
                    return false;
                }
            }    
            catch (Exception ex) {
                ProcessLog("DB test "+ex.Message);
                runspace.Close();
                return false;
            }
            return true;
        }

        // ***** update job_status to DB *****
        static bool DBupdateStatus(string TestStatus) {
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();

            try {
                if( TestStatus == " ") {
                    pipeline.Commands.AddScript(TMDIRECTORY+"DBupdateStatus.ps1");
                }
                else {
                    pipeline.Commands.AddScript(TMDIRECTORY+"DBupdateStatus.ps1 "+TestStatus);
                }
                var result = pipeline.Invoke();
                if(result[0].ToString() == "Unconnected_")
                {
                    runspace.Close();
                    return false;
                }
            }    
            catch (Exception ex) {
                ProcessLog("Update "+ex.Message);
                runspace.Close();
                return false;
            }

            runspace.Close();
            return true;
        }

        // ***** ExecuteDll *****
        static void ExecuteDll(string dllPath) {
            string callingDomainName = AppDomain.CurrentDomain.FriendlyName;
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            AppDomain ad = AppDomain.CreateDomain("TestManager DLL");
            ProxyObject obj = (ProxyObject)ad.CreateInstanceFromAndUnwrap(basePath+callingDomainName, "TM1007b1.ProxyObject");
            try {
                ProcessLog("Loading.."+dllPath);
                obj.LoadAssembly(dllPath);
            }
            catch (System.IO.FileNotFoundException) {
                ProcessLog("!!! Can't find out Dll");
                return;
            }

//(EdisonLin-20240110-1)             // Read Reboot status    
//(EdisonLin-20240110-1)             string ftpJson = System.IO.File.ReadAllText(TR);
//(EdisonLin-20240110-1)             JObject fjson = JObject.Parse(ftpJson);
//(EdisonLin-20240110-1)             bool Reboot = ( (int)fjson["Reboot"] > 0 ); //Reboot?

            // Start Stopwatch
            ItemWatch = new Stopwatch();
            ItemWatch.Start();

            /* Setup() */
            try {
//(EdisonLin-20240110-1)                 if(!Reboot) {
                obj.Invoke("Setup", null);
                ProcessLog("Setup() - Success! ");
//(EdisonLin-20240110-1)                 }    
            }
            catch (Exception ex)
            {
                ProcessLog("--- Setup() Exception caught ---");
                ProcessLog("Exception type:  " + ex.GetType().Name);
                ProcessLog("error message:  " + ex.Message);
                ProcessLog("stack trace:  " + ex.StackTrace);
            }

            /* Run() */
            try {
                obj.Invoke("Run", null);
                ProcessLog("Run() - Success! ");
            }
            catch (Exception ex)
            {
                ProcessLog("--- Run() Exception caught ---");
                ProcessLog("Exception type:  " + ex.GetType().Name);
                ProcessLog("error message:  " + ex.Message);
                ProcessLog("stack trace:  " + ex.StackTrace);
            }

            /* UpdateResults() */
            try {
                obj.Invoke("UpdateResults", null);
                ProcessLog("UpdateResults() - Success! ");
            }
            catch (Exception ex)
            {
                ProcessLog("--- UpdateResults() Exception caught ---");
                ProcessLog("Exception type:  " + ex.GetType().Name);
                ProcessLog("error message:  " + ex.Message);
                ProcessLog("stack trace:  " + ex.StackTrace);
            }

            /* TearDown() */
            try {
                obj.Invoke("TearDown", null);
                ProcessLog("TearDown() - Success! ");
            }
            catch (Exception ex)
            {
                ProcessLog("--- TearDown() Exception caught ---");
                ProcessLog("Exception type:  " + ex.GetType().Name);
                ProcessLog("error message:  " + ex.Message);
                ProcessLog("stack trace:  " + ex.StackTrace);
            }

            // Stop Stopwatch
            ItemWatch.Stop();
            AppDomain.Unload(ad);
            obj = null;
        }

        // ******* New Thread to monitor TimeOut *********
        static void MonitorExecutionTime() {
            bool NewWatch = true;
            do {
                Thread.Sleep(1000);

                if( NewWatch ) {
                    if(ItemWatch.Elapsed.TotalSeconds >= timeout) {
                        Console.WriteLine("\n==========================================");
                        Console.WriteLine($"    Time-out <{ItemWatch.Elapsed.TotalSeconds}> time exceeded !!!!");
                        Console.WriteLine("==========================================\n");

                        ProcessLog("==========================================");
                        ProcessLog($"    Time-out <{ItemWatch.Elapsed.TotalSeconds}> time exceeded !!!!");
                        ProcessLog("==========================================\n");

                        NewWatch = false;

                        /* Reset TestManager */
                        if( false ) {
                            string exePath = Assembly.GetExecutingAssembly().Location;  // Get the path of the currently executing program
                            string arguments;
                            if (arg0 != null) {
                                arguments = string.Format("{0} {1}", arg0, arg1);
                            } else {
                                arguments = null;
                            }  
                            Process.Start(exePath, arguments);                        
                            Environment.Exit(0);                                        // Close current program
                        }    
                    }
                }
                else {
                    if(ItemWatch.Elapsed.TotalSeconds < timeout)
                    {
                        NewWatch = true;
                    }                    
                }
            } while (!(Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Escape));
            
        }

        // ============== MAIN ==============
        static void Main(string[] args) {
            string JobList;
            DateTime startTime;
            DateTime endTime;
            TimeSpan timeSpan;
            Thread monitoringThread = new Thread(MonitorExecutionTime);     // Another Thread to watch timeout
            monitoringThread.Start();

            if(args.Length == 0) {
                do {

                    if( !CheckTR() ) {
                        Thread.Sleep(2000);
                        continue;
                    } else {
                        Console.WriteLine("TR_Result.json verify OK..");
                    };

                    CreateDirectoryAndFile();
                    while(!DBtest()) {
                        ProcessLog("Waiting 3 sec for database connection... ");
                        Thread.Sleep(3000);
                    };

//(EdisonLin-20240110-1)                     // Read Reboot status    
//(EdisonLin-20240110-1)                     string ftpJson = System.IO.File.ReadAllText(TR);
//(EdisonLin-20240110-1)                     JObject fjson = JObject.Parse(ftpJson);
//(EdisonLin-20240110-1)                     bool Reboot = ( (int)fjson["Reboot"] > 0 ); //Reboot?
                    // if( UpgradeCheck() )
                    if( false )
                    {
                        ProcessLog("Found a new TestManager version on FTP, trying to upgrade! ");
                        UpgradTestManager();
                    }   

                    JobList = DBimage();
                    if(JobList == null)
                    {
                        JobList = DBjob();

                    }
                    if(JobList == null)
                    {
                        Thread.Sleep(2000);
                        continue;
                    } 
                    else if (JobList == "Unconnected_")
                    {
                        ProcessLog("Waiting 5 sec for DB connected !!!");
                        Thread.Sleep(5000);
                        continue;
                    }    

                    // step 1. Got Job then downloading
                    ProcessLog("<<Step 1>> Got Job then downloading");
//(EdisonLin-20240110-1)                     if(!Reboot) {
                    if(!FTPdownload(JobList)) {
                        ProcessLog(" <<Drop!>> MD5 check failed");
                        ProcessLog("=================Completed================");
                        DBupdateStatus("Drop");
                        continue;
                    }
//(EdisonLin-20240110-1)                     }    
//(EdisonLin-20240110-1)                     else    
//(EdisonLin-20240110-1)                         ProcessLog("Skip downloading after reboot");

                    startTime = DateTime.Now;

                    // Read TR_Result.json timeout
                    string jsonString = System.IO.File.ReadAllText(TR);
                    JObject json = JObject.Parse(jsonString);
                    try {
                        timeout = (int)json["Test_TimeOut"];
                    }
                    catch (Exception ex)
                    {
                        ProcessLog("error message:  " + ex.Message);
                        ProcessLog("Parsing fail:  TR_Result.json [Test_TimeOut] ");
                    }    

                    try {
                        // step 2. Execute Dll
                        ProcessLog("<<Step 2>> Executing "+ITEMDOWNLOAD+JobList);
                        if( (JobList != "common_bios_pxeboot_default.dll") && (JobList != "common_image_pxeboot_default.dll") ) {
                            DBupdateStatus("Running");
                        }
                        ExecuteDll(ITEMDOWNLOAD+JobList);
                        if( (JobList == "common_bios_pxeboot_default.dll") || (JobList == "common_image_pxeboot_default.dll") ) {
                            ProcessLog("=================Completed================");
                            continue; //Don't update status for these DLL
                        }    
                    }
                    catch (Exception ex)
                    {
                        ProcessLog("--- ExecuteDll Exception caught ---");
                        ProcessLog("Exception type:  " + ex.GetType().Name);
                        ProcessLog("error message:  " + ex.Message);
                        ProcessLog("stack trace:  " + ex.StackTrace);
                    }
                    // step 3. update test status to DB
                    ProcessLog("<<Step 3>>  update test status to DB ");
                    DBupdateStatus(" ");   //Update test status as "Done"

                    // step 4. Listening job status
                    ProcessLog("<<Step 4>>  Keep listening job status");
                    endTime = DateTime.Now;
                    timeSpan = endTime - startTime;
                    ProcessLog("Spend " + timeSpan.Minutes + " Minutes " + timeSpan.Seconds + " Senconds");

                    // step 5. upload log to FTP
                    ProcessLog("<<Step 5>>  upload log to FTP");
                    ProcessLog("=================Completed================");
                    FTPupload();

//(EdisonLin-20240116-1)                 } while (!(Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Escape));
                } while (true);
            }  
            // Test DLL only          
            else {

                if( !CheckTR() ) {
                    ProcessLog("**** TR_Result.json error ****");            
                    Environment.Exit(0);            
                };

                arg0 = args[0];
//(EdisonLin-20240110-1)                 arg1 = args[1];
                CreateDirectoryAndFile();
                if( args.Length > 1 ) {
                    arg1 = args[1];                    
                    string strNumber = args[1];
                    try {
                        timeout = int.Parse(strNumber);
                    }
                    catch {
                        timeout = int.MaxValue;
                    }

                    Console.WriteLine("Timeout: " + timeout.ToString() + " seconds");
                }
                else {
                    timeout = int.MaxValue;    
                }

                try {
                    ProcessLog("<<Step 1>> Executing "+args[0]+" TimeOut: "+timeout+" seconds");
                    ExecuteDll(ITEMDOWNLOAD+args[0]);
                }
				catch (Exception ex)
				{
                    ProcessLog("--- ExecuteDll Exception caught ---");
				    ProcessLog("Exception caught:");
				    ProcessLog("Exception type:  " + ex.GetType().Name);
				    ProcessLog("error message:  " + ex.Message);
				    ProcessLog("stack trace:  " + ex.StackTrace);
				}
            }   // Test DLL only 

            // Close window
            ProcessLog("**** Exit ****");            
            Environment.Exit(0);            
        }
    }

    class ProxyObject : MarshalByRefObject {
        Assembly assembly = null;
        public void LoadAssembly(string myDllPath) {
            assembly = Assembly.LoadFile(myDllPath);
        }
        public bool Invoke(string methodName, params Object[] args) {
            if (assembly == null)
                return false;
            var cName=assembly.GetTypes().First(m=>!m.IsAbstract && m.IsClass);
            string fullClassName = cName.FullName;
            Type tp = assembly.GetType(fullClassName);
            if (tp == null)
                return false;
            MethodInfo method = tp.GetMethod(methodName);
            if (method == null)
                return false;
            Object obj = Activator.CreateInstance(tp);
            method.Invoke(obj, args);
            return true;
        }
    }    
}
