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

namespace TM1004 {
    public class TestManager {
        private const string TMDIRECTORY = "C:\\TestManager\\";
        private const string ITEMDOWNLOAD = "C:\\TestManager\\ItemDownload\\";
        private const string TMLOG = "C:\\TestManager\\MyLog\\TestManager.log";
        private const string TR = "C:\\TestManager\\TR_Result.json";
        static Stopwatch ItemWatch = new Stopwatch();
        private static int timeout = int.MaxValue;

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
            string scriptCommand = $"Start-Process powershell -ArgumentList '-NoExit -File C:\\TestManager\\UpgradTestManager.ps1' -WindowStyle Hidden; Stop-Process -Id {currentProcessId}";

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
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();

            try {
                // pipeline.Commands.AddScript(TMDIRECTORY+"RunAs.ps1");
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

        // ***** update job_status to DB *****
        static bool DBupdateStatus() {
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();

            try {
                pipeline.Commands.AddScript(TMDIRECTORY+"DBupdateStatus.ps1");
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
            ProxyObject obj = (ProxyObject)ad.CreateInstanceFromAndUnwrap(basePath+callingDomainName, "TM1004.ProxyObject");
            try {
                ProcessLog(".... Loading LoadDll.dll ....");
                obj.LoadAssembly(TMDIRECTORY+"LoadDll.dll");
            }
            catch (System.IO.FileNotFoundException) {
                ProcessLog("!!! Can't find out LoadDll.dll");
                return;
            }

            // Start Stopwatch
            ItemWatch = new Stopwatch();
            ItemWatch.Start();

            ProcessLog(".... Loading "+dllPath+" ....");
            Object[] p = new object[]{ dllPath };
            obj.Invoke("RunTestItem",p);

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
                        Console.WriteLine("\n======================================================");
                        Console.WriteLine($"Has been executed for <{ItemWatch.Elapsed.TotalSeconds}> seconds, Time-out time exceeded !!!!");
                        Console.WriteLine("========================================================\n");
                        NewWatch = false;
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
                    CreateDirectoryAndFile();

                    if( UpgradeCheck() )
                    //if( false )
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
                    // if (!File.Exists(ITEMDOWNLOAD+"DoneDll.txt"))
                        if(!FTPdownload(JobList)) {
                            ProcessLog(" <Job abort!>  MD5 check of FTP download failed");
                            // Update TR_Result.json 
                            string ftpJson = System.IO.File.ReadAllText(TR);
                            JObject fjson = JObject.Parse(ftpJson);
                            fjson["TestStatus"] = "DONE";
                            string updatedJson = fjson.ToString();
                            System.IO.File.WriteAllText(TR, updatedJson);
                            DBupdateStatus();
                            continue;
                        }
                    // else    
                    //     ProcessLog("Skip downloading again for reboot");

                    startTime = DateTime.Now;

                    // Read TR_Result.json timeout
                    string jsonString = System.IO.File.ReadAllText(TR);
                    JObject json = JObject.Parse(jsonString);
                    timeout = (int)json["Test_TimeOut"];

                    try {
                        // step 2. Execute Dll
                        ProcessLog("<<Step 2>> Executing "+ITEMDOWNLOAD+JobList);
                        ExecuteDll(ITEMDOWNLOAD+JobList);
                    }
                    catch (Exception ex) {
                        ProcessLog("Run test Error!!! " + ex.Message);
                    }
                    // step 3. update test status to DB
                    ProcessLog("<<Step 3>>  update test status to DB ");
                    DBupdateStatus();   //Update test result

                    // step 5. Job_List的PowerShell程式都完成，繼續Listening job status
                    ProcessLog("<<Step 4>>  Keep listening job status");
                    endTime = DateTime.Now;
                    timeSpan = endTime - startTime;
                    ProcessLog("Spend " + timeSpan.Minutes + " Minutes " + timeSpan.Seconds + " Senconds");

                    // step 5. upload log to FTP
                    ProcessLog("<<Step 5>>  upload log to FTP");
                    ProcessLog("=================Completed================");

                    FTPupload();

                } while (!(Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Escape));
            }  
            // Test DLL only          
            else {
                CreateDirectoryAndFile();
                if( args.Length > 1 ) {
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
                catch (Exception ex) {
                    ProcessLog("Run test Error!!! " + ex.Message);
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
            var r = method.Invoke(obj, args);
            if(r.ToString() == "True") 
                return true;
            else 
                return false;
        }
    }    
}
