using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;             //手動加入參考
using System.Management.Automation.Runspaces;   //手動加入參考
using System.Threading;
using System.Reflection;
using System.Diagnostics;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TM1002
{
    public class Program
    {
        private static string currentDirectory = Directory.GetCurrentDirectory() + '\\';
        private static string ItemDownload = currentDirectory+"ItemDownload\\";
        private static string log_file = currentDirectory+"MyLog\\TestManager.log";
        static string TR_FilePath = currentDirectory+"TR_Result.json";
        static Stopwatch ItemWatch = new Stopwatch();
        private static int timeout = 9999;

        // **** Check Need Update ****
        static bool Check_Need_Update()
        {
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();
            try
            {
                pipeline.Commands.AddScript(currentDirectory + "Check_Update.ps1");
                var result = pipeline.Invoke();
                runspace.Close();  
                foreach (var psObject in result)
                {
                    if(psObject != null)
                    {
                        psObject.ToString();
                        Console.WriteLine(psObject.ToString()); 
                        if(psObject.ToString() == "True") 
                            return true;
                        else 
                            return false;
                    }    
                    else
                        return false;
                }
            }
            catch
            {
                runspace.Close();
                process_log("Waiting 2 sec for ready");
                Thread.Sleep(2000);
                return false;
            }
            return false;
        }    

        // **** Update TestManager ****
        static void UpdateT()
        {
            // 取得目前程式的進程 ID
            int currentProcessId = Process.GetCurrentProcess().Id;

            // PowerShell 指令
            string scriptCommand = $"Start-Process powershell -ArgumentList '-NoExit -File C:\\TestManager\\UpdateT.ps1' -WindowStyle Hidden; Stop-Process -Id {currentProcessId}";

            // 建立 ProcessStartInfo object，設定要啟動的進程資訊
             ProcessStartInfo psi = new ProcessStartInfo
             {
                 FileName = "powershell.exe", // 指定要啟動的程式（PowerShell）
                 Arguments = $"-Command \"{scriptCommand}\"",
                 RedirectStandardOutput = true,
                 UseShellExecute = false,
                 CreateNoWindow = true
             };

             // 建立 Process object，表示要啟動的進程
             Process process = new Process
             {
                 StartInfo = psi
             };

             // 啟動進程
             process.Start();

             // 等待 PowerShell 執行完畢
             process.WaitForExit();

             // 讀取 PowerShell 輸出
             string output = process.StandardOutput.ReadToEnd();
             Console.WriteLine(output);

             // 關閉行程
             process.Close();

             // 關閉目前程式
             Environment.Exit(0);       

        }

        // **** 創建log file ****
        static void CreateDirectoryAndFile()
        {
            try
            {
                // 檢查目錄是否存在，如果不存在則建立
                if (!Directory.Exists(ItemDownload))
                {
                    Directory.CreateDirectory(ItemDownload);
                }                
                if (!Directory.Exists(currentDirectory+"MyLog\\"))
                {
                    Directory.CreateDirectory(currentDirectory+"MyLog\\");
                }                
                if (!Directory.Exists(currentDirectory+"TestLog\\"))
                {
                    Directory.CreateDirectory(currentDirectory+"TestLog\\");
                }                

                // 檢查檔案是否存在，如果不存在則建立，檔案存在內容就清空
                if (!File.Exists(log_file))
                {
                    using (FileStream fs = File.Create(log_file));
                }
                else
                {
                    // 清空內容
                    // using (FileStream fs = new FileStream(log_file, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    // {
                    //     fs.SetLength(0);
                    // }                    
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error!!! " + ex.Message);
            }
        }

        // **** TestManager.log ****
        public static void process_log(string content)
        {
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
                Console.WriteLine("Error!!! " + ex.Message);
            }
        }      
        
        // ***** get jobs from DB *****
        static string Get_PXE()
        {
            string job_list = null;
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();
            try
            {
                pipeline.Commands.AddScript(currentDirectory + "RunAs.ps1");
                pipeline.Commands.AddScript(currentDirectory + "Get_PXE.ps1");
                var result = pipeline.Invoke();

                foreach (var psObject in result)
                {
                    if(psObject != null)
                        job_list = psObject.ToString();
                    else
                        job_list = null;
                }
                runspace.Close();                
            }
            catch
            {
                runspace.Close();
                process_log("Waiting 2 sec for Get_PXE ready");
                Thread.Sleep(2000);
                return null;
            }
            return job_list;
        }

        // ***** get jobs from DB *****
        static string Get_Job()
        {
            string job_list = "";
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();
            try
            {
                pipeline.Commands.AddScript(currentDirectory+"RunAs.ps1");
                pipeline.Commands.AddScript(currentDirectory+"Get_Job.ps1");
                var result = pipeline.Invoke();
                runspace.Close();
                foreach (var psObject in result)
                {
                    if(psObject != null)
                        job_list = psObject.ToString();
                    else
                        job_list = null;
                }
            }    
            catch
            {
                runspace.Close();
                process_log("Waiting 2 sec for Get_JOB ready");
                Thread.Sleep(2000);
                return null;
            }
            return job_list;
        }

        // ***** FTP_Download *****
        static void FTP_Download(string job_list)
        {
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();

            try
            {
                pipeline.Commands.AddScript(currentDirectory+"RunAs.ps1");
                pipeline.Commands.AddScript("$remoteFile = \""+job_list+"\"");
                pipeline.Commands.AddScript(currentDirectory+"Download.ps1");
                var result = pipeline.Invoke();
            }    
            catch (Exception ex)
            {
                process_log("Error!!! Downloading "+ex.Message);
                runspace.Close();
                return;
            }

            runspace.Close();
            return;
        }
        // ***** upload a program from FTP *****
        static void FTP_Upload()
        {
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();

            try
            {
                pipeline.Commands.AddScript(currentDirectory+"Upload.ps1 ");
                var result = pipeline.Invoke();
            }    
            catch (Exception ex)
            {
                process_log("Upload "+ex.Message);
                runspace.Close();
                return;
            }

            runspace.Close();
            return;
        }

        // ***** update job_status to DB *****
        static bool Update_Job_Status()
        {
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();

            try
            {
                pipeline.Commands.AddScript(currentDirectory+"RunAs.ps1");
                pipeline.Commands.AddScript(currentDirectory+"Update_Job_Status.ps1");
                var result = pipeline.Invoke();
                if(result[0].ToString() == "Unconnected_")
                {
                    runspace.Close();
                    return false;
                }

            }    
            catch (Exception ex)
            {
                process_log("Update "+ex.Message);
                runspace.Close();
                return false;
            }

            runspace.Close();
            return true;
        }

        // ***** Execute_dll *****
        static bool Execute_dll(string dllPath)
        {
            string callingDomainName = AppDomain.CurrentDomain.FriendlyName;//Thread.GetDomain().FriendlyName;
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            AppDomain ad = AppDomain.CreateDomain("TestManager DLL");
            ProxyObject obj = (ProxyObject)ad.CreateInstanceFromAndUnwrap(basePath+callingDomainName, "TM1002.ProxyObject");
            try
            {
                process_log(".... Loading Common.dll ....");
                obj.LoadAssembly(currentDirectory+"Common.dll");
            }
            catch (System.IO.FileNotFoundException)
            {
                process_log("!!! 找不到 Common.dll");
                return false;
            }

            // 啟動計時器
            ItemWatch = new Stopwatch();
            ItemWatch.Start();

            process_log(".... Loading "+dllPath+" ....");
            Object[] p = new object[]{ dllPath, new object[]{}, new object[]{}, new object[]{}, new object[]{} };
            var result = obj.Invoke("RunTestItem",p);

            // 停止計時器
            ItemWatch.Stop();
			
            AppDomain.Unload(ad);
            obj = null;
            if(result.ToString() == "True") return true;
            else return false;
        }

        // ******* New Thread to monitor TimeOut *********
        static void MonitorExecutionTime()
        {
            bool NewWatch = true;
            do
            {
                Thread.Sleep(1000);
                // Console.WriteLine($"{ItemWatch.Elapsed.TotalSeconds} .... {NewWatch}");

                if( NewWatch )
                {
                    if(ItemWatch.Elapsed.TotalSeconds >= timeout)
                    {
                        Console.WriteLine($"Time Out!!!!: {ItemWatch.Elapsed.TotalSeconds} seconds");
                        NewWatch = false;
                    }
                }
                else
                {
                    if(ItemWatch.Elapsed.TotalSeconds < timeout)
                    {
                        NewWatch = true;
                    }                    
                }
            } while (!(Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Escape));
            
        }


        // ============== MAIN ==============

        static void Main(string[] args)
        {
            if( Check_Need_Update() )
            {
                Console.WriteLine("Need update!!");
                UpdateT();
            }    


//(EdisonLin-20231205-)>>
            // Console.WriteLine("**** Exit ****"); 
            // Console.ReadKey();
            // Environment.Exit(0);            
//(EdisonLin-20231205-)<<

            string Job_List;
            DateTime startTime, endTime;
            TimeSpan timeSpan;
            bool result = true;
            CreateDirectoryAndFile();
            // 啟動一個執行緒來監測主程式的執行時間
            Thread monitoringThread = new Thread(MonitorExecutionTime);
            monitoringThread.Start();
            do
            {
                // step 1. Listening job status from DB
                Job_List = Get_PXE();
                if(Job_List == null)
                {
                    Job_List = Get_Job();
                }
                if(Job_List == null)
                {
                    Thread.Sleep(2000);
                    continue;
                } 
                else if (Job_List == "Unconnected_")
                {
                    process_log("Waiting 5 sec for DB connected !!!");
                    Thread.Sleep(5000);
                    continue;
                }    

                // step 1. Got Job then downloading
                process_log("<<Step 1>> Got Job then downloading");
                if (!File.Exists(ItemDownload+"DoneDll.txt"))
                    FTP_Download(Job_List);
                else    
                    process_log("Skip downloading again for reboot");

                startTime = DateTime.Now;

                // 讀取 TR_Result.json timeout 内容
                string jsonString = System.IO.File.ReadAllText(TR_FilePath);
                JObject json = JObject.Parse(jsonString);
                timeout = (int)json["Test_TimeOut"];

                try
                {
                    // step 2. 執行Dll程式
                    process_log("<<Step 2>> Executing "+ItemDownload+Job_List);
                    result = Execute_dll(ItemDownload+Job_List);
                }
                catch (Exception ex)
                {
                    process_log("Run test Error!!! " + ex.Message);
                }
                // step 3. update test status to DB
                process_log("<<Step 3>>  update test status to DB ");
                Update_Job_Status();   //Update test result

                // step 4. upload log to FTP
                process_log("<<Step 4>>  upload log to FTP:  "+Job_List);
                FTP_Upload();

                // step 5. Job_List的PowerShell程式都完成，繼續Listening job status
                process_log("<<Step 5>>  Keep listening job status");
                endTime = DateTime.Now;
                timeSpan = endTime - startTime;
                // 输出时间间隔
                process_log("執行花費時間: " + timeSpan.Minutes + "分鐘 " + timeSpan.Seconds + "秒");
                process_log("=================Completed================");
                // Thread.Sleep(1000);
            } while (!(Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Escape));
            
            // Close window
            process_log("**** Exit ****");            
            Environment.Exit(0);            
        }
    }

    class ProxyObject : MarshalByRefObject
    {
        Assembly assembly = null;
        public void LoadAssembly(string myDllPath)
        {
            assembly = Assembly.LoadFile(myDllPath);
        }
        public bool Invoke(string methodName, params Object[] args)
        {
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
