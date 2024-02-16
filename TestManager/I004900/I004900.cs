using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;
using System.Xml.Linq;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading;
using CaptainWin.CommonAPI;
using Microsoft.Win32;

using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;

namespace I004900 {
    public class MyI004900 {
        private const string TR = "C:\\TestManager\\TR_Result.json";
        static Stopwatch ItemWatch = new Stopwatch();

        public static void TestResult(string TestResult) {
            try {
                string ftpJson = System.IO.File.ReadAllText(TR);
                JObject fjson = JObject.Parse(ftpJson);
                fjson["TestResult"] = TestResult;
                string updatedJson = fjson.ToString();
                System.IO.File.WriteAllText(TR, updatedJson);
            }
            catch (Exception ex) {
                Console.WriteLine($"Write TR.json error occurred: {ex.Message}");
            }
        }
        static void PlayMP4() {

            string videoPath = @"c:\TestManager\ItemDownload\20240131_090904.mp4";

            // 使用Process.Start啟動Windows Media Player應用程式
            Process playerProcess = new Process();
            playerProcess.StartInfo.FileName = "wmplayer.exe";
            playerProcess.StartInfo.Arguments = videoPath;
            playerProcess.Start();
            while (ItemWatch.IsRunning) {
                Thread.Sleep(1000);
                Console.WriteLine($"{ItemWatch.Elapsed.TotalSeconds} sec");
            };

            // 搜尋並關閉Windows Media Player應用程式
            Process[] playerProcesses = Process.GetProcessesByName("wmplayer");

            foreach (Process MyProcess in playerProcesses)
            {
                MyProcess.CloseMainWindow();
                MyProcess.WaitForExit();
                MyProcess.Dispose();
            }            

            Console.WriteLine("-- Video playback ends");
        }     


        public static void setPreferences() {
            try
            {
                // 设置.reg文件的路径
                string regFilePath = "c:\\TestManager\\ItemDownload\\Preferences.reg";

                // 创建一个ProcessStartInfo对象
                ProcessStartInfo processStartInfo = new ProcessStartInfo
                {
                    FileName = "regedit.exe",
                    Arguments = $"/s \"{regFilePath}\"", // /s 参数使regedit在无提示模式下运行
                    UseShellExecute = true,
                    Verb = "runas", // 请求管理员权限
                    CreateNoWindow = true, // 不创建新窗口
                };

                // 启动进程
                Process process = Process.Start(processStartInfo);

                // 等待进程结束
                process.WaitForExit();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            bool result = CommonDevicesStatusCheck.CheckDeviceStatus();
            File.Move(@"c:\TestManager\ItemDownload\DeviceStatusCheck.txt", @"c:\TestManager\ItemDownload\DeviceBefore.txt");

            if (!result) {
                TestResult("Fail");
                return;
            }


//(EdisonLin-20240215-)>>
            string keyPath = @"Software\\Microsoft\\MediaPlayer\\Preferences";
            string valueName = "AcceptedPrivacyStatement";
            bool Preferences = false;
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath))
            {
                if (key != null)
                {
                    object value = key.GetValue(valueName);
                    // 如果value不为null，则表示找到了值
                    Preferences = (value != null);
                }
                if (!Preferences) {
                    setPreferences();           
                }
            }
//(EdisonLin-20240215-)<<

            // Start Stopwatch
            ItemWatch = new Stopwatch();
            ItemWatch.Start();

            Thread wmpThread = new Thread(PlayMP4);
            wmpThread.Start();


            Console.WriteLine($"Wait 5 seconds to enter S4");
            Thread.Sleep(5000);
            // DoSleep.Sleep(4, 1);

            Thread.Sleep(10000);
            ItemWatch.Stop();
            Thread.Sleep(2000);


//(EdisonLin-20240215-)>>
        if( !Preferences ) {
            // 打开指定的注册表键，需要写权限
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath, writable: true))
            {
                if (key == null)
                {
                    Console.WriteLine($"Key not found: {keyPath}");
                    return;
                }

                try
                {
                    // 删除指定的值
                    key.DeleteValue(valueName);
                }
                catch (ArgumentException)
                {
                    Console.WriteLine($"Value '{valueName}' does not exist.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                }
            }
        }    
//(EdisonLin-20240215-)<<

            result = CommonDevicesStatusCheck.CheckDeviceStatus();
            if (!result) {
                TestResult("Fail");
                return;
            }

            File.Move(@"c:\TestManager\ItemDownload\DeviceStatusCheck.txt", @"c:\TestManager\ItemDownload\DeviceAfter.txt");

//(EdisonLin-20240215-)>>
            // 文件路徑
            string filePath1 = @"c:\TestManager\ItemDownload\DeviceBefore.txt";
            string filePath2 = @"c:\TestManager\ItemDownload\DeviceAfter.txt";

            string[] file1Lines = File.ReadAllLines(filePath1);
            string[] file2Lines = File.ReadAllLines(filePath2);

            // 比較檔案行數
            result = true;
            if (file1Lines.Length != file2Lines.Length)
            {
                // 文件行數不同，因此文件內容不同。
                result = false;
            }
            else
            {
                // 逐行比較
                for (int i = 0; i < file1Lines.Length; i++)
                {
                    if (file1Lines[i] != file2Lines[i])
                    {
                        result = false;
                    }
                }
            }

            File.Delete(filePath1);
            File.Delete(filePath2);            
//(EdisonLin-20240215-)<<

            if (result) {
                TestResult("Pass");
            } else {
                TestResult("Fail");
            }
        }

        public static void UpdateResults() {
        }

        public static void Setup() {
        }

        public static void TearDown() {
        }
    } //Class1
} //namespace
