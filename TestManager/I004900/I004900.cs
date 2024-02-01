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
using WMPLib; // 引用 Windows Media Player COM 库

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

        private static void Wmp_PlayStateChange(int NewState)
        {
            WMPPlayState playState = (WMPPlayState)NewState;

            switch (playState)
            {
                case WMPPlayState.wmppsPlaying:
                    Console.WriteLine("Playing....");
                    break;
                case WMPPlayState.wmppsPaused:
                    Console.WriteLine("Pause");
                    break;
                case WMPPlayState.wmppsStopped:
                    Console.WriteLine("Stop");
                    break;
                default:
                    Console.WriteLine("Other status: " + playState.ToString());
                    break;
            }
        }

        // ******* New Thread to monitor TimeOut *********
        static void PlayWMP() {
            // Media Player 
            WindowsMediaPlayer wmp = new WindowsMediaPlayer();
            // wmp.URL = @"c:\TestManager\ItemDownload\Soul.mp3";
            wmp.URL = @"c:\TestManager\ItemDownload\n.mp4";
            wmp.PlayStateChange += new _WMPOCXEvents_PlayStateChangeEventHandler(Wmp_PlayStateChange);
            while (ItemWatch.IsRunning) {
                // wmp.controls.pause();
                wmp.controls.play();
                Thread.Sleep(2000);
                Console.WriteLine($"{ItemWatch.Elapsed.TotalSeconds} sec");
            };
            wmp.controls.stop();
        }

//(EdisonLin-20240131-)>>
        static void PlayMP4() {

            string videoPath = @"m:\20240131_090904.mp4";

            // 使用Process.Start啟動Windows Media Player應用程式
            Process playerProcess = new Process();
            playerProcess.StartInfo.FileName = "wmplayer.exe";
            playerProcess.StartInfo.Arguments = videoPath; // 設定視頻路徑作為命令行參數
            playerProcess.Start();
            while (ItemWatch.IsRunning) {
                Thread.Sleep(1000);
                Console.WriteLine($"{ItemWatch.Elapsed.TotalSeconds} sec");
            };

            // 搜尋並關閉Windows Media Player應用程式
            Process[] playerProcesses = Process.GetProcessesByName("wmplayer");

            foreach (Process MyProcess in playerProcesses)
            {
                MyProcess.CloseMainWindow(); // 嘗試使用主窗口關閉
                MyProcess.WaitForExit(); // 等待應用程式退出
                MyProcess.Dispose(); // 釋放資源
            }            

            Console.WriteLine("視頻播放結束。");
        }     
//(EdisonLin-20240131-)<<

        public static void Run()
        {







//             string oldFilePath = @"c:\TestManager\ItemDownload\DeviceStatusCheck.txt";
//             string newFilePath = @"c:\TestManager\ItemDownload\DeviceBefore.txt";
//             if (System.IO.File.Exists(oldFilePath)) {
//                  File.Delete(oldFilePath);
//             }
//             if (System.IO.File.Exists(newFilePath)) {
//                  File.Delete(newFilePath);
//             }

//             bool result = CommonDevicesStatusCheck.CheckDeviceStatus();
//             if (!result) {
//                 TestResult("Fail");
//                 return;
//             }
//             try
//             {
//                 File.Move(oldFilePath, newFilePath);
//             }
//             catch (Exception ex)
//             {
//                 Console.WriteLine($"Error: {ex.Message}");
//             }            


// //(EdisonLin-20240129-)>>





       WindowsMediaPlayer wmp = new WindowsMediaPlayer();

        // 設定視頻檔案的路徑
        wmp.URL = @"m:\20240131_090904.mp4";

        // 開始播放
        wmp.controls.play();

        // 檢查撥放狀態
        while (true)
        {
            Console.WriteLine("撥放狀態: " + wmp.playState.ToString());

            if (wmp.playState == WMPPlayState.wmppsStopped)
            {
                Console.WriteLine("播放結束");
                break;
            }
            else if (wmp.playState == WMPPlayState.wmppsPaused)
            {
                Console.WriteLine("已暫停");
            }

            // 等待一段時間再檢查狀態（可以根據需要調整時間間隔）
            System.Threading.Thread.Sleep(1000);
        }

        // 停止播放
        wmp.controls.stop();















            // // Start Stopwatch
            // ItemWatch = new Stopwatch();
            // ItemWatch.Start();

            // Thread wmpThread = new Thread(PlayMP4);
            // wmpThread.Start();


            // Console.WriteLine($"Wait 5 seconds to enter S4");
            // Thread.Sleep(5000);
            // DoSleep.Sleep(3, 1);

            // Console.ReadKey();

            // // Stop Stopwatch
            // ItemWatch.Stop();
            // Console.ReadKey();


        // GetSystemInfo.GetDiskPartition();
        // GetSystemInfo.MediaType();

//(EdisonLin-20240129-)<<


            // result = CommonDevicesStatusCheck.CheckDeviceStatus();
            // if (!result) {
            //     TestResult("Fail");
            //     return;
            // }

            // try
            // {
            //     string content1 = File.ReadAllText(oldFilePath);
            //     string content2 = File.ReadAllText(newFilePath);

            //     bool areEqual = content1.Equals(content2, StringComparison.OrdinalIgnoreCase);

            //     if (areEqual)
            //     {
            //         result = true;
            //     }
            //     else
            //     {
            //         result = false;
            //     }
            // }
            // catch (Exception ex)
            // {
            //     Console.WriteLine($"Error: {ex.Message}");
            //     result = false;
            // }            

            // if (System.IO.File.Exists(oldFilePath)) {
            //      File.Delete(oldFilePath);
            // }
            // if (System.IO.File.Exists(newFilePath)) {
            //      File.Delete(newFilePath);
            // }


            // if (result) {
            //     TestResult("Pass");
            // } else {
            //     TestResult("Fail");
            // }
        }

        public static void UpdateResults() {
        }

        public static void Setup() {
        }

        public static void TearDown() {
        }
    } //Class1
} //namespace
