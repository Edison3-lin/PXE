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

        static void getProcess(int ProcessId) {
            try
            {
                // 通过进程ID获取进程对象
                Process targetProcess = Process.GetProcessById(ProcessId);

                if (targetProcess != null)
                {
                    // 获取进程名称
                    string processName = targetProcess.ProcessName;
                    Console.WriteLine($"进程ID {ProcessId} 的进程名称为: {processName}");
                }
                else
                {
                    Console.WriteLine($"未找到ID为 {ProcessId} 的进程。");
                }
            }
            catch (ArgumentException)
            {
                Console.WriteLine($"ID为 {ProcessId} 的进程不存在。");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"发生错误: {ex.Message}");
            }
        }


        public static void processById(int processId) {
            // 假设要获取进程ID为12345的进程的信息
            // int processId = 12345;

            try
            {
                Process process = Process.GetProcessById(processId);

                Console.WriteLine("进程ID: " + process.Id);
                Console.WriteLine("进程名称: " + process.ProcessName);
                Console.WriteLine("启动时间: " + process.StartTime);
                Console.WriteLine("基本优先级: " + process.BasePriority);
                Console.WriteLine("总处理时间: " + process.TotalProcessorTime);
                Console.WriteLine("是否正在响应: " + process.Responding);

                // 可以根据需要获取更多信息
                // 例如，使用 process.Modules 获取进程的模块信息等

                process.Close(); // 记得关闭进程对象
            }
            catch (ArgumentException)
            {
                Console.WriteLine("找不到具有指定ID的进程。");
            }
            catch (InvalidOperationException)
            {
                Console.WriteLine("指定ID的进程无法访问。");
            }
        }

        static bool IsProcessRunning(string processName)
        {
            Process[] processes = Process.GetProcessesByName(processName);
            return processes.Length > 0;
        }

static bool IsMusicPlaying(WindowsMediaPlayer wmp)
    {
        try
        {
            // Console.WriteLine(a.ToString());
            return wmp.playState == WMPPlayState.wmppsPlaying;
        }
        catch (Exception)
        {
            // 如果 Windows Media Player 未安装或不可用，可能会引发异常
            Console.WriteLine("dddddddddddd");
            return false;
        }
    }
        public static void playM() {
                string mediaFilePath = @"c:\Users\edison\Downloads\Soul.mp3";
                if (System.IO.File.Exists(mediaFilePath))
                {
                    try
                    {
                        Process musicPlayerProcess = Process.Start("wmplayer.exe", mediaFilePath);
                        if (musicPlayerProcess != null)
                        {
                            Console.WriteLine($"Process ID: {musicPlayerProcess.Id}");
                            Console.WriteLine($"Process Name: {musicPlayerProcess.ProcessName}");
                            Console.WriteLine($"Start time: {musicPlayerProcess.StartTime}");
                            Console.WriteLine($"Responding: {musicPlayerProcess.Responding}");
                        }

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Can't start Media player：" + ex.Message);
                    }
                }
                else
                {
                    Console.WriteLine("The specified file doesn't exist：" + mediaFilePath);
                }
        }

        public static void Run()
        {
            string oldFilePath = @"c:\TestManager\ItemDownload\DeviceStatusCheck.txt";
            string newFilePath = @"c:\TestManager\ItemDownload\DeviceBefore.txt";
            if (System.IO.File.Exists(oldFilePath)) {
                 File.Delete(oldFilePath);
            }
            if (System.IO.File.Exists(newFilePath)) {
                 File.Delete(newFilePath);
            }

            bool result = CommonDevicesStatusCheck.CheckDeviceStatus();
            if (!result) {
                TestResult("Fail");
                return;
            }
            try
            {
                File.Move(oldFilePath, newFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }            



//(EdisonLin-20240129-)>>
 // 启动 Windows Media Player 进程
        Process wmpProcess = new Process();
        wmpProcess.StartInfo.FileName = "wmplayer.exe"; // Windows Media Player 的可执行文件路径
        wmpProcess.Start();

        // 等待一段时间以确保 WMP 启动
        System.Threading.Thread.Sleep(5000); // 例如，等待5秒钟

        // 创建 Windows Media Player COM 对象
        WindowsMediaPlayer wmp = new WindowsMediaPlayer();

        bool musicIsPlaying = true;

        while (musicIsPlaying)
        {
            WMPPlayState playState = wmp.playState;

            switch (playState)
            {
                case WMPPlayState.wmppsPlaying:
                    Console.WriteLine("音乐正在播放。");
                    break;
                case WMPPlayState.wmppsPaused:
                    Console.WriteLine("音乐已暂停。");
                    break;
                case WMPPlayState.wmppsStopped:
                    Console.WriteLine("音乐已停止。");
                    musicIsPlaying = false; // 停止循环
                    break;
                default:
                    Console.WriteLine("音乐处于其他状态。");
                    break;
            }

            // 等待一段时间再进行下一次状态检查
            System.Threading.Thread.Sleep(1000); // 例如，等待1秒钟
        }

//(EdisonLin-20240129-)<<





        // if (IsMusicPlaying(wmp))
        // {
        //     Console.WriteLine("音乐正在播放。");
        //     // 在这里执行相关操作
        // }
        // else
        // {
        //     Console.WriteLine("音乐未在播放。");
        // }
            // playM();
            // processById(11176);
            // DoSleep.Sleep(4, 1);

            result = CommonDevicesStatusCheck.CheckDeviceStatus();
            if (!result) {
                TestResult("Fail");
                return;
            }

            try
            {
                string content1 = File.ReadAllText(oldFilePath);
                string content2 = File.ReadAllText(newFilePath);

                bool areEqual = content1.Equals(content2, StringComparison.OrdinalIgnoreCase);

                if (areEqual)
                {
                    result = true;
                }
                else
                {
                    result = false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                result = false;
            }            

            if (System.IO.File.Exists(oldFilePath)) {
                 File.Delete(oldFilePath);
            }
            if (System.IO.File.Exists(newFilePath)) {
                 File.Delete(newFilePath);
            }


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
