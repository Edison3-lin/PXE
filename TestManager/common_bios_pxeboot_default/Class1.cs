using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;
//using System.Text.Json;
using System.Xml.Linq;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading;

namespace common_bios_pxeboot_default
{
    public class common_bios_pxeboot_default
    {

        // public static void Setup() 
        // {
        //     Console.WriteLine("Setup");
        // }
        public static void Run()
        {
            //string currentDirectory1 = Environment.CurrentDirectory;
            string currentDirectory1 = @"c:\TestManager\ItemDownload\";
            Console.WriteLine(currentDirectory1);
            // Specify the path to the executable you want to run
            string exePath = $"{currentDirectory1}" + "\\" + "Abst64_unsign.exe";
            Console.WriteLine(exePath);
            // Specify the arguments to pass to the executable
            string arguments = "/password 0 /set \"Boot Priority Order=7,2,16,17,255\"";
            string arguments1 = "/password 0 /set \"Network Boot = 1\"";
            string arguments2 = "/password 0 /set \"Network Boot = 0\"";

            string command = "shutdown";
            string arguments3 = "-r -t 0";

            // Create a new process start info
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = arguments,
                UseShellExecute = false, // Required when redirecting output
                RedirectStandardOutput = true,
                CreateNoWindow = false
            };

            ProcessStartInfo startInfo1 = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = arguments1,
                UseShellExecute = false, // Required when redirecting output
                RedirectStandardOutput = true,
                CreateNoWindow = false
            };

            ProcessStartInfo startInfo2 = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = arguments1,
                UseShellExecute = false, // Required when redirecting output
                RedirectStandardOutput = true,
                CreateNoWindow = false
            };

            // Create a new process start info
            ProcessStartInfo startInfo3 = new ProcessStartInfo
            {
                FileName = command,
                Arguments = arguments3,
                UseShellExecute = false, // Required when redirecting output
                RedirectStandardOutput = true,
                CreateNoWindow = false
            };

            string filePath = @"c:\\TestManager\\TR_Result.json"; // 將路徑替換為你的JSON文件的實際路徑
            // 讀取JSON文件內容
            string jsonContent = File.ReadAllText(filePath);
            // 將JSON字串解析為JObject
            JObject jsonObject = JObject.Parse(jsonContent);
            // 讀取"TestStatus"的值
            string test_status = (string)jsonObject["TestStatus"];
            Console.WriteLine("TestStatus is: " + test_status);
            if (test_status == "New")//new, first to process pxe boot
            {
                try
                {
                    // Start the process
                    using (Process process = new Process())
                    {
                        process.StartInfo = startInfo;
                        process.Start();

                        // Optionally, read the standard output
                        string output = process.StandardOutput.ReadToEnd();
                        Console.WriteLine("Output:\n" + output);
                        Console.WriteLine("---------------End of Output----------------");
                        if (output.IndexOf("Get BIOS options success") > 0) //success to pxe boot, status ----> pxe boot
                        {
                            Console.WriteLine("修改 \"TestStatus\" 內容");
                            // 讀取JSON文件內容
                            string jsonContent1 = File.ReadAllText(filePath);
                            // 將JSON字串解析為JObject
                            JObject jsonObject1 = JObject.Parse(jsonContent1);
                            // 修改 "site" 內容
                            jsonObject1["TestStatus"] = "pxe boot"; // 在這裡將新的值賦給 "site" 屬性
                                                                // 將修改後的JObject轉換回JSON字符串
                            string modifiedJson1 = jsonObject1.ToString();
                            // 將修改後的JSON字串保存回文件
                            File.WriteAllText(filePath, modifiedJson1);
                            Console.WriteLine("TestStatus is: " + test_status);
                        }
                        else // fail to pxe boot , status --> Done
                        {
                            Console.WriteLine("修改 \"TestStatus\" 內容");
                            // 讀取JSON文件內容
                            string jsonContent1 = File.ReadAllText(filePath);
                            // 將JSON字串解析為JObject
                            JObject jsonObject1 = JObject.Parse(jsonContent1);
                            // 修改 "site" 內容
                            jsonObject1["TestStatus"] = "Done"; // 在這裡將新的值賦給 "site" 屬性
                                                                    // 將修改後的JObject轉換回JSON字符串
                            jsonObject1["TestResult"] = "Fail"; // 在這裡將新的值賦給 "site" 屬性
                                                                // 將修改後的JObject轉換回JSON字符串
                            string modifiedJson1 = jsonObject1.ToString();
                            // 將修改後的JSON字串保存回文件
                            File.WriteAllText(filePath, modifiedJson1);
                            Console.WriteLine("TestStatus is: " + test_status);
                        }
                        // Optionally, wait for the process to exit
                        process.WaitForExit();

                        // Optionally, retrieve the exit code
                        int exitCode = process.ExitCode;
                        Console.WriteLine($"Process exited with code: {exitCode}");
                    }
                    
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error running executable: {ex.Message}");

                }

                try
                {
                    // Start the process
                    using (Process process1 = new Process())
                    {
                        process1.StartInfo = startInfo3;
                        process1.Start();
                        while (true)
                        {
                            //loop here 
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error running executable: {ex.Message}");
                }
            }
            else if (test_status == "pxe boot")
            {
                Console.WriteLine("修改 \"TestStatus\" 內容");
                // 讀取JSON文件內容
                string jsonContent1 = File.ReadAllText(filePath);
                // 將JSON字串解析為JObject
                JObject jsonObject1 = JObject.Parse(jsonContent1);

                ProcessStartInfo ps1 = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-File Compare_Bios.ps1",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                };

                using (Process process = new Process { StartInfo = ps1 })
                {
                    process.Start();
                    string output = process.StandardOutput.ReadToEnd().Trim();

                    // Output will be "True" or "False"
                    bool biosVersionMatched = bool.Parse(output);

                    // Use the value as needed
                    if (biosVersionMatched)
                    {
                        Console.WriteLine("Bios version matched!");
                        jsonObject1["TestResult"] = "Pass"; // 在這裡將新的值賦給 "site" 屬性
                                                            // 將修改後的JObject轉換回JSON字符串
                    }
                    else
                    {
                        Console.WriteLine("Bios version did not match.");
                        jsonObject1["TestResult"] = "Fail"; // 在這裡將新的值賦給 "site" 屬性
                                                            // 將修改後的JObject轉換回JSON字符串
                    }
                }

                string modifiedJson1 = jsonObject1.ToString();
                // 將修改後的JSON字串保存回文件
                File.WriteAllText(filePath, modifiedJson1);
                Console.WriteLine("TestStatus is: " + test_status);
            }
        }


        public static void Setup() 
        {
            Console.WriteLine("Setup");
        }

        public static void UpdateResults() 
        {
            Console.WriteLine("UpdateResults");
        }
        public static void TearDown() 
        {
            //Disable Network boot to default
            string currentDirectory1 = @"c:\TestManager\ItemDownload\";
            string exePath = $"{currentDirectory1}" + "\\" + "Abst64_unsign.exe";
            Console.WriteLine(exePath);
            // Specify the arguments to pass to the executable
            string arguments = "/password 0 /set \"Network Boot = 0\"";

            // Create a new process start info
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = arguments,
                UseShellExecute = false, // Required when redirecting output
                RedirectStandardOutput = true,
                CreateNoWindow = false
            };
            try
            {
                // Start the process
                using (Process process1 = new Process())
                {
                    process1.StartInfo = startInfo;
                    process1.Start();
                    // Optionally, read the standard output
                    string output = process1.StandardOutput.ReadToEnd();
                    Console.WriteLine("Output:\n" + output);

                    // Optionally, wait for the process to exit
                    process1.WaitForExit();

                    // Optionally, retrieve the exit code
                    int exitCode = process1.ExitCode;
                    Console.WriteLine($"Process exited with code: {exitCode}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error running executable: {ex.Message}");
            }
            Console.WriteLine("TearDown");
        }
    }
}
