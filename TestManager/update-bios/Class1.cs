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

namespace common_bios_pxeboot_default
{
    public class common_bios_pxeboot_default
    {
        /// <summary>
        /// Setup, Empty
        /// </summary>
        public static void Setup() {
            Console.WriteLine("Setup");
        }

        /// <summary>
        /// If the test status is "New", then set the test status to "pxe boot" and reboot the system.
        /// If the test status is "pxe boot", then run the powershell script to check the bios version.
        /// If the bios version is matched, then set the test result to "Pass".
        /// If the bios version is not matched, then set the test result to "Fail".
        /// </summary>
        public static void Run() {
            string currentDirectory1 = @"c:\TestManager\ItemDownload\";
            Console.WriteLine(currentDirectory1);

            string exePath = $"{currentDirectory1}" + "\\" + "Abst64_unsign.exe";
            Console.WriteLine(exePath);

            string arguments = "/password 0 /set \"Network Boot=1\"";
            string arguments1 = "/password 0 /set \"Boot Priority Order=7,2,16,17,255\"";

            string command = "shutdown";
            string arguments2 = "-r -t 0";

            // Start info for boot priority
            ProcessStartInfo startInfo = new ProcessStartInfo {
                FileName = exePath,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = false
            };

            // Start info for network boot
            ProcessStartInfo startInfo1 = new ProcessStartInfo {
                FileName = exePath,
                Arguments = arguments1,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = false
            };

            // Start info for reboot
            ProcessStartInfo startInfo2 = new ProcessStartInfo {
                FileName = command,
                Arguments = arguments2,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = false
            };

            string filePath = @"c:\\TestManager\\TR_Result.json";
            string jsonContent = File.ReadAllText(filePath);
            JObject jsonObject = JObject.Parse(jsonContent);
            string test_status = (string)jsonObject["TestStatus"];
            Console.WriteLine("TestStatus is: " + test_status);

            if (test_status == "New") {
                jsonObject["TestStatus"] = "pxe boot";
                string modifiedJson1 = jsonObject.ToString();
                File.WriteAllText(filePath, modifiedJson1);
                Console.WriteLine("TestStatus is: " + test_status);

                try {
                    // Start the process
                    using (Process process = new Process()) {
                        process.StartInfo = startInfo;
                        process.Start();
                        process.WaitForExit();
                        int exitCode = process.ExitCode;
                    }

                    using (Process process = new Process()) {
                        process.StartInfo = startInfo1;
                        process.Start();
                        process.WaitForExit();
                        int exitCode = process.ExitCode;
                    }
                }
                catch (Exception ex) {
                    Console.WriteLine($"Error running executable: {ex.Message}");
                    jsonObject["TestStatus"] = "ABST Error";
                    jsonObject["TestResult"] = "Fail";
                    modifiedJson1 = jsonObject.ToString();
                    File.WriteAllText(filePath, modifiedJson1);
                }

                try {
                    //reboot
                    using (Process process1 = new Process()) {
                        process1.StartInfo = startInfo2;
                        process1.Start();
                    }
                }
                catch (Exception ex) {
                    Console.WriteLine($"Error running executable: {ex.Message}");
                    jsonObject["TestStatus"] = "Reboot Command Error";
                    jsonObject["TestResult"] = "Fail";
                    modifiedJson1 = jsonObject.ToString();
                    File.WriteAllText(filePath, modifiedJson1);
                }
            }
            else if (test_status == "pxe boot") {

                ProcessStartInfo ps1 = new ProcessStartInfo {
                    FileName = "powershell.exe",
                    Arguments = $"-File Compare_Bios.ps1",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                };

                using (Process process = new Process { StartInfo = ps1 }) {
                    process.Start();
                    string output = process.StandardOutput.ReadToEnd().Trim();
                    bool biosVersionMatched = bool.Parse(output);
                    if (biosVersionMatched) {
                        Console.WriteLine("Bios version matched!");
                        jsonObject["TestResult"] = "Pass";
                    }
                    else {
                        Console.WriteLine("Bios version did not match.");
                        jsonObject["TestResult"] = "Fail";
                    }
                }
                jsonObject["TestStatus"] = "Done";
                string modifiedJson1 = jsonObject.ToString();
                File.WriteAllText(filePath, modifiedJson1);
                Console.WriteLine("TestStatus is: " + test_status);
            }
        }

        /// <summary>
        /// TearDown, Disable network boot
        /// </summary>
        public static void TearDown() 
        {
            string currentDirectory1 = @"c:\TestManager\ItemDownload\";
            string exePath = $"{currentDirectory1}" + "\\" + "Abst64_unsign.exe";
            Console.WriteLine(exePath);
            string arguments = "/password 0 /set \"Network Boot=0\"";

            ProcessStartInfo startInfo = new ProcessStartInfo {
                FileName = exePath,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = false
            };

            try {
                using (Process process1 = new Process()) {
                    process1.StartInfo = startInfo;
                    process1.Start();
                    string output = process1.StandardOutput.ReadToEnd();
                    Console.WriteLine("Output:\n" + output);
                    process1.WaitForExit();
                    int exitCode = process1.ExitCode;
                    Console.WriteLine($"Process exited with code: {exitCode}");
                }
            }
            catch (Exception ex) {
                Console.WriteLine($"Error running executable: {ex.Message}");
            }
            Console.WriteLine("TearDown");
        }
    }
}
