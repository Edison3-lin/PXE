/*
* CaptainWin.Common - Common API for test items
* TestOperation.cs - Common test operations for test items
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Bencool Lin <Bencool.Lin@quantatw.com>
*  Edison Lin  <Edison.Lin@quantatw.com>
*  Jimmy Chen  <Jimmychen3@quantatw.com>
*  Jacky Kao   <Jacky.Kao@quantatw.com>
*/

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace CaptainWin.CommonAPI {

    /// <summary>
    /// This class contains common test operations for test items
    /// </summary>
    public class TestOperation {
        /// <summary>
        /// Check if the current process is running as administrator 
        /// </summary>
        /// <returns>True if the current process is running as administrator</returns>
        public bool IsProcessAdmin() {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        /// <summary>
        /// Check if the current user is administrator
        /// </summary>
        /// <returns>True if the current user is administrator</returns>
        public bool IsUserAdmin() {
            WindowsIdentity user = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(user);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        /// <summary>
        /// Run a file as administrator. Can run batch, cmd, exe...etc.
        /// </summary>
        /// <param name="args">Path of the file to run and its arguments</param>
        /// <returns>Output of the file</returns>
        public string Run(string[] args) {
            try {
                ProcessStartInfo processInfo = new ProcessStartInfo(args[0]) {
                    FileName = args[1],
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    Verb = "RunAs"
                };
                if (args.Length > 2) {
                    processInfo.Arguments = string.Join(" ", args.Skip(2).Select(arg => $"\"{arg}\""));
                }
                int loop = int.Parse(args[0]);
                string output = "";
                for (int i = 0; i < loop; i++) {
                    using (Process process = new Process { StartInfo = processInfo }) {
                        process.Start();
                        output = string.Join("\n", process.StandardOutput.ReadToEnd());
                    }
                }
                return output;
            }
            catch (Exception ex) {
                return ex.Message;
            }
        }

        /// <summary>
        /// Run a file as administrator. Can run batch, cmd, exe...etc.
        /// And wait for the process to exit.
        /// </summary>
        /// <param name="args">Path of the file to run and its arguments</param>
        /// <returns>Output of the file</returns>
        public string RunWait(string[] args) {
            foreach (string arg in args) {
                Console.WriteLine(arg);
            }
            try {
                ProcessStartInfo processInfo = new ProcessStartInfo(args[0]) {
                    FileName = args[1],
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    Verb = "RunAs"
                };
                if (args.Length > 2) {
                    processInfo.Arguments = string.Join(" ", args.Skip(2).Select(arg => $"\"{arg}\""));
                }
                int loop = int.Parse(args[0]);
                string output = "";
                for (int i = 0; i < loop; i++) {
                    using (Process process = new Process { StartInfo = processInfo }) {
                        process.Start();
                        process.WaitForExit();
                        output = string.Join("\n", process.StandardOutput.ReadToEnd());
                    }
                }
                return output;
            }
            catch (Exception ex) {
                return ex.Message;
            }
        }

        /// <summary>
        /// Run a powershell script as administrator.
        /// </summary>
        /// <param name="args">Path of the powershell script to run and its arguments</param>
        /// <returns>Output of the powershell script</returns>
        public string RunPS1(string[] args) {
            try {
                string filePath = args[1];
                string arguments = $"-File \"{filePath}\"";
                if (args.Length > 2) {
                    arguments = string.Join(" ", args.Skip(2).Select(arg => $"\"{arg}\""));
                }
                ProcessStartInfo processInfo = new ProcessStartInfo(args[0]) {
                    FileName = "powershell.exe",
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    Verb = "RunAs"
                };

                int loop = int.Parse(args[0]);
                string output = "";
                for (int i = 0; i < loop; i++) {
                    using (Process process = new Process { StartInfo = processInfo }) {
                        process.Start();
                        output = string.Join("\n", process.StandardOutput.ReadToEnd());
                    }
                }
                return output;
            }
            catch (Exception ex) {
                return ex.Message;
            }
        }

        /// <summary>
        /// Run a powershell script as administrator.
        /// And wait for the process to exit.
        /// </summary>
        /// <param name="args">Path of the powershell script to run and its arguments</param>
        /// <returns>Output of the powershell script</returns>
        public string RunPS1Wait(string[] args) {
            try {
                string filePath = args[1];
                string arguments = $"-File \"{filePath}\"";
                if (args.Length > 2) {
                    arguments = string.Join(" ", args.Skip(2).Select(arg => $"\"{arg}\""));
                }
                ProcessStartInfo processInfo = new ProcessStartInfo(args[0]) {
                    FileName = "powershell.exe",
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    Verb = "RunAs"
                };

                int loop = int.Parse(args[0]);
                string output = "";
                for (int i = 0; i < loop; i++) {
                    using (Process process = new Process { StartInfo = processInfo }) {
                        process.Start();
                        process.WaitForExit();
                        output = string.Join("\n", process.StandardOutput.ReadToEnd());
                    }
                }
                return output;
            }
            catch (Exception ex) {
                return ex.Message;
            }
        }

        /// <summary>
        /// Reboot the system
        /// </summary>
        /// <param name="sec">Seconds to wait before reboot</param>
        public void Reboot (int sec) {
            string shutdownCommand = $"shutdown /r /t {sec}";
            // Set up the process start info
            ProcessStartInfo processStartInfo = new ProcessStartInfo {
                FileName = "cmd.exe",
                RedirectStandardInput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            Process process = new Process { StartInfo = processStartInfo };
            process.Start();
            process.StandardInput.WriteLine(shutdownCommand);
            process.StandardInput.Flush();
            process.StandardInput.Close();
            process.WaitForExit();
        }
    }
}
