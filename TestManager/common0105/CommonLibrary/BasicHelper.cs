/*
* BasicHelper.cs 
* Collection for some common functions for other common functions to use.
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Bencool   <Bencool.lin@quantatw.com>
*/

//using IWshRuntimeLibrary;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CaptainWin.CommonAPI
{
    /// <summary>
    ///  for some common functions that for multiple helper, modules used.
    /// </summary>
    public static  class BasicHelper
    {
        /// <summary>
        /// Function for list the .lnk files for the folder path pass in 
        /// </summary>
        /// <returns>List of string, Each String combine of {fileName}, {fileSize} bytes, {programName}</returns>
        public static List<string> ListShortcut(string filePath)
        {
           // string taskbarPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), _g_USER_QUICK_LAUNCH_PATH);

            List<string> pinnedPrograms = new List<string>();

            foreach (string lnkFile in Directory.GetFiles(filePath, "*.lnk"))
            {
                // Extract information from the shortcut
                var programInfo = ShortcutHelper.GetShortcutInfo(lnkFile);
               // var programInfo = GetProgramInfoFromShortcut(lnkFile);
                if (!string.IsNullOrEmpty(programInfo))
                {
                    pinnedPrograms.Add(programInfo);
                }
            }

            return pinnedPrograms;
        }

        //private static string GetProgramInfoFromShortcut(string shortcutPath)
        //{
        //    try
        //    {
        //        WshShell shell = new WshShell();
        //        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);

        //        string fileName = Path.GetFileName(shortcutPath);
        //        long fileSize = new FileInfo(shortcutPath).Length;
        //        string programName = Path.GetFileNameWithoutExtension(shortcut.TargetPath);

        //        return $"{fileName}, {fileSize} bytes, {programName}";
        //    }
        //    catch (Exception)
        //    {
        //        // Handle exceptions, e.g., if the shortcut is invalid
        //        return null;
        //    }
        //}

        /// <summary>
        /// Using power shell Wmiobject command to get the monitor actual resoultion setting.
        /// (if using screen bound to get the resoultion. this resoultion may effect by monitor ratio setting).  
        /// </summary>
        /// <returns>wide and height resolution</returns>
        public static (int Width, int Height) GetScreenResolution()
        {
            // Create the process start info
            var psi = new ProcessStartInfo
            {
                FileName = "powershell",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                RedirectStandardInput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                Arguments = "Get-WmiObject Win32_VideoController | Select-Object CurrentHorizontalResolution, CurrentVerticalResolution"
            };

            // Start the process
            using (var process = new Process { StartInfo = psi })
            {
                process.Start();

                // Read the output and error
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();

                // Wait for the process to exit
                process.WaitForExit();

                // Check for errors
                if (!string.IsNullOrEmpty(error))
                {
                    throw new Exception($"PowerShell command failed with error: {error}");
                }

                // Parse the output
                string[] lines = output.Split('\n', (char)StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length < 4)
                {
                    throw new Exception("Unexpected output format from PowerShell command");
                }

                var values = lines[3].Split(' ', (char)StringSplitOptions.RemoveEmptyEntries)
                                .Where(s => int.TryParse(s, out _))
                                .Select(int.Parse)
                                .ToArray();

                return (values[0], values[1]);
            }

        }

        /// <summary>
        /// Get the monitor ratio setting.  
        /// </summary>
        /// <returns>wide and height Ratio setting</returns>
        public static (float X_Ratio,float Y_Ratio) GetMonitorSettingRatio()
        {
            float xratio = 1.0f;
            float yratio = 1.0f;
            int resolutionX = 0;
            int resolutionY = 0;
            (resolutionX, resolutionY) = GetScreenResolution();
            Screen screen = Screen.PrimaryScreen;
            xratio = (float)resolutionX / screen.Bounds.Width;
            yratio = (float)resolutionY / screen.Bounds.Height;

            return (xratio,yratio);
        }
        /// <summary>
        /// Close windows setting window.  
        /// </summary>
        public static void CloseWindowsSettingWindow()
        {
            Process[] processes = Process.GetProcesses();
            foreach (Process process in processes)
            {
             //   Console.WriteLine(process.ProcessName);
                if (process.ProcessName == "SystemSettings")
                {
                    try
                    {
                        process.Kill();
                        break;
                    }
                    catch (Exception)
                    {
                        //handle any exception here
                    }
                }
            }
        }
    }
}
