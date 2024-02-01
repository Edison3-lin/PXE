/*
* AllAppsHelper.cs 
* For AllApps testing 
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Bencool   <Bencool.lin@quantatw.com>
*/

using System;
using System.Collections.Generic;
using System.IO;
using System.Management;

namespace CaptainWin.CommonAPI
{
    /// <summary>
    ///  A class help you to get Allapps app list and check a speic app existence 
    /// </summary>
    public static class AllAppsHelper
    {

        /// <summary>
        ///  To get all apps installed list by WMI query 
        /// </summary>
        /// <returns>List of class:InstalledApplication object which include app name,Version and vendor infomation.</returns>
        public static List<InstalledApplication> GetInstalledApplications()
        {
            List<InstalledApplication> installedApps = new List<InstalledApplication>();

            try
            {
                // Query WMI for installed applications
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Product");
                ManagementObjectCollection queryCollection = searcher.Get();

                foreach (ManagementObject m in queryCollection)
                {
                    InstalledApplication app = new InstalledApplication
                    {
                        Name = m["Name"]?.ToString(),
                        Version = m["Version"]?.ToString(),
                        Publisher = m["Vendor"]?.ToString()
                    };

                    installedApps.Add(app);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving installed applications: {ex.Message}");
            }

            return installedApps;
        }


        /// <summary>
        ///  To get all apps .lnk files from three place(including subfolder) : 
        ///  1. C:\ProgramData\Microsoft\Windows\Start Menu\Programs
        ///  2. C:\ProgramData\Microsoft\Windows\Start Menu\Acer
        ///  3. C:\Users\$User Name\AppData\Roaming\Microsoft\Windows\Start Menu\Programs
        /// </summary>
        /// <returns>List of string for each string contain: lnkFile,fileSize,appName</returns>
        public static List<string> GetAllAppsFolder()
        {
            List<string> allApps = new List<string>();

            // ProgramData Start Menu Folder
            string programDataPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Programs");
            AddAppsFromFolder(allApps, programDataPath);

            programDataPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Acer");
            AddAppsFromFolder(allApps, programDataPath);

            // AppData Start Menu Folder
            string appDataPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs");
            AddAppsFromFolder(allApps, appDataPath);

            return allApps;
        }


        /// <summary>
        ///  Function to check if a specific program is in All Apps, check by folder .lnk finding.
        /// </summary>
        /// <param name="programName">The program name(.lnk file name) that need to check</param>
        /// <returns>True of False, True means found the programName</returns>
        public static bool IsProgramInAllAPPS_Folder(string programName)
        {
            List<string> allApps = new List<string>();
            allApps = GetAllAppsFolder();
            foreach (string app in allApps)
            {
                if (app.ToUpper().Contains(programName.ToUpper()))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        ///  Function to check if a specific program is in All Apps, check by WMI app install query.
        /// </summary>
        /// <param name="programName">The program name that need to check</param>
        /// <returns>True of False, True means found the programName</returns>
        public static bool IsProgramInAllAPPS_WMI(string programName)
        {
            List<InstalledApplication> installedApps = new List<InstalledApplication>();
            installedApps = GetInstalledApplications();
            foreach (InstalledApplication app in installedApps)
            {
                if (app.Name != null)
                {
                    if (app.Name.ToUpper() == programName.ToUpper())
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private static void AddAppsFromFolder(List<string> appsList, string folderPath)
        {
            try
            {
                // Check if the folder exists
                if (Directory.Exists(folderPath))
                {
                    // Get all lnk files in the current folder
                    string[] lnkFiles = Directory.GetFiles(folderPath, "*.lnk");

                    foreach (string lnkFile in lnkFiles)
                    {
                        // Process each lnk file
                        string appName = Path.GetFileNameWithoutExtension(lnkFile);
                        string fileSize = GetFileSize(lnkFile);
                        string appInfo = $"{Path.GetFileName(lnkFile)},{fileSize},{appName}";
                        appsList.Add(appInfo);
                    }

                    // Get all subdirectories
                    string[] subDirectories = Directory.GetDirectories(folderPath);

                    // Recursively process each subdirectory
                    foreach (string subDirectory in subDirectories)
                    {
                        AddAppsFromFolder(appsList, subDirectory);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading folder {folderPath}: {ex.Message}");
            }
        }

        private static string GetFileSize(string filePath)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(filePath);
                long sizeInBytes = fileInfo.Length;
                double sizeInKb = sizeInBytes / 1024.0;

                return $"{sizeInKb:F2} KB";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting file size for {filePath}: {ex.Message}");
                return "Unknown";
            }
        }
    }

    /// <summary>
    ///  A data model class contain an application info 
    /// </summary>
    public class InstalledApplication
    {
        public string Name { get; set; }
        public string Version { get; set; }
        public string Publisher { get; set; }

        public override string ToString()
        {
            return $"{Name} (Version: {Version}, Publisher: {Publisher})";
        }
    }

}
