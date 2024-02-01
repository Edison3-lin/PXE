/*
* TaskbarHelper.cs 
* 
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

namespace CaptainWin.CommonAPI
{
    public static class TaskbarHelper
    {
        private const string _g_USER_QUICK_LAUNCH_PATH = @"Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar";
        private const string _g_TaskBarProgramClassName = "Shell_TrayWnd";
        private const string _g_TaskBarItemsClassName = "Taskbar.TaskListButtonAutomationPeer";


        /// <summary>
        /// Use automation Ellement library to find all taskbar items and return the names 
        /// </summary>
        /// <returns>List of string, names</returns>
        public static List<string> ListTaskBar_AutomationUI()
        {
            return AutomationUIHelper.GetNamesFromClassName(_g_TaskBarProgramClassName, _g_TaskBarItemsClassName);
        }

        /// <summary>
        /// Function to list all user-pinned .lnk files in the specified folder(Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar)
        /// </summary>
        /// <returns>List of string, Each String combine of {fileName}, {fileSize} bytes, {programName}</returns>
        public static List<string> ListTaskBar_ShortcutFolder()
        {
            string taskbarPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), _g_USER_QUICK_LAUNCH_PATH);
            List<string> pinnedPrograms = new List<string>();
            pinnedPrograms = BasicHelper.ListShortcut(taskbarPath);
            return pinnedPrograms;
        }

        /// <summary>
        ///  Function to check if a specific program is pinned to the toolbar
        /// </summary>
        /// <param name="programName">The program name that need to check</param>
        /// <returns>True of False</returns>
        public static bool IsProgramPinned(string programName)
        {
            //check from AutomationUI
            List<string> TaskBarItems = ListTaskBar_AutomationUI();
            foreach (string Find_Pinned_Item in TaskBarItems)
            {
                if (Find_Pinned_Item == programName)
                {
                    return true;
                }
            }
            //check shortCut
            List<string> pinnedPrograms = ListTaskBar_ShortcutFolder();
            foreach(string Find_Pinned_Item in pinnedPrograms)
            {
                if(Find_Pinned_Item.ToUpper().Contains(programName.ToUpper()))
                {
                    return true;                
                }
            }

            return false;
        }
    

        /// <summary>
        /// Use Win API handler : IsWindowVisible to check if a specific program is running and has an icon on the taskbar
        /// </summary>
        /// <param name="programName">The program name that need to check</param>
        /// <returns>True of False</returns>
        public static bool IsRunningOnTaskbar(string programName)
        {
            var processes = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(programName));

            return processes.Any(process => _IsWindowVisible(process.MainWindowHandle));
        }

        // Helper function to check if a window is visible
        private static bool _IsWindowVisible(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero)
                return false;

            return IsWindowVisible(hWnd);
        }

        /// <summary>
        ///  Win API handler : IsWindowVisible
        /// </summary>
        /// <returns>True of False</returns>
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);



    }
}
