/*
* SystemTrayHelper.cs 
* 
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Bencool   <Bencool.lin@quantatw.com>
*/

using System.Collections.Generic;
using System.Threading;


namespace CaptainWin.CommonAPI
{
    public static class SystemTrayHelper
    {
        private const string _g_SystemTrayProgramClassName = "Shell_TrayWnd";
        private const string _g_SystemTrayClassName = "SystemTray.NormalButton";
        private const string _g_SystemTrayClickProgramClassName = "TopLevelWindowForOverflowXamlIsland";
        private const string _g_SystemTrayClickClassName = "SystemTray.NormalButton";
        private static UserInput _keyboardMouse = new UserInput();

        /// <summary>
        ///  Use Automation Element get system tray menu click and get all the element's name back.
        /// </summary>
        /// <returns>List of string for system tray element showing name</returns>
        public static List<string> GetSystemTrayItems()
        {
            List<string> items = new List<string>();
            float fx = 0, fy = 0;
            //step1. Get monitor setting ratio
            (fx, fy) = BasicHelper.GetMonitorSettingRatio();

            //step2. Get systemTray positions
            MousePosition sT = AutomationUIHelper.GetAutomationElementPosition_ClassName(_g_SystemTrayProgramClassName, _g_SystemTrayClassName);
            //   sT.adjustByRatio(fx,fy);

            //step3. click system tray to expand the system tray window
            _keyboardMouse.Mouse_Move(sT.X, sT.Y);
            Thread.Sleep(300);
            _keyboardMouse.Mouse_LeftClick(200);
            Thread.Sleep(1000);

            //step4. Get the System Tray items
            // List<string> ss = AutomationUIHelper.ListRootChildrenClassName();
            items = AutomationUIHelper.GetNamesFromClassName(_g_SystemTrayClickProgramClassName, _g_SystemTrayClickClassName);
            _keyboardMouse.Mouse_Move(sT.X, sT.Y);
            Thread.Sleep(300);
            _keyboardMouse.Mouse_LeftClick(200);
            Thread.Sleep(1000);
            return items;
        }

        /// <summary>
        ///  Function to check if a specific program is in system tray. Use Automation Element to check.
        /// </summary>
        /// <param name="programName">The program name that need to check</param>
        /// <returns>True of False</returns>
        public static bool IsProgramInSystemTray(string programName)
        {
            List<string> st = new List<string>();
            st = GetSystemTrayItems();
            foreach (string app in st)
            {
                if (app.ToUpper().Contains(programName.ToUpper()))
                {
                    return true;
                }
            }
            return false;
        }

    }
}
