/*
* StartMenuHelper.cs 
* Test logic for Windows Start Menu.
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Bencool   <Bencool.lin@quantatw.com>
*/


using System;
using System.Collections.Generic;
using System.IO;
using System.Linq.Expressions;
using System.Threading;
using System.Windows.Forms;
using static CaptainWin.CommonAPI.ImageHelper;



namespace CaptainWin.CommonAPI
{
    public static class StartMenuHelper
    {
        private const string _g_StartMenuClassName = "Shell_TrayWnd";
        private const string _g_StartMenuItemsClassName = "StartButton";
        //  private const int _g_MouseScrollDown3Line = 3;
        private static UserInput _keyboardMouse = new UserInput();
        /// <summary>
        ///    Print Screen for start menu and use OCR to recongnize text and return all the recognize text.
        /// </summary>
        /// <returns>string for all text in the screenshot that OCR recognized</returns>
        public static string ListStartMenuItems_OCR()
        {
            //check resolution
            //Get monitor setting ratio
            float X_Ratio = 1, Y_Ratio = 1;
            (X_Ratio, Y_Ratio) = BasicHelper.GetMonitorSettingRatio();

            if (X_Ratio != 1.5)
            {
                ImageHelper.ChangeMonitorRatio_Win11((int)MonitorScalingOption.Percent150);

            }
            string ocrResult = "";
            _keyboardMouse.KB_PressWin(100);
            string startMenuScreenShotFilePath = ImageHelper.PrintScreenStartMenu("StartMenu1");
            ocrResult = ImageHelper.PerformOCR(startMenuScreenShotFilePath);
            Console.WriteLine(ocrResult);
            Scroll_Down_StatrMenu();
            startMenuScreenShotFilePath = ImageHelper.PrintScreenStartMenu("StartMenu2");
            ocrResult += ImageHelper.PerformOCR(startMenuScreenShotFilePath);
            Console.WriteLine(ocrResult);
            _keyboardMouse.KB_PressWin(100);

            if (X_Ratio != 1.5)
            {
                int oriRatio = Convert.ToInt32((Y_Ratio - 1) / 0.25);
                ImageHelper.ChangeMonitorRatio_Win11(oriRatio);
            }


            return ocrResult;
        }
        private static void Scroll_Down_StatrMenu()
        {
            // MouseHelper.ScrollDown(_g_MouseScrollDown3Line);
            SendKeys.SendWait("{TAB}");
            Thread.Sleep(200);
            // Simulate pressing Down arrow key
            for (int i = 0; i < 3; i++)
            {
                SendKeys.SendWait("{DOWN}");
                Thread.Sleep(200);
            }
        }
        /// <summary>
        ///  Function to check if a specific program is pinned to Start Menu(use OCR method)
        /// </summary>
        /// <param name="program">The program name that need to check</param>
        /// <returns>True of False</returns>
        public static bool IsProgramInStartMenu(string program)
        {
            string ocrCheck = ListStartMenuItems_OCR();
            if (ocrCheck.Contains(program))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        //public static List<string> ListStartMenuItems_AutomationUI()
        //{
        //    KeyboardHelper.WinKeyPress();
        //    MousePosition startBottonPosition = AutomationUIHelper.GetAutomationElementPosition(_g_StartMenuClassName, _g_StartMenuItemsClassName);
        //    MouseHelper.Click(startBottonPosition.X + 10, (startBottonPosition.Y - 10));
        //    Thread.Sleep(2000);
        //    MouseHelper.Click(startBottonPosition.X + 10, (startBottonPosition.Y - 10));

        //    return AutomationUIHelper.GetStartMenuItems();
        //}

        /// <summary>
        ///   Check 
        /// </summary>
        /// <returns>string for all text in the screenshot that OCR recognized</returns>
        //public static List<string> GetStartMenuItems_Folder()
        //{
        //    List<string> returnLnkFileName = new List<string>();
        //    string startMenuPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs");
        //    returnLnkFileName = BasicHelper.ListShortcut(startMenuPath);
        //    if (Directory.Exists(startMenuPath))
        //    {
        //        Console.WriteLine("Pinned items in the Start Menu:");

        //        foreach (string lnkFile in Directory.GetFiles(startMenuPath, "*.lnk", SearchOption.AllDirectories))
        //        {
        //            Console.WriteLine(lnkFile);
        //            returnLnkFileName.Add(lnkFile);
        //        }
        //    }
        //    else
        //    {
        //        Console.WriteLine("Start Menu folder not found.");
        //    }
        //    return returnLnkFileName;
        //}

        //public static List<string> GetStartMenuItems_Regitry()
        //{
        //    List<string> returnPinned_Regidtry = new List<string>();

        //    string startMenuPath = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu);
        //    string userStartMenuPath = Path.Combine(startMenuPath, "Programs");

        //    // Registry path for pinned Start Menu items
        //    string registryPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage";

        //    try
        //    {
        //        using (RegistryKey key = Registry.CurrentUser.OpenSubKey(registryPath))
        //        {
        //            if (key != null)
        //            {
        //                // Retrieve the pinned items value
        //                string pinnedItems = key.GetValue("Favorites") as string;

        //                if (!string.IsNullOrEmpty(pinnedItems))
        //                {
        //                    Console.WriteLine("Pinned items in the Start Menu:");
        //                    Console.WriteLine(pinnedItems);
        //                    returnPinned_Regidtry.Add(pinnedItems);
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Error reading registry: {ex.Message}");
        //    }
        //    return returnPinned_Regidtry;
        //}



    }
}
