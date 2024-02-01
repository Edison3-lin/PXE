/*
* ImageHelper.cs 
* -V1.0 Basic functions for
* -Print Screen
* -OCR (using Tesseract library)
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Bencool   <Bencool.lin@quantatw.com>
*/

using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Tesseract;


namespace CaptainWin.CommonAPI
{
    /// <summary>
    ///  Can take the ScreenShot, change monitor showing text and windows's resolution, 
    ///  and OCR to get text from the images.
    /// </summary>
    public static class ImageHelper
    {
        private const string _g_WindowsSettingClassName = "ApplicationFrameWindow";
        private const string _g_MonitorSettingClassName = "NamedContainerAutomationPeer";
        private const string _g_MonitorSettingMaximumAutomationID = "Maximize";
        private const string _g_MonitorSettingScaleAutomationID = "SystemSettings_Display_Scaling_ItemSizeOverride_ComboBox";
        private static UserInput _keyboardMouse = new UserInput();


        /// <summary>
        /// PrintScreen for screen.Bounds width and height
        /// </summary>
        /// <returns>Bitmap for printScreen</returns>
        public static Bitmap PrintScreen()
        {
          // Get the primary screen, using primary screen may affect by monitor ratio setting
          //  Screen screen = Screen.PrimaryScreen; 
          //Get actual resolution from power sheel script
            (int rWide, int rHeight) = BasicHelper.GetScreenResolution();
            Bitmap screenshot = new Bitmap(rWide, rHeight, PixelFormat.Format32bppArgb);
            using (Graphics graphics = Graphics.FromImage(screenshot))
            {
                // Capture the screen to the bitmap
                graphics.CopyFromScreen(0, 0, 0, 0, screenshot.Size);
            }
            return screenshot;
        }
        /// <summary>
        /// PrintScreen for screen.Bounds width and height,
        /// save the image to the FolderPath with timestamp in file name and return the final file path
        /// </summary>
        /// <param name="FolderPath">the folder path to save image</param>
        /// <returns>file path</returns>
        public static string PrintScreen(string FolderPath)
        {
            string fileName = $"ScreenShot_{DateTime.Now:yyyyMMddHHmmss}.png";
            string filePath = Path.Combine(FolderPath, fileName);
          //  Screen screen = Screen.PrimaryScreen;
            (int rWide, int rHeight) = BasicHelper.GetScreenResolution();
            Bitmap screenshot = new Bitmap(rWide, rHeight, PixelFormat.Format32bppArgb);

            using (Graphics graphics = Graphics.FromImage(screenshot))
            {
                // Capture the screen to the bitmap
                graphics.CopyFromScreen(0, 0, 0, 0, screenshot.Size);
            }

            screenshot.Save(filePath);

            return filePath;
        }


        /// <summary>
        ///    screenshot for start menu, will cut the left and right 25% width
        /// </summary>
        /// <param name="filenamePrefix">FileName prefix for the output image file</param>
        /// <returns>Filename indicate the saved print screen(in the same .exe folder)</returns>
        public static string PrintScreenStartMenu(string filenamePrefix)
        {

            // Get the combined bounds of all screens
            Rectangle bounds = Screen.PrimaryScreen.Bounds;
            for (int i = 1; i < Screen.AllScreens.Length; i++)
            {
                bounds = Rectangle.Union(bounds, Screen.AllScreens[i].Bounds);
            }

            (int rWide, int rHeight) = BasicHelper.GetScreenResolution();
            // Capture the screen to the bitmap
            Bitmap screenshot = new Bitmap(rWide, rHeight, PixelFormat.Format32bppArgb);
            //  Bitmap screenshot = new Bitmap(2048, 1200, PixelFormat.Format32bppArgb);

            using (Graphics graphics = Graphics.FromImage(screenshot))
            {
                int captureWidth = (int)(rWide * 0.5); // Capturing 50% of the width
                int excludedWidth = (rWide - captureWidth) / 2; // Excluding 25% on each side

                graphics.CopyFromScreen(bounds.Left + excludedWidth, bounds.Top, captureWidth, 0, new Size(captureWidth, rHeight));
            }

            // Save the screenshot to a file
            string filePath = $"{filenamePrefix}_{DateTime.Now:yyyyMMddHHmmss}.png";
            screenshot.Save(filePath);

            return filePath;
        }



        /// <summary>
        ///   Passing an image file path and use tesseract OCR module to do the OCR 
        ///   and return all the text capture(Currently only setting to recognize for English) 
        /// </summary>
        /// <param name="imagePath">Target image file path to do the OCR</param>
        /// <returns>OCR Results</returns>
        public static string PerformOCR(string imagePath)
        {
            //using (var engine = new TesseractEngine(@"tessdata", "eng+chi_tra", EngineMode.Default))
            //  using (var engine = new TesseractEngine(@"tessdata", "chi_tra", EngineMode.Default))
            using (var engine = new TesseractEngine(@"tessdata", "eng", EngineMode.Default))
            {
                using (var img = Pix.LoadFromFile(imagePath))
                {
                    using (var page = engine.Process(img))
                    {
                        return page.GetText();
                    }
                }
            }
        }


        /// <summary>
        ///    Call out windows setting window, then use automation Element to find win11 setting monitor element position.
        ///    and after a serial of mouse and keyboard actions,
        ///    can change the monitor showing ratio to the parameter that indicated.    
        /// </summary>
        /// <param name="percentage">
        ///    win11 setting->monitor->ratio selections, please use the follow enum to pass in parameter:
        /// <code>
        ///    ImageHelper.ChangeMonitorRatio_Win11((int)MonitorScalingOption.Percent100);
        ///    ImageHelper.ChangeMonitorRatio_Win11((int)MonitorScalingOption.Percent125);
        ///    ImageHelper.ChangeMonitorRatio_Win11((int)MonitorScalingOption.Percent150);
        ///    ImageHelper.ChangeMonitorRatio_Win11((int)MonitorScalingOption.Percent175);
        /// </code>
        /// </param>
        public static void ChangeMonitorRatio_Win11(int percentage)
        {
            //step1. open Setting menu and maximum the window
            // Open the Settings window in a new process
            Process settingsProcess = new Process();
            settingsProcess.StartInfo.FileName = "ms-settings:";
            settingsProcess.Start();
            Thread.Sleep(2000);

            MousePosition maxWindowPosition = AutomationUIHelper.GetAutomationElementPosition_AutomationID(_g_WindowsSettingClassName, _g_MonitorSettingMaximumAutomationID);
            _keyboardMouse.Mouse_Move(maxWindowPosition.X, maxWindowPosition.Y);
            Thread.Sleep(300);
            _keyboardMouse.Mouse_LeftClick(200);
            Thread.Sleep(100);
            //step2. tab * 5 to Monitor setting

            MousePosition monitorSettingPosition = AutomationUIHelper.GetAutomationElementPosition_ClassName(_g_WindowsSettingClassName, _g_MonitorSettingClassName, 1);
            _keyboardMouse.Mouse_Move(monitorSettingPosition.X, monitorSettingPosition.Y);
            Thread.Sleep(300);
            _keyboardMouse.Mouse_LeftClick(200);
            Thread.Sleep(1000);

            MousePosition monitorSettingScalePosition = AutomationUIHelper.GetAutomationElementPosition_AutomationID(_g_WindowsSettingClassName, _g_MonitorSettingScaleAutomationID);
            _keyboardMouse.Mouse_Move(monitorSettingScalePosition.X, monitorSettingScalePosition.Y);
            Thread.Sleep(300);
            _keyboardMouse.Mouse_LeftClick(200);
            Thread.Sleep(1000);


            //step4. percentage setting
            //move to last one first
            for (int i = 0; i < 4; i++)
            {
                SendKeys.SendWait("{UP}");
                Thread.Sleep(100);
            }
            // UserInput ui = new UserInput();
            for (int i = 0; i < (int)percentage; i++)
            {
                SendKeys.SendWait("{DOWN}");
                Thread.Sleep(100);
            }
            //ui.KB_PressSpace(2);
            SendKeys.SendWait(" ");
            //  SendKeys.SendWait("{ENTER}");

            //step5. close setting window
            BasicHelper.CloseWindowsSettingWindow();
            Thread.Sleep(1000);
        }

        public enum MonitorScalingOption
        {
            Percent100 = 0,
            Percent125 = 1,
            Percent150 = 2,
            Percent175 = 3
        }

        //public enum TesseractLanguage
        //{
        //    English,
        //    Spanish,
        //    French,
        //    German,
        //    Italian,
        //    Portuguese,
        //    Dutch,
        //    Russian,
        //    ChineseSimplified,
        //    ChineseTraditional,
        //    Japanese,
        //    Korean,
        //    Arabic,
        //    // Add more languages as needed
        //}
    }
}

