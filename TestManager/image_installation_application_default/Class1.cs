using Microsoft.Office.Interop.Excel;
using Microsoft.SqlServer.Server;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Linq;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Common;

namespace image_installation_application_default
{
    public class image_installation_application_default
    {
        public int Setup()
        {
            // common.Setup
            Testflow.Setup("xxx");
            return 11;
        }

        public static bool Run()
        {
            //For Get all Installed applications
            
            List<string> gInstalledSoftware = new List<string>();
            List<string> checktable = new List<string>();
            List<string> result_applications_list = new List<string>();//List for check test result

            string AcerCareCenterV4 = "not check";//0
            string AcerConfigurationManager = "not check";
            string Acerlegalportal = "not check";
            string AcerProductRegistration = "not check";
            string AcerSense = "not check";
            string AmazonWeblink = "not check";
            string AppExplorer = "not check";
            string BaiduNetDisk = "not check";
            string BookingWeblink = "not check";
            string BookingFavorite = "not check";
            string Hao123Weblink = "not check";
            string Hao123Favorite = "not check";
            string BaiduWeblink = "not check";
            string BaiduFavorite = "not check";
            string SoftwarePicks = "not check";
            string DropboxPromotion = "not check";
            string Evernote = "not check";//10
            string ForgeofEmpires = "not check";
            string GooglePlayGames = "not check";
            string IntelMultiDeviceExperience = "not check";
            string McAfeeLiveSafeP1 = "not check";
            string MicrosoftOffice = "not check";
            string MyOfficesuite = "not check";
            string WPSChina = "not check";
            string MicrosoftOfficeInstaller = "not check";
            string MorphoCameraEffects = "not check";
            string Planet9 = "not check";
            string Solitaire = "not check";//20
            string Spades = "not check";
            string StorageUtilities = "not check";
            string PowerManagementSetting = "not check";
            string UserBehaviorTrackingFramework = "not check";
            string Yandex = "not check";//25
            int fail = 0;
            //Build all apps check list table
            string[,] appschecklist = new string[40, 3];
            appschecklist[0,0] = "Acer Care Center V4"; appschecklist[0,1] = "not check"; appschecklist[0, 2] = "en-US";
            appschecklist[1,0] = "Acer Configuration Manager"; appschecklist[1,1] = "not check"; appschecklist[1,2] = "en-US";
            appschecklist[2,0] = "Acer legal portal"; appschecklist[2,1] = "not check"; appschecklist[2,2] = "en-US";
            appschecklist[3,0] = "Acer Product Registration"; appschecklist[3,1] = "not check"; appschecklist[3,2] = "en-US";
            appschecklist[4,0] = "AcerSense"; appschecklist[4,1] = "not check"; appschecklist[4,2] = "en-US";
            appschecklist[5,0] = "AcerWarrantyRegistration for India"; appschecklist[5,1] = "not check"; appschecklist[5, 2] = "en-IN";//India
            appschecklist[6,0] = "Agoda Weblink"; appschecklist[6,1] = "not check"; appschecklist[6,2] = "eu-EN";
            appschecklist[7,0] = "Agoda_Favorite"; appschecklist[7, 1] = "not check"; appschecklist[7,2] = "eu-EN";
            appschecklist[8,0] = "AmazonWeblink"; appschecklist[8,1] = "not check"; appschecklist[8,2] = "en-US";
            appschecklist[9,0] = "AppExplorer"; appschecklist[9,1] = "not check"; appschecklist[9,2] = "en-US";
            appschecklist[10,0] = "Baidu Net Disk"; appschecklist[10,1] = "not check"; appschecklist[10,2] = "zh-CN";//China
            appschecklist[11,0] = "Baidu Weblink"; appschecklist[11,1] = "not check"; appschecklist[11,2] = "zh-CN";//China
            appschecklist[12,0] = "Baidu_Favorite"; appschecklist[12, 1] = "not check"; appschecklist[12,2] = "zh-CN";//China
            appschecklist[13,0] = "Booking.com Weblink"; appschecklist[13,1] = "not check"; appschecklist[13,2] = "en-US";
            appschecklist[14,0] = "Booking.com_Favorite"; appschecklist[14,1] = "not check"; appschecklist[14,2] = "en-US";
            appschecklist[15,0] = "Dropbox Promotion"; appschecklist[15,1] = "not check"; appschecklist[15,2] = "en-US";
            appschecklist[16,0] = "Evernote"; appschecklist[16,1] = "not check"; appschecklist[16,2] = "en-US";
            appschecklist[17,0] = "Forge of Empires"; appschecklist[17,1] = "not check"; appschecklist[17,2] = "en-US";
            appschecklist[18,0] = "Google Play Games"; appschecklist[18,1] = "not check"; appschecklist[18, 2] = "en-US";
            appschecklist[19,0] = "Hao123 Weblink"; appschecklist[19,1] = "not check"; appschecklist[19,2] = "zh-CN";//China
            appschecklist[20,0] = "Hao123_Favorite"; appschecklist[20,1] = "not check"; appschecklist[20,2] = "zh-CN";//China
            appschecklist[21,0] = "Intel Multi-Device Experience"; appschecklist[21,1] = "not check"; appschecklist[21,2] = "en-US";
            appschecklist[22,0] = "McAfee LiveSafe P1"; appschecklist[22,1] = "not check"; appschecklist[22, 2] = "en-US";
            appschecklist[23,0] = "Microsoft Office"; appschecklist[23,1] = "not check"; appschecklist[23,2] = "en-US";
            appschecklist[24,0] = "Microsoft Office Installer"; appschecklist[24,1] = "not check"; appschecklist[24,2] = "en-US";
            appschecklist[25,0] = "Morpho Camera Effects"; appschecklist[25,1] = "not check"; appschecklist[25,2] = "en-US";
            appschecklist[26,0] = "MyOffice_Spreadsheet"; appschecklist[26,1] = "not check"; appschecklist[26,2] = "ru-RU";//Russian
            appschecklist[27,0] = "MyOffice_Text"; appschecklist[27,1] = "not check"; appschecklist[27,2] = "ru-RU";//Russian
            appschecklist[28,0] = "Planet9"; appschecklist[28,1] = "not check"; appschecklist[28,2] = "en-US";
            appschecklist[29,0] = "Power Management Setting"; appschecklist[29,1] = "not check"; appschecklist[29,2] = "en-US";
            appschecklist[30,0] = "Software Picks"; appschecklist[30,1] = "not check"; appschecklist[30,2] = "zh-CN";
            appschecklist[31,0] = "MicrosoftOfficeInstaller"; appschecklist[31,1] = "not check"; appschecklist[31,2] = "en-US";
            appschecklist[32,0] = "Solitaire"; appschecklist[32,1] = "not check"; appschecklist[32,2] = "en-US";
            appschecklist[33,0] = "Spades"; appschecklist[33,1] = "not check"; appschecklist[33,2] = "en-US";         
            appschecklist[34,0] = "Storage Utilities"; appschecklist[34,1] = "not check"; appschecklist[34,2] = "en-US";  
            appschecklist[35,0] = "User Behavior Tracking Framework"; appschecklist[35,1] = "not check"; appschecklist[35,2] = "en-US";
            appschecklist[36,0] = "WPS China"; appschecklist[36,1] = "not check"; appschecklist[36,2] = "zh-CN";//China
            appschecklist[37,0] = "Yandex Browser"; appschecklist[37,1] = "not check"; appschecklist[37,2] = "ru-RU";//Russian
            appschecklist[38,0] = "Yandex Weblink"; appschecklist[38,1] = "not check"; appschecklist[38,2] = "ru-RU";//Russian
            appschecklist[39,0] = "Yandex_Favorite"; appschecklist[39,1] = "not check"; appschecklist[39,2] = "ru-RU";//Russian


            CultureInfo ci = CultureInfo.InstalledUICulture;
            Console.WriteLine("Default Language Info:");
            Console.WriteLine("* Name: {0}", ci.Name);

            //Get project name and PC name
            string[] PCInformation = new string[2];
            PCInformation = setup1();
            Console.WriteLine(PCInformation[0] + " " + PCInformation[1]);

            //Get Tag files name, counts in PC, return values in a List, AppidExistListInPC
            List<string> AppidExistListInPC = new List<string>();
            AppidExistListInPC = GetAcerTagfiles("C:\\OEM\\Preload\\InstalledApps");
            int AppidExistListInPC_size = AppidExistListInPC.Count;
            Console.WriteLine("-------------Count if AppidExistListInPC_size: " + AppidExistListInPC_size);


            string userName = Environment.UserName;
            // 設定Excel檔案的路徑
            string root_path = "C:\\Users\\" + userName + "\\Downloads\\";
            string root_patth = @"C:\TestManager\ItemDownload\";
            string excelFileName = "SCD_RV07RC.xls";
            string excelFilePath = root_path + excelFileName;
            //string excelFilePath = @"C:\\Users\\k\\Downloads\\SCL_Aspire_Twix_ADN_WNNOP64W11_SV2_MAYN_Generic_RV03RC_Office (Trial)_BNM000035446_-.xls"; // Replace with the path to your Excel file
            Console.WriteLine(excelFilePath);

            int area_row_item_type = 0;
            int area_col_item_type = 0;

            // 建立一個新的Excel Application物件
            Excel.Application excelApp = new Excel.Application();

            // 打開Excel檔案
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);

            // Excel檔案有N個工作表，使用索引"SCL Content"來取得該工作表
            Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets["SCL Content"];

            // 讀取資料
            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;
            string driver_cellValue = null;
            string version_cellValue = null;
            string appid_cellValue = null;
            string application_name_cellValue = "*";
            string application_version_cellValue = null;
            string applicationID_version_cellValue = null;
            string sub_brand_name_cellValue = null;
            List<string> drivers_list_SCL = new List<string>();
            List<string> applications_list_SCL = new List<string>();
            List<string> New_applications_list_SCL = new List<string>();
            List<string> full_apps_table = new List<string>();

            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    // 使用Cells物件來取得單元格的值
                    Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, col];
                    string cellValue = cell.Value != null ? cell.Value.ToString() : "";
                    if (cellValue == "Sub Brand")
                    {
                        int brand_base_row = row;
                        int brand_base_col = col;
                        Excel.Range Item_Desc_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[brand_base_row, brand_base_col + 1];
                        sub_brand_name_cellValue = Item_Desc_cell.Value != null ? Item_Desc_cell.Value.ToString() : "";
                        Console.WriteLine("sub brand: " + sub_brand_name_cellValue);
                    }

                    if (cellValue == "Application")
                    {
                        //Console.Write(cellValue);
                        //Console.WriteLine();
                        //Excel.Range cColumn = sheet.get_Range("B", null);
                        int driver_base_row = row;
                        int driver_base_col = col;

                        //Console.WriteLine("@@@Driver row: {0}" + " " + "@@@Driver col: {1}", row, col);
                        area_row_item_type = driver_base_row + 2;
                        area_col_item_type = driver_base_col + 2;
                        if (sub_brand_name_cellValue == PCInformation[0])
                        {
                            do
                            {
                                Excel.Range Item_Desc_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[area_row_item_type, area_col_item_type];
                                application_name_cellValue = Item_Desc_cell.Value != null ? Item_Desc_cell.Value.ToString() : "";

                                Excel.Range APP_Version_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[area_row_item_type, area_col_item_type + 1];
                                application_version_cellValue = APP_Version_cell.Value != null ? APP_Version_cell.Value.ToString() : "";

                                Excel.Range APPID_Version_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[area_row_item_type, area_col_item_type + 18];
                                applicationID_version_cellValue = APPID_Version_cell.Value != null ? APPID_Version_cell.Value.ToString() : "";
                                

                                full_apps_table.Add(appNamingReprocess(application_name_cellValue, application_version_cellValue) + "," + applicationID_version_cellValue);

                                if (AppidExistListInPC.Find(s => s == applicationID_version_cellValue) != null)
                                {
                                    New_applications_list_SCL.Add(appNamingReprocess(application_name_cellValue, application_version_cellValue));
                                }

                                GetInstalledAppsListByLanguage(ci.Name, 
                                    application_name_cellValue,
                                    application_version_cellValue,
                                    applicationID_version_cellValue,
                                    applications_list_SCL,
                                    result_applications_list);


                                area_row_item_type += 1;

                            } while (application_name_cellValue != "");
                            if (applications_list_SCL.Count > 0)
                            {
                                applications_list_SCL.RemoveAt(applications_list_SCL.Count - 1);
                                result_applications_list.RemoveAt(result_applications_list.Count - 1);
                            }

                            Console.WriteLine("Fininshed !!!");
                            //Console.WriteLine();
                        }
                        else
                        {
                            Console.WriteLine("SCL Brand Name does not match project Name of PC !!!!");
                            Console.WriteLine("Stop checking, exit process ....");
                        }
                    }
                }
                //Console.WriteLine();
            }

            // 關閉Excel檔案
            workbook.Close();
            excelApp.Quit();

            // 釋放資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            //DUMP check list
            foreach (string apps in result_applications_list)
            {
                Console.WriteLine(apps);
            }
            //DUMP full apps list
            foreach (string elements in full_apps_table) 
            {
                Console.WriteLine($"{elements}");
            }

            GetinstalledinformationFromPC(gInstalledSoftware);

            // if (sub_brand_name_cellValue == PCInformation[0])
            // {
            //     Console.WriteLine("----- DUMP tag exist and match apps in this PC ------");
            //     foreach (string items in New_applications_list_SCL)
            //     {
            //         Console.WriteLine(items);
            //     }
            //     Console.WriteLine();
            // }


            // //Dump all applications info in OS
            // StreamWriter sw = new StreamWriter(root_path + "installedapplist.txt");
            // foreach (var key in gInstalledSoftware)
            // {
            //     sw.Write(key.ToString());
            //     sw.Write('\n');
            // }
            // sw.Close();

            // //Dump all application list in SCL 
            // Console.WriteLine("DUMP ALL applications in SCL file");
            // Console.WriteLine("---------------------------------");
            
            // string correctString = " ";
            // //Try to do comparing between SCL list and all applications list
            // string filePath = root_path + "installedapplist.txt"; // Replace with the path to your text file
                                                                  
            // foreach (string appname in New_applications_list_SCL)
            // {
            //     Console.WriteLine("----------------------");
            //     Console.WriteLine($"{appname}");

            //     if (!string.IsNullOrWhiteSpace(appname))
            //     {
            //         string[] applist = appname.Split(',');
            //         //Mapping Application name ....
            //         if (applist[0] == "Dropbox Promotion")
            //         {
            //             correctString = applist[0].Replace("Dropbox Promotion", "DropboxOEM");
            //         }
            //         else if (applist[0] == "Acer Product Registration")
            //         {
            //             correctString = applist[0].Replace("Acer Product Registration", "AcerRegistration");
            //             if (applist[1].IndexOf(".00") >= 0)
            //             {
            //                 applist[1] = applist[1].Replace(".00", ".0");
            //             }
            //         }
            //         else if (applist[0] == "Acer Care Center V4")
            //         {
            //             correctString = applist[0].Replace("Acer Care Center V4", "Care Center");
            //             int skip_char = applist[1].IndexOf("_");
            //             if (skip_char > 0)
            //             {
            //                 applist[1] = applist[1].Substring(0, skip_char);
            //             }
            //         }
            //         else if (applist[0] == "Acer Configuration Manager")
            //         {
            //             correctString = applist[0].Replace("Acer Configuration Manager", "Acer Configuration Manager");
            //             int skip_char = applist[1].IndexOf("r");
            //             if (skip_char > 0)
            //             {
            //                 applist[1] = applist[1].Substring(0, skip_char);
            //                 //Console.WriteLine(applist[0] + " " + applist[1]);
            //             }
            //         }
            //         else if (applist[0] == "McAfee LiveSafe P1")
            //         {
            //             correctString = applist[0].Replace("McAfee LiveSafe P1", "McAfee LiveSafe");
            //             applist[1] = applist[1].Replace(".R", " R");
            //             //Console.WriteLine(applist[1]);
            //         }
            //         else if (applist[0] == "User Behavior Tracking Framework")
            //         {
            //             if (applist[1].IndexOf(".00") >= 0)
            //             {
            //                 applist[1] = applist[1].Replace(".00", ".0");
            //             }
            //             correctString = applist[0].Replace("User Behavior Tracking Framework", "UserExperienceImprovementProgram");
            //         }
            //         else if (applist[0] == "Google Play Games")
            //         {
            //             correctString = applist[0].Replace("Google Play Games", "Google Play Games");
            //         }
            //         else if (applist[0] == "Booking.com Weblink")//Booking suite
            //         {
            //             correctString = applist[0].Replace("Booking.com Weblink", "Booking.com.url");
            //         }
            //         else if (applist[0] == "Booking.com_Favorite")//Booking suite
            //         {
            //             correctString = applist[0].Replace("Booking.com_Favorite", "Booking.com.url");
            //         }
            //         else if (applist[0] == "Hao123 Weblink" || applist[0] == "Hao123_Favorite")
            //         {
            //             correctString = applist[0].Replace("Hao123 Weblink", "hao123");
            //             correctString = applist[0].Replace("Hao123_Favorite", "hao123");
            //         }
            //         else if (applist[0] == "Acer legal portal")
            //         {
            //             correctString = applist[0].Replace("Acer legal portal", "Acer Legal Information");
            //         }
            //         else if (applist[0] == "Baidu Net Disk")
            //         {
            //             correctString = applist[0].Replace("Baidu Net Disk", "百度网盘");
            //         }
            //         else if (applist[0] == "Software Picks")
            //         {
            //             correctString = applist[0].Replace("Software Picks", "软件精选");
            //         }
            //         else if (applist[0] == "Amazon Weblink")
            //         {
            //             correctString = applist[0].Replace("Amazon Weblink", "Amazon.url");
            //         }
            //         else if (applist[0] == "Intel Multi-Device Experience")
            //         {
            //             correctString = applist[0].Replace("Intel Multi-Device Experience", "IntelTechnologyMDE");
            //         }
            //         else if (applist[0] == "Microsoft Office")//MS Office suite
            //         {
            //             correctString = applist[0].Replace("Microsoft Office", "Microsoft 365");
            //         }
            //         else if (applist[0] == "Microsoft Office Installer")//MS Office suite
            //         {
            //             correctString = applist[0].Replace("Microsoft Office Installer", "Microsoft OneNote");
            //         }
            //         else if (applist[0] == "MyOffice_Spreadsheet")//MyOffice suite
            //         {
            //             correctString = applist[0].Replace("MyOffice_Spreadsheet", "MyOffice");
            //         }
            //         else if (applist[0] == "MyOffice_Text")//MyOffice suite
            //         {
            //             correctString = applist[0].Replace("MyOffice_Text", "MyOffice");
            //         }
            //         else if (applist[0] == "Morpho Camera Effects")
            //         {
            //             correctString = applist[0].Replace("Morpho Camera Effects", "Morpho Inc.");
            //         }
            //         else if (applist[0] == "Power Management Setting")
            //         {
            //             correctString = applist[0].Replace("Power Management Setting", "PowerManagementSetting.tag");
            //         }
            //         else if (applist[0] == "Storage Utilities")
            //         {
            //             correctString = applist[0].Replace("Storage Utilities", "StorageUtilities.tag");
            //         }
            //         else if (applist[0] == "Planet9")
            //         {
            //             correctString = applist[0].Replace("Planet9", "Planet9_Version.txt");
            //         }
            //         else if (applist[0] == "Solitaire")
            //         {
            //             correctString = applist[0].Replace("Solitaire", "RandomSaladGamesLLC");
            //         }
            //         else if (applist[0] == "WPS China")
            //         {
            //             correctString = applist[0].Replace("WPS China", "WPS Office");
            //             int skip_char = applist[1].IndexOf("r");
            //             if (skip_char > 0)
            //             {
            //                 applist[1] = applist[1].Substring(0, skip_char);
            //             }
            //         }
            //         else if (applist[0] == "Baidu Weblink" || applist[0] == "Baidu_Favorite")
            //         {
            //             correctString = applist[0].Replace("Baidu Weblink", "百度一下");
            //             correctString = applist[0].Replace("Baidu_Favorite", "百度一下");

            //         }
            //         else if (applist[0] == "Yandex Browser")
            //         {
            //             correctString = applist[0].Replace("Yandex Browser", "Yandex");
            //         }
            //         else if (applist[0] == "Yandex Weblink")
            //         {
            //             correctString = applist[0].Replace("Yandex Weblink", "Yandex");
            //         }
            //         else if (applist[0] == "Yandex_Favorite")
            //         {
            //             correctString = applist[0].Replace("Yandex_Favorite", "Yandex");
            //         }
            //         else
            //         {
            //             correctString = applist[0];
            //         }

            //         try
            //         {
            //             // Read the file line by line
            //             using (StreamReader reader = new StreamReader(filePath))
            //             {
            //                 int lineNumber = 1;
            //                 string line;

            //                 while ((line = reader.ReadLine()) != null)
            //                 {
            //                     if (line.IndexOf(correctString, StringComparison.CurrentCultureIgnoreCase) >= 0)
            //                     {
            //                         if (correctString == "Care Center" ||
            //                             correctString == "Acer Configuration Manager" ||
            //                             correctString == "AcerRegistration" ||
            //                             correctString == "AcerSense" ||
            //                             correctString == "App Explorer" ||
            //                             correctString == "百度网盘" ||
            //                             correctString == "软件精选" ||
            //                             correctString == "WPS Office" ||
            //                             correctString == "DropboxOEM" ||
            //                             correctString == "Evernote" ||
            //                             correctString == "Google Play Games" ||
            //                             correctString == "IntelTechnologyMDE" ||
            //                             correctString == "McAfee LiveSafe" ||
            //                             correctString == "Microsoft 365" ||
            //                             correctString == "MyOffice" ||
            //                             correctString == "Spades" ||
            //                             correctString == "UserExperienceImprovementProgram")
            //                         {
            //                             if (line.IndexOf(applist[1]) >= 0)//check version correct
            //                             {
            //                                 int index = -1;//index pos for string "not check" or "checked"
            //                                 switch (correctString)
            //                                 {
            //                                     case "Care Center":
            //                                         AcerCareCenterV4 = "checked";
            //                                         index = result_applications_list.FindIndex(s => s == "Acer Care Center V4,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "Acer Care Center V4,checked";
            //                                         }
            //                                         //result_applications_list.Add("AcerCareCenterV4" + "," + AcerCareCenterV4);
            //                                         break;
            //                                     case "Acer Configuration Manager":
            //                                         AcerConfigurationManager = "checked";
            //                                         index = result_applications_list.FindIndex(s => s == "Acer Configuration Manager,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "Acer Configuration Manager,checked";
            //                                         }
            //                                         break;
            //                                     case "AcerRegistration":
            //                                         AcerProductRegistration = "checked";
            //                                         index = result_applications_list.FindIndex(s => s == "Acer Product Registration,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "Acer Product Registration,checked";
            //                                         }
            //                                         break;
            //                                     case "AcerSense":
            //                                         AcerSense = "checked";
            //                                         index = result_applications_list.FindIndex(s => s == "AcerSense,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "AcerSense,checked";
            //                                         }
            //                                         break;
            //                                     case "App Explorer":
            //                                         AppExplorer = "checked";
            //                                         index = result_applications_list.FindIndex(s => s == "App Explorer,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "App Explorer,checked";
            //                                         }
            //                                         break;
            //                                     case "百度网盘":
            //                                         BaiduNetDisk = "checked";
            //                                         index = result_applications_list.FindIndex(s => s == "Baidu Net Disk,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "Baidu Net Disk,checked";
            //                                         }
            //                                         break;
            //                                     case "软件精选":
            //                                         SoftwarePicks = "checked";
            //                                         index = result_applications_list.FindIndex(s => s == "Software Picks,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "Software Picks,checked";
            //                                         }
            //                                         break;
            //                                     case "DropboxOEM":
            //                                         DropboxPromotion = "checked";
            //                                         index = result_applications_list.FindIndex(s => s == "Dropbox Promotion,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "Dropbox Promotion,checked";
            //                                         }
            //                                         break;
            //                                     case "Evernote":
            //                                         Evernote = "checked";
            //                                         index = result_applications_list.FindIndex(s => s == "Evernote,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "Evernote,checked";
            //                                         }
            //                                         break;
            //                                     case "Google Play Games":
            //                                         GooglePlayGames = "checked";
            //                                         index = result_applications_list.FindIndex(s => s == "Google Play Games,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "Google Play Games,checked";
            //                                         }
            //                                         break;
            //                                     case "IntelTechnologyMDE":
            //                                         IntelMultiDeviceExperience = "checked";
            //                                         index = result_applications_list.FindIndex(s => s == "Intel Multi-Device Experience,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "Intel Multi-Device Experience,checked";
            //                                         }
            //                                         break;
            //                                     case "McAfee LiveSafe":
            //                                         McAfeeLiveSafeP1 = "checked";
            //                                         index = result_applications_list.FindIndex(s => s == "McAfee LiveSafe P1,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "McAfee LiveSafe P1,checked";
            //                                         }
            //                                         break;
            //                                     case "Microsoft 365":
            //                                         MicrosoftOffice = "checked";
            //                                         index = result_applications_list.FindIndex(s => s == "Microsoft Office,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "Microsoft Office,checked";
            //                                         }
            //                                         break;
            //                                     case "MyOffice":
            //                                         MyOfficesuite = "checked";
            //                                         index = result_applications_list.FindIndex(s => s == "MyOffice_Spreadsheet,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "MyOffice_Spreadsheet,checked";
            //                                         }

            //                                         index = result_applications_list.FindIndex(s => s == "MyOffice_Text,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "MyOffice_Text,checked";
            //                                         }
            //                                         break;
            //                                     case "WPS Office":
            //                                         WPSChina = "checked";
            //                                         index = result_applications_list.FindIndex(s => s == "WPS China,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "WPS China,checked";
            //                                         }
            //                                         break;
            //                                     case "Spades":
            //                                         Spades = "checked";
            //                                         index = result_applications_list.FindIndex(s => s == "Spades,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "Spades,checked";
            //                                         }
            //                                         break;
            //                                     case "UserExperienceImprovementProgram":
            //                                         UserBehaviorTrackingFramework = "checked";
            //                                         index = result_applications_list.FindIndex(s => s == "User Behavior Tracking Framework,not check");
            //                                         if (index != -1)
            //                                         {
            //                                             // Modify the value at the found index
            //                                             result_applications_list[index] = "User Behavior Tracking Framework,checked";
            //                                         }
            //                                         break;
            //                                     default://no case mapping, do nothing here
            //                                         break;
            //                                 };

            //                                 Console.WriteLine($"Found '{applist[0]}'-->mapping name '{correctString}', '{applist[1]}' in line {lineNumber}: {line}");
            //                             }
            //                         }
            //                         else if (correctString == "Acer Legal Information" ||
            //                             correctString == "Amazon.url" ||
            //                             correctString == "Booking.com.url" ||
            //                             correctString == "hao123" ||
            //                             correctString == "百度一下" ||
            //                             correctString == "Forge of Empires" ||
            //                             correctString == "Microsoft OneNote" ||
            //                             correctString == "Morpho Inc." ||
            //                             correctString == "Planet9_Version.txt" ||
            //                             correctString == "PowerManagementSetting.tag" ||
            //                             correctString == "StorageUtilities.tag")
            //                         {
            //                             int index = -1;
            //                             switch (correctString)//case does not need to check version
            //                             {
            //                                 case "Acer Legal Information":
            //                                     Acerlegalportal = "checked";
            //                                     index = result_applications_list.FindIndex(s => s == "Acer legal portal,not check");
            //                                     if (index != -1)
            //                                     {
            //                                         // Modify the value at the found index
            //                                         result_applications_list[index] = "Acer legal portal,checked";
            //                                     }
            //                                     break;
            //                                 case "Amazon.url":
            //                                     AmazonWeblink = "checked";
            //                                     index = result_applications_list.FindIndex(s => s == "Amazon Weblink,not check");
            //                                     if (index != -1)
            //                                     {
            //                                         // Modify the value at the found index
            //                                         result_applications_list[index] = "Amazon Weblink,checked";
            //                                     }
            //                                     break;
            //                                 case "Booking.com.url":
            //                                     BookingWeblink = "checked";
            //                                     index = result_applications_list.FindIndex(s => s == "Booking.com Weblink,not check");
            //                                     if (index != -1)
            //                                     {
            //                                         // Modify the value at the found index
            //                                         result_applications_list[index] = "Booking.com Weblink,checked";
            //                                     }
            //                                     BookingFavorite = "checked";
            //                                     index = result_applications_list.FindIndex(s => s == "Booking.com_Favorite,not check");
            //                                     if (index != -1)
            //                                     {
            //                                         // Modify the value at the found index
            //                                         result_applications_list[index] = "Booking.com_Favorite,checked";
            //                                     }
            //                                     break;
            //                                 case "hao123":
            //                                     Hao123Weblink = "checked";
            //                                     index = result_applications_list.FindIndex(s => s == "Hao123 Weblink,not check");
            //                                     if (index != -1)
            //                                     {
            //                                         // Modify the value at the found index
            //                                         result_applications_list[index] = "Hao123 Weblink,checked";
            //                                     }
            //                                     Hao123Favorite = "checked";
            //                                     index = result_applications_list.FindIndex(s => s == "Hao123_Favorite,not check");
            //                                     if (index != -1)
            //                                     {
            //                                         // Modify the value at the found index
            //                                         result_applications_list[index] = "Hao123_Favorite,checked";
            //                                     }
            //                                     break;
            //                                 case "百度一下":
            //                                     BaiduWeblink = "checked";
            //                                     index = result_applications_list.FindIndex(s => s == "Baidu Weblink,not check");
            //                                     if (index != -1)
            //                                     {
            //                                         // Modify the value at the found index
            //                                         result_applications_list[index] = "Baidu Weblink,checked";
            //                                     }
            //                                     BaiduFavorite = "checked";
            //                                     index = result_applications_list.FindIndex(s => s == "Baidu_Favorite,not check");
            //                                     if (index != -1)
            //                                     {
            //                                         // Modify the value at the found index
            //                                         result_applications_list[index] = "Baidu_Favorite,checked";
            //                                     }
            //                                     break;
            //                                 case "Forge of Empires":
            //                                     ForgeofEmpires = "checked";
            //                                     index = result_applications_list.FindIndex(s => s == "Forge of Empires,not check");
            //                                     if (index != -1)
            //                                     {
            //                                         // Modify the value at the found index
            //                                         result_applications_list[index] = "Forge of Empires,checked";
            //                                     }
            //                                     break;
            //                                 case "Microsoft OneNote":
            //                                     MicrosoftOfficeInstaller = "checked";
            //                                     index = result_applications_list.FindIndex(s => s == "Microsoft Office Installer,not check");
            //                                     if (index != -1)
            //                                     {
            //                                         // Modify the value at the found index
            //                                         result_applications_list[index] = "Microsoft Office Installer,checked";
            //                                     }
            //                                     break;
            //                                 case "Morpho Inc.":
            //                                     MorphoCameraEffects = "checked";
            //                                     index = result_applications_list.FindIndex(s => s == "Morpho Camera Effects,not check");
            //                                     if (index != -1)
            //                                     {
            //                                         // Modify the value at the found index
            //                                         result_applications_list[index] = "Morpho Camera Effects,checked";
            //                                     }
            //                                     break;
            //                                 case "Planet9_Version.txt":
            //                                     Planet9 = "checked";
            //                                     index = result_applications_list.FindIndex(s => s == "Planet9,not check");
            //                                     if (index != -1)
            //                                     {
            //                                         // Modify the value at the found index
            //                                         result_applications_list[index] = "Planet9,checked";
            //                                     }
            //                                     break;
            //                                 case "PowerManagementSetting.tag":
            //                                     PowerManagementSetting = "checked";
            //                                     index = result_applications_list.FindIndex(s => s == "Power Management Setting,not check");
            //                                     if (index != -1)
            //                                     {
            //                                         // Modify the value at the found index
            //                                         result_applications_list[index] = "Power Management Setting,checked";
            //                                     }
            //                                     break;
            //                                 case "StorageUtilities.tag":
            //                                     StorageUtilities = "checked";
            //                                     index = result_applications_list.FindIndex(s => s == "Storage Utilities,not check");
            //                                     if (index != -1)
            //                                     {
            //                                         // Modify the value at the found index
            //                                         result_applications_list[index] = "Storage Utilities,checked";
            //                                     }
            //                                     break;
            //                                 default://no case mapping, do nothing here 
            //                                     break;
            //                             };
            //                             Console.WriteLine($"Found '{applist[0]}'-->mapping name '{correctString}' in line {lineNumber}: {line}");
            //                         }
            //                         //language/Region is not en-US 
            //                         else if (correctString == "RandomSaladGamesLLC")//Solitaire, ru-RU
            //                         {
            //                             if (line.IndexOf(applist[1]) >= 0)
            //                             {
            //                                 Solitaire = "checked";
            //                                 int index = -1;
            //                                 index = result_applications_list.FindIndex(s => s == "Solitaire,not check");
            //                                 if (index != -1)
            //                                 {
            //                                     // Modify the value at the found index
            //                                     result_applications_list[index] = "Solitaire,checked";
            //                                 }
            //                                 Console.WriteLine($"Found '{applist[0]}'-->mapping name '{correctString}', '{applist[1]}' in line {lineNumber}: {line}");
            //                             }
            //                         }
            //                         else if (correctString == "Yandex" && ci.Name == "ru-RU")
            //                         {
            //                             Yandex = "checked";
            //                             int index = -1; int index2 = -1; int index3 = -1;
            //                             index = result_applications_list.FindIndex(s => s == "Yandex Browser,not check");
            //                             if (index != -1)
            //                             {
            //                                 // Modify the value at the found index
            //                                 result_applications_list[index] = "Yandex Browser,checked";
            //                             }
            //                             index2 = result_applications_list.FindIndex(s => s == "Yandex Weblink,not check");
            //                             if (index2 != -1)
            //                             {
            //                                 // Modify the value at the found index
            //                                 result_applications_list[index2] = "Yandex Weblink,checked";
            //                             }
            //                             index3 = result_applications_list.FindIndex(s => s == "Yandex_Favorite,not check");
            //                             if (index3 != -1)
            //                             {
            //                                 // Modify the value at the found index
            //                                 result_applications_list[index3] = "Yandex_Favorite,checked";
            //                             }
            //                         }
            //                         else
            //                         {
            //                             Console.WriteLine("Error !! Not in the installed applist" + " " + correctString);
            //                         }

            //                     }
            //                     lineNumber++;
            //                 }
            //             }
            //         }
            //         catch (Exception ex)
            //         {
            //             Console.WriteLine($"An error occurred: {ex.Message}");
            //         }
            //     }
            // }

            // foreach (string item in result_applications_list)
            // {
            //     if (item.IndexOf("not check") >0)
            //     {
            //         fail++;
            //         Console.WriteLine($"{item} error!");
            //     }

            // }

            // if (fail > 0)
            // {
            //     Console.WriteLine("Not all apps checked OK Fail!");
            //     return false;
            // }
            // else 
            // {
            //     //Console.WriteLine($"fail:{fail}");
            //     Console.WriteLine("All apps checked OK Success !!");
            //     return true;
            // }
            return true;
        }
        

        public static void UpdateResults() 
        {
            Console.WriteLine("UpdateResults");
        }
        public static void TearDown() 
        {
            Console.WriteLine("TearDown");
        }

        static string[] setup1()
        {
            string project_name = null;
            string computerName = Environment.MachineName;
            //Console.WriteLine("Computer Name: " + computerName);
            string registryKeyPath = @"HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS"; // Replace with the actual registry key path
            string valueName = "SystemProductName"; // Replace with the name of the specific value you want to retrieve
            String[] setupvalues = new string[2];
            try
            {
                // Use Registry.GetValue to retrieve the value of the specified key
                object value = Registry.GetValue(registryKeyPath, valueName, null);

                if (value != null)
                {
                    project_name = value.ToString();
                    int index = project_name.IndexOf(" ");
                    if (index > 0)
                    {
                        project_name = project_name.Substring(0, index);
                    }

                    //Console.WriteLine($"Registry Value ({valueName}): {value} : {project_name}");
                    setupvalues[0] = project_name;
                    setupvalues[1] = computerName;
                }
                else
                {
                    Console.WriteLine($"Value '{valueName}' not found in the registry key.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
            return setupvalues;
        }

        static void GetInstalledAppsListByLanguage(string language_Region,
            string application_name_cellValue,
            string application_version_cellValue,
            string applicationID_version_cellValue,
            List<string> applications_list_SCL,
            List<string> result_applications_list)
        {
            string application_version_SCL = "";
            if (language_Region == "en-US")//check language and region
            {
                if (application_name_cellValue == "AcerWarrantyRegistration for India") //India only
                { }
                else if (application_name_cellValue == "Agoda Weblink (2023.Q2)") //global
                { }
                else if (application_name_cellValue == "Agoda_Favorite (2023.Q2)")//global
                { }
                else if (application_name_cellValue == "Baidu Net Disk (2023.Q2)")//China
                { }
                else if (application_name_cellValue == "Baidu Weblink (2023.Q2)")//China
                { }
                else if (application_name_cellValue == "Baidu_Favorite (2023.Q2)")//China
                { }
                else if (application_name_cellValue == "Hao123 Weblink (2023.Q2)")//zh-CN
                { }
                else if (application_name_cellValue == "Hao123_Favorite (2023.Q2)")//zh-CN
                { }
                else if (application_name_cellValue == "WPS China (2023.Q2)")//China
                { }
                else if (application_name_cellValue == "Yandex Browser (2023.Q2)")//ru-RU
                { }
                else if (application_name_cellValue == "Yandex Weblink (2023.Q2)")//ru-RU
                { }
                else if (application_name_cellValue == "Yandex_Favorite (2023.Q2)")//ru-RU
                { }
                else if (application_name_cellValue == "MyOffice_Spreadsheet (2023.Q2)")//ru-RU
                { }
                else if (application_name_cellValue == "MyOffice_Text (2023.Q2)")//ru-RU
                { }
                else if (application_name_cellValue == "Software Picks (2023.Q2)")//China
                { }
                else//other applications will be installed in the os
                {
                    if (application_version_cellValue.Length > 1)
                    {
                        application_version_SCL = application_version_cellValue.Substring(1);
                    }
                    int index_skip = application_name_cellValue.IndexOf('(');//Skip read string after "(" char
                    int index_skip_version = application_version_SCL.IndexOf("(");
                    //int index_skip_version2 = application_version_SCL.IndexOf("_");
                    string result_version = " ";
                    string result_version2 = "";
                    if (index_skip >= 0)//find '(' char in cell, application_name_cellValue
                    {
                        string result = application_name_cellValue.Substring(0, index_skip);//skip '(' char
                        string result2 = result.Substring(0, result.Length - 1);//skip " " for string format

                        if (index_skip_version >= 0)//find '(' char in cell, application_version_SCL
                        {
                            result_version = application_version_SCL.Substring(0, index_skip_version);//skip '(' char

                        }
                        else
                        {
                            result_version = application_version_SCL;//not contain '(' or '_' , just add
                        }
                        //combine name + version, split by ','
                        applications_list_SCL.Add(result2 + "," + result_version + "," + applicationID_version_cellValue);
                        result_applications_list.Add(result2 + "," + "not check");

                    }
                    else // case not find '(' char in cell, application_name_cellValue
                    {
                        //name without '(' + version without '(', '_'
                        result_version2 = application_version_SCL.Replace(" ", "");
                        applications_list_SCL.Add(application_name_cellValue + "," + result_version2 + "," + applicationID_version_cellValue);
                        result_applications_list.Add(application_name_cellValue + "," + "not check");
                    }
                }
            }//end of lanaguage en-US
            if (language_Region == "ru-RU")
            {
                if (application_name_cellValue == "AcerWarrantyRegistration for India") //India only
                { }
                else if (application_name_cellValue == "Agoda Weblink (2023.Q2)") //global
                { }
                else if (application_name_cellValue == "Agoda_Favorite (2023.Q2)")//global
                { }
                else if (application_name_cellValue == "Baidu Net Disk (2023.Q2)")//China
                { }
                else if (application_name_cellValue == "Baidu Weblink (2023.Q2)")//China
                { }
                else if (application_name_cellValue == "Baidu_Favorite (2023.Q2)")//China
                { }
                else if (application_name_cellValue == "Hao123 Weblink (2023.Q2)")//zh-CN
                { }
                else if (application_name_cellValue == "Hao123_Favorite (2023.Q2)")//zh-CN
                { }
                else if (application_name_cellValue == "WPS China (2023.Q2)")//China
                { }
                else if (application_name_cellValue == "Software Picks (2023.Q2)")//China
                { }
                else if (application_name_cellValue == "Amazon Weblink (2023.Q2)")//en-US
                { }
                else if (application_name_cellValue == "Google Play Games (2023.Q2)")//en-US
                { }
                else//other applications will be installed in the os
                {
                    if (application_version_cellValue.Length > 1)
                    {
                        application_version_SCL = application_version_cellValue.Substring(1);
                    }
                    int index_skip = application_name_cellValue.IndexOf('(');//Skip read string after "(" char
                    int index_skip_version = application_version_SCL.IndexOf("(");
                    //int index_skip_version2 = application_version_SCL.IndexOf("_");
                    string result_version = " ";
                    string result_version2 = "";
                    if (index_skip >= 0)//find '(' char in cell, application_name_cellValue
                    {
                        string result = application_name_cellValue.Substring(0, index_skip);//skip '(' char
                        string result2 = result.Substring(0, result.Length - 1);//skip " " for string format

                        if (index_skip_version >= 0)//find '(' char in cell, application_version_SCL
                        {
                            result_version = application_version_SCL.Substring(0, index_skip_version);//skip '(' char

                        }
                        else
                        {
                            result_version = application_version_SCL;//not contain '(' or '_' , just add
                        }
                        //combine name + version, split by ','
                        applications_list_SCL.Add(result2 + "," + result_version + "," + applicationID_version_cellValue);
                        result_applications_list.Add(result2 + "," + "not check");

                    }
                    else // case not find '(' char in cell, application_name_cellValue
                    {

                        result_version2 = application_version_SCL.Replace(" ", "");

                        applications_list_SCL.Add(application_name_cellValue + "," + result_version2 + "," + applicationID_version_cellValue);
                        result_applications_list.Add(application_name_cellValue + "," + "not check");
                    }
                }
            }//end of lanaguage ru-RU
            if (language_Region == "zh-CN")
            {
                if (application_name_cellValue == "AcerWarrantyRegistration for India") //India only
                { }
                else if (application_name_cellValue == "Agoda Weblink (2023.Q2)") //global
                { }
                else if (application_name_cellValue == "Agoda_Favorite (2023.Q2)")//global
                { }
                else if (application_name_cellValue == "Amazon Weblink (2023.Q2)")//en-US
                { }
                else if (application_name_cellValue == "Google Play Games (2023.Q2)")//en-US
                { }
                else if (application_name_cellValue == "MyOffice_Spreadsheet (2023.Q2)")//ru-RU
                { }
                else if (application_name_cellValue == "MyOffice_Text (2023.Q2)")//ru-RU
                { }
                else if (application_name_cellValue == "Yandex Browser (2023.Q2)")//ru-RU
                { }
                else if (application_name_cellValue == "Yandex Weblink (2023.Q2)")//ru-RU
                { }
                else if (application_name_cellValue == "Yandex_Favorite (2023.Q2)")//ru-RU
                { }
                else if (application_name_cellValue == "McAfee LiveSafe P1 (2023.Q2)")//en-US
                { }
                else if (application_name_cellValue == "Evernote (2023.Q2)")//en-US
                { }
                else if (application_name_cellValue == "Forge of Empires (2023.Q2)")//en-US
                { }
                else if (application_name_cellValue == "Dropbox Promotion (2023.Q2)")//en-US
                { }
                else if (application_name_cellValue == "Planet9 (For Consumer)")//en-US
                { }
                else if (application_name_cellValue == "Solitaire (2023.Q2)")//en-US
                { }
                else if (application_name_cellValue == "Spades (2023.Q2)")//en-US
                { }
                else//other applications will be installed in the os
                {
                    if (application_version_cellValue.Length > 1)
                    {
                        application_version_SCL = application_version_cellValue.Substring(1);
                    }
                    int index_skip = application_name_cellValue.IndexOf('(');//Skip read string after "(" char
                    int index_skip_version = application_version_SCL.IndexOf("(");
                    int index_skip_version_r = application_version_SCL.IndexOf("r");
                    //int index_skip_version2 = application_version_SCL.IndexOf("_");
                    string result_version = " ";
                    string result_version2 = "";
                    if (index_skip >= 0)//find '(' char in cell, application_name_cellValue
                    {
                        string result = application_name_cellValue.Substring(0, index_skip);//skip '(' char
                        string result2 = result.Substring(0, result.Length - 1);//skip " " for string format

                        if (index_skip_version >= 0)//find '(' char in cell, application_version_SCL
                        {
                            result_version = application_version_SCL.Substring(0, index_skip_version);//skip '(' char
                        }
                        else if (index_skip_version_r >= 0)
                        {
                            result_version = application_version_SCL.Substring(0, index_skip_version_r);
                            Console.WriteLine($"WPS China refine --> version --> {result_version}");
                        }
                        else
                        {
                            result_version = application_version_SCL;//not contain '(' or '_' , just add
                        }

                        //combine name + version, split by ','
                        applications_list_SCL.Add(result2 + "," + result_version + "," + applicationID_version_cellValue);
                        result_applications_list.Add(result2 + "," + "not check");

                    }
                    else // case not find '(' char in cell, application_name_cellValue
                    {

                        result_version2 = application_version_SCL.Replace(" ", "");

                        applications_list_SCL.Add(application_name_cellValue + "," + result_version2 + "," + applicationID_version_cellValue);
                        result_applications_list.Add(application_name_cellValue + "," + "not check");
                    }
                }
            }//end of lanaguage zh-CN
        }

        static List<string> GetinstalledinformationFromPC(List<string> gInstalledSoftware)
        {
            string displayName;
            string app_version;
            string app_vendor;
            string ItemName;
            string ItemFavIconFile;
            string strSystemComponent;

            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", false))
            {
                foreach (String keyName in key.GetSubKeyNames())
                {
                    RegistryKey subkey = key.OpenSubKey(keyName);
                    displayName = subkey.GetValue("DisplayName") as string;
                    app_version = subkey.GetValue("DisplayVersion") as string;
                    app_vendor = subkey.GetValue("Publisher") as string;
                    strSystemComponent = subkey.GetValue("SystemComponent") as string;

                    if (string.IsNullOrEmpty(displayName))
                        continue;

                    gInstalledSoftware.Add(app_vendor + "," + displayName + "," + app_version + "," + strSystemComponent);

                }
            }

            using (var localMachine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            {
                var key = localMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", false);
                foreach (String keyName in key.GetSubKeyNames())
                {
                    RegistryKey subkey = key.OpenSubKey(keyName);
                    displayName = subkey.GetValue("DisplayName") as string;
                    app_version = subkey.GetValue("DisplayVersion") as string;
                    app_vendor = subkey.GetValue("Publisher") as string;
                    strSystemComponent = subkey.GetValue("SystemComponent") as string;

                    if (string.IsNullOrEmpty(displayName))
                        continue;

                    gInstalledSoftware.Add(app_vendor + "," + displayName + "," + app_version + "," + strSystemComponent);


                }
            }

            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall", false))
            {
                foreach (String keyName in key.GetSubKeyNames())
                {
                    RegistryKey subkey = key.OpenSubKey(keyName);
                    displayName = subkey.GetValue("DisplayName") as string;
                    app_version = subkey.GetValue("DisplayVersion") as string;
                    app_vendor = subkey.GetValue("Publisher") as string;
                    strSystemComponent = subkey.GetValue("SystemComponent") as string;
                    //Console.WriteLine("strSystemComponent: {0}", strSystemComponent);
                    if (string.IsNullOrEmpty(displayName))
                        continue;

                    gInstalledSoftware.Add(app_vendor + "," + displayName + "," + app_version + "," + strSystemComponent);

                }
            }

            using (RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\PackageRepository\Packages"))
            {
                if (registryKey != null)
                {
                    // Get the names of the subkeys
                    string[] subKeyNames = registryKey.GetSubKeyNames();
                    // Display the subkey names
                    foreach (string subKeyName in subKeyNames)
                    {
                        gInstalledSoftware.Add(subKeyName);
                    }
                }
                else
                {
                    Console.WriteLine("Registry Key not found.");
                }
            }

            using (RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(@"SYSTEM\DriverDatabase\DriverPackages"))
            {
                if (registryKey != null)
                {
                    // Get the names of the subkeys
                    string[] subKeyNames = registryKey.GetSubKeyNames();
                    // Display the subkey names
                    foreach (string subKeyName in subKeyNames)
                    {

                        RegistryKey subkey = registryKey.OpenSubKey(subKeyName);
                        string Provider = subkey.GetValue("Provider") as string;
                        //Console.WriteLine($"{subKeyName}" + "," +$"{Provider}" );
                        if (Provider == "Morpho Inc.")
                        {
                            gInstalledSoftware.Add(subKeyName + "," + Provider);
                        }
                    }
                }
            }

            using (var localMachine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            {
                var key = localMachine.OpenSubKey(@"SOFTWARE\Microsoft\MicrosoftEdge\Main\FavoriteBarItems", false);
                foreach (String keyName in key.GetSubKeyNames())
                {
                    RegistryKey subkey = key.OpenSubKey(keyName);
                    ItemName = subkey.GetValue("ItemName") as string;
                    ItemFavIconFile = subkey.GetValue("ItemFavIconFile") as string;


                    //Console.WriteLine("strSystemComponent: {0}", strSystemComponent);
                    if (string.IsNullOrEmpty(ItemName))
                        continue;

                    gInstalledSoftware.Add(ItemName + "," + ItemFavIconFile);


                }
            }

            string startMenuPath = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu);

            string directoryPath = @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"; // Replace with the directory path you want to search.
            string directoryPath2 = @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Acer";
            string directoryPath3 = @"C:\OEM\Preload";
            string directoryPath4 = @"C:\OEM\";

            try
            {
                string[] fileNames = Directory.GetFiles(directoryPath);
                string[] fileNames2 = Directory.GetFiles(directoryPath2);
                string[] fileNames3 = Directory.GetFiles(directoryPath3);
                string[] fileNames4 = Directory.GetFiles(directoryPath4);
                foreach (string fileName in fileNames)
                {
                    gInstalledSoftware.Add(fileName);
                }

                foreach (string fileName in fileNames2)
                {
                    gInstalledSoftware.Add(fileName);
                }

                foreach (string fileName in fileNames3)
                {
                    gInstalledSoftware.Add(fileName);
                }

                foreach (string fileName in fileNames4)
                {
                    gInstalledSoftware.Add(fileName);

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }

            return gInstalledSoftware;
        }

        static List<string> GetAcerTagfiles(string directoryPath)
        {
            //string directoryPath = @"C:\YourDirectory"; // Replace with the path to your specific directory
            List<string> apptagslist = new List<string>();
            if (Directory.Exists(directoryPath))
            {
                string[] tagFiles = Directory.GetFiles(directoryPath, "*.tag");

                if (tagFiles.Length > 0)
                {
                    Console.WriteLine("Files with .tag extension:");

                    foreach (string file in tagFiles)
                    {
                        //Console.WriteLine(file);
                        string refine1 = file.Replace("C:\\OEM\\Preload\\InstalledApps\\", "");
                        string refine2 = refine1.Replace(".tag", "");
                        apptagslist.Add(refine2);
                    }
                }
                else
                {
                    Console.WriteLine("No .tag files found in the directory.");
                }
            }
            else
            {
                Console.WriteLine("The specified directory does not exist.");
            }
            return apptagslist;
        }
        static string appNamingReprocess(string application_name_cellValue, string application_version_cellValue)
        {
            string application_version_SCL = "";

            if (application_version_cellValue.Length > 1)
            {
                application_version_SCL = application_version_cellValue.Substring(1);
            }
            int index_skip = application_name_cellValue.IndexOf('(');//Skip read string after "(" char
            int index_skip_version = application_version_SCL.IndexOf("(");
            //int index_skip_version2 = application_version_SCL.IndexOf("_");
            string result_version = " ";
            string result_version2 = "";
            if (index_skip >= 0)//find '(' char in cell, application_name_cellValue
            {
                string result = application_name_cellValue.Substring(0, index_skip);//skip '(' char
                string result2 = result.Substring(0, result.Length - 1);//skip " " for string format

                if (index_skip_version >= 0)//find '(' char in cell, application_version_SCL
                {
                    result_version = application_version_SCL.Substring(0, index_skip_version);//skip '(' char

                }

                else
                {
                    result_version = application_version_SCL;//not contain '(' or '_' , just add
                }
                //combine name + version, split by ','
                Console.WriteLine(result2 + "," + result_version);
                return result2 + "," + result_version;

            }
            else // case not find '(' char in cell, application_name_cellValue
            {
                //name without '(' + version without '(', '_'
                Console.WriteLine(application_version_SCL);
                result_version2 = application_version_SCL.Replace(" ", "");
                Console.WriteLine(result_version2);
                //applications_list_SCL.Add(application_name_cellValue + "," + result_version2);
                Console.WriteLine(application_name_cellValue + "," + result_version2);
                return application_name_cellValue + "," + result_version2;

            }
        }
        static void ListStartMenuItems(string directory)
        {
            try
            {
                foreach (string file in Directory.GetFiles(directory, "*.lnk"))
                {
                    Console.WriteLine(Path.GetFileNameWithoutExtension(file));
                }

                foreach (string subDir in Directory.GetDirectories(directory))
                {
                    ListStartMenuItems(subDir);
                }
            }
            catch (UnauthorizedAccessException)
            {
                // Handle any permission-related issues, if needed
            }
        }
    }
}
