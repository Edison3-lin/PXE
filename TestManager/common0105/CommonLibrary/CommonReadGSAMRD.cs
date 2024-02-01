using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Net;

namespace CaptainWin.CommonAPI{
    internal class CommonReadGSAMRD{
        //17+1 for align row of metro_lang_check_table
        public static string[,] metro_table = new string[18, 2];
        public static string[,] metro_lang_check_table = new string[18, 15];
        //11+1 for align row of desktop_lang_check_table
        public static string[,] desktop_table = new string[12, 2];
        public static string[,] desktop_lang_check_table = new string[12, 15];
        //8+1 for align row of desktop_bar_lang_check_table
        public static string[,] desktop_bar_table = new string[9, 2];
        public static string[,] desktop_bar_lang_check_table = new string[9, 15];
        //21+1 for align row of shortCut_lang_chekc_table
        public static string[,] shortCut_table = new string[22, 2];
        public static string[,] shortCut_lang_check_table = new string[22, 15];
        //5+1 for align row of browser_lang_check_table
        public static string[,] browser_table = new string[6, 2];
        public static string[,] browser_lang_check_table = new string[6, 15];

        public static string[,] OOBE_table = new String[2, 2];
        public static string[,] OOBE_lang_check_table = new string[2, 15];

        public static string[,] next_recommended_table = new string[3, 2];
        public static string[,] next_recommended_lang_check_table = new string[3, 15];

        //list of metro apps
        public static List<string> metro_app_list = new List<string>();
        public static List<string> us_metro_app_list = new List<string>();
        public static List<string> ca_metro_app_list = new List<string>();
        public static List<string> latam_metro_app_list = new List<string>();
        public static List<string> frfr_metro_app_list = new List<string>();
        public static List<string> dede_metro_app_list = new List<string>();
        public static List<string> gb_metro_app_list = new List<string>();
        public static List<string> nordic_metro_app_list = new List<string>();
        public static List<string> ru_metro_app_list = new List<string>();
        public static List<string> emea1_metro_app_list = new List<string>();
        public static List<string> au_metro_app_list = new List<string>();
        public static List<string> jp_metro_app_list = new List<string>();
        public static List<string> kr_metro_app_list = new List<string>();
        public static List<string> aap1_metro_app_list = new List<string>();
        public static List<string> cn_metro_app_list = new List<string>();
        public static List<string> tw_metro_app_list = new List<string>();
        //list of desktop apps
        public static List<string> desktop_app_list = new List<string>();
        public static List<string> us_desktop_app_list = new List<string>();
        public static List<string> ca_desktop_app_list = new List<string>();
        public static List<string> latam_desktop_app_list = new List<string>();
        public static List<string> frfr_desktop_app_list = new List<string>();
        public static List<string> dede_desktop_app_list = new List<string>();
        public static List<string> gb_desktop_app_list = new List<string>();
        public static List<string> nordic_desktop_app_list = new List<string>();
        public static List<string> ru_desktop_app_list = new List<string>();
        public static List<string> emea1_desktop_app_list = new List<string>();
        public static List<string> au_desktop_app_list = new List<string>();
        public static List<string> jp_desktop_app_list = new List<string>();
        public static List<string> kr_desktop_app_list = new List<string>();
        public static List<string> aap1_desktop_app_list = new List<string>();
        public static List<string> cn_desktop_app_list = new List<string>();
        public static List<string> tw_desktop_app_list = new List<string>();
        //List of desktop_bar_app_list
        public static List<string> desktop_bar_app_list = new List<string>();
        public static List<string> us_desktop_bar_app_list = new List<string>();
        public static List<string> ca_desktop_bar_app_list = new List<string>();
        public static List<string> latam_desktop_bar_app_list = new List<string>();
        public static List<string> frfr_desktop_bar_app_list = new List<string>();
        public static List<string> dede_desktop_bar_app_list = new List<string>();
        public static List<string> gb_desktop_bar_app_list = new List<string>();
        public static List<string> nordic_desktop_bar_app_list = new List<string>();
        public static List<string> ru_desktop_bar_app_list = new List<string>();
        public static List<string> emea1_desktop_bar_app_list = new List<string>();
        public static List<string> au_desktop_bar_app_list = new List<string>();
        public static List<string> jp_desktop_bar_app_list = new List<string>();
        public static List<string> kr_desktop_bar_app_list = new List<string>();
        public static List<string> aap1_desktop_bar_app_list = new List<string>();
        public static List<string> cn_desktop_bar_app_list = new List<string>();
        public static List<string> tw_desktop_bar_app_list = new List<string>();

        //list of desktop apps
        public static List<string> shortCut_app_list = new List<string>();
        public static List<string> us_shortCut_app_list = new List<string>();
        public static List<string> ca_shortCut_app_list = new List<string>();
        public static List<string> latam_shortCut_app_list = new List<string>();
        public static List<string> frfr_shortCut_app_list = new List<string>();
        public static List<string> dede_shortCut_app_list = new List<string>();
        public static List<string> gb_shortCut_app_list = new List<string>();
        public static List<string> nordic_shortCut_app_list = new List<string>();
        public static List<string> ru_shortCut_app_list = new List<string>();
        public static List<string> emea1_shortCut_app_list = new List<string>();
        public static List<string> au_shortCut_app_list = new List<string>();
        public static List<string> jp_shortCut_app_list = new List<string>();
        public static List<string> kr_shortCut_app_list = new List<string>();
        public static List<string> aap1_shortCut_app_list = new List<string>();
        public static List<string> cn_shortCut_app_list = new List<string>();
        public static List<string> tw_shortCut_app_list = new List<string>();

        //list of desktop apps
        public static List<string> browser_app_list = new List<string>();
        public static List<string> us_browser_app_list = new List<string>();
        public static List<string> ca_browser_app_list = new List<string>();
        public static List<string> latam_browser_app_list = new List<string>();
        public static List<string> frfr_browser_app_list = new List<string>();
        public static List<string> dede_browser_app_list = new List<string>();
        public static List<string> gb_browser_app_list = new List<string>();
        public static List<string> nordic_browser_app_list = new List<string>();
        public static List<string> ru_browser_app_list = new List<string>();
        public static List<string> emea1_browser_app_list = new List<string>();
        public static List<string> au_browser_app_list = new List<string>();
        public static List<string> jp_browser_app_list = new List<string>();
        public static List<string> kr_browser_app_list = new List<string>();
        public static List<string> aap1_browser_app_list = new List<string>();
        public static List<string> cn_browser_app_list = new List<string>();
        public static List<string> tw_browser_app_list = new List<string>();

        //list of OOBE Integration
        public static List<string> OOBE_integration_list = new List<string>();
        public static List<string> us_OOBE_integration_list = new List<string>();
        public static List<string> ca_OOBE_integration_list = new List<string>();
        public static List<string> latam_OOBE_integration_list = new List<string>();
        public static List<string> frfr_OOBE_integration_list = new List<string>();
        public static List<string> dede_OOBE_integration_list = new List<string>();
        public static List<string> gb_OOBE_integration_list = new List<string>();
        public static List<string> nordic_OOBE_integration_list = new List<string>();
        public static List<string> ru_OOBE_integration_list = new List<string>();
        public static List<string> emea1_OOBE_integration_list = new List<string>();
        public static List<string> au_OOBE_integration_list = new List<string>();
        public static List<string> jp_OOBE_integration_list = new List<string>();
        public static List<string> kr_OOBE_integration_list = new List<string>();
        public static List<string> aap1_OOBE_integration_list = new List<string>();
        public static List<string> cn_OOBE_integration_list = new List<string>();
        public static List<string> tw_OOBE_integration_list = new List<string>();

        //list of OOBE Integration
        public static List<string> Windows_Next_Recommended_list = new List<string>();
        public static List<string> us_next_recommended_list = new List<string>();
        public static List<string> ca_next_recommended_list = new List<string>();
        public static List<string> latam_next_recommended_list = new List<string>();
        public static List<string> frfr_next_recommended_list = new List<string>();
        public static List<string> dede_next_recommended_list = new List<string>();
        public static List<string> gb_next_recommended_list = new List<string>();
        public static List<string> nordic_next_recommended_list = new List<string>();
        public static List<string> ru_next_recommended_list = new List<string>();
        public static List<string> emea1_next_recommended_list = new List<string>();
        public static List<string> au_next_recommended_list = new List<string>();
        public static List<string> jp_next_recommended_list = new List<string>();
        public static List<string> kr_next_recommended_list = new List<string>();
        public static List<string> aap1_next_recommended_list = new List<string>();
        public static List<string> cn_next_recommended_list = new List<string>();
        public static List<string> tw_next_recommended_list = new List<string>();

        public static int area_row_item_type = 0;
        public static int area_col_item_type = 0;
        public static int lang_area_col_item_type = 0;
        public static int lang_area_row_item_type = 0;
        //language cell value
        public static string LANG_cellValue = null;

        ///--------------------------------------------------------------------
        /// <summary>
        /// Get path of SCL file on FTP server, will download SCL to local side
        /// </summary>
        /// <param name="QuantaProjectName">
        /// project name defined by Quanta
        /// </param>
        /// <returns>
        ///none
        /// </returns>
        ///--------------------------------------------------------------------
        public void GetFTPFilePath(string QuantaProjectName){
            // Replace these values with your FTP server details
            string ftpServer = "ftp://172.0.0.9";
            string username = "sit001";
            string password = "sit1234";
            // Specify the directory path on the server
            string directoryPath = "/Image/SCD/" + QuantaProjectName;
            try{
                // Create the FTP request
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(new Uri($"{ftpServer}{directoryPath}"));
                request.Method = WebRequestMethods.Ftp.ListDirectory;
                request.Credentials = new NetworkCredential(username, password);

                // Get the response and read the directory listing
                using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
                using (Stream responseStream = response.GetResponseStream())
                using (StreamReader reader = new StreamReader(responseStream)){
                    string line;
                    while ((line = reader.ReadLine()) != null){
                        Console.WriteLine($"File path on server: {directoryPath}{line}");
                    }
                }
            }catch (WebException ex){
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        ///--------------------------------------------------------------------
        /// <summary>
        /// Setup all searching tables for GSA MRD check
        /// </summary>
        /// <param name="">
        /// none
        /// </param>
        /// <returns>
        ///none
        /// </returns>
        ///--------------------------------------------------------------------
        public void SetupFullTable(string excelFile){
            // Get the Windows user account name
            string userName = Environment.UserName;
            // Set Excel file path
            string root_path = "C:\\Users\\" + userName + "\\Documents\\";
            string excelFileName = excelFile;
            string excelFilePath = root_path + excelFileName;
            Console.WriteLine(excelFilePath);
            // Build a new Excel Application object
            Excel.Application excelApp = new Excel.Application();

            // Open Excel file
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);

            // Use index "Slotting" to get "GSA MRD" worksheet
            Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets["Slotting"];

            // Read the area of data to get number of row and column
            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;

            //app name cell value
            string AppName_cellValue = null;

            int app_id = 0;
            int desktop_app_id = 0;
            int desktop_bar_app_id = 0;
            int shortCut_app_id = 0;
            int browser_app_id = 0;
            int OOBE_app_id = 0;
            int Next_Recommended_id = 0;
            // Declare a two dimensional array.
            string[,] multiDimensionalArray1 = new string[17, 18];
            for (int row = 1; row <= rowCount; row++){
                for (int col = 1; col <= colCount; col++){
                    //use Cells object to get cellValue
                    Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, col];
                    string cellValue = cell.Value != null ? cell.Value.ToString() : "";

                    //build Metro app table
                    if (cellValue == "Metro Apps"){
                        //get row index offset of "Driver"
                        int metro_base_row = row;
                        //get col index offset of "Driver"
                        int metro_base_col = col;

                        area_row_item_type = metro_base_row + 1;
                        area_col_item_type = metro_base_col + 1;
                        Console.WriteLine();
                        do{
                            //Read Category cell
                            Excel.Range AppName_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[area_row_item_type, metro_base_col + 1];
                            AppName_cellValue = AppName_cell.Value != null ? AppName_cell.Value.ToString() : "";
                            //Console.WriteLine($"AppName_cellValue: {AppName_cellValue}, app_id: {app_id}");
                            metro_app_list.Add(AppName_cellValue + "," + app_id);
                            app_id++;
                            area_row_item_type += 1;
                        }while (AppName_cellValue != "Spades (2023.Q2)");
                    }

                    //Read supported Metro apps and build a check table
                    ReadSupportDataFromExcel(cellValue, row, col, 0, 17, worksheet, 0);//17


                    //build Desktop app table
                    if (cellValue == "Desktop Apps"){
                        //get row index offset of "Driver"
                        int desktop_base_row = row;
                        //get col index offset of "Driver"
                        int desktop_base_col = col;

                        area_row_item_type = desktop_base_row + 1;
                        area_col_item_type = desktop_base_col + 1;

                        do{
                            //Read Category cell
                            Excel.Range AppName_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[area_row_item_type, desktop_base_col + 1];
                            AppName_cellValue = AppName_cell.Value != null ? AppName_cell.Value.ToString() : "";
                            //Console.WriteLine($"DesktopAPPName: {AppName_cellValue}, app_id: {desktop_app_id}");
                            desktop_app_list.Add(AppName_cellValue + "," + desktop_app_id);
                            desktop_app_id++;
                            area_row_item_type += 1;
                        }while (AppName_cellValue != "Acer Product Registration");
                    }
                    //Read supported Desktop Apps apps and build a check table
                    ReadSupportDataFromExcel(cellValue, row, col, 18, 11, worksheet, 1);//11

                    //build Desktop Taskbar Pin table
                    if (cellValue.IndexOf("Desktop Taskbar Pin") >= 0){
                        //get row index offset of "Driver"
                        int desktop_base_row = row;
                        //get col index offset of "Driver"
                        int desktop_base_col = col;

                        area_row_item_type = desktop_base_row + 1;
                        area_col_item_type = desktop_base_col;

                        do{
                            //Read Category cell
                            Excel.Range AppName_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[area_row_item_type, desktop_base_col];
                            AppName_cellValue = AppName_cell.Value != null ? AppName_cell.Value.ToString() : "";
                            //Console.WriteLine($"DesktopTaskbarPin AppName: {AppName_cellValue}, app_id: {desktop_bar_app_id}");
                            desktop_bar_app_list.Add(AppName_cellValue + "," + desktop_bar_app_id);
                            desktop_bar_app_id++;
                            area_row_item_type += 1;
                        }while (AppName_cellValue != "Yandex Browser (2023.Q2)");
                    }

                    //Read supported Desktop Bar pin apps and build a check table
                    //set app count 8, offset 30 to locate cell position in excel file
                    ReadSupportDataFromExcel(cellValue, row, col, 30, 8, worksheet, 2);


                    //build All Apps Shortcut table
                    if (cellValue == "All Apps Shortcut (1 : present. 0 : not present.)"){
                        //get row index offset of "Driver"
                        int shortCut_base_row = row;
                        //get col index offset of "Driver"
                        int shortCut_base_col = col;

                        area_row_item_type = shortCut_base_row + 1;
                        area_col_item_type = shortCut_base_col;

                        do{
                            //Read Category cell
                            Excel.Range AppName_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[area_row_item_type, shortCut_base_col];
                            AppName_cellValue = AppName_cell.Value != null ? AppName_cell.Value.ToString() : "";
                            //Console.WriteLine($"ShortCut AppName: {AppName_cellValue}, app_id: {shortCut_app_id}");
                            shortCut_app_list.Add(AppName_cellValue + "," + shortCut_app_id);
                            shortCut_app_id++;
                            area_row_item_type += 1;
                        }while (AppName_cellValue != "Baidu Weblink (2023.Q2)");
                    }

                    //Read supported all Apps shortcut apps and build a check table
                    //apps count 21, offset 39
                    ReadSupportDataFromExcel(cellValue, row, col, 39, 21, worksheet, 3);

                    //build Browser Favorites table
                    if (cellValue == "Browser Favorites (1 : present. 0 : not present.)"){
                        int browser_base_row = row;//get row index offset of "Driver"
                        int browser_base_col = col;//get col index offset of "Driver"

                        area_row_item_type = browser_base_row + 1;
                        area_col_item_type = browser_base_col + 1;
                        Console.WriteLine();
                        do{
                            //Read Category cell
                            Excel.Range AppName_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[area_row_item_type, browser_base_col];
                            AppName_cellValue = AppName_cell.Value != null ? AppName_cell.Value.ToString() : "";
                            //Console.WriteLine($"Browser AppName: {AppName_cellValue}, app_id: {browser_app_id}");
                            browser_app_list.Add(AppName_cellValue + "," + browser_app_id);
                            browser_app_id++;
                            area_row_item_type += 1;
                        }while (AppName_cellValue != "Agoda_Favorite (2023.Q2)");
                    }

                    //Read supported Browser Favorites apps and build a check table
                    ReadSupportDataFromExcel(cellValue, row, col, 61, 5, worksheet, 4);//5

                    //Build Browser Favorites table
                    if (cellValue.IndexOf("OOBE Integration") > 0){
                        int OOBE_base_row = row;//get row index offset of "Driver"
                        int OOBE_base_col = col;//get col index offset of "Driver"

                        area_row_item_type = OOBE_base_row + 1;
                        area_col_item_type = OOBE_base_col + 1;
                        Console.WriteLine();
                        do{
                            //Read Category cell
                            Excel.Range AppName_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[area_row_item_type, OOBE_base_col];
                            AppName_cellValue = AppName_cell.Value != null ? AppName_cell.Value.ToString() : "";
                            //Console.WriteLine($"Browser AppName: {AppName_cellValue}, app_id: {OOBE_app_id}");
                            OOBE_integration_list.Add(AppName_cellValue + "," + OOBE_app_id);
                            OOBE_app_id++;
                            area_row_item_type += 1;
                        }while (AppName_cellValue != "McAfee LiveSafe P1 (2023.Q2)");
                    }

                    ReadSupportDataFromExcel(cellValue, row, col, 67, 1, worksheet, 5);//1

                    //build Browser Favorites table
                    if (cellValue == "Windows Next Recommended (1 : present. 0 : not present.)"){
                        int Next_Recommended_row = row;//get row index offset of "Driver"
                        int Next_Recommended_col = col;//get col index offset of "Driver"

                        area_row_item_type = Next_Recommended_row + 1;
                        area_col_item_type = Next_Recommended_col + 1;
                        Console.WriteLine();
                        do{
                            //Read Category cell
                            Excel.Range AppName_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[area_row_item_type, Next_Recommended_col];
                            AppName_cellValue = AppName_cell.Value != null ? AppName_cell.Value.ToString() : "";
                            //Console.WriteLine($"Windows Next Recommended: {AppName_cellValue}, Next_Recommended_id: {Next_Recommended_id}");
                            Windows_Next_Recommended_list.Add(AppName_cellValue + "," + Next_Recommended_id);
                            Next_Recommended_id++;
                            area_row_item_type += 1;
                        }while (AppName_cellValue != "WPS China (2023.Q2)");
                    }
                    ReadSupportDataFromExcel(cellValue, row, col, 69, 2, worksheet, 6);//2
                }

            }

            // 關閉Excel檔案
            workbook.Close();
            excelApp.Quit();

            // 釋放資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            BuildMetroTable();
            BuildDesktopTable();
            BuildDesktopBarPINTable();
            BuildAllAppsShortCutTable();
            BuildBrowserTable();
            BuildOOBETable();
            BuildWindowsNextRecommendedTable();
        }
        public void ReadSupportDataFromExcel(string cellValue,
            int row,
            int col,
            int offset,
            int apps_count,
            Excel.Worksheet worksheet,
            int table_id){
            if (cellValue == "US"){
                int language_base_row = row;
                int language_base_col = col;
                int app_index = 0;
                lang_area_row_item_type = language_base_row + 2;
                lang_area_col_item_type = language_base_col;
                int desktop_app_index = 0;
                int desktop_bar_app_index = 0;
                int shortCut_app_index = 0;
                int browser_app_index = 0;
                int OOBE_index = 0;
                int Next_Recommended_index = 0;
                int iter = 0;
                do{
                    //Read Category cell
                    Excel.Range US_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[lang_area_row_item_type + offset, language_base_col];
                    LANG_cellValue = US_cell.Value != null ? US_cell.Value.ToString() : "";
                    switch (table_id){
                        case 0:
                            //Console.WriteLine($"US_cell: {LANG_cellValue}, app_index: {app_index}");
                            us_metro_app_list.Add(LANG_cellValue + "," + app_index);
                            app_index++;
                            break;
                        case 1:
                            us_desktop_app_list.Add(LANG_cellValue + "," + desktop_app_index);
                            desktop_app_index++;
                            break;
                        case 2:
                            us_desktop_bar_app_list.Add(LANG_cellValue + "," + desktop_bar_app_index);
                            desktop_bar_app_index++;
                            break;
                        case 3:
                            us_shortCut_app_list.Add(LANG_cellValue + "," + shortCut_app_index);
                            shortCut_app_index++;
                            break;
                        case 4:
                            us_browser_app_list.Add(LANG_cellValue + "," + browser_app_index);
                            browser_app_index++;
                            break;
                        case 5:
                            us_OOBE_integration_list.Add(LANG_cellValue + "," + OOBE_index);
                            OOBE_index++;
                            break;
                        case 6:
                            us_next_recommended_list.Add(LANG_cellValue + "," + Next_Recommended_index);
                            Next_Recommended_index++;
                            break;
                    }
                    lang_area_row_item_type += 1;
                    iter++;
                }while (iter < apps_count);
            }
            if (cellValue == "CA"){
                int language_base_row = row;
                int language_base_col = col;
                int app_index = 0;
                lang_area_row_item_type = language_base_row + 2;
                lang_area_col_item_type = language_base_col;
                int desktop_app_index = 0;
                int desktop_bar_app_index = 0;
                int shortCut_app_index = 0;
                int browser_app_index = 0;
                int OOBE_index = 0;
                int Next_Recommended_index = 0;
                int iter = 0;
                do{
                    //Read Category cell
                    Excel.Range CA_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[lang_area_row_item_type + offset, language_base_col];
                    LANG_cellValue = CA_cell.Value != null ? CA_cell.Value.ToString() : "";
                    switch (table_id){
                        case 0:
                            //Console.WriteLine($"CA_cell:{LANG_cellValue}, app_index: {app_index}");
                            ca_metro_app_list.Add(LANG_cellValue + "," + app_index);
                            app_index++;
                            break;
                        case 1:
                            ca_desktop_app_list.Add(LANG_cellValue + "," + desktop_app_index);
                            desktop_app_index++;
                            break;
                        case 2:
                            ca_desktop_bar_app_list.Add(LANG_cellValue + "," + desktop_bar_app_index);
                            desktop_bar_app_index++;
                            break;
                        case 3:
                            ca_shortCut_app_list.Add(LANG_cellValue + "," + shortCut_app_index);
                            shortCut_app_index++;
                            break;
                        case 4:
                            ca_browser_app_list.Add(LANG_cellValue + "," + browser_app_index);
                            browser_app_index++;
                            break;
                        case 5:
                            ca_OOBE_integration_list.Add(LANG_cellValue + "," + OOBE_index);
                            OOBE_index++;
                            break;
                        case 6:
                            ca_next_recommended_list.Add(LANG_cellValue + "," + Next_Recommended_index);
                            Next_Recommended_index++;
                            break;
                    }
                    lang_area_row_item_type += 1;
                    iter++;
                }while (iter < apps_count);
            }
            if (cellValue == "LATAM"){
                int language_base_row = row;
                int language_base_col = col;
                int app_index = 0;
                lang_area_row_item_type = language_base_row + 2;
                lang_area_col_item_type = language_base_col;
                int desktop_app_index = 0;
                int desktop_bar_app_index = 0;
                int shortCut_app_index = 0;
                int browser_app_index = 0;
                int OOBE_index = 0;
                int Next_Recommended_index = 0;
                int iter = 0;
                do{
                    //Read Category cell
                    Excel.Range LATAM_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[lang_area_row_item_type + offset, language_base_col];
                    LANG_cellValue = LATAM_cell.Value != null ? LATAM_cell.Value.ToString() : "";
                    switch (table_id){
                        case 0:
                            //Console.WriteLine($"LATAM_cell:{LANG_cellValue}, app_index: {app_index}");
                            latam_metro_app_list.Add(LANG_cellValue + "," + app_index);
                            app_index++;
                            break;
                        case 1:
                            latam_desktop_app_list.Add(LANG_cellValue + "," + desktop_app_index);
                            desktop_app_index++;
                            break;
                        case 2:
                            latam_desktop_bar_app_list.Add(LANG_cellValue + "," + desktop_bar_app_index);
                            desktop_bar_app_index++;
                            break;
                        case 3:
                            latam_shortCut_app_list.Add(LANG_cellValue + "," + shortCut_app_index);
                            shortCut_app_index++;
                            break;
                        case 4:
                            latam_browser_app_list.Add(LANG_cellValue + "," + browser_app_index);
                            browser_app_index++;
                            break;
                        case 5:
                            latam_OOBE_integration_list.Add(LANG_cellValue + "," + OOBE_index);
                            OOBE_index++;
                            break;
                        case 6:
                            latam_next_recommended_list.Add(LANG_cellValue + "," + Next_Recommended_index);
                            Next_Recommended_index++;
                            break;
                    }
                    lang_area_row_item_type += 1;
                    iter++;
                }while (iter < apps_count);
            }
            if (cellValue == "FRFR"){
                int language_base_row = row;
                int language_base_col = col;
                int app_index = 0;
                lang_area_row_item_type = language_base_row + 2;
                lang_area_col_item_type = language_base_col;
                int desktop_app_index = 0;
                int desktop_bar_app_index = 0;
                int shortCut_app_index = 0;
                int browser_app_index = 0;
                int OOBE_index = 0;
                int Next_Recommended_index = 0;
                int iter = 0;
                do{
                    //Read Category cell
                    Excel.Range FRFR_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[lang_area_row_item_type + offset, language_base_col];
                    LANG_cellValue = FRFR_cell.Value != null ? FRFR_cell.Value.ToString() : "";
                    switch (table_id){
                        case 0:
                            //Console.WriteLine($"FRFR_cell:{LANG_cellValue}, app_index: {app_index}");
                            frfr_metro_app_list.Add(LANG_cellValue + "," + app_index);
                            app_index++;
                            break;
                        case 1:
                            frfr_desktop_app_list.Add(LANG_cellValue + "," + desktop_app_index);
                            desktop_app_index++;
                            break;
                        case 2:
                            frfr_desktop_bar_app_list.Add(LANG_cellValue + "," + desktop_bar_app_index);
                            desktop_bar_app_index++;
                            break;
                        case 3:
                            frfr_shortCut_app_list.Add(LANG_cellValue + "," + shortCut_app_index);
                            shortCut_app_index++;
                            break;
                        case 4:
                            frfr_browser_app_list.Add(LANG_cellValue + "," + browser_app_index);
                            browser_app_index++;
                            break;
                        case 5:
                            frfr_OOBE_integration_list.Add(LANG_cellValue + "," + OOBE_index);
                            OOBE_index++;
                            break;
                        case 6:
                            frfr_next_recommended_list.Add(LANG_cellValue + "," + Next_Recommended_index);
                            Next_Recommended_index++;
                            break;
                    }
                    lang_area_row_item_type += 1;
                    iter++;
                }while (iter < apps_count);
            }
            if (cellValue == "DEDE"){
                int language_base_row = row;
                int language_base_col = col;
                int app_index = 0;
                lang_area_row_item_type = language_base_row + 2;
                lang_area_col_item_type = language_base_col;
                int desktop_app_index = 0;
                int desktop_bar_app_index = 0;
                int shortCut_app_index = 0;
                int browser_app_index = 0;
                int OOBE_index = 0;
                int Next_Recommended_index = 0;
                int iter = 0;
                do{
                    //Read Category cell
                    Excel.Range DEDE_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[lang_area_row_item_type + offset, language_base_col];
                    LANG_cellValue = DEDE_cell.Value != null ? DEDE_cell.Value.ToString() : "";
                    switch (table_id){
                        case 0:
                            //Console.WriteLine($"DEDE_cell:{LANG_cellValue}, app_index: {app_index}");
                            dede_metro_app_list.Add(LANG_cellValue + "," + app_index);
                            app_index++;
                            break;
                        case 1:
                            dede_desktop_app_list.Add(LANG_cellValue + "," + desktop_app_index);
                            desktop_app_index++;
                            break;
                        case 2:
                            dede_desktop_bar_app_list.Add(LANG_cellValue + "," + desktop_bar_app_index);
                            desktop_bar_app_index++;
                            break;
                        case 3:
                            dede_shortCut_app_list.Add(LANG_cellValue + "," + shortCut_app_index);
                            shortCut_app_index++;
                            break;
                        case 4:
                            dede_browser_app_list.Add(LANG_cellValue + "," + browser_app_index);
                            browser_app_index++;
                            break;
                        case 5:
                            dede_OOBE_integration_list.Add(LANG_cellValue + "," + OOBE_index);
                            OOBE_index++;
                            break;
                        case 6:
                            dede_next_recommended_list.Add(LANG_cellValue + "," + Next_Recommended_index);
                            Next_Recommended_index++;
                            break;
                    }
                    lang_area_row_item_type += 1;
                    iter++;
                }while (iter < apps_count);
            }
            if (cellValue == "GB"){
                int language_base_row = row;
                int language_base_col = col;
                int app_index = 0;
                lang_area_row_item_type = language_base_row + 2;
                lang_area_col_item_type = language_base_col;
                int desktop_app_index = 0;
                int desktop_bar_app_index = 0;
                int shortCut_app_index = 0;
                int browser_app_index = 0;
                int OOBE_index = 0;
                int Next_Recommended_index = 0;
                int iter = 0;
                do{
                    //Read Category cell
                    Excel.Range GB_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[lang_area_row_item_type + offset, language_base_col];
                    LANG_cellValue = GB_cell.Value != null ? GB_cell.Value.ToString() : "";
                    switch (table_id){
                        case 0:
                            //Console.WriteLine($"GB_cell:{LANG_cellValue}, app_index: {app_index}");
                            gb_metro_app_list.Add(LANG_cellValue + "," + app_index);
                            app_index++;
                            break;
                        case 1:
                            gb_desktop_app_list.Add(LANG_cellValue + "," + desktop_app_index);
                            desktop_app_index++;
                            break;
                        case 2:
                            gb_desktop_bar_app_list.Add(LANG_cellValue + "," + desktop_bar_app_index);
                            desktop_bar_app_index++;
                            break;
                        case 3:
                            gb_shortCut_app_list.Add(LANG_cellValue + "," + shortCut_app_index);
                            shortCut_app_index++;
                            break;
                        case 4:
                            gb_browser_app_list.Add(LANG_cellValue + "," + browser_app_index);
                            browser_app_index++;
                            break;
                        case 5:
                            gb_OOBE_integration_list.Add(LANG_cellValue + "," + OOBE_index);
                            OOBE_index++;
                            break;
                        case 6:
                            gb_next_recommended_list.Add(LANG_cellValue + "," + Next_Recommended_index);
                            Next_Recommended_index++;
                            break;
                    }
                    lang_area_row_item_type += 1;
                    iter++;
                }while (iter < apps_count);
            }
            if (cellValue == "NORDIC"){
                int language_base_row = row;
                int language_base_col = col;
                int app_index = 0;
                lang_area_row_item_type = language_base_row + 2;
                lang_area_col_item_type = language_base_col;
                int desktop_app_index = 0;
                int desktop_bar_app_index = 0;
                int shortCut_app_index = 0;
                int browser_app_index = 0;
                int OOBE_index = 0;
                int Next_Recommended_index = 0;
                int iter = 0;
                do{
                    //Read Category cell
                    Excel.Range NORDIC_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[lang_area_row_item_type + offset, language_base_col];
                    LANG_cellValue = NORDIC_cell.Value != null ? NORDIC_cell.Value.ToString() : "";
                    switch (table_id){
                        case 0:
                            //Console.WriteLine($"NORDIC_cell:{LANG_cellValue}, app_index: {app_index}");
                            nordic_metro_app_list.Add(LANG_cellValue + "," + app_index);
                            app_index++;
                            break;
                        case 1:
                            nordic_desktop_app_list.Add(LANG_cellValue + "," + desktop_app_index);
                            desktop_app_index++;
                            break;
                        case 2:
                            nordic_desktop_bar_app_list.Add(LANG_cellValue + "," + desktop_bar_app_index);
                            desktop_bar_app_index++;
                            break;
                        case 3:
                            nordic_shortCut_app_list.Add(LANG_cellValue + "," + shortCut_app_index);
                            shortCut_app_index++;
                            break;
                        case 4:
                            nordic_browser_app_list.Add(LANG_cellValue + "," + browser_app_index);
                            browser_app_index++;
                            break;
                        case 5:
                            nordic_OOBE_integration_list.Add(LANG_cellValue + "," + OOBE_index);
                            OOBE_index++;
                            break;
                        case 6:
                            nordic_next_recommended_list.Add(LANG_cellValue + "," + Next_Recommended_index);
                            Next_Recommended_index++;
                            break;
                    }
                    lang_area_row_item_type += 1;
                    iter++;
                }while (iter < apps_count);
            }
            if (cellValue == "RU"){
                int language_base_row = row;
                int language_base_col = col;
                int app_index = 0;
                lang_area_row_item_type = language_base_row + 2;
                lang_area_col_item_type = language_base_col;
                int desktop_app_index = 0;
                int desktop_bar_app_index = 0;
                int shortCut_app_index = 0;
                int browser_app_index = 0;
                int OOBE_index = 0;
                int Next_Recommended_index = 0;
                int iter = 0;
                do{
                    //Read Category cell
                    Excel.Range RU_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[lang_area_row_item_type + offset, language_base_col];
                    LANG_cellValue = RU_cell.Value != null ? RU_cell.Value.ToString() : "";
                    switch (table_id){
                        case 0:
                            //Console.WriteLine($"RU_cell:{LANG_cellValue}, app_index: {app_index}");
                            ru_metro_app_list.Add(LANG_cellValue + "," + app_index);
                            app_index++;
                            break;
                        case 1:
                            ru_desktop_app_list.Add(LANG_cellValue + "," + desktop_app_index);
                            desktop_app_index++;
                            break;
                        case 2:
                            ru_desktop_bar_app_list.Add(LANG_cellValue + "," + desktop_bar_app_index);
                            desktop_bar_app_index++;
                            break;
                        case 3:
                            ru_shortCut_app_list.Add(LANG_cellValue + "," + shortCut_app_index);
                            shortCut_app_index++;
                            break;
                        case 4:
                            ru_browser_app_list.Add(LANG_cellValue + "," + browser_app_index);
                            browser_app_index++;
                            break;
                        case 5:
                            ru_OOBE_integration_list.Add(LANG_cellValue + "," + OOBE_index);
                            OOBE_index++;
                            break;
                        case 6:
                            ru_next_recommended_list.Add(LANG_cellValue + "," + Next_Recommended_index);
                            Next_Recommended_index++;
                            break;
                    }
                    lang_area_row_item_type += 1;
                    iter++;
                }while (iter < apps_count);
            }
            if (cellValue == "EMEA1"){
                int language_base_row = row;
                int language_base_col = col;
                int app_index = 0;
                lang_area_row_item_type = language_base_row + 2;
                lang_area_col_item_type = language_base_col;
                int desktop_app_index = 0;
                int desktop_bar_app_index = 0;
                int shortCut_app_index = 0;
                int browser_app_index = 0;
                int OOBE_index = 0;
                int Next_Recommended_index = 0;
                int iter = 0;
                do{
                    //Read Category cell
                    Excel.Range EMEA1_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[lang_area_row_item_type + offset, language_base_col];
                    LANG_cellValue = EMEA1_cell.Value != null ? EMEA1_cell.Value.ToString() : "";
                    switch (table_id){
                        case 0:
                            //Console.WriteLine($"EMEA1_cell:{LANG_cellValue}, app_index: {app_index}");
                            emea1_metro_app_list.Add(LANG_cellValue + "," + app_index);
                            app_index++;
                            break;
                        case 1:
                            emea1_desktop_app_list.Add(LANG_cellValue + "," + desktop_app_index);
                            desktop_app_index++;
                            break;
                        case 2:
                            emea1_desktop_bar_app_list.Add(LANG_cellValue + "," + desktop_bar_app_index);
                            desktop_bar_app_index++;
                            break;
                        case 3:
                            emea1_shortCut_app_list.Add(LANG_cellValue + "," + shortCut_app_index);
                            shortCut_app_index++;
                            break;
                        case 4:
                            emea1_browser_app_list.Add(LANG_cellValue + "," + browser_app_index);
                            browser_app_index++;
                            break;
                        case 5:
                            emea1_OOBE_integration_list.Add(LANG_cellValue + "," + OOBE_index);
                            OOBE_index++;
                            break;
                        case 6:
                            emea1_next_recommended_list.Add(LANG_cellValue + "," + Next_Recommended_index);
                            Next_Recommended_index++;
                            break;
                    }
                    lang_area_row_item_type += 1;
                    iter++;
                }while (iter < apps_count);
            }
            if (cellValue == "AU"){
                int language_base_row = row;
                int language_base_col = col;
                int app_index = 0;
                lang_area_row_item_type = language_base_row + 2;
                lang_area_col_item_type = language_base_col;
                int desktop_app_index = 0;
                int desktop_bar_app_index = 0;
                int shortCut_app_index = 0;
                int browser_app_index = 0;
                int OOBE_index = 0;
                int Next_Recommended_index = 0;
                int iter = 0;
                do{
                    //Read Category cell
                    Excel.Range AU_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[lang_area_row_item_type + offset, language_base_col];
                    LANG_cellValue = AU_cell.Value != null ? AU_cell.Value.ToString() : "";
                    switch (table_id){
                        case 0:
                            //Console.WriteLine($"AU_cell:{LANG_cellValue}, app_index: {app_index}");
                            au_metro_app_list.Add(LANG_cellValue + "," + app_index);
                            app_index++;
                            break;
                        case 1:
                            au_desktop_app_list.Add(LANG_cellValue + "," + desktop_app_index);
                            desktop_app_index++;
                            break;
                        case 2:
                            au_desktop_bar_app_list.Add(LANG_cellValue + "," + desktop_bar_app_index);
                            desktop_bar_app_index++;
                            break;
                        case 3:
                            au_shortCut_app_list.Add(LANG_cellValue + "," + shortCut_app_index);
                            shortCut_app_index++;
                            break;
                        case 4:
                            au_browser_app_list.Add(LANG_cellValue + "," + browser_app_index);
                            browser_app_index++;
                            break;
                        case 5:
                            au_OOBE_integration_list.Add(LANG_cellValue + "," + OOBE_index);
                            OOBE_index++;
                            break;
                        case 6:
                            au_next_recommended_list.Add(LANG_cellValue + "," + Next_Recommended_index);
                            Next_Recommended_index++;
                            break;
                    }
                    lang_area_row_item_type += 1;
                    iter++;
                }while (iter < apps_count);
            }
            if (cellValue == "JP"){
                int language_base_row = row;
                int language_base_col = col;
                int app_index = 0;
                lang_area_row_item_type = language_base_row + 2;
                lang_area_col_item_type = language_base_col;
                int desktop_app_index = 0;
                int desktop_bar_app_index = 0;
                int shortCut_app_index = 0;
                int browser_app_index = 0;
                int OOBE_index = 0;
                int Next_Recommended_index = 0;
                int iter = 0;
                do{
                    //Read Category cell
                    Excel.Range JP_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[lang_area_row_item_type + offset, language_base_col];
                    LANG_cellValue = JP_cell.Value != null ? JP_cell.Value.ToString() : "";
                    switch (table_id){
                        case 0:
                            //Console.WriteLine($"JP_cell:{LANG_cellValue}, app_index: {app_index}");
                            jp_metro_app_list.Add(LANG_cellValue + "," + app_index);
                            app_index++;
                            break;
                        case 1:
                            jp_desktop_app_list.Add(LANG_cellValue + "," + desktop_app_index);
                            desktop_app_index++;
                            break;
                        case 2:
                            jp_desktop_bar_app_list.Add(LANG_cellValue + "," + desktop_bar_app_index);
                            desktop_bar_app_index++;
                            break;
                        case 3:
                            jp_shortCut_app_list.Add(LANG_cellValue + "," + shortCut_app_index);
                            shortCut_app_index++;
                            break;
                        case 4:
                            jp_browser_app_list.Add(LANG_cellValue + "," + browser_app_index);
                            browser_app_index++;
                            break;
                        case 5:
                            jp_OOBE_integration_list.Add(LANG_cellValue + "," + OOBE_index);
                            OOBE_index++;
                            break;
                        case 6:
                            jp_next_recommended_list.Add(LANG_cellValue + "," + Next_Recommended_index);
                            Next_Recommended_index++;
                            break;
                    }
                    lang_area_row_item_type += 1;
                    iter++;
                }while (iter < apps_count);
            }
            if (cellValue == "KR"){
                int language_base_row = row;
                int language_base_col = col;
                int app_index = 0;
                lang_area_row_item_type = language_base_row + 2;
                lang_area_col_item_type = language_base_col;
                int desktop_app_index = 0;
                int desktop_bar_app_index = 0;
                int shortCut_app_index = 0;
                int browser_app_index = 0;
                int OOBE_index = 0;
                int Next_Recommended_index = 0;
                int iter = 0;
                do{
                    //Read Category cell
                    Excel.Range KR_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[lang_area_row_item_type + offset, language_base_col];
                    LANG_cellValue = KR_cell.Value != null ? KR_cell.Value.ToString() : "";
                    switch (table_id){
                        case 0:
                            //Console.WriteLine($"KR_cell:{LANG_cellValue}, app_index: {app_index}");
                            kr_metro_app_list.Add(LANG_cellValue + "," + app_index);
                            app_index++;
                            break;
                        case 1:
                            kr_desktop_app_list.Add(LANG_cellValue + "," + desktop_app_index);
                            desktop_app_index++;
                            break;
                        case 2:
                            kr_desktop_bar_app_list.Add(LANG_cellValue + "," + desktop_bar_app_index);
                            desktop_bar_app_index++;
                            break;
                        case 3:
                            kr_shortCut_app_list.Add(LANG_cellValue + "," + shortCut_app_index);
                            shortCut_app_index++;
                            break;
                        case 4:
                            kr_browser_app_list.Add(LANG_cellValue + "," + browser_app_index);
                            browser_app_index++;
                            break;
                        case 5:
                            kr_OOBE_integration_list.Add(LANG_cellValue + "," + OOBE_index);
                            OOBE_index++;
                            break;
                        case 6:
                            kr_next_recommended_list.Add(LANG_cellValue + "," + Next_Recommended_index);
                            Next_Recommended_index++;
                            break;
                    }
                    lang_area_row_item_type += 1;
                    iter++;
                }while (iter < apps_count);
            }
            if (cellValue == "AAP1"){
                int language_base_row = row;
                int language_base_col = col;
                int app_index = 0;
                lang_area_row_item_type = language_base_row + 2;
                lang_area_col_item_type = language_base_col;
                int desktop_app_index = 0;
                int desktop_bar_app_index = 0;
                int shortCut_app_index = 0;
                int browser_app_index = 0;
                int OOBE_index = 0;
                int Next_Recommended_index = 0;
                int iter = 0;
                do{
                    //Read Category cell
                    Excel.Range AAP1_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[lang_area_row_item_type + offset, language_base_col];
                    LANG_cellValue = AAP1_cell.Value != null ? AAP1_cell.Value.ToString() : "";
                    switch (table_id){
                        case 0:
                            //Console.WriteLine($"AAP1_cell:{LANG_cellValue}, app_index: {app_index}");
                            aap1_metro_app_list.Add(LANG_cellValue + "," + app_index);
                            app_index++;
                            break;
                        case 1:
                            aap1_desktop_app_list.Add(LANG_cellValue + "," + desktop_app_index);
                            desktop_app_index++;
                            break;
                        case 2:
                            aap1_desktop_bar_app_list.Add(LANG_cellValue + "," + desktop_bar_app_index);
                            desktop_bar_app_index++;
                            break;
                        case 3:
                            aap1_shortCut_app_list.Add(LANG_cellValue + "," + shortCut_app_index);
                            shortCut_app_index++;
                            break;
                        case 4:
                            aap1_browser_app_list.Add(LANG_cellValue + "," + browser_app_index);
                            browser_app_index++;
                            break;
                        case 5:
                            aap1_OOBE_integration_list.Add(LANG_cellValue + "," + OOBE_index);
                            OOBE_index++;
                            break;
                        case 6:
                            aap1_next_recommended_list.Add(LANG_cellValue + "," + Next_Recommended_index);
                            Next_Recommended_index++;
                            break;
                    }
                    lang_area_row_item_type += 1;
                    iter++;
                }while (iter < apps_count);
            }
            if (cellValue == "CN"){
                int language_base_row = row;
                int language_base_col = col;
                int app_index = 0;
                lang_area_row_item_type = language_base_row + 2;
                lang_area_col_item_type = language_base_col;
                int desktop_app_index = 0;
                int desktop_bar_app_index = 0;
                int shortCut_app_index = 0;
                int browser_app_index = 0;
                int OOBE_index = 0;
                int Next_Recommended_index = 0;
                int iter = 0;
                do{
                    //Read Category cell
                    Excel.Range CN_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[lang_area_row_item_type + offset, language_base_col];
                    LANG_cellValue = CN_cell.Value != null ? CN_cell.Value.ToString() : "";
                    switch (table_id){
                        case 0:
                            //Console.WriteLine($"CN_cell:{LANG_cellValue}, app_index: {app_index}");
                            cn_metro_app_list.Add(LANG_cellValue + "," + app_index);
                            app_index++;
                            break;
                        case 1:
                            cn_desktop_app_list.Add(LANG_cellValue + "," + desktop_app_index);
                            desktop_app_index++;
                            break;
                        case 2:
                            cn_desktop_bar_app_list.Add(LANG_cellValue + "," + desktop_bar_app_index);
                            desktop_bar_app_index++;
                            break;
                        case 3:
                            cn_shortCut_app_list.Add(LANG_cellValue + "," + shortCut_app_index);
                            shortCut_app_index++;
                            break;
                        case 4:
                            cn_browser_app_list.Add(LANG_cellValue + "," + browser_app_index);
                            browser_app_index++;
                            break;
                        case 5:
                            cn_OOBE_integration_list.Add(LANG_cellValue + "," + OOBE_index);
                            OOBE_index++;
                            break;
                        case 6:
                            cn_next_recommended_list.Add(LANG_cellValue + "," + Next_Recommended_index);
                            Next_Recommended_index++;
                            break;
                    }
                    lang_area_row_item_type += 1;
                    iter++;
                }while (iter < apps_count);
            }
            if (cellValue == "TW"){
                int language_base_row = row;
                int language_base_col = col;
                int app_index = 0;
                lang_area_row_item_type = language_base_row + 2;
                lang_area_col_item_type = language_base_col;
                int desktop_app_index = 0;
                int desktop_bar_app_index = 0;
                int shortCut_app_index = 0;
                int browser_app_index = 0;
                int OOBE_index = 0;
                int Next_Recommended_index = 0;
                int iter = 0;
                do{
                    //Read Category cell
                    Excel.Range TW_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[lang_area_row_item_type + offset, language_base_col];
                    LANG_cellValue = TW_cell.Value != null ? TW_cell.Value.ToString() : "";
                    switch (table_id){
                        case 0:
                            //Console.WriteLine($"TW_cell:{LANG_cellValue}, app_index: {app_index}");
                            tw_metro_app_list.Add(LANG_cellValue + "," + app_index);
                            app_index++;
                            break;
                        case 1:
                            tw_desktop_app_list.Add(LANG_cellValue + "," + desktop_app_index);
                            desktop_app_index++;
                            break;
                        case 2:
                            tw_desktop_bar_app_list.Add(LANG_cellValue + "," + desktop_bar_app_index);
                            desktop_bar_app_index++;
                            break;
                        case 3:
                            tw_shortCut_app_list.Add(LANG_cellValue + "," + shortCut_app_index);
                            shortCut_app_index++;
                            break;
                        case 4:
                            tw_browser_app_list.Add(LANG_cellValue + "," + browser_app_index);
                            browser_app_index++;
                            break;
                        case 5:
                            tw_OOBE_integration_list.Add(LANG_cellValue + "," + OOBE_index);
                            OOBE_index++;
                            break;
                        case 6:
                            tw_next_recommended_list.Add(LANG_cellValue + "," + Next_Recommended_index);
                            Next_Recommended_index++;
                            break;
                    }
                    lang_area_row_item_type += 1;
                    iter++;
                }while (iter < apps_count);
            }
        }
        public void BuildMetroTable(){
            int irow = 1;
            int isubrow = 0;
            int ilangrow = 0;
            int ilangcol = 0;
            metro_table[0, 0] = "Metro Apps";
            metro_table[0, 1] = "";
            foreach (string item in metro_app_list){
                string[] var = item.Split(',');
                //Console.WriteLine($"metro_table [{var[0]}],[{var[1]}] irow:{irow}");
                metro_table[irow, 0] = var[0];
                metro_table[irow, 1] = var[1];
                //Console.WriteLine($"{metro_table[irow, 0]} {metro_table[irow, 1]}");

                irow++;
            }

            // Combine multiple lists into tuples
            var combinedLists = us_metro_app_list.Zip(ca_metro_app_list, (item1, item2) => (item1, item2))
                                    .Zip(latam_metro_app_list, (tuple, item3) => (tuple.item1, tuple.item2, item3))
                                    .Zip(frfr_metro_app_list, (tuple, item4) => (tuple.item1, tuple.item2, tuple.item3, item4))
                                    .Zip(dede_metro_app_list, (tuple, item5) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, item5))
                                    .Zip(gb_metro_app_list, (tuple, item6) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, item6))
                                    .Zip(nordic_metro_app_list, (tuple, item7) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, item7))
                                    .Zip(ru_metro_app_list, (tuple, item8) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, item8))
                                    .Zip(emea1_metro_app_list, (tuple, item9) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, item9))
                                    .Zip(au_metro_app_list, (tuple, item10) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, item10))
                                    .Zip(jp_metro_app_list, (tuple, item11) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, item11))
                                    .Zip(kr_metro_app_list, (tuple, item12) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, item12))
                                    .Zip(aap1_metro_app_list, (tuple, item13) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, item13))
                                    .Zip(cn_metro_app_list, (tuple, item14) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, tuple.item13, item14))
                                    .Zip(tw_metro_app_list, (tuple, item15) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, tuple.item13, tuple.item14, item15));

            //Fullfill all language on the top row of table
            metro_lang_check_table[0, 0] = "US"; metro_lang_check_table[0, 1] = "CA"; metro_lang_check_table[0, 2] = "LATAM";
            metro_lang_check_table[0, 3] = "FRFR"; metro_lang_check_table[0, 4] = "DEDE"; metro_lang_check_table[0, 5] = "GB";
            metro_lang_check_table[0, 6] = "NORDIC"; metro_lang_check_table[0, 7] = "RU"; metro_lang_check_table[0, 8] = "EMEA1";
            metro_lang_check_table[0, 9] = "AU"; metro_lang_check_table[0, 10] = "JP"; metro_lang_check_table[0, 11] = "KR";
            metro_lang_check_table[0, 12] = "AAP1"; metro_lang_check_table[0, 13] = "CN"; metro_lang_check_table[0, 14] = "TW";
            // Iterate over the combined lists
            foreach (var (item1, item2, item3, item4, item5, item6, item7, item8, item9, item10, item11, item12, item13, item14, item15) in combinedLists){
                //Console.WriteLine($"List1: {item1}  List2: {item2}  List3: {item3}  List4: {item4} List5: {item5}");
                string[] us = item1.Split(',');
                string[] ca = item2.Split(',');
                string[] latam = item3.Split(',');
                string[] frfr = item4.Split(',');
                string[] dede = item5.Split(',');
                string[] gb = item6.Split(',');
                string[] nordic = item7.Split(',');
                string[] ru = item8.Split(',');  
                string[] emea1 = item9.Split(',');
                string[] au = item10.Split(',');
                string[] jp = item11.Split(',');
                string[] kr = item12.Split(',');
                string[] aap1 = item13.Split(',');
                string[] cn = item14.Split(',');
                string[] tw = item15.Split(',');

                //Console.WriteLine($"ilangrow: {ilangrow}, ilangcol: {ilangcol}");
                metro_lang_check_table[ilangrow + 1, ilangcol] = us[0];
                metro_lang_check_table[ilangrow + 1, ilangcol + 1] = ca[0];
                metro_lang_check_table[ilangrow + 1, ilangcol + 2] = latam[0];
                metro_lang_check_table[ilangrow + 1, ilangcol + 3] = frfr[0];
                metro_lang_check_table[ilangrow + 1, ilangcol + 4] = dede[0];
                metro_lang_check_table[ilangrow + 1, ilangcol + 5] = gb[0];
                metro_lang_check_table[ilangrow + 1, ilangcol + 6] = nordic[0];
                metro_lang_check_table[ilangrow + 1, ilangcol + 7] = ru[0];
                metro_lang_check_table[ilangrow + 1, ilangcol + 8] = emea1[0];
                metro_lang_check_table[ilangrow + 1, ilangcol + 9] = au[0];
                metro_lang_check_table[ilangrow + 1, ilangcol + 10] = jp[0];
                metro_lang_check_table[ilangrow + 1, ilangcol + 11] = kr[0];
                metro_lang_check_table[ilangrow + 1, ilangcol + 12] = aap1[0];
                metro_lang_check_table[ilangrow + 1, ilangcol + 13] = cn[0];
                metro_lang_check_table[ilangrow + 1, ilangcol + 14] = tw[0];
                ilangrow++;
            }
        }
        public void BuildDesktopTable(){
            int irow = 1;
            int isubrow = 0;
            int ilangrow = 0;
            int ilangcol = 0;
            desktop_table[0, 0] = "Desktop Apps";
            desktop_table[0, 1] = "";
            foreach (string item in desktop_app_list){
                string[] var = item.Split(',');
                //Console.WriteLine($"desktop_table [{var[0]}],[{var[1]}] irow:{irow}");
                desktop_table[irow, 0] = var[0];
                desktop_table[irow, 1] = var[1];
                //Console.WriteLine($"desktopApps: {desktop_table[irow, 0]} {desktop_table[irow, 1]}");
                irow++;
            }

            // Combine multiple lists into tuples
            var combinedLists = us_desktop_app_list.Zip(ca_desktop_app_list, (item1, item2) => (item1, item2))
                                    .Zip(latam_desktop_app_list, (tuple, item3) => (tuple.item1, tuple.item2, item3))
                                    .Zip(frfr_desktop_app_list, (tuple, item4) => (tuple.item1, tuple.item2, tuple.item3, item4))
                                    .Zip(dede_desktop_app_list, (tuple, item5) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, item5))
                                    .Zip(gb_desktop_app_list, (tuple, item6) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, item6))
                                    .Zip(nordic_desktop_app_list, (tuple, item7) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, item7))
                                    .Zip(ru_desktop_app_list, (tuple, item8) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, item8))
                                    .Zip(emea1_desktop_app_list, (tuple, item9) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, item9))
                                    .Zip(au_desktop_app_list, (tuple, item10) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, item10))
                                    .Zip(jp_desktop_app_list, (tuple, item11) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, item11))
                                    .Zip(kr_desktop_app_list, (tuple, item12) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, item12))
                                    .Zip(aap1_desktop_app_list, (tuple, item13) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, item13))
                                    .Zip(cn_desktop_app_list, (tuple, item14) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, tuple.item13, item14))
                                    .Zip(tw_desktop_app_list, (tuple, item15) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, tuple.item13, tuple.item14, item15));

            desktop_lang_check_table[0, 0] = "US"; desktop_lang_check_table[0, 1] = "CA"; desktop_lang_check_table[0, 2] = "LATAM";
            desktop_lang_check_table[0, 3] = "FRFR"; desktop_lang_check_table[0, 4] = "DEDE"; desktop_lang_check_table[0, 5] = "GB";
            desktop_lang_check_table[0, 6] = "NORDIC"; desktop_lang_check_table[0, 7] = "RU"; desktop_lang_check_table[0, 8] = "EMEA1";
            desktop_lang_check_table[0, 9] = "AU"; desktop_lang_check_table[0, 10] = "JP"; desktop_lang_check_table[0, 11] = "KR";
            desktop_lang_check_table[0, 12] = "AAP1"; desktop_lang_check_table[0, 13] = "CN"; desktop_lang_check_table[0, 14] = "TW";
            // Iterate over the combined lists
            foreach (var (item1, item2, item3, item4, item5, item6, item7, item8, item9, item10, item11, item12, item13, item14, item15) in combinedLists){
                //Console.WriteLine($"List1: {item1}  List2: {item2}  List3: {item3}  List4: {item4} List5: {item5}");
                string[] us = item1.Split(',');
                string[] ca = item2.Split(',');
                string[] latam = item3.Split(',');
                string[] frfr = item4.Split(',');
                string[] dede = item5.Split(',');
                string[] gb = item6.Split(',');
                string[] nordic = item7.Split(',');
                string[] ru = item8.Split(',');
                string[] emea1 = item9.Split(',');
                string[] au = item10.Split(',');
                string[] jp = item11.Split(',');
                string[] kr = item12.Split(',');
                string[] aap1 = item13.Split(',');
                string[] cn = item14.Split(',');
                string[] tw = item15.Split(',');

                desktop_lang_check_table[ilangrow + 1, ilangcol] = us[0];
                desktop_lang_check_table[ilangrow + 1, ilangcol + 1] = ca[0];
                desktop_lang_check_table[ilangrow + 1, ilangcol + 2] = latam[0];
                desktop_lang_check_table[ilangrow + 1, ilangcol + 3] = frfr[0];
                desktop_lang_check_table[ilangrow + 1, ilangcol + 4] = dede[0];
                desktop_lang_check_table[ilangrow + 1, ilangcol + 5] = gb[0];
                desktop_lang_check_table[ilangrow + 1, ilangcol + 6] = nordic[0];
                desktop_lang_check_table[ilangrow + 1, ilangcol + 7] = ru[0];
                desktop_lang_check_table[ilangrow + 1, ilangcol + 8] = emea1[0];
                desktop_lang_check_table[ilangrow + 1, ilangcol + 9] = au[0];
                desktop_lang_check_table[ilangrow + 1, ilangcol + 10] = jp[0];
                desktop_lang_check_table[ilangrow + 1, ilangcol + 11] = kr[0];
                desktop_lang_check_table[ilangrow + 1, ilangcol + 12] = aap1[0];
                desktop_lang_check_table[ilangrow + 1, ilangcol + 13] = cn[0];
                desktop_lang_check_table[ilangrow + 1, ilangcol + 14] = tw[0];
                ilangrow++;
            }
        }
        public void BuildDesktopBarPINTable(){
            int irow = 1;
            int isubrow = 0;
            int ilangrow = 0;
            int ilangcol = 0;
            desktop_bar_table[0, 0] = "Desktop Taskbar Pin";
            desktop_table[0, 1] = "";
            foreach (string item in desktop_bar_app_list){
                string[] var = item.Split(',');
                //Console.WriteLine($"desktop_bar_table [{var[0]}],[{var[1]}] irow:{irow}");
                desktop_bar_table[irow, 0] = var[0];
                desktop_bar_table[irow, 1] = var[1];
                //Console.WriteLine($"desktop_bar_Apps: {desktop_bar_table[irow, 0]} {desktop_bar_table[irow, 1]}");
                irow++;
            }

            // Combine multiple lists into tuples
            var combinedLists = us_desktop_bar_app_list.Zip(ca_desktop_bar_app_list, (item1, item2) => (item1, item2))
                                    .Zip(latam_desktop_bar_app_list, (tuple, item3) => (tuple.item1, tuple.item2, item3))
                                    .Zip(frfr_desktop_bar_app_list, (tuple, item4) => (tuple.item1, tuple.item2, tuple.item3, item4))
                                    .Zip(dede_desktop_bar_app_list, (tuple, item5) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, item5))
                                    .Zip(gb_desktop_bar_app_list, (tuple, item6) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, item6))
                                    .Zip(nordic_desktop_bar_app_list, (tuple, item7) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, item7))
                                    .Zip(ru_desktop_bar_app_list, (tuple, item8) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, item8))
                                    .Zip(emea1_desktop_bar_app_list, (tuple, item9) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, item9))
                                    .Zip(au_desktop_bar_app_list, (tuple, item10) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, item10))
                                    .Zip(jp_desktop_bar_app_list, (tuple, item11) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, item11))
                                    .Zip(kr_desktop_bar_app_list, (tuple, item12) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, item12))
                                    .Zip(aap1_desktop_bar_app_list, (tuple, item13) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, item13))
                                    .Zip(cn_desktop_bar_app_list, (tuple, item14) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, tuple.item13, item14))
                                    .Zip(tw_desktop_bar_app_list, (tuple, item15) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, tuple.item13, tuple.item14, item15));

            desktop_bar_lang_check_table[0, 0] = "US"; desktop_bar_lang_check_table[0, 1] = "CA"; desktop_bar_lang_check_table[0, 2] = "LATAM";
            desktop_bar_lang_check_table[0, 3] = "FRFR"; desktop_bar_lang_check_table[0, 4] = "DEDE"; desktop_bar_lang_check_table[0, 5] = "GB";
            desktop_bar_lang_check_table[0, 6] = "NORDIC"; desktop_bar_lang_check_table[0, 7] = "RU"; desktop_bar_lang_check_table[0, 8] = "EMEA1";
            desktop_bar_lang_check_table[0, 9] = "AU"; desktop_bar_lang_check_table[0, 10] = "JP"; desktop_bar_lang_check_table[0, 11] = "KR";
            desktop_bar_lang_check_table[0, 12] = "AAP1"; desktop_bar_lang_check_table[0, 13] = "CN"; desktop_bar_lang_check_table[0, 14] = "TW";
            // Iterate over the combined lists
            foreach (var (item1, item2, item3, item4, item5, item6, item7, item8, item9, item10, item11, item12, item13, item14, item15) in combinedLists){
                //Console.WriteLine($"List1: {item1}  List2: {item2}  List3: {item3}  List4: {item4} List5: {item5}");
                string[] us = item1.Split(',');
                string[] ca = item2.Split(',');
                string[] latam = item3.Split(',');
                string[] frfr = item4.Split(',');
                string[] dede = item5.Split(',');
                string[] gb = item6.Split(',');
                string[] nordic = item7.Split(',');
                string[] ru = item8.Split(',');
                string[] emea1 = item9.Split(',');
                string[] au = item10.Split(',');
                string[] jp = item11.Split(',');
                string[] kr = item12.Split(',');
                string[] aap1 = item13.Split(',');
                string[] cn = item14.Split(',');
                string[] tw = item15.Split(',');

                desktop_bar_lang_check_table[ilangrow + 1, ilangcol] = us[0];
                desktop_bar_lang_check_table[ilangrow + 1, ilangcol + 1] = ca[0];
                desktop_bar_lang_check_table[ilangrow + 1, ilangcol + 2] = latam[0];
                desktop_bar_lang_check_table[ilangrow + 1, ilangcol + 3] = frfr[0];
                desktop_bar_lang_check_table[ilangrow + 1, ilangcol + 4] = dede[0];
                desktop_bar_lang_check_table[ilangrow + 1, ilangcol + 5] = gb[0];
                desktop_bar_lang_check_table[ilangrow + 1, ilangcol + 6] = nordic[0];
                desktop_bar_lang_check_table[ilangrow + 1, ilangcol + 7] = ru[0];
                desktop_bar_lang_check_table[ilangrow + 1, ilangcol + 8] = emea1[0];
                desktop_bar_lang_check_table[ilangrow + 1, ilangcol + 9] = au[0];
                desktop_bar_lang_check_table[ilangrow + 1, ilangcol + 10] = jp[0];
                desktop_bar_lang_check_table[ilangrow + 1, ilangcol + 11] = kr[0];
                desktop_bar_lang_check_table[ilangrow + 1, ilangcol + 12] = aap1[0];
                desktop_bar_lang_check_table[ilangrow + 1, ilangcol + 13] = cn[0];
                desktop_bar_lang_check_table[ilangrow + 1, ilangcol + 14] = tw[0];
                ilangrow++;
            }
        }
        public void BuildAllAppsShortCutTable(){
            int irow = 1;
            int isubrow = 0;
            int ilangrow = 0;
            int ilangcol = 0;
            shortCut_table[0, 0] = "All Apps Shortcut";
            shortCut_table[0, 1] = "";
            foreach (string item in shortCut_app_list){
                string[] var = item.Split(',');
                //Console.WriteLine($"shortCut_table [{var[0]}],[{var[1]}] irow:{irow}");
                shortCut_table[irow, 0] = var[0];
                shortCut_table[irow, 1] = var[1];
                //Console.WriteLine($"shortCut_Apps: {shortCut_table[irow, 0]} {shortCut_table[irow, 1]}");
                irow++;
            }
            // Combine multiple lists into tuples
            var combinedLists = us_shortCut_app_list.Zip(ca_shortCut_app_list, (item1, item2) => (item1, item2))
                                    .Zip(latam_shortCut_app_list, (tuple, item3) => (tuple.item1, tuple.item2, item3))
                                    .Zip(frfr_shortCut_app_list, (tuple, item4) => (tuple.item1, tuple.item2, tuple.item3, item4))
                                    .Zip(dede_shortCut_app_list, (tuple, item5) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, item5))
                                    .Zip(gb_shortCut_app_list, (tuple, item6) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, item6))
                                    .Zip(nordic_shortCut_app_list, (tuple, item7) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, item7))
                                    .Zip(ru_shortCut_app_list, (tuple, item8) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, item8))
                                    .Zip(emea1_shortCut_app_list, (tuple, item9) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, item9))
                                    .Zip(au_shortCut_app_list, (tuple, item10) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, item10))
                                    .Zip(jp_shortCut_app_list, (tuple, item11) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, item11))
                                    .Zip(kr_shortCut_app_list, (tuple, item12) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, item12))
                                    .Zip(aap1_shortCut_app_list, (tuple, item13) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, item13))
                                    .Zip(cn_shortCut_app_list, (tuple, item14) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, tuple.item13, item14))
                                    .Zip(tw_shortCut_app_list, (tuple, item15) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, tuple.item13, tuple.item14, item15));

            shortCut_lang_check_table[0, 0] = "US"; shortCut_lang_check_table[0, 1] = "CA"; shortCut_lang_check_table[0, 2] = "LATAM";
            shortCut_lang_check_table[0, 3] = "FRFR"; shortCut_lang_check_table[0, 4] = "DEDE"; shortCut_lang_check_table[0, 5] = "GB";
            shortCut_lang_check_table[0, 6] = "NORDIC"; shortCut_lang_check_table[0, 7] = "RU"; shortCut_lang_check_table[0, 8] = "EMEA1";
            shortCut_lang_check_table[0, 9] = "AU"; shortCut_lang_check_table[0, 10] = "JP"; shortCut_lang_check_table[0, 11] = "KR";
            shortCut_lang_check_table[0, 12] = "AAP1"; shortCut_lang_check_table[0, 13] = "CN"; shortCut_lang_check_table[0, 14] = "TW";
            // Iterate over the combined lists
            foreach (var (item1, item2, item3, item4, item5, item6, item7, item8, item9, item10, item11, item12, item13, item14, item15) in combinedLists){
                //Console.WriteLine($"List1: {item1}  List2: {item2}  List3: {item3}  List4: {item4} List5: {item5}");
                string[] us = item1.Split(',');
                string[] ca = item2.Split(',');
                string[] latam = item3.Split(',');
                string[] frfr = item4.Split(',');
                string[] dede = item5.Split(',');
                string[] gb = item6.Split(',');
                string[] nordic = item7.Split(',');
                string[] ru = item8.Split(',');
                string[] emea1 = item9.Split(',');
                string[] au = item10.Split(',');
                string[] jp = item11.Split(',');
                string[] kr = item12.Split(',');
                string[] aap1 = item13.Split(',');
                string[] cn = item14.Split(',');
                string[] tw = item15.Split(',');

                shortCut_lang_check_table[ilangrow + 1, ilangcol] = us[0];
                shortCut_lang_check_table[ilangrow + 1, ilangcol + 1] = ca[0];
                shortCut_lang_check_table[ilangrow + 1, ilangcol + 2] = latam[0];
                shortCut_lang_check_table[ilangrow + 1, ilangcol + 3] = frfr[0];
                shortCut_lang_check_table[ilangrow + 1, ilangcol + 4] = dede[0];
                shortCut_lang_check_table[ilangrow + 1, ilangcol + 5] = gb[0];
                shortCut_lang_check_table[ilangrow + 1, ilangcol + 6] = nordic[0];
                shortCut_lang_check_table[ilangrow + 1, ilangcol + 7] = ru[0];
                shortCut_lang_check_table[ilangrow + 1, ilangcol + 8] = emea1[0];
                shortCut_lang_check_table[ilangrow + 1, ilangcol + 9] = au[0];
                shortCut_lang_check_table[ilangrow + 1, ilangcol + 10] = jp[0];
                shortCut_lang_check_table[ilangrow + 1, ilangcol + 11] = kr[0];
                shortCut_lang_check_table[ilangrow + 1, ilangcol + 12] = aap1[0];
                shortCut_lang_check_table[ilangrow + 1, ilangcol + 13] = cn[0];
                shortCut_lang_check_table[ilangrow + 1, ilangcol + 14] = tw[0];
                ilangrow++;
            }
        }
        public void BuildBrowserTable(){
            int irow = 1;
            int isubrow = 0;
            int ilangrow = 0;
            int ilangcol = 0;
            metro_table[0, 0] = "Browser Favorites";
            metro_table[0, 1] = "";
            foreach (string item in browser_app_list){
                string[] var = item.Split(',');
                //Console.WriteLine($"browser_table [{var[0]}],[{var[1]}] irow:{irow}");
                browser_table[irow, 0] = var[0];
                browser_table[irow, 1] = var[1];
                //Console.WriteLine($"{browser_table[irow, 0]} {browser_table[irow, 1]}");
                irow++;
            }

            // Combine multiple lists into tuples
            var combinedLists = us_browser_app_list.Zip(ca_browser_app_list, (item1, item2) => (item1, item2))
                                    .Zip(latam_browser_app_list, (tuple, item3) => (tuple.item1, tuple.item2, item3))
                                    .Zip(frfr_browser_app_list, (tuple, item4) => (tuple.item1, tuple.item2, tuple.item3, item4))
                                    .Zip(dede_browser_app_list, (tuple, item5) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, item5))
                                    .Zip(gb_browser_app_list, (tuple, item6) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, item6))
                                    .Zip(nordic_browser_app_list, (tuple, item7) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, item7))
                                    .Zip(ru_browser_app_list, (tuple, item8) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, item8))
                                    .Zip(emea1_browser_app_list, (tuple, item9) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, item9))
                                    .Zip(au_browser_app_list, (tuple, item10) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, item10))
                                    .Zip(jp_browser_app_list, (tuple, item11) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, item11))
                                    .Zip(kr_browser_app_list, (tuple, item12) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, item12))
                                    .Zip(aap1_browser_app_list, (tuple, item13) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, item13))
                                    .Zip(cn_browser_app_list, (tuple, item14) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, tuple.item13, item14))
                                    .Zip(tw_browser_app_list, (tuple, item15) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, tuple.item13, tuple.item14, item15));

            //Fullfill all language on the top row of table
            browser_lang_check_table[0, 0] = "US"; browser_lang_check_table[0, 1] = "CA"; browser_lang_check_table[0, 2] = "LATAM";
            browser_lang_check_table[0, 3] = "FRFR"; browser_lang_check_table[0, 4] = "DEDE"; browser_lang_check_table[0, 5] = "GB";
            browser_lang_check_table[0, 6] = "NORDIC"; browser_lang_check_table[0, 7] = "RU"; browser_lang_check_table[0, 8] = "EMEA1";
            browser_lang_check_table[0, 9] = "AU"; browser_lang_check_table[0, 10] = "JP"; browser_lang_check_table[0, 11] = "KR";
            browser_lang_check_table[0, 12] = "AAP1"; browser_lang_check_table[0, 13] = "CN"; browser_lang_check_table[0, 14] = "TW";
            // Iterate over the combined lists
            foreach (var (item1, item2, item3, item4, item5, item6, item7, item8, item9, item10, item11, item12, item13, item14, item15) in combinedLists){
                //Console.WriteLine($"List1: {item1}  List2: {item2}  List3: {item3}  List4: {item4} List5: {item5}");
                string[] us = item1.Split(',');
                string[] ca = item2.Split(',');
                string[] latam = item3.Split(',');
                string[] frfr = item4.Split(',');
                string[] dede = item5.Split(',');
                string[] gb = item6.Split(',');
                string[] nordic = item7.Split(',');
                string[] ru = item8.Split(',');
                string[] emea1 = item9.Split(',');
                string[] au = item10.Split(',');
                string[] jp = item11.Split(',');
                string[] kr = item12.Split(',');
                string[] aap1 = item13.Split(',');
                string[] cn = item14.Split(',');
                string[] tw = item15.Split(',');

                //Console.WriteLine($"ilangrow: {ilangrow}, ilangcol: {ilangcol}");
                browser_lang_check_table[ilangrow + 1, ilangcol] = us[0];
                browser_lang_check_table[ilangrow + 1, ilangcol + 1] = ca[0];
                browser_lang_check_table[ilangrow + 1, ilangcol + 2] = latam[0];
                browser_lang_check_table[ilangrow + 1, ilangcol + 3] = frfr[0];
                browser_lang_check_table[ilangrow + 1, ilangcol + 4] = dede[0];
                browser_lang_check_table[ilangrow + 1, ilangcol + 5] = gb[0];
                browser_lang_check_table[ilangrow + 1, ilangcol + 6] = nordic[0];
                browser_lang_check_table[ilangrow + 1, ilangcol + 7] = ru[0];
                browser_lang_check_table[ilangrow + 1, ilangcol + 8] = emea1[0];
                browser_lang_check_table[ilangrow + 1, ilangcol + 9] = au[0];
                browser_lang_check_table[ilangrow + 1, ilangcol + 10] = jp[0];
                browser_lang_check_table[ilangrow + 1, ilangcol + 11] = kr[0];
                browser_lang_check_table[ilangrow + 1, ilangcol + 12] = aap1[0];
                browser_lang_check_table[ilangrow + 1, ilangcol + 13] = cn[0];
                browser_lang_check_table[ilangrow + 1, ilangcol + 14] = tw[0];
                ilangrow++;
            }
        }

        public void BuildOOBETable(){
            int irow = 1;
            //int isubrow = 0;
            int ilangrow = 0;
            int ilangcol = 0;
            OOBE_table[0, 0] = "Windows Welcome/OOBE Integration";
            OOBE_table[0, 1] = "";
            foreach (string item in OOBE_integration_list){
                string[] var = item.Split(',');
                //Console.WriteLine($"OOBE_table [{var[0]}],[{var[1]}] irow:{irow}");
                OOBE_table[irow, 0] = var[0];
                OOBE_table[irow, 1] = var[1];
                //Console.WriteLine($"{OOBE_table[irow, 0]} {OOBE_table[irow, 1]}");
                irow++;
            }

            // Combine multiple lists into tuples
            var combinedLists = us_OOBE_integration_list.Zip(ca_OOBE_integration_list, (item1, item2) => (item1, item2))
                                    .Zip(latam_OOBE_integration_list, (tuple, item3) => (tuple.item1, tuple.item2, item3))
                                    .Zip(frfr_OOBE_integration_list, (tuple, item4) => (tuple.item1, tuple.item2, tuple.item3, item4))
                                    .Zip(dede_OOBE_integration_list, (tuple, item5) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, item5))
                                    .Zip(gb_OOBE_integration_list, (tuple, item6) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, item6))
                                    .Zip(nordic_OOBE_integration_list, (tuple, item7) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, item7))
                                    .Zip(ru_OOBE_integration_list, (tuple, item8) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, item8))
                                    .Zip(emea1_OOBE_integration_list, (tuple, item9) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, item9))
                                    .Zip(au_OOBE_integration_list, (tuple, item10) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, item10))
                                    .Zip(jp_OOBE_integration_list, (tuple, item11) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, item11))
                                    .Zip(kr_OOBE_integration_list, (tuple, item12) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, item12))
                                    .Zip(aap1_OOBE_integration_list, (tuple, item13) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, item13))
                                    .Zip(cn_OOBE_integration_list, (tuple, item14) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, tuple.item13, item14))
                                    .Zip(tw_OOBE_integration_list, (tuple, item15) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, tuple.item13, tuple.item14, item15));

            //Fullfill all language on the top row of table
            OOBE_lang_check_table[0, 0] = "US"; OOBE_lang_check_table[0, 1] = "CA"; OOBE_lang_check_table[0, 2] = "LATAM";
            OOBE_lang_check_table[0, 3] = "FRFR"; OOBE_lang_check_table[0, 4] = "DEDE"; OOBE_lang_check_table[0, 5] = "GB";
            OOBE_lang_check_table[0, 6] = "NORDIC"; OOBE_lang_check_table[0, 7] = "RU"; OOBE_lang_check_table[0, 8] = "EMEA1";
            OOBE_lang_check_table[0, 9] = "AU"; OOBE_lang_check_table[0, 10] = "JP"; OOBE_lang_check_table[0, 11] = "KR";
            OOBE_lang_check_table[0, 12] = "AAP1"; OOBE_lang_check_table[0, 13] = "CN"; OOBE_lang_check_table[0, 14] = "TW";
            // Iterate over the combined lists
            foreach (var (item1, item2, item3, item4, item5, item6, item7, item8, item9, item10, item11, item12, item13, item14, item15) in combinedLists){
                //Console.WriteLine($"List1: {item1}  List2: {item2}  List3: {item3}  List4: {item4} List5: {item5}");

                string[] us = item1.Split(',');
                string[] ca = item2.Split(',');
                string[] latam = item3.Split(',');
                string[] frfr = item4.Split(',');
                string[] dede = item5.Split(',');
                string[] gb = item6.Split(',');
                string[] nordic = item7.Split(',');
                string[] ru = item8.Split(',');
                string[] emea1 = item9.Split(',');
                string[] au = item10.Split(',');
                string[] jp = item11.Split(',');
                string[] kr = item12.Split(',');
                string[] aap1 = item13.Split(',');
                string[] cn = item14.Split(',');
                string[] tw = item15.Split(',');

                //Console.WriteLine($"ilangrow: {ilangrow}, ilangcol: {ilangcol}");
                OOBE_lang_check_table[ilangrow + 1, ilangcol] = us[0];
                OOBE_lang_check_table[ilangrow + 1, ilangcol + 1] = ca[0];
                OOBE_lang_check_table[ilangrow + 1, ilangcol + 2] = latam[0];
                OOBE_lang_check_table[ilangrow + 1, ilangcol + 3] = frfr[0];
                OOBE_lang_check_table[ilangrow + 1, ilangcol + 4] = dede[0];
                OOBE_lang_check_table[ilangrow + 1, ilangcol + 5] = gb[0];
                OOBE_lang_check_table[ilangrow + 1, ilangcol + 6] = nordic[0];
                OOBE_lang_check_table[ilangrow + 1, ilangcol + 7] = ru[0];
                OOBE_lang_check_table[ilangrow + 1, ilangcol + 8] = emea1[0];
                OOBE_lang_check_table[ilangrow + 1, ilangcol + 9] = au[0];
                OOBE_lang_check_table[ilangrow + 1, ilangcol + 10] = jp[0];
                OOBE_lang_check_table[ilangrow + 1, ilangcol + 11] = kr[0];
                OOBE_lang_check_table[ilangrow + 1, ilangcol + 12] = aap1[0];
                OOBE_lang_check_table[ilangrow + 1, ilangcol + 13] = cn[0];
                OOBE_lang_check_table[ilangrow + 1, ilangcol + 14] = tw[0];
                ilangrow++;
            }
        }
        public void BuildWindowsNextRecommendedTable(){
            int irow = 1;
            int isubrow = 0;
            int ilangrow = 0;
            int ilangcol = 0;
            next_recommended_table[0, 0] = "Windows Next Recommended";
            next_recommended_table[0, 1] = "";
            foreach (string item in Windows_Next_Recommended_list){
                string[] var = item.Split(',');
                //Console.WriteLine($"OOBE_table [{var[0]}],[{var[1]}] irow:{irow}");
                next_recommended_table[irow, 0] = var[0];
                next_recommended_table[irow, 1] = var[1];
                //Console.WriteLine($"{next_recommended_table[irow, 0]} {next_recommended_table[irow, 1]}");

                irow++;
            }

            // Combine multiple lists into tuples
            var combinedLists = us_next_recommended_list.Zip(ca_next_recommended_list, (item1, item2) => (item1, item2))
                                    .Zip(latam_next_recommended_list, (tuple, item3) => (tuple.item1, tuple.item2, item3))
                                    .Zip(frfr_next_recommended_list, (tuple, item4) => (tuple.item1, tuple.item2, tuple.item3, item4))
                                    .Zip(dede_next_recommended_list, (tuple, item5) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, item5))
                                    .Zip(gb_next_recommended_list, (tuple, item6) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, item6))
                                    .Zip(nordic_next_recommended_list, (tuple, item7) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, item7))
                                    .Zip(ru_next_recommended_list, (tuple, item8) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, item8))
                                    .Zip(emea1_next_recommended_list, (tuple, item9) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, item9))
                                    .Zip(au_next_recommended_list, (tuple, item10) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, item10))
                                    .Zip(jp_next_recommended_list, (tuple, item11) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, item11))
                                    .Zip(kr_next_recommended_list, (tuple, item12) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, item12))
                                    .Zip(aap1_next_recommended_list, (tuple, item13) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, item13))
                                    .Zip(cn_next_recommended_list, (tuple, item14) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, tuple.item13, item14))
                                    .Zip(tw_next_recommended_list, (tuple, item15) => (tuple.item1, tuple.item2, tuple.item3, tuple.item4, tuple.item5, tuple.item6, tuple.item7, tuple.item8, tuple.item9, tuple.item10, tuple.item11, tuple.item12, tuple.item13, tuple.item14, item15));

            //Fullfill all language on the top row of table
            next_recommended_lang_check_table[0, 0] = "US"; next_recommended_lang_check_table[0, 1] = "CA"; next_recommended_lang_check_table[0, 2] = "LATAM";
            next_recommended_lang_check_table[0, 3] = "FRFR"; next_recommended_lang_check_table[0, 4] = "DEDE"; next_recommended_lang_check_table[0, 5] = "GB";
            next_recommended_lang_check_table[0, 6] = "NORDIC"; next_recommended_lang_check_table[0, 7] = "RU"; next_recommended_lang_check_table[0, 8] = "EMEA1";
            next_recommended_lang_check_table[0, 9] = "AU"; next_recommended_lang_check_table[0, 10] = "JP"; next_recommended_lang_check_table[0, 11] = "KR";
            next_recommended_lang_check_table[0, 12] = "AAP1"; next_recommended_lang_check_table[0, 13] = "CN"; next_recommended_lang_check_table[0, 14] = "TW";
            // Iterate over the combined lists
            foreach (var (item1, item2, item3, item4, item5, item6, item7, item8, item9, item10, item11, item12, item13, item14, item15) in combinedLists){
                //Console.WriteLine($"List1: {item1}  List2: {item2}  List3: {item3}  List4: {item4} List5: {item5}");

                string[] us = item1.Split(',');
                string[] ca = item2.Split(',');
                string[] latam = item3.Split(',');
                string[] frfr = item4.Split(',');
                string[] dede = item5.Split(',');
                string[] gb = item6.Split(',');
                string[] nordic = item7.Split(',');
                string[] ru = item8.Split(',');
                string[] emea1 = item9.Split(',');
                string[] au = item10.Split(',');
                string[] jp = item11.Split(',');
                string[] kr = item12.Split(',');
                string[] aap1 = item13.Split(',');
                string[] cn = item14.Split(',');
                string[] tw = item15.Split(',');

                //Console.WriteLine($"ilangrow: {ilangrow}, ilangcol: {ilangcol}");
                next_recommended_lang_check_table[ilangrow + 1, ilangcol] = us[0];
                next_recommended_lang_check_table[ilangrow + 1, ilangcol + 1] = ca[0];
                next_recommended_lang_check_table[ilangrow + 1, ilangcol + 2] = latam[0];
                next_recommended_lang_check_table[ilangrow + 1, ilangcol + 3] = frfr[0];
                next_recommended_lang_check_table[ilangrow + 1, ilangcol + 4] = dede[0];
                next_recommended_lang_check_table[ilangrow + 1, ilangcol + 5] = gb[0];
                next_recommended_lang_check_table[ilangrow + 1, ilangcol + 6] = nordic[0];
                next_recommended_lang_check_table[ilangrow + 1, ilangcol + 7] = ru[0];
                next_recommended_lang_check_table[ilangrow + 1, ilangcol + 8] = emea1[0];
                next_recommended_lang_check_table[ilangrow + 1, ilangcol + 9] = au[0];
                next_recommended_lang_check_table[ilangrow + 1, ilangcol + 10] = jp[0];
                next_recommended_lang_check_table[ilangrow + 1, ilangcol + 11] = kr[0];
                next_recommended_lang_check_table[ilangrow + 1, ilangcol + 12] = aap1[0];
                next_recommended_lang_check_table[ilangrow + 1, ilangcol + 13] = cn[0];
                next_recommended_lang_check_table[ilangrow + 1, ilangcol + 14] = tw[0];
                ilangrow++;
            }
        }

        ///--------------------------------------------------------------------
        /// <summary>
        /// Fetch supported apps list of Merto App by Language
        /// ex: call GetSupportMetroAppList("US"), GetSupportMetroAppList("FRFR")
        /// </summary>
        /// <param name="language">
        /// Language of different country or region
        /// support: US|CA|LATAM|FRFR|DEDE|GB|NORDIC|RU|EMEA1|AU|JP|KR|AAP1|CN|TW
        /// </param>
        /// <returns>
        /// support app list and remark char in the table of SCL file, slotting sheet
        /// </returns>
        ///--------------------------------------------------------------------
        public static List<string> GetSupportMetroAppList(string language){
            var list = new List<string>();
            int lang_col_index = -1;
            for (int row = 0; row < 18; row++){
                for (int col = 0; col < 15; col++){
                    if (metro_lang_check_table[row, col] == language){
                        Console.WriteLine($"GetSupportMetroAppList===>Language:{language} row:{row}, col:{col}");
                        lang_col_index = col;
                    }
                }
            }

            for (int fetch_index = 0; fetch_index < 18; fetch_index++){
                if (metro_lang_check_table[fetch_index, lang_col_index] != "" && fetch_index > 0){
                    Console.WriteLine($"{metro_table[fetch_index, 0]}, {metro_lang_check_table[fetch_index, lang_col_index]}");
                    list.Add(metro_table[fetch_index, 0] + "," + metro_lang_check_table[fetch_index, lang_col_index]);
                }
                //Console.WriteLine(metro_lang_check_table[fetch_index, lang_col_index]);
            }
            return list;
        }
        ///--------------------------------------------------------------------
        /// <summary>
        /// Fetch supported apps list of Desktop App by Language
        /// ex: call GetSupportdesktopAppList("US"), GetSupportdesktopAppList("FRFR")
        /// </summary>
        /// <param name="language">
        /// Language of different country or region
        /// support: US|CA|LATAM|FRFR|DEDE|GB|NORDIC|RU|EMEA1|AU|JP|KR|AAP1|CN|TW
        /// </param>
        /// <returns>
        /// support app list and remark char in the table of slotting sheet, SCL file
        /// </returns>
        ///--------------------------------------------------------------------
        public static List<string> GetSupportdesktopAppList(string language){
            var list = new List<string>();
            int lang_col_index = -1;
            for (int row = 0; row < 12; row++){
                for (int col = 0; col < 15; col++){
                    if (desktop_lang_check_table[row, col] == language){
                        Console.WriteLine($"GetSupportDesktopAppList===>Language:{language} row:{row}, col:{col}");
                        lang_col_index = col;
                    }
                }
            }

            for (int fetch_index = 0; fetch_index < 12; fetch_index++){
                if (desktop_lang_check_table[fetch_index, lang_col_index] != "" && fetch_index > 0){
                    Console.WriteLine($"{desktop_table[fetch_index, 0]}, {desktop_lang_check_table[fetch_index, lang_col_index]}");
                    list.Add(desktop_table[fetch_index, 0] + "," + desktop_lang_check_table[fetch_index, lang_col_index]);
                }
                //Console.WriteLine(metro_lang_check_table[fetch_index, lang_col_index]);
            }
            return list;
        }
        ///--------------------------------------------------------------------
        /// <summary>
        /// Fetch supported apps list of Desktop bar App by Language
        /// ex: call GetSupportdesktopBarAppList("US"), GetSupportdesktopBarAppList("FRFR")
        /// </summary>
        /// <param name="language">
        /// Language of different country or region
        /// support: US|CA|LATAM|FRFR|DEDE|GB|NORDIC|RU|EMEA1|AU|JP|KR|AAP1|CN|TW
        /// </param>
        /// <returns>
        /// support app list and remark char in the table of slotting sheet, SCL file
        /// </returns>
        ///--------------------------------------------------------------------
        public static List<string> GetSupportdesktopBarAppList(string language){
            var list = new List<string>();
            int lang_col_index = -1;
            for (int row = 0; row < 9; row++){
                for (int col = 0; col < 15; col++){
                    if (desktop_bar_lang_check_table[row, col] == language){
                        Console.WriteLine($"GetSupportDesktopBarAppList===>Language:{language} row:{row}, col:{col}");
                        lang_col_index = col;
                    }
                }
            }

            for (int fetch_index = 0; fetch_index < 9; fetch_index++){
                if (desktop_bar_lang_check_table[fetch_index, lang_col_index] != "" && fetch_index > 0){
                    Console.WriteLine($"{desktop_bar_table[fetch_index, 0]}, {desktop_bar_lang_check_table[fetch_index, lang_col_index]}");
                    list.Add(desktop_bar_table[fetch_index, 0] + "," + desktop_bar_lang_check_table[fetch_index, lang_col_index]);
                }
                //Console.WriteLine(metro_lang_check_table[fetch_index, lang_col_index]);
            }
            return list;
        }
        ///--------------------------------------------------------------------
        /// <summary>
        /// Fetch supported apps list of shortcut App by Language
        /// ex: call GetSupportShortCutAppList("US"), GetSupportShortCutAppList("FRFR")
        /// </summary>
        /// <param name="language">
        /// Language of different country or region
        /// support: US|CA|LATAM|FRFR|DEDE|GB|NORDIC|RU|EMEA1|AU|JP|KR|AAP1|CN|TW
        /// </param>
        /// <returns>
        /// support app list and remark char in the table of slotting sheet, SCL file
        /// </returns>
        ///--------------------------------------------------------------------
        public static List<string> GetSupportShortCutAppList(string language){
            var list = new List<string>();
            int lang_col_index = -1;
            for (int row = 0; row < 22; row++){
                for (int col = 0; col < 15; col++){
                    if (shortCut_lang_check_table[row, col] == language){
                        Console.WriteLine($"GetSupportShortCutAppList===>Language:{language} row:{row}, col:{col}");
                        lang_col_index = col;
                    }
                }
            }

            for (int fetch_index = 0; fetch_index < 22; fetch_index++){
                if (shortCut_lang_check_table[fetch_index, lang_col_index] != "" && fetch_index > 0){
                    Console.WriteLine($"{shortCut_table[fetch_index, 0]}, {shortCut_lang_check_table[fetch_index, lang_col_index]}");
                    list.Add(shortCut_table[fetch_index, 0] + "," + shortCut_lang_check_table[fetch_index, lang_col_index]);
                }
                //Console.WriteLine(metro_lang_check_table[fetch_index, lang_col_index]);
            }
            return list;
        }
        ///--------------------------------------------------------------------
        /// <summary>
        /// Fetch supported apps list of browser App by Language
        /// ex: call GetSupportBrowserAppList("US"), GetSupportBrowserAppList("FRFR")
        /// </summary>
        /// <param name="language">
        /// Language of different country or region
        /// support: US|CA|LATAM|FRFR|DEDE|GB|NORDIC|RU|EMEA1|AU|JP|KR|AAP1|CN|TW
        /// </param>
        /// <returns>
        /// support app list and remark char in the table of slotting sheet, SCL file
        /// </returns>
        ///--------------------------------------------------------------------
        public static List<string> GetSupportBrowserAppList(string language){
            var list = new List<string>();
            int lang_col_index = -1;
            for (int row = 0; row < 6; row++){
                for (int col = 0; col < 15; col++){
                    if (browser_lang_check_table[row, col] == language){
                        Console.WriteLine($"GetSupportBrowserAppList===>Language:{language} row:{row}, col:{col}");
                        lang_col_index = col;
                    }
                }
            }

            for (int fetch_index = 0; fetch_index < 6; fetch_index++){
                if (browser_lang_check_table[fetch_index, lang_col_index] != "" && fetch_index > 0){
                    Console.WriteLine($"{browser_table[fetch_index, 0]}, {browser_lang_check_table[fetch_index, lang_col_index]}");
                    list.Add(browser_table[fetch_index, 0] + "," + browser_lang_check_table[fetch_index, lang_col_index]);
                }
                //Console.WriteLine(metro_lang_check_table[fetch_index, lang_col_index]);
            }
            return list;
        }
        ///--------------------------------------------------------------------
        /// <summary>
        /// Fetch supported apps list of OOBE Integration App by Language
        /// ex: call GetSupportOOBEIntegrationList("US"), GetSupportOOBEIntegrationList("FRFR")
        /// </summary>
        /// <param name="language">
        /// Language of different country or region
        /// support: US|CA|LATAM|FRFR|DEDE|GB|NORDIC|RU|EMEA1|AU|JP|KR|AAP1|CN|TW
        /// </param>
        /// <returns>
        /// support app list and remark char in the table of slotting sheet, SCL file
        /// </returns>
        ///--------------------------------------------------------------------
        public static List<string> GetSupportOOBEIntegrationList(string language){
            var list = new List<string>();
            int lang_col_index = -1;
            for (int row = 0; row < 2; row++){
                for (int col = 0; col < 15; col++){
                    if (OOBE_lang_check_table[row, col] == language){
                        Console.WriteLine($"GetSupportOOBEIntegrationList===>Language:{language} row:{row}, col:{col}");
                        lang_col_index = col;
                    }
                }
            }

            for (int fetch_index = 0; fetch_index < 2; fetch_index++){
                if (OOBE_lang_check_table[fetch_index, lang_col_index] != "" && fetch_index > 0){
                    Console.WriteLine($"{OOBE_table[fetch_index, 0]}, {OOBE_lang_check_table[fetch_index, lang_col_index]}");
                    list.Add(OOBE_table[fetch_index, 0] + "," + OOBE_lang_check_table[fetch_index, lang_col_index]);
                }
                //Console.WriteLine(metro_lang_check_table[fetch_index, lang_col_index]);
            }
            return list;
        }
        ///--------------------------------------------------------------------
        /// <summary>
        /// Fetch supported apps list of Windows Next Recommended App by Language
        /// ex: call GetSupportWindowsNextRecommendedList("US"), GetSupportWindowsNextRecommendedList("FRFR")
        /// </summary>
        /// <param name="language">
        /// Language of different country or region
        /// support: US|CA|LATAM|FRFR|DEDE|GB|NORDIC|RU|EMEA1|AU|JP|KR|AAP1|CN|TW
        /// </param>
        /// <returns>
        /// support app list and remark char in the table of slotting sheet, SCL file
        /// </returns>
        ///--------------------------------------------------------------------
        public static List<string> GetSupportWindowsNextRecommendedList(string language){
            var list = new List<string>();
            int lang_col_index = -1;
            for (int row = 0; row < 3; row++){
                for (int col = 0; col < 15; col++){
                    if (next_recommended_lang_check_table[row, col] == language){
                        Console.WriteLine($"GetSupportWindowsNextRecommendedList===>Language:{language} row:{row}, col:{col}");
                        lang_col_index = col;
                    }
                }
            }

            for (int fetch_index = 0; fetch_index < 3; fetch_index++){
                if (next_recommended_lang_check_table[fetch_index, lang_col_index] != "" && fetch_index > 0){
                    Console.WriteLine($"{next_recommended_table[fetch_index, 0]}, {next_recommended_lang_check_table[fetch_index, lang_col_index]}");
                    list.Add(next_recommended_table[fetch_index, 0] + "," + next_recommended_lang_check_table[fetch_index, lang_col_index]);
                }
                //Console.WriteLine(next_recommended_lang_check_table[fetch_index, lang_col_index]);
            }
            return list;
        }
    }
}
