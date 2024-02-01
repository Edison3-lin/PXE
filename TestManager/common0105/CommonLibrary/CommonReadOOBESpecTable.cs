using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaptainWin.CommonAPI{
    public class CommonReadOOBESpecTable{
        /// <summary>
        /// Read timeZone in OS and return a string of Display time 
        /// </summary>
        /// <returns></returns>
        public static string GetTimeZone(){
            // Get the local time zone
            TimeZoneInfo localTimeZone = TimeZoneInfo.Local;

            // Display time zone information
            Console.WriteLine($"Time Zone Id: {localTimeZone.Id}");
            Console.WriteLine($"Display Name: {localTimeZone.DisplayName}");
            Console.WriteLine($"Standard Time Name: {localTimeZone.StandardName}");
            Console.WriteLine($"Daylight Time Name: {localTimeZone.DaylightName}");
            Console.WriteLine($"UTC Offset: {localTimeZone.BaseUtcOffset}");

            // Optionally, you can also display information about the current daylight saving time rules
            if (localTimeZone.SupportsDaylightSavingTime){
                TimeZoneInfo.AdjustmentRule daylightSavingRule = localTimeZone.GetAdjustmentRules()[0];
                Console.WriteLine($"Daylight Saving Time Start: {daylightSavingRule.DaylightTransitionStart}");
                Console.WriteLine($"Daylight Saving Time End: {daylightSavingRule.DaylightTransitionEnd}");
            }

            return localTimeZone.DisplayName;
        }

        public static string[,] ConvertListToArray(List<string> list, int columns){
            // Calculate the number of rows needed in the array
            //int rows = (int)Math.Ceiling((double)list.Count / columns);
            int irow = 0;
            string[,] OOBE_SPEC_table = new string[list.Count, columns];
            foreach (string item in list){
                string[] mitem = item.Split('|');
                OOBE_SPEC_table[irow, 0] = mitem[0];
                OOBE_SPEC_table[irow, 1] = mitem[1];
                OOBE_SPEC_table[irow, 2] = mitem[2];
                OOBE_SPEC_table[irow, 3] = mitem[3];
                irow++;
            }
            return OOBE_SPEC_table;
        }
        /// <summary>
        /// Read Excel data and build a table for query
        /// </summary>
        public static void Setup(){

            // Get the Windows user account name
            string userName = Environment.UserName;


            // 設定Excel檔案的路徑
            string root_path = @"C:\\Users\\" + userName + "\\Documents\\";
            string excelFileName = "Win11_OOBE_SPEC.xlsx";
            string excelFilePath = root_path + excelFileName;
            Console.WriteLine(excelFilePath);
            // 建立一個新的Excel Application物件
            Excel.Application excelApp = new Excel.Application();

            // 打開Excel檔案
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);

            // 假設Excel檔案只有一個工作表，使用索引Lang_Region_Keyboard_Timezone來取得該OOBE SPEC工作表
            Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets["Lang_Region_Keyboard_Timezone"];

            // 讀取資料
            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;


            int GAIA_id = 0;
            int SWBOM_id = 0;
            int LIP_id = 0;
            //app name cell value
            string TimeZone_cellValue = null;
            string SWBOM_cellValue = null;
            string LIP_cellValue = null;

            for (int row = 1; row <= rowCount; row++){
                for (int col = 1; col <= colCount; col++){
                    // 使用Cells物件來取得單元格的值
                    Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, col];
                    string cellValue = cell.Value != null ? cell.Value.ToString() : "";

                    //build SWBOM app table
                    if (cellValue.IndexOf("SWBOM") >= 0){
                        //get row index offset of "Driver"
                        int swbom_base_row = row;
                        //get col index offset of "Driver"
                        int swbom_base_col = col;

                        area_row_item_type = swbom_base_row + 1;
                        area_col_item_type = swbom_base_col;
                        Console.WriteLine();
                        do{
                            //Read Category cell
                            Excel.Range TimeZone_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[area_row_item_type, swbom_base_col + 5];
                            TimeZone_cellValue = TimeZone_cell.Value != null ? TimeZone_cell.Value.ToString() : "";
                            //Console.WriteLine($"TimeZone_cellValue: {TimeZone_cellValue}| GAIA_id: {GAIA_id}");

                            Excel.Range SWBOM_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[area_row_item_type, swbom_base_col];
                            SWBOM_cellValue = SWBOM_cell.Value != null ? SWBOM_cell.Value.ToString() : "";
                            //Console.WriteLine($"SWBOM_cellValue: {SWBOM_cellValue}| SWBOM_id: {SWBOM_id}");

                            Excel.Range LIP_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[area_row_item_type, swbom_base_col + 1];
                            LIP_cellValue = LIP_cell.Value != null ? LIP_cell.Value.ToString() : "";
                            //Console.WriteLine($"LIP_cellValue: {LIP_cellValue}| LIP_cellValue: {LIP_id}");

                            OOBE_SPEC_list.Add(SWBOM_cellValue + "|" + LIP_cellValue + "|" + TimeZone_cellValue + "|" + GAIA_id);
                            GAIA_id++;
                            SWBOM_id++;
                            LIP_id++;
                            area_row_item_type += 1;
                        }while (SWBOM_cellValue != "ENGC");//end of table
                        Console.WriteLine(area_row_item_type + " " + GAIA_id);
                    }
                }
            }
        }
        /// <summary>
        /// Get OOBE Spec data and build a table, will return OOBE query reult
        /// </summary>
        /// <param name="lang"></param>
        /// <returns></returns>
        public string GetOOBESpec(string lang){
            //Build OOBE SPEC searching table
            string[,] mOOBESPECTable = ConvertListToArray(OOBE_SPEC_list, 4);
            for (int i = 0; i < OOBE_SPEC_list.Count; i++){
                if (mOOBESPECTable[i, 0] == lang){
                    Console.WriteLine(mOOBESPECTable[i, 2]);
                    return mOOBESPECTable[i, 2];
                }
            }
            //no match in the table, it should be something wrong happen...
            return null;
        }
        public static int area_row_item_type = 0;
        public static int area_col_item_type = 0;
        public static List<string> OOBE_SPEC_list = new List<string>();
    }
}
