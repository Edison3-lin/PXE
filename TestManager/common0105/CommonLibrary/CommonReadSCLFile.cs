using System;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Management;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaptainWin.CommonAPI{
    public class CommonReadSCLFile{
        public static int area_row_item_type = 0;
        public static int area_col_item_type = 0;
        public static string finalfilename = "";
        /// <summary>
        /// To get DUT UUID by using wmi infterface api
        /// </summary>
        /// <returns>
        /// return UUID of DUT
        /// </returns>
        public static string GetWindowsUUID(){
            string uuid = string.Empty;
            string l_uuid = "";
            try{
                string ComputerName = "localhost";
                ManagementScope Scope;
                Scope = new ManagementScope(String.Format("\\\\{0}\\root\\CIMV2", ComputerName), null);
                Scope.Connect();
                ObjectQuery Query = new ObjectQuery("SELECT UUID FROM Win32_ComputerSystemProduct");
                ManagementObjectSearcher Searcher = new ManagementObjectSearcher(Scope, Query);

                foreach (ManagementObject WmiObject in Searcher.Get()){
                    Console.WriteLine("{0,-35} {1,-40}", "UUID", WmiObject["UUID"]);// String
                    l_uuid = WmiObject["UUID"].ToString();
                }
            }
            catch (Exception e){
                Console.WriteLine(String.Format("Exception {0} Trace {1}", e.Message, e.StackTrace));
            }
            return l_uuid;
        }
        /// <summary>
        /// Read files name of a specific path on FTP server
        /// </summary>
        /// <param name="ftpServer"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <param name="remoteDirectory"></param>
        /// <returns>
        /// return a array of string which conatin files name in the speific path
        /// </returns>
        public static string[] GetFileNamesFromFTP(string ftpServer, string username, string password, string remoteDirectory){
            try{
                // Create the FTP request
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create($"{ftpServer}/{remoteDirectory}");
                request.Method = WebRequestMethods.Ftp.ListDirectory;
                request.Credentials = new NetworkCredential(username, password);

                // Get the FTP response
                using (FtpWebResponse response = (FtpWebResponse)request.GetResponse()){
                    // Read the response stream
                    using (StreamReader reader = new StreamReader(response.GetResponseStream())){
                        // Read and split the file names
                        string fileNamesString = reader.ReadToEnd();
                        Console.WriteLine(fileNamesString);
                        string[] fileNames = fileNamesString.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                        return fileNames;
                    }
                }
            }catch (Exception ex){
                // Handle FTP-related exceptions
                Console.WriteLine($"Error: {ex.Message}");
                return new string[0]; // Return an empty array on error
            }
        }
        /// <summary>
        /// To call this function to download SCL file from FTP server by input project name
        /// </summary>
        /// <param name="QuantaProjectName"></param>
        public static void FTPDownloadFile(string QuantaProjectName){
            // FTP server details
            string ftpServerUrl = "ftp://127.0.0.9:2121";
            string ftpFilePath = $"/Image/SCD/{QuantaProjectName}/{finalfilename}";
            string ftpUsername = "sit001";
            string ftpPassword = "sit1234";

            // Local file path to save the downloaded file
            string localFilePath = $"C:\\TestManager\\{finalfilename}";

            try{
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create($"{ftpServerUrl}{ftpFilePath}");
                request.Method = WebRequestMethods.Ftp.DownloadFile;
                request.Credentials = new NetworkCredential(ftpUsername, ftpPassword);

                using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
                using (Stream responseStream = response.GetResponseStream())
                using (FileStream fileStream = File.Create(localFilePath))
                {
                    responseStream.CopyTo(fileStream);
                }

                Console.WriteLine($"File downloaded to: {localFilePath}");
            }catch (WebException ex){
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        /// <summary>
        /// Query project name of DUT in the MS database by using UUID of DUT
        /// </summary>
        /// <param name="local_DUT_UUID"></param>
        /// <returns>
        /// will return project name of DUT if data exist, return empty if data does not exist
        /// </returns>
        public static string AutoTestDataBase(string local_DUT_UUID){
            string cellValue = "";
            // Replace "YourConnectionString" with the actual connection string for your SQL Server database
            string connectionString = "Data Source=172.0.0.9;Initial Catalog=SIT_TEST;User ID=Captain001;Password=Captaintest2023@SIT;";

            // Replace "YourTableName" with the actual table name
            string tableName = "DUT_Profile";

            // Replace "ColumnName" and "ConditionColumn" with the actual column names
            string columnName = "DP_Project_Code";
            string conditionColumn = "DP_UUID";
            string conditionValue = local_DUT_UUID; // Replace with the desired condition value

            // Create a SELECT query with a WHERE clause to filter based on a condition
            string selectQuery = $"SELECT {columnName} FROM {tableName} WHERE {conditionColumn} = '{conditionValue}'";

            // Connect to the database and execute the query
            using (SqlConnection connection = new SqlConnection(connectionString)){
                connection.Open();

                using (SqlCommand command = new SqlCommand(selectQuery, connection)){
                    // Add a parameter for the condition value to prevent SQL injection
                    command.Parameters.AddWithValue("@ConditionValue", conditionValue);

                    using (SqlDataReader reader = command.ExecuteReader()){
                        // Check if there are rows
                        if (reader.HasRows){
                            // Read the first row (assuming there is only one matching row)
                            reader.Read();

                            // Replace "ColumnName" with the actual column name you want to retrieve
                            cellValue = reader[columnName].ToString();
                            Console.WriteLine($"Cell Value: {cellValue}");
                        }else{
                            Console.WriteLine("No matching rows found.");
                        }
                        return cellValue;
                    }
                }
            }
        }
        public static FileInfo GetLargestFile(string directoryPath){
            try{
                DirectoryInfo directory = new DirectoryInfo(directoryPath);

                // Get all files in the directory
                FileInfo[] files = directory.GetFiles();

                if (files.Length > 0){
                    // Find the file with the largest size
                    FileInfo largestFile = files.OrderByDescending(f => f.Length).First();
                    return largestFile;
                }else{
                    return null;
                }
            }catch (Exception ex){
                // Handle directory access or other exceptions
                Console.WriteLine($"Error: {ex.Message}");
                return null;
            }
        }

        public static string ReadIniFile(string filePath, string sectionName, string key){
            try{
                // Read all lines from the INI file
                string[] lines = File.ReadAllLines(filePath);

                // Flag to indicate if the current line is inside the target section
                bool insideSection = false;

                // Iterate through each line
                foreach (string line in lines){
                    // Trim leading and trailing whitespace
                    string trimmedLine = line.Trim();

                    // Check if the line is empty or a comment
                    if (string.IsNullOrWhiteSpace(trimmedLine) || trimmedLine.StartsWith(";")){
                        continue; // Skip empty lines or comments
                    }

                    // Check if the line is the start of the target section
                    if (trimmedLine.StartsWith($"[{sectionName}]")){
                        insideSection = true;
                        continue; // Skip the section header line
                    }

                    // Check if the line is outside the target section
                    if (!insideSection){
                        continue;
                    }

                    // Check if the line contains the specified key
                    if (trimmedLine.StartsWith(key + "=")){
                        // Extract the value part and return it
                        return trimmedLine.Substring(key.Length + 1);
                    }
                }
                // Key not found
                return null;
            }catch (Exception ex){
                // Handle file reading or other exceptions
                Console.WriteLine($"Error: {ex.Message}");
                return null;
            }
        }

        public static string GetLargestFile(){
            // Replace "your-directory-path" with the actual directory path
            string directoryPath = "C:\\OEM\\Preload\\Command\\PAP";

            FileInfo largestFile = GetLargestFile(directoryPath);

            if (largestFile != null){
                Console.WriteLine($"Largest File: {largestFile.FullName}");
                //OpenFile(largestFile.FullName);
                return largestFile.FullName;
            }else{
                Console.WriteLine("No files found in the specified directory.");
                return "";
            }
        }
        /// <summary>
        /// Get and Read inf file in DUT local path to get the SCL version in the DUT
        /// </summary>
        /// <returns>
        /// return SCL file version or return empty string if read value fail
        /// </returns>
        public static string GetLocalDUTSCLFileVersion(){
            string iniFilePath = GetLargestFile();
            string sectionName = "Main";
            string keyToRetrieve = "Image Version";

            string value = ReadIniFile(iniFilePath, sectionName, keyToRetrieve);

            if (value != null){
                Console.WriteLine($"Value for {keyToRetrieve}: {value}");
                return value;
            }else{
                Console.WriteLine($"Key '{keyToRetrieve}' not found in section '{sectionName}'");
                return "";
            }
        }
        /// <summary>
        /// Search value of keyword and return the value in a string format
        /// user must keyin sheetname of excel file and the keyword which want to get 
        /// set parameter "search_by_col" to choose "search by column" or "search by row"
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="searchKeyword"></param>
        /// <param name="search_by_col"></param>
        /// <returns>
        /// </returns>
        public static string SearchKeyListinSCLFile(string sheetName, string searchKeyword, bool search_by_col){

            // Get the Windows user account name
            string userName = Environment.UserName;


            // set Excel file path
            //string root_path = "C:\\Users\\" + userName + "\\Downloads\\";
            string root_path = @"C:\\Users\\" + userName + "\\Documents\\";
            string excelFileName = "TEST_SCD_RV07RC.xls";
            string excelFilePath = root_path + excelFileName;
            Console.WriteLine(excelFilePath);
            // new a Excel Application object
            Excel.Application excelApp = new Excel.Application();

            // open excel file
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);

            // use index "Lang_Region_Keyboard_Timezone" to get "OOBE SPEC" worksheet
            Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[sheetName];

            // read data
            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;
            int row_base = 0;
            int col_base = 0;
            int append_index = 0;
            string searchcellValue = null;
            for (int row = 1; row <= rowCount; row++){
                for (int col = 1; col <= colCount; col++){
                    // get "cellValue" by using "cell" object
                    Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, col];
                    string cellValue = cell.Value != null ? cell.Value.ToString() : "";

                    //build Metro app table
                    if (cellValue.IndexOf(searchKeyword) >= 0){
                        row_base = row;
                        col_base = col;
                        if (search_by_col == true){
                            do{
                                Excel.Range searchcell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row_base, col + append_index];
                                searchcellValue = searchcell.Value != null ? searchcell.Value.ToString() : "";
                                Console.WriteLine($"searchcellValue{searchcellValue}");
                                append_index++;
                            }while (searchcellValue != "");
                        }else {
                            do{
                                Excel.Range searchcell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row_base + append_index, col];
                                searchcellValue = searchcell.Value != null ? searchcell.Value.ToString() : "";
                                Console.WriteLine(searchcellValue);
                                append_index++;
                            } while (searchcellValue != "");
                        }
                    }else{
                        searchcellValue = "not found";
                    }
                }
            }
            return searchcellValue;
        }
        /// <summary>
        /// To call the function can tell you if the file exists in the specific path
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns>
        /// true, file exist in the path
        /// false, file does not exist in the path
        /// </returns>
        public static bool CheckSCLFileExist(string filePath)
        {
            //string filePath = "C:\\TestManager\\SCD_RV07RC.xls";

            // Check if the file exists
            if (File.Exists(filePath))
            {
                Console.WriteLine($"The file at '{filePath}' exists.");
                return true;
            }
            else
            {
                Console.WriteLine($"The file at '{filePath}' does not exist.");
                return false;
            }
        }
        /// <summary>
        /// Download SCL file from FTP Server
        /// This function will call GetWindowsUUID to get UUID of DUT and use the UUID to query MS database
        /// By calling AutoTestDataBase(), we can get project name of DUT. Base on the project name, we can
        /// get path of SCL file on FTP server and download the SCL file to local side.
        /// </summary>
        public static void DownloadSCLFileFromFTP(){
            //string DUT_UUID = "651976C6-6AF1-8E4E-B2FD-AD489F3C76AF";
            string DUT_UUID = GetWindowsUUID();
            //Use UUID to query project name stored in database
            string projectName = AutoTestDataBase(DUT_UUID);
            //FTP server ip and port location
            string ftpServer = "ftp://127.0.0.9:2121";
            // Replace "username" and "password" with FTP credentials
            string username = "sit001";
            string password = "sit1234";
            // Specify the remote directory on the FTP server
            string remoteDirectory = $"/Image/SCD/{projectName}";

            // Connect to the FTP server and retrieve the list of file names
            string[] fileNames = GetFileNamesFromFTP(ftpServer, username, password, remoteDirectory);

            // Display the list of file names
            foreach (string fileName in fileNames){
                int index = fileName.IndexOf('/');
                if (index != -1){
                    // Use Substring to get the portion of the string after the delimiter
                    finalfilename = fileName.Substring(index + 1);
                    Console.WriteLine(finalfilename);
                }
                Console.WriteLine(fileName);
                FTPDownloadFile(projectName);
            }
        }
    }
}
