using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Common;

namespace image_installation_driver_default
{
    public class image_installation_driver_default
    {
        // Define constants and structures
        private const int DIGCF_PRESENT = 0x000000002;
        private const int DIGCF_ALLCLASSES = 0x000000004;

        [StructLayout(LayoutKind.Sequential)]
        public struct SP_DEVINFO_DATA
        {
            public int cbSize;
            public Guid ClassGuid;
            public int DevInst;
            public IntPtr Reserved;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct SP_DRVINFO_DATA
        {
            public uint cbSize;
            public uint DriverType;
            public IntPtr Reserved;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
            public string Description;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
            public string MfgName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
            public string ProviderName;
            public System.Runtime.InteropServices.ComTypes.FILETIME DriverDate;
            public uint DriverVersion;
        }

        [DllImport("setupapi.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SetupDiGetClassDevs(
            ref Guid ClassGuid,
            string Enumerator,
            IntPtr hwndParent,
            int Flags
        );

        [DllImport("setupapi.dll", CharSet = CharSet.Auto)]
        public static extern bool SetupDiEnumDeviceInfo(
            IntPtr DeviceInfoSet,
            int MemberIndex,
            ref SP_DEVINFO_DATA DeviceInfoData
        );

        [DllImport("setupapi.dll", CharSet = CharSet.Auto)]
        public static extern bool SetupDiGetDeviceRegistryProperty(
            IntPtr DeviceInfoSet,
            ref SP_DEVINFO_DATA DeviceInfoData,
            int Property,
            out int PropertyRegDataType,
            IntPtr PropertyBuffer,
            int PropertyBufferSize,
            out int RequiredSize
        );

        [DllImport("setupapi.dll", SetLastError = true)]
        public static extern bool SetupDiEnumDriverInfo(
            IntPtr DeviceInfoSet,
            SP_DEVINFO_DATA DeviceInfoData,
            uint DriverType, uint MemberIndex,
            SP_DRVINFO_DATA DriverInfoData
        );


        [DllImport("setupapi.dll")]
        private static extern Boolean SetupDiDestroyDeviceInfoList(IntPtr DeviceInfoSet);

        public const int SPDRP_DEVICEDESC = 0x00000000;
        public const int SPDRP_DRIVER = 0x00000009;
        public const int SPDRP_INSTALL_STATE = 0x00000022;  // Device Install State (R)
        const int SPDRP_HARDWAREID = 0x00000001; // Hardware ID
        const int SPDIT_COMPATDRIVER = 0x00000002; // Get driver information for compatible drivers.

        [DllImport("setupapi.dll", SetLastError = true)]
        public static extern bool SetupDiBuildDriverInfoList(
            IntPtr DeviceInfoSet,
            SP_DEVINFO_DATA DeviceInfoData,
            uint DriverType
        );

        // P/Invoke declarations
        [DllImport("setupapi.dll")]
        public static extern int CM_Locate_DevNode(
            out uint pdnDevInst,
            string pDeviceID,
            uint ulFlags
        );

        [DllImport("setupapi.dll")]
        public static extern int CM_Get_DevNode_Status(
            out uint pulStatus,
            out uint pulProblemNumber,
            int dnDevInst,
            uint ulFlags
        );


        public static string GetDeviceName(IntPtr deviceInfoSet, SP_DEVINFO_DATA deviceInfoData)
        {
            int requiredSize = 0;
            SetupDiGetDeviceRegistryProperty(deviceInfoSet, ref deviceInfoData, SPDRP_DEVICEDESC, out int regDataType, IntPtr.Zero, 0, out requiredSize);
            if (requiredSize == 0)
                return string.Empty;

            IntPtr propertyBuffer = Marshal.AllocHGlobal(requiredSize);
            if (SetupDiGetDeviceRegistryProperty(deviceInfoSet, ref deviceInfoData, SPDRP_DEVICEDESC, out regDataType, propertyBuffer, requiredSize, out requiredSize))
            {
                string deviceName = Marshal.PtrToStringAuto(propertyBuffer);
                Marshal.FreeHGlobal(propertyBuffer);
                return deviceName;
            }

            return string.Empty;
        }

        public static string GetDriverVersion(IntPtr deviceInfoSet, SP_DEVINFO_DATA deviceInfoData)
        {
            int requiredSize = 0;
            SetupDiGetDeviceRegistryProperty(deviceInfoSet, ref deviceInfoData, SPDRP_DRIVER, out int regDataType, IntPtr.Zero, 0, out requiredSize);
            if (requiredSize == 0)
                return string.Empty;

            IntPtr propertyBuffer = Marshal.AllocHGlobal(requiredSize);
            if (SetupDiGetDeviceRegistryProperty(deviceInfoSet, ref deviceInfoData, SPDRP_DRIVER, out regDataType, propertyBuffer, requiredSize, out requiredSize))
            {
                string driverVersion = Marshal.PtrToStringAuto(propertyBuffer);
                Marshal.FreeHGlobal(propertyBuffer);

                // Specify the registry key path and value name you want to read.
                string keyPath = @"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Class\" + driverVersion;
                string valueName = "DriverVersion";

                // Use the Registry.GetValue method to read the registry value.
                object value = Registry.GetValue(keyPath, valueName, null);

                if (value != null)
                {
                    //Console.WriteLine($"Value of {valueName} in {keyPath}: {value}");
                }
                else
                {
                    Console.WriteLine($"Registry value {valueName} not found in {keyPath}");
                }

                return value.ToString();
            }

            return string.Empty;
        }

        public static string GetHardwareID(IntPtr deviceInfoSet, SP_DEVINFO_DATA deviceInfoData)
        {
            int requiredSize = 0;
            SetupDiGetDeviceRegistryProperty(deviceInfoSet, ref deviceInfoData, SPDRP_HARDWAREID, out int regDataType, IntPtr.Zero, 0, out requiredSize);
            if (requiredSize == 0)
                return string.Empty;

            IntPtr propertyBuffer = Marshal.AllocHGlobal(requiredSize);
            if (SetupDiGetDeviceRegistryProperty(deviceInfoSet, ref deviceInfoData, SPDRP_HARDWAREID, out regDataType, propertyBuffer, requiredSize, out requiredSize))
            {
                string hardwareid = Marshal.PtrToStringAuto(propertyBuffer);
                Marshal.FreeHGlobal(propertyBuffer);
                return hardwareid;
            }

            return string.Empty;
        }

        public static string GetDevicesStatusAndProblemCode(IntPtr deviceInfoSet, SP_DEVINFO_DATA deviceInfoData)
        {
            // Find the device node for the specified hardware ID.
            uint devInst = 0;

            uint status = 0;
            uint problemNumber = 0;

            int result = CM_Locate_DevNode(out devInst, null, 0);

            if (result == 0)
            {
                // Device node found, now get its status.

                result = CM_Get_DevNode_Status(out status, out problemNumber, deviceInfoData.DevInst, 0);

                if (result == 0)
                {
                    //Console.WriteLine("Device Status: " + status);
                    //Console.WriteLine("Device Problem Code: " + problemNumber);
                }
                else
                {
                    Console.WriteLine("Failed to get device status.");
                }
                return problemNumber.ToString();
            }
            else
            {
                Console.WriteLine("Device not found or error locating the device node.");
                return string.Empty;
            }
        }

        static List<string> ReadInfFiles(string directoryPath)
        {
            List<string> infContents = new List<string>();

            try
            {
                // Get a list of all .inf files in the specified directory
                string[] infFiles = Directory.GetFiles(directoryPath, "*.inf");

                // Loop through each .inf file and read its content
                foreach (string infFile in infFiles)
                {
                    // Read the file line by line
                    foreach (string line in File.ReadLines(infFile))
                    {
                        // Check if the line contains the target string
                        if (line.Contains("DriverVer ="))
                        {
                            infContents.Add(line);
                            break; // Optionally, break after the first match if needed
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                // Handle any exceptions (e.g., directory not found, permission issues)
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            return infContents;
        }

        public static string GetCpuManufacturer()
        {
            try
            {
                using (RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(@"HARDWARE\DESCRIPTION\System\CentralProcessor\0", RegistryKeyPermissionCheck.ReadSubTree))
                {
                    if (registryKey != null)
                    {
                        object value = registryKey.GetValue("ProcessorNameString");

                        if (value != null)
                        {
                            string processorName = value.ToString();

                            if (processorName.Contains("AMD"))
                            {
                                return "AMD";
                            }
                            else if (processorName.Contains("Intel"))
                            {
                                return "Intel";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle any exceptions that may occur during registry access
                Console.WriteLine("Error: " + ex.Message);
            }

            return "Unknown";
        }

        static string[] setup()
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

        public int Setup()
        {
            // common.Setup
            Testflow.Setup("xxx");
            return 11;
        }

        public static bool Run()
        {
            //Get all Device info
            string displayName;
            string app_version;
            string app_vendor;
            string strSystemComponent;
            string project_name = null;
            List<string> checkflaglist = new List<string>();
            //Flag of Check list for device driver
            string Had_touchpad = "not check";
            string AcerAirplaneModeController = "not check";
            string ApplicationBasedriver = "not check";
            string RealtekAudio = "not check";
            string AcerPurifiedVoiceConsole = "not check";
            string RealtekAudioConsoleUWP = "not check";
            string IntelWirelessBluetooth = "not check";
            string RealtekPCIECardReader = "not check";
            string DTSXUltra = "not check";
            string DTSConsoleUWP = "not check";
            string DTSsoundUWP = "not check";
            string AcerDeviceEnablingSevice = "not check";
            string FingerPrint = "not check";
            string IntelDPTF = "not check";
            int IntelDPTF_count = 0;
            string IntelGNA = "not check";
            string IntelRapidStorageWinPeReDriver = "not check";
            string InteliRST = "not check";
            string IntelISST = "not check";
            int IntelISST_count = 0;
            string IntelSerialIO = "not check";
            int IntelSerialIO_count = 0;
            string IntelSerialIOWinPEREDrivers = "not check";
            string UMA = "not check";
            string IntelSMBus = "not check";
            string Had_external_VGA = "not check";
            string Had_NVIDIA_Utility = "not check";
            int Had_NVDIA_Utility_count = 0;
            string IntelManagementEngineInterface = "not check";
            string KillerWiFi6EAX1675i = "not check";
            string WirelessLAN_MWinREDrivers = "not check";
            string KillerControlCenter = "not check";
            //build list for check
            checkflaglist.Add("Had_touchpad" + "," + "not check");
            checkflaglist.Add("AcerAirplaneModeController" + "," + "not check");
            checkflaglist.Add("ApplicationBasedriver" + "," + "not check");
            checkflaglist.Add("RealtekAudio" + "," + "not check");
            checkflaglist.Add("AcerPurifiedVoiceConsole" + "," + "not check");
            checkflaglist.Add("RealtekAudioConsoleUWP" + "," + "not check");
            checkflaglist.Add("IntelWirelessBluetooth" + "," + "not check");
            checkflaglist.Add("RealtekPCIECardReader" + "," + "not check");
            checkflaglist.Add("DTSXUltra" + "," + "not check");
            checkflaglist.Add("DTSConsoleUWP" + "," + "not check");
            checkflaglist.Add("DTSsoundUWP" + "," + "not check");
            checkflaglist.Add("AcerDeviceEnablingSevice" + "," + "not check");
            checkflaglist.Add("FingerPrint" + "," + "not check");
            checkflaglist.Add("IntelDPTF" + "," + "not check");
            checkflaglist.Add("IntelGNA" + "," + "not check");
            checkflaglist.Add("IntelRapidStorageWinPeReDriver" + "," + "not check");
            checkflaglist.Add("InteliRST" + "," + "not check");
            checkflaglist.Add("IntelISST" + "," + "not check");
            checkflaglist.Add("IntelSerialIO" + "," + "not check");
            checkflaglist.Add("IntelSerialIOWinPEREDrivers" + "," + "not check");
            checkflaglist.Add("UMA" + "," + "not check");
            checkflaglist.Add("IntelSMBus" + "," + "not check");
            checkflaglist.Add("Had_external_VGA" + "," + "not check");
            checkflaglist.Add("Had_NVIDIA_Utility" + "," + "not check");
            checkflaglist.Add("IntelManagementEngineInterface" + "," + "not check");
            checkflaglist.Add("KillerWiFi6EAX1675i" + "," + "not check");
            checkflaglist.Add("WirelessLAN_MWinREDrivers" + "," + "not check");
            checkflaglist.Add("KillerControlCenter" + "," + "not check");

            List<string> AlldevicesInfoInDM = new List<string>();//List for store all device info read from OS
            CultureInfo ci = CultureInfo.InstalledUICulture;
            Console.WriteLine("Default Language Info:");
            Console.WriteLine("* Name: {0}", ci.Name);
            Console.WriteLine(GetCpuManufacturer());

            //Get project name and PC name
            string[] PCInformation = new string[2];
            PCInformation = setup();
            Console.WriteLine(PCInformation[0] + " " + PCInformation[1]);

            string computerName = Environment.MachineName;
            Console.WriteLine("Computer Name: " + computerName);
            string registryKeyPath = @"HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS"; // Replace with the actual registry key path
            string valueName = "SystemProductName"; // Replace with the name of the specific value you want to retrieve

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

                    Console.WriteLine($"Registry Value ({valueName}): {value} : {project_name}");
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
            string userName = Environment.UserName;
            // 設定Excel檔案的路徑
            //string root_path = "C:\\Users\\" + userName + "\\Downloads\\";
            string root_path = @"C:\TestManager\ItemDownload\";
            string excelFileName = "SCD_RV07RC.xls";
            string excelFilePath = root_path + excelFileName;

            Console.WriteLine(excelFilePath);

            int area_row_item_type = 0;
            int area_col_item_type = 0;

            // 建立一個新的Excel Application物件
            Excel.Application excelApp = new Excel.Application();

            // 打開Excel檔案
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);

            // 假設Excel檔案只有一個工作表，直接使用索引1來取得該工作表
            Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets["SCL Content"];

            // 讀取資料
            int rowCount = worksheet.UsedRange.Rows.Count;
            int colCount = worksheet.UsedRange.Columns.Count;
            string drivername_cellValue = null;
            string Category_cellValue = null;
            string Provider_cellValue = null;
            string version_cellValue = null;
            string sub_brand_name_cellValue = null;
            string checkflag = "not check";
            List<string> drivers_list_SCL = new List<string>();
            //List<String> HadcheckTable = new List<String>();
            //string[] mapping_name = null;
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

                    if (cellValue == "Driver")
                    {
                        Console.Write(cellValue + "\t");
                        //Excel.Range cColumn = sheet.get_Range("B", null);
                        int driver_base_row = row;//get row index offset of "Driver"
                        int driver_base_col = col;//get col index offset of "Driver"

                        area_row_item_type = driver_base_row + 2;
                        area_col_item_type = driver_base_col + 2;

                        if (sub_brand_name_cellValue == PCInformation[0])
                        {
                            do
                            {
                                //Read Category cell
                                Excel.Range Category_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[area_row_item_type, driver_base_col];
                                Category_cellValue = Category_cell.Value != null ? Category_cell.Value.ToString() : "";
                                //Read Provider cell
                                Excel.Range Provider_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[area_row_item_type, driver_base_col + 1];
                                Provider_cellValue = Provider_cell.Value != null ? Provider_cell.Value.ToString() : "";
                                //Read Item Type cell
                                Excel.Range Item_Type_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[area_row_item_type, area_col_item_type];
                                drivername_cellValue = Item_Type_cell.Value != null ? Item_Type_cell.Value.ToString() : "";
                                //Read Driver version
                                Excel.Range Driver_Version_cell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[area_row_item_type, area_col_item_type + 2];
                                version_cellValue = Driver_Version_cell.Value != null ? Driver_Version_cell.Value.ToString() : "";
                                version_cellValue = version_cellValue.Substring(1);

                                Console.WriteLine(Category_cellValue + " " + Provider_cellValue + " " + drivername_cellValue + " " + version_cellValue + "\t");

                                if (Category_cellValue == "Audio Codec_M" && Provider_cellValue == "REALTEK")
                                {

                                    string mapping_name = "";
                                    mapping_name = "Realtek Audio";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);

                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (Category_cellValue == "Audio Driver Utility" && Provider_cellValue == "Acer")
                                {
                                    string mapping_name = "";
                                    mapping_name = "AcerPurifiedVoiceConsole";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);

                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (Category_cellValue == "Bluetooth_M" && Provider_cellValue == "Killer")
                                {
                                    string mapping_name = "";
                                    mapping_name = "Intel(R) Wireless Bluetooth";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);

                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (Category_cellValue == "Wireless LAN_M" && Provider_cellValue == "Killer" && drivername_cellValue == "1675i")
                                {
                                    string mapping_name = "Killer(R) Wi-Fi 6E AX1675i 160MHz Wireless Network Adapter";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);

                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (Category_cellValue == "Audio Driver Utility" && Provider_cellValue == "Realtek")
                                {
                                    string mapping_name = "";
                                    mapping_name = "RealtekAudioControl";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    //HadcheckTable.Add(mapping_name);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (Category_cellValue == "Intel DPTF" && Provider_cellValue == "Intel")
                                {
                                    string mapping_name = "";
                                    mapping_name = "Intel(R) Innovation Platform Framework Generic Participant";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                    mapping_name = "Intel(R) Innovation Platform Framework Manager";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                    mapping_name = "Intel(R) Innovation Platform Framework Processor Participant";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (Category_cellValue == "Application Base driver" && Provider_cellValue == "Acer")
                                {
                                    string mapping_name = "";
                                    mapping_name = "Acer Application Base Driver";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (Category_cellValue == "FUB" && Provider_cellValue == "Acer")
                                {
                                    string mapping_name = "";
                                    mapping_name = "Acer Device Enabling Sevice";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (Category_cellValue == "TouchPad" && Provider_cellValue == "Synaptics")
                                {
                                    string mapping_name = "";
                                    mapping_name = "Synaptics";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (Category_cellValue == "TouchPad" && Provider_cellValue == "Elantech")
                                {
                                    string mapping_name = "";
                                    mapping_name = "ELAN";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (Category_cellValue == "Card Reader Chip" && Provider_cellValue == "Realtek")
                                {
                                    string mapping_name = "";
                                    mapping_name = "Realtek PCIE CardReader";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (drivername_cellValue == "DTSX/Ultra" && Category_cellValue == "DTS")
                                {
                                    string mapping_name = "";
                                    mapping_name = "DTS APO4x Service";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (drivername_cellValue == "DTS Console UWP" && Category_cellValue == "DTS Utility")
                                {
                                    string mapping_name = "";
                                    mapping_name = "DTSXUltra";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (drivername_cellValue == "DTS sound UWP" && Category_cellValue == "DTS Utility")
                                {
                                    string mapping_name = "";
                                    mapping_name = "DTSSoundUnbound";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (Category_cellValue == "Finger Print_M" && Provider_cellValue == "Carewe")
                                {
                                    string mapping_name = "";
                                    mapping_name = "Fingerprint";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (Category_cellValue == "Intel Gaussian and Neural Accelerator" && Provider_cellValue == "Intel")
                                {
                                    string mapping_name = "";
                                    mapping_name = "Intel(R) GNA";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (Category_cellValue == "Intel Rapid Storage Technology" && Provider_cellValue == "Intel")
                                {
                                    string mapping_name = "";
                                    mapping_name = "Intel RST VMD";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (Category_cellValue == "Intel SST" && Provider_cellValue == "Intel")
                                {
                                    if (version_cellValue.Contains(".00"))
                                    {
                                        version_cellValue = version_cellValue.Replace(".00", ".0");
                                    }
                                    string mapping_name = "";
                                    mapping_name = "Smart Sound Technology BUS";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                    mapping_name = "Smart Sound Technology OED";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (drivername_cellValue == "Serial I/O" && Provider_cellValue == "Intel")
                                {
                                    string mapping_name = "";
                                    mapping_name = "Intel(R) Serial IO";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (Category_cellValue == "Intel VGA" && Provider_cellValue == "Intel")
                                {
                                    string mapping_name = "";
                                    mapping_name = "Graphics";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (Category_cellValue == "NB_Chipset_M" && Provider_cellValue == "Intel")
                                {
                                    string mapping_name = "";
                                    mapping_name = "Intel(R) SMBus";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (drivername_cellValue == "GN20-P0" && Provider_cellValue == "NVIDIA")
                                {
                                    string mapping_name = "";
                                    mapping_name = "NVIDIA GeForce RTX 3050";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (drivername_cellValue == "GN21-X2" && Provider_cellValue == "NVIDIA")
                                {
                                    string mapping_name = "";
                                    mapping_name = "NVIDIA GeForce RTX 4050";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                //looking for all NVIDIA apps in control panel
                                else if (Category_cellValue == "NVIDIA VGA Utility" && Provider_cellValue == "NVIDIA")
                                {
                                    string mapping_name = "";
                                    if (project_name == "Swift")
                                    {
                                        mapping_name = "NVIDIA Canvas";
                                        drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                        Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                    }
                                    mapping_name = "NVIDIA FrameView SDK";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                    mapping_name = "NVIDIA GeForce Experience";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                    mapping_name = "NVIDIA Graphics Driver";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                    mapping_name = "NVIDIA HD Audio Driver";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                    mapping_name = "NVIDIA PhysX System Software";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                    mapping_name = "NVIDIAControlPanel";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (drivername_cellValue == "Manageability Engine Code" && Provider_cellValue == "Intel")
                                {
                                    string mapping_name = "";
                                    mapping_name = "Intel(R) Management Engine Interface";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (drivername_cellValue == "WinPERE Drivers" && Provider_cellValue == "Intel" && Category_cellValue == "Intel Rapid Storage Technology")
                                {
                                    //WinRe driver do onthing here ....
                                    string mapping_name = "Intel Rapid Storage WinPERE Drivers";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");

                                }
                                else if (drivername_cellValue == "WinPERE Drivers" && Provider_cellValue == "Intel" && Category_cellValue == "Intel Serial I/O")
                                {
                                    //WinRe driver do onthing here ....
                                    string mapping_name = "Intel Serial I/O WinPERE Drivers";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");

                                }
                                else if (drivername_cellValue == "Killer control center" && Provider_cellValue == "Killer")
                                {
                                    string mapping_name = "";
                                    mapping_name = "KillerControlCenter";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");
                                }
                                else if (drivername_cellValue == "WinRE Drivers" && Category_cellValue == "Wireless LAN_M" && Provider_cellValue == "Killer")
                                {
                                    string mapping_name = "Wireless LAN_M WinRE Drivers";
                                    drivers_list_SCL.Add(mapping_name + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                    Console.WriteLine($"SCL  name: {drivername_cellValue}  ---> mapping name: {mapping_name} ");

                                }
                                else
                                {
                                    drivers_list_SCL.Add(drivername_cellValue + "," + version_cellValue + "," + Category_cellValue + "," + Provider_cellValue + "," + checkflag);
                                }

                                area_row_item_type += 1;

                            } while (drivername_cellValue != "Killer control center");

                            Console.WriteLine("Fininshed !!!");
                        }
                        else
                        {
                            Console.WriteLine($"sub brand name: {sub_brand_name_cellValue} does not match PC namme: {PCInformation[0]}");
                            Console.WriteLine("Please put correct SCL file, Stop to check installed device driver .....");
                        }

                    }
                }
            }

            // 關閉Excel檔案
            workbook.Close();
            excelApp.Quit();

            // 釋放資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            //Get installed device driver list from registry
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", false))
            {
                foreach (String keyName in key.GetSubKeyNames())
                {
                    RegistryKey subkey = key.OpenSubKey(keyName);
                    displayName = subkey.GetValue("DisplayName") as string;
                    app_version = subkey.GetValue("DisplayVersion") as string;
                    app_vendor = subkey.GetValue("Publisher") as string;
                    strSystemComponent = subkey.GetValue("SystemComponent") as string;
                    //Console.WriteLine(strSystemComponent);
                    if (string.IsNullOrEmpty(displayName))
                        continue;
                    AlldevicesInfoInDM.Add(app_vendor + "," + displayName + "," + app_version + "," + strSystemComponent);
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
                    //Console.WriteLine("strSystemComponent: ", strSystemComponent);
                    if (string.IsNullOrEmpty(displayName))
                        continue;
                    AlldevicesInfoInDM.Add(app_vendor + "," + displayName + "," + app_version + "," + strSystemComponent);
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
                    AlldevicesInfoInDM.Add(app_vendor + "," + displayName + "," + app_version + "," + strSystemComponent);
                }
            }

            //Get Registry key value for device driver
            using (RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\PackageRepository\Packages"))
            {
                if (registryKey != null)
                {
                    // Get the names of the subkeys
                    string[] subKeyNames = registryKey.GetSubKeyNames();
                    // Display the subkey names
                    foreach (string subKeyName in subKeyNames)
                    {
                        AlldevicesInfoInDM.Add(subKeyName);
                    }
                }
                else
                {
                    Console.WriteLine("Registry Key not found.");
                }
            }

            //Query all devices in DM
            Guid guid = Guid.Empty; // List all devices
            IntPtr deviceInfoSet = SetupDiGetClassDevs(ref guid, null, IntPtr.Zero, DIGCF_PRESENT | DIGCF_ALLCLASSES);

            if (deviceInfoSet != IntPtr.Zero)
            {
                SP_DEVINFO_DATA deviceInfoData = new SP_DEVINFO_DATA();
                deviceInfoData.cbSize = Marshal.SizeOf(typeof(SP_DEVINFO_DATA));
                int index = 0;
                while (SetupDiEnumDeviceInfo(deviceInfoSet, index, ref deviceInfoData))
                {
                    string deviceName = GetDeviceName(deviceInfoSet, deviceInfoData);
                    string driverVersion = GetDriverVersion(deviceInfoSet, deviceInfoData);
                    string hardwareid = GetHardwareID(deviceInfoSet, deviceInfoData);
                    string devicestatus = GetDevicesStatusAndProblemCode(deviceInfoSet, deviceInfoData);
                    if (deviceName != null && driverVersion != null)
                    {
                        string result = deviceName + "," + driverVersion + "," + hardwareid + "," + devicestatus;
                        AlldevicesInfoInDM.Add(result);
                    }

                    index++;
                }

                // Clean up
                Marshal.FreeHGlobal(deviceInfoSet);
            }

            if (sub_brand_name_cellValue == PCInformation[0])
            {
                //WinRe/WinPeRe driver check here: .....
                Console.WriteLine("WinRe or WinPeRe driver check here .............");
                string directoryPath = @"C:\Windows\INF"; // Replace with the actual directory path

                List<string> infContents = ReadInfFiles(directoryPath);
                StreamWriter sw = new StreamWriter(root_path + "\\installedDriverlist.txt");
                foreach (var key in AlldevicesInfoInDM)
                {
                    Console.WriteLine(key.ToString());
                    sw.Write(key.ToString());
                    sw.Write('\n');
                }
                sw.Close();
                //Try to do comparing between SCL list and all applications list
                string filePath = "C:\\installedDriverlist.txt"; // Replace with the path to your text file

                foreach (string drivername in drivers_list_SCL)
                {
                    string[] driverInfo_SCL = null;
                    driverInfo_SCL = drivername.Split(',');
                    Console.WriteLine(driverInfo_SCL[0] + " " + driverInfo_SCL[1]);

                    if (driverInfo_SCL[0] != " ")
                    {
                        try
                        {
                            // Read the file line by line
                            using (StreamReader reader = new StreamReader(filePath))
                            {
                                int lineNumber = 1;
                                string line;

                                while ((line = reader.ReadLine()) != null)
                                {

                                    if (driverInfo_SCL[0] == "NVIDIA Canvas" ||
                                        driverInfo_SCL[0] == "NVIDIA FrameView SDK" ||
                                        driverInfo_SCL[0] == "NVIDIA GeForce Experience" ||
                                        driverInfo_SCL[0] == "NVIDIA Graphics Driver" ||
                                        driverInfo_SCL[0] == "NVIDIA HD Audio Driver" ||
                                        driverInfo_SCL[0] == "NVIDIA PhysX System Software" ||
                                        driverInfo_SCL[0] == "NVIDIAControlPanel")
                                    {
                                        if (driverInfo_SCL[0] == "NVIDIA PhysX System Software")
                                        {
                                            if (line.IndexOf("PhysX", StringComparison.CurrentCultureIgnoreCase) >= 0)
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                
                                                Had_NVDIA_Utility_count += 1;
                                            }

                                        }
                                        if (driverInfo_SCL[0] == "NVIDIA HD Audio Driver")
                                        {
                                            if (line.IndexOf("HDAUDIO", StringComparison.CurrentCultureIgnoreCase) >= 0)
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                
                                                Had_NVDIA_Utility_count += 1;
                                            }
                                        }
                                        if (driverInfo_SCL[0] == "NVIDIA Graphics Driver")
                                        {
                                            if (line.IndexOf("Laptop GPU", StringComparison.CurrentCultureIgnoreCase) >= 0)
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                
                                                Had_NVDIA_Utility_count += 1;
                                            }

                                        }
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                            
                                            Had_NVDIA_Utility_count += 1;
                                        }


                                        if ((Had_NVDIA_Utility_count == 7 && project_name == "Swift") || Had_NVDIA_Utility_count == 6)
                                        {
                                            Had_NVIDIA_Utility = "Checked";
                                            int index = checkflaglist.FindIndex(s => s == "Had_NVIDIA_Utility,not check");
                                            if (index != -1)
                                            {
                                                // Modify the value at the found index
                                                checkflaglist[index] = "Had_NVIDIA_Utility,checked";
                                            }
                                        }

                                    }
                                    else if (driverInfo_SCL[0] == "Synaptics" && Had_touchpad == "not check")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                Had_touchpad = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "Had_touchpad,not check");
                                                if (index != -1)
                                                {
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "Had_touchpad,checked";
                                                }
                                                Console.WriteLine($"--------> Had_touchpad {Had_touchpad}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "ELAN" && Had_touchpad == "not check")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                Had_touchpad = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "Had_touchpad,not check");
                                                if (index != -1)
                                                {
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "Had_touchpad,checked";
                                                }
                                                Console.WriteLine($"----------> Had_touchpad {Had_touchpad}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "Acer Airplane Mode Controller")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                AcerAirplaneModeController = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "AcerAirplaneModeController,not check");
                                                if (index != -1)
                                                {
                                                    
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "AcerAirplaneModeController,checked";
                                                }
                                                Console.WriteLine($"----------> AcerAirplaneModeController {AcerAirplaneModeController}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "Acer Application Base Driver")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                ApplicationBasedriver = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "ApplicationBasedriver,not check");
                                                if (index != -1)
                                                {
                                                    
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "ApplicationBasedriver,checked";
                                                }
                                                Console.WriteLine($"----------> ApplicationBasedriver {ApplicationBasedriver}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "Realtek Audio")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                RealtekAudio = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "RealtekAudio,not check");
                                                if (index != -1)
                                                {
                                                    
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "RealtekAudio,checked";
                                                }
                                                Console.WriteLine($"---------> RealtekAudio {RealtekAudio}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "AcerPurifiedVoiceConsole")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                AcerPurifiedVoiceConsole = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "AcerPurifiedVoiceConsole,not check");
                                                if (index != -1)
                                                {
                                                    
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "AcerPurifiedVoiceConsole,checked";
                                                }
                                                Console.WriteLine($"-------> AcerPurifiedVoiceConsole {AcerPurifiedVoiceConsole}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "RealtekAudioControl")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                RealtekAudioConsoleUWP = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "RealtekAudioConsoleUWP,not check");
                                                if (index != -1)
                                                {
                                            
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "RealtekAudioConsoleUWP,checked";
                                                }
                                                Console.WriteLine($"-------> RealtekAudioConsoleUWP {RealtekAudioConsoleUWP}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "Intel(R) Wireless Bluetooth")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                IntelWirelessBluetooth = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "IntelWirelessBluetooth,not check");
                                                if (index != -1)
                                                {
                                               
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "IntelWirelessBluetooth,checked";
                                                }
                                                Console.WriteLine($"-------> RealtekAudioConsoleUWP {IntelWirelessBluetooth}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "Realtek PCIE CardReader")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                RealtekPCIECardReader = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "RealtekPCIECardReader,not check");
                                                if (index != -1)
                                                {
                                               
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "RealtekPCIECardReader,checked";
                                                }
                                                Console.WriteLine($"-------> RealtekPCIECardReader {RealtekPCIECardReader}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "DTS APO4x Service")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                DTSXUltra = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "DTSXUltra,not check");
                                                if (index != -1)
                                                {
                                                
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "DTSXUltra,checked";
                                                }
                                                Console.WriteLine($"-------> DTSXUltra {DTSXUltra}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "DTSXUltra")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                DTSConsoleUWP = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "DTSConsoleUWP,not check");
                                                if (index != -1)
                                                {
                                                   
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "DTSConsoleUWP,checked";
                                                }
                                                Console.WriteLine($"-------> DTSConsoleUWP {DTSConsoleUWP}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "DTSSoundUnbound")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                DTSsoundUWP = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "DTSsoundUWP,not check");
                                                if (index != -1)
                                                {
                                                  
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "DTSsoundUWP,checked";
                                                }
                                                Console.WriteLine($"-------> DTSsoundUWP {DTSsoundUWP}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "Acer Device Enabling Sevice")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                AcerDeviceEnablingSevice = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "AcerDeviceEnablingSevice,not check");
                                                if (index != -1)
                                                {
                                                  
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "AcerDeviceEnablingSevice,checked";
                                                }
                                                Console.WriteLine($"-------> AcerDeviceEnablingSevice {AcerDeviceEnablingSevice}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "Fingerprint")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                FingerPrint = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "FingerPrint,not check");
                                                if (index != -1)
                                                {
                                                    
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "FingerPrint,checked";
                                                }
                                                Console.WriteLine($"-------> FingerPrint {FingerPrint}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "Intel(R) Innovation Platform Framework Generic Participant" ||
                                        driverInfo_SCL[0] == "Intel(R) Innovation Platform Framework Manager" ||
                                        driverInfo_SCL[0] == "Intel(R) Innovation Platform Framework Processor Participant")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[1]) >= 0)
                                        {
                                            IntelDPTF_count += 1;
                                        }

                                        if (IntelDPTF_count == 3)
                                        {
                                            IntelDPTF = "checked";
                                            int index = checkflaglist.FindIndex(s => s == "IntelDPTF,not check");
                                            if (index != -1)
                                            {
                                                
                                                // Modify the value at the found index
                                                checkflaglist[index] = "IntelDPTF,checked";
                                            }
                                        }

                                    }
                                    else if (driverInfo_SCL[0] == "Intel(R) GNA")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                IntelGNA = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "IntelGNA,not check");
                                                if (index != -1)
                                                {
                                                    
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "IntelGNA,checked";
                                                }
                                                Console.WriteLine($"-------> IntelGNA {IntelGNA}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "Intel RST VMD")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                InteliRST = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "InteliRST,not check");
                                                if (index != -1)
                                                {
                                                    
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "InteliRST,checked";
                                                }
                                                Console.WriteLine($"-------> InteliRST {InteliRST}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "Smart Sound Technology BUS" ||
                                        driverInfo_SCL[0] == "Smart Sound Technology OED")
                                    {
                                        if (driverInfo_SCL[0] == "Smart Sound Technology OED")
                                        {
                                            if (line.IndexOf("OED", StringComparison.CurrentCultureIgnoreCase) >= 0)
                                            {
                                                if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                                {
                                                    Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                    IntelISST_count += 1;
                                                }
                                            }
                                        }
                                        if (driverInfo_SCL[0] == "Smart Sound Technology BUS")
                                        {
                                            if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                            {
                                                if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                                {
                                                    Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                    IntelISST_count += 1;
                                                }
                                            }
                                        }
                                        if (IntelISST_count == 2)
                                        {
                                            IntelISST = "checked";
                                            int index = checkflaglist.FindIndex(s => s == "IntelISST,not check");
                                            if (index != -1)
                                            {
                                                
                                                // Modify the value at the found index
                                                checkflaglist[index] = "IntelISST,checked";
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "Intel(R) Serial IO GPIO" ||
                                        driverInfo_SCL[0] == "Intel(R) Serial IO I2C")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                IntelSerialIO_count += 1;
                                            }

                                            if (IntelSerialIO_count == 2)
                                            {
                                                IntelSerialIO = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "IntelSerialIO,not check");
                                                if (index != -1)
                                                {
                                                    
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "IntelSerialIO,checked";
                                                }
                                                Console.WriteLine($"-------> IntelSerialIO checked !!");
                                                
                                            }
                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "Graphics")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                UMA = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "UMA,not check");
                                                if (index != -1)
                                                {
                                                    
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "UMA,checked";
                                                }
                                                Console.WriteLine($"-------> UMA {UMA}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "Intel(R) SMBus")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                IntelSMBus = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "IntelSMBus,not check");
                                                if (index != -1)
                                                {
                                                    
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "IntelSMBus,checked";
                                                }
                                                Console.WriteLine($"-------> IntelSMBus {IntelSMBus}");
                                                
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "Intel(R) Management Engine Interface")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                IntelManagementEngineInterface = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "IntelManagementEngineInterface,not check");
                                                if (index != -1)
                                                {
                                                    Console.WriteLine($"####{driverInfo_SCL[0]}####");
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "IntelManagementEngineInterface,checked";
                                                }
                                                Console.WriteLine($"-------> IntelManagementEngineInterface {IntelManagementEngineInterface}");
                                                Console.WriteLine("------------------------------------------------------");
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "Killer(R) Wi-Fi 6E AX1675i 160MHz Wireless Network Adapter")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                KillerWiFi6EAX1675i = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "KillerWiFi6EAX1675i,not check");
                                                if (index != -1)
                                                {
                                                   
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "KillerWiFi6EAX1675i,checked";
                                                }
                                                Console.WriteLine($"-------> KillerWiFi6EAX1675i {KillerWiFi6EAX1675i}");
                                                Console.WriteLine("------------------------------------------------------");
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "KillerControlCenter")
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                KillerControlCenter = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "KillerControlCenter,not check");
                                                if (index != -1)
                                                {
                                                  
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "KillerControlCenter,checked";
                                                }
                                                Console.WriteLine($"-------> KillerControlCenter {KillerControlCenter}");
                                                Console.WriteLine("------------------------------------------------------");
                                            }

                                        }
                                    }
                                    else if (driverInfo_SCL[0] == "NVIDIA GeForce RTX 3050" ||
                                        driverInfo_SCL[0] == "NVIDIA GeForce RTX 4050" && Had_external_VGA == "not check")
                                    {

                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        //found Touchpad device of ELAN, skip check other Touchpad device in list
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                Had_external_VGA = "checked";
                                                int index = checkflaglist.FindIndex(s => s == "Had_external_VGA,not check");
                                                if (index != -1)
                                                {
                                                    //Console.WriteLine($"####{driverInfo_SCL[0]}####");
                                                    // Modify the value at the found index
                                                    checkflaglist[index] = "Had_external_VGA,checked";
                                                }
                                                Console.WriteLine($"------------------> Had_external_VGA {Had_external_VGA}");
                                                Console.WriteLine("------------------------------------------------------");
                                            }

                                        }
                                    }
                                    else
                                    {
                                        if (line.IndexOf(driverInfo_SCL[0], StringComparison.CurrentCultureIgnoreCase) >= 0)
                                        {
                                            if (line.IndexOf(driverInfo_SCL[1]) >= 0)//Check Driver Version
                                            {
                                                Console.WriteLine("Case not include !!!!!!!!!!!!!!!!");
                                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in line {lineNumber}: {line}");
                                                Console.WriteLine("------------------------------------------------------");
                                            }

                                        }
                                    }
                                    lineNumber++;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"An error occurred: {ex.Message}");
                        }
                        //Check WinRe driver
                        if (driverInfo_SCL[0] == "Intel Rapid Storage WinPERE Drivers")
                        {
                            if (infContents.IndexOf(driverInfo_SCL[1]) >= 0)
                            {
                                IntelRapidStorageWinPeReDriver = "checked";
                                int index = checkflaglist.FindIndex(s => s == "IntelRapidStorageWinPeReDriver,not check");
                                if (index != -1)
                                {
                              
                                    // Modify the value at the found index
                                    checkflaglist[index] = "IntelRapidStorageWinPeReDriver,checked";
                                }
                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in infContents list, DriverVer is {driverInfo_SCL[1]}");
                                Console.WriteLine("------------------------------------------------------");
                            }

                        }
                        else if (driverInfo_SCL[0] == "Intel Serial I/O WinPERE Drivers")
                        {
                            if (infContents.IndexOf(driverInfo_SCL[1]) >= 0)
                            {
                                IntelSerialIOWinPEREDrivers = "checked";
                                int index = checkflaglist.FindIndex(s => s == "IntelSerialIOWinPEREDrivers,not check");
                                if (index != -1)
                                {
                                   
                                    // Modify the value at the found index
                                    checkflaglist[index] = "IntelSerialIOWinPEREDrivers,checked";
                                }
                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in infContents list, DriverVer is {driverInfo_SCL[1]}");
                                Console.WriteLine("------------------------------------------------------");
                            }

                        }
                        else if (driverInfo_SCL[0] == "Wireless LAN_M WinRE Drivers")
                        {
                            if (infContents.IndexOf(driverInfo_SCL[1]) >= 0)
                            {
                                WirelessLAN_MWinREDrivers = "checked";
                                int index = checkflaglist.FindIndex(s => s == "WirelessLAN_MWinREDrivers,not check");
                                if (index != -1)
                                {
                                   
                                    // Modify the value at the found index
                                    checkflaglist[index] = "WirelessLAN_MWinREDrivers,checked";
                                }
                                Console.WriteLine($"Found '{driverInfo_SCL[0]}' in infContents list, DriverVer is {driverInfo_SCL[1]}");
                                Console.WriteLine("------------------------------------------------------");
                            }

                        }
                        else 
                        {
                            //else
                            //{
                             //   Console.WriteLine("Had searched all Driver list, Finished!!! ---------------------");
                            //}
                        }
                    }
                }
            }

            int fail = 0;
            int success = 0;
            Console.WriteLine();
            Console.WriteLine("DUMP Result:(Only show fail items) ");
            foreach (string item in checkflaglist)
            { 
                string[] checkcell = item.Split(',');
                if (checkcell[1] != "checked")
                {
                    Console.WriteLine(item);
                    fail++;
                    
                }
                else 
                {
                    success++;
                    
                }

            }

            if (success ==28)
            {
                Console.WriteLine("All driver device found in PC, Success");
                return true;
            }
            if (fail != 0)
            {
                Console.WriteLine("Not all driver device found in PC, Fail");
                return false;
            }
            return false;
        }

        public static void UpdateResults() 
        {
            Console.WriteLine("UpdateResults");
        }
        public static void TearDown() 
        {
            Console.WriteLine("TearDown");
        }

    }
}
