using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace CaptainWin.CommonAPI{
    public class CommonDevicesStatusCheck{
        // Constants for SetupAPI functions
        const int DIGCF_PRESENT = 0x00000002;
        const int SPDRP_DEVICEDESC = 0x00000000;
        //const uint SPDRP_STATUS = 0x00000017;
        // Define constants and structures
        private const int DIGCF_ALLCLASSES = 0x000000004;
        // Hardware ID
        const int SPDRP_HARDWAREID = 0x00000001; 
        public const int SPDRP_DRIVER = 0x00000009;

        [StructLayout(LayoutKind.Sequential)]
        public struct SP_DEVINFO_DATA{
            public int cbSize;
            public Guid ClassGuid;
            public int DevInst;
            public IntPtr Reserved;
        }

        [DllImport("setupapi.dll", SetLastError = true)]
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
        public static extern bool SetupDiDestroyDeviceInfoList(IntPtr DeviceInfoSet);

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

        [DllImport("setupapi.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SetupDiGetClassDevs(
                ref Guid ClassGuid,
                string Enumerator,
                IntPtr hwndParent,
                int Flags
        );
        /// <summary>
        /// Get the name of devices in Device Manager
        /// </summary>
        /// <param name="deviceInfoSet"></param>
        /// <param name="deviceInfoData"></param>
        /// <returns>
        /// resturn a string of name of device
        /// </returns>
        public static string GetDeviceName(IntPtr deviceInfoSet, SP_DEVINFO_DATA deviceInfoData){
            int requiredSize = 0;
            SetupDiGetDeviceRegistryProperty(deviceInfoSet,
                ref deviceInfoData,
                SPDRP_DEVICEDESC,
                out int regDataType,
                IntPtr.Zero,
                0,
                out requiredSize);
            if (requiredSize == 0)
                return string.Empty;

            IntPtr propertyBuffer = Marshal.AllocHGlobal((IntPtr)requiredSize);
            if (SetupDiGetDeviceRegistryProperty(deviceInfoSet,
                ref deviceInfoData,
                SPDRP_DEVICEDESC,
                out regDataType,
                propertyBuffer,
                requiredSize,
                out requiredSize)){
                string deviceName = Marshal.PtrToStringAuto(propertyBuffer);
                Marshal.FreeHGlobal(propertyBuffer);
                return deviceName;
            }
            return string.Empty;
        }
        /// <summary>
        /// Get driver version of devices in Device Manager
        /// </summary>
        /// <param name="deviceInfoSet"></param>
        /// <param name="deviceInfoData"></param>
        /// <returns>
        /// return a string of version of device driver in device manager
        /// </returns>
        public static string GetDriverVersion(IntPtr deviceInfoSet, SP_DEVINFO_DATA deviceInfoData){
            int requiredSize = 0;
            SetupDiGetDeviceRegistryProperty(deviceInfoSet,
                ref deviceInfoData,
                SPDRP_DRIVER,
                out int regDataType,
                IntPtr.Zero,
                0,
                out requiredSize);
            if (requiredSize == 0)
                return string.Empty;

            IntPtr propertyBuffer = Marshal.AllocHGlobal(requiredSize);
            if (SetupDiGetDeviceRegistryProperty(deviceInfoSet,
                ref deviceInfoData,
                SPDRP_DRIVER,
                out regDataType,
                propertyBuffer,
                requiredSize,
                out requiredSize)){
                string driverVersion = Marshal.PtrToStringAuto(propertyBuffer);
                Marshal.FreeHGlobal(propertyBuffer);

                // Specify the registry key path and value name you want to read.
                string keyPath = @"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Class\" + driverVersion;
                string valueName = "DriverVersion";

                // Use the Registry.GetValue method to read the registry value.
                object value = Registry.GetValue(keyPath, valueName, null);

                if (value != null){
                    //Console.WriteLine($"Value of {valueName} in {keyPath}: {value}");
                }else{
                    Console.WriteLine($"Registry value {valueName} not found in {keyPath}");
                }

                return value.ToString();
            }

            return string.Empty;
        }
        /// <summary>
        /// Get hardware id for each devices in Device Manager and return the id by string
        /// </summary>
        /// <param name="deviceInfoSet"></param>
        /// <param name="deviceInfoData"></param>
        /// <returns>
        /// return a string of HWID info
        /// </returns>
        public static string GetHardwareID(IntPtr deviceInfoSet, SP_DEVINFO_DATA deviceInfoData){
            int requiredSize = 0;
            SetupDiGetDeviceRegistryProperty(deviceInfoSet,
                ref deviceInfoData,
                SPDRP_HARDWAREID,
                out int regDataType,
                IntPtr.Zero,
                0,
                out requiredSize);
            if (requiredSize == 0)
                return string.Empty;

            IntPtr propertyBuffer = Marshal.AllocHGlobal((IntPtr)requiredSize);
            if (SetupDiGetDeviceRegistryProperty(deviceInfoSet,
                ref deviceInfoData,
                SPDRP_HARDWAREID,
                out regDataType,
                propertyBuffer,
                requiredSize,
                out requiredSize)){
                string hardwareid = Marshal.PtrToStringAuto(propertyBuffer);
                Marshal.FreeHGlobal(propertyBuffer);
                return hardwareid;
            }
            return string.Empty;
        }
        /// <summary>
        /// Get Device status code and problem code 
        /// </summary>
        /// <param name="deviceInfoSet"></param>
        /// <param name="deviceInfoData"></param>
        /// <returns>
        /// return a string of devices status which contains HWID and problem code
        /// </returns>
        public static string GetDevicesStatusAndProblemCode(IntPtr deviceInfoSet, SP_DEVINFO_DATA deviceInfoData){
            // Find the device node for the specified hardware ID.
            uint devInst = 0;

            uint status = 0;
            uint problemNumber = 0;

            int result = CM_Locate_DevNode(out devInst, null, 0);

            if (result == 0){
                // Device node found, now get its status.
                result = CM_Get_DevNode_Status(out status, out problemNumber, (int)deviceInfoData.DevInst, 0);

                if (result == 0){
                    //Console.WriteLine("Device Status: " + status);
                    //Console.WriteLine("Device Problem Code: " + problemNumber);
                }else{
                    Console.WriteLine("Failed to get device status.");
                    return string.Empty;
                }
                return problemNumber.ToString();
            }else{
                Console.WriteLine("Device not found or error locating the device node.");
                return string.Empty;
            }
        }
        /// <summary>
        /// Check all devices status in Device Manager, pass will retrun true, fail will return false
        /// </summary>
        /// <returns></returns>
        public static bool CheckDeviceStatus(){
            List<string> CheckStatusList = new List<string>();
            // Replace "searchString" with the string you want to find
            string searchString = "DeviceStatusError";
            // string currentDirectory = Directory.GetCurrentDirectory() + '\\';
            string currentDirectory = @"C:\TestManager\ItemDownload\";
            string filePath = currentDirectory + "DeviceStatusCheck.txt";
            Console.WriteLine(filePath);
            //Query all devices in DM
            string result = null;
            Guid guid = Guid.Empty; // List all devices
            IntPtr deviceInfoSet = SetupDiGetClassDevs(ref guid, null, IntPtr.Zero, DIGCF_PRESENT | DIGCF_ALLCLASSES);

            if (deviceInfoSet != IntPtr.Zero){
                SP_DEVINFO_DATA deviceInfoData = new SP_DEVINFO_DATA();
                deviceInfoData.cbSize = Marshal.SizeOf(typeof(SP_DEVINFO_DATA));
                int index = 0;

                while (SetupDiEnumDeviceInfo(deviceInfoSet, index, ref deviceInfoData)){
                    string deviceName = GetDeviceName(deviceInfoSet, deviceInfoData);
                    string driverVersion = GetDriverVersion(deviceInfoSet, deviceInfoData);
                    string hardwareid = GetHardwareID(deviceInfoSet, deviceInfoData);
                    string devicestatus = GetDevicesStatusAndProblemCode(deviceInfoSet, deviceInfoData);

                    if (deviceName != null && driverVersion != null){
                        if (devicestatus != "0"){
                            devicestatus = devicestatus + " " + "DeviceStatusError";
                            //Console.WriteLine(devicestatus);
                        }
                        result = deviceName + "," + driverVersion + "," + hardwareid + "," + devicestatus;
                        CheckStatusList.Add(result);
                    }
                    //Console.WriteLine($"result==> {CheckStatusList}");
                    index++;
                }
                // Clean up
                Marshal.FreeHGlobal(deviceInfoSet);
                string writefilePath = currentDirectory + "DeviceStatusCheck.txt";
                try{
                    // Write data from the list to the file, overwriting its current content
                    File.WriteAllLines(writefilePath, CheckStatusList);

                    //Console.WriteLine("Data written to the file successfully.");
                }catch (Exception ex){
                    Console.WriteLine($"Error: {ex.Message}");
                }

            }


            try{
                // Read all lines from the file
                string[] lines = File.ReadAllLines(filePath);

                // Search for the string in each line
                for (int lineNumber = 0; lineNumber < lines.Length; lineNumber++)
                {
                    if (lines[lineNumber].Contains(searchString))
                    {
                        Console.WriteLine($"Found '{searchString}' in line {lineNumber + 1}: {lines[lineNumber]}");
                        Console.WriteLine("Device status check FAIL");
                        return false;
                    }
                }
            }catch (Exception ex){
                Console.WriteLine($"Error: {ex.Message}");
            }
            Console.WriteLine("Device status check PASS");
            return true;
        }
    }
}
