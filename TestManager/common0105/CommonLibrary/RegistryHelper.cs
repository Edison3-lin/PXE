using System;
using Microsoft.Win32;
using System.Management;
using System.IO;


namespace CaptainWin.CommonAPI
{
    /// <summary>
    ///  Provide read/write/create function to a registry path and item.
    /// </summary>
    public class RegistryHelper
    {
        /// <summary>
        /// Read a registry path and item's value (not support for multi string type)
        /// </summary>
        /// <param name="hive">The group of the registry, selection for RegistryHive :
        /// RegistryHive.ClassesRoot : HKEY_CLASSES_ROOT
        /// RegistryHive.CurrentUser : HKEY_CURRENT_USER
        /// RegistryHive.LocalMachine : HKEY_LOCAL_MACHINE
        /// RegistryHive.Users : HKEY_USERS
        /// RegistryHive.CurrentConfig : HKEY_CURRENT_CONFIG
        /// </param>
        /// <param name="keyPath">Path to the folder</param>
        /// <param name="itemName">items need to read back under the path</param>
        /// <returns>Find or not, and the read value or string</returns>
        public static (bool isFind, string getValue) ReadRegistryValue(RegistryHive hive, string keyPath, string itemName)
        {
            try
            {
                using (var baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Registry64))
                using (var key = baseKey.OpenSubKey(keyPath))
                {
                    if (key != null)
                    {
                        var value = key.GetValue(itemName);
                        return (true, value?.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading registry value: {ex.Message}");
            }

            return (false, "");
        }
        /// <summary>
        /// Write value to the registry path item.
        /// </summary>
        /// <param name="hive">The group of the registry, seletion for RegistryHive</param>
        /// <param name="keyPath">Path to the folder</param>
        /// <param name="itemName">items under the keyPath</param>
        /// <param name="dataToWrite">data write to itemName</param>
        /// <returns>success or not(true or false)</returns>
        public static bool WriteRegistryValue(RegistryHive hive, string keyPath, string itemName, object dataToWrite)
        {
            try
            {
                using (var baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default))
                using (var key = baseKey.CreateSubKey(keyPath))
                {
                    if (key != null)
                    {
                        key.SetValue(itemName, dataToWrite);
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing registry value: {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// create a registry path folder.
        /// </summary>
        /// <param name="hive">The group of the registry, seletion for RegistryHive</param>
        /// <param name="keyPath">Registry Path to the folder</param> 
        /// <returns>success or not(true or false)</returns>
        public static bool CreateRegistryKey(RegistryHive hive, string keyPath)
        {
            try
            {
                using (var baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default))
                {
                    baseKey.CreateSubKey(keyPath);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating registry key: {ex.Message}");
            }

            return false;
        }
        /// <summary>
        /// create a Dword item with given value to the given registry path.
        /// </summary>
        /// <param name="hive">The group of the registry, seletion for RegistryHive</param>
        /// <param name="keyPath">Registry Path to the folder</param> 
        /// <param name="itemName">Dword item name</param> 
        /// <param name="writeValue">Dword item assigned value</param> 
        /// <returns>success or not(true or false)</returns>
        public static bool CreateDWordValue(RegistryHive hive, string keyPath, string itemName, int writeValue)
        {
            try
            {
                using (var baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default))
                using (var key = baseKey.OpenSubKey(keyPath, true) ?? baseKey.CreateSubKey(keyPath))
                {
                    key.SetValue(itemName, writeValue, RegistryValueKind.DWord);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating DWORD registry value: {ex.Message}");
            }

            return false;
        }
        /// <summary>
        /// create a string item with given value to the given registry path.
        /// </summary>
        /// <param name="hive">The group of the registry, seletion for RegistryHive</param>
        /// <param name="keyPath">Registry Path to the folder</param> 
        /// <param name="itemName">string item name</param> 
        /// <param name="writeValue">string item assigned value</param> 
        /// <returns>success or not(true or false)</returns>
        public static bool CreateStringValue(RegistryHive hive, string keyPath, string itemName, string writeValue)
        {
            try
            {
                using (var baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default))
                using (var key = baseKey.OpenSubKey(keyPath, true) ?? baseKey.CreateSubKey(keyPath))
                {
                    key.SetValue(itemName, writeValue, RegistryValueKind.String);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating String registry value: {ex.Message}");
            }

            return false;
        }
        /// <summary>
        /// create a string array item with given value to the given registry path.
        /// </summary>
        /// <param name="hive">The group of the registry, seletion for RegistryHive</param>
        /// <param name="keyPath">Registry Path to the folder</param> 
        /// <param name="itemName">string array item name</param> 
        /// <param name="writeValue">string array item assigned value</param> 
        /// <returns>success or not(true or false)</returns>
        public static bool CreateMultiStringValue(RegistryHive hive, string keyPath, string itemName, string[] writeValue)
        {
            try
            {
                using (var baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default))
                using (var key = baseKey.OpenSubKey(keyPath, true) ?? baseKey.CreateSubKey(keyPath))
                {
                    key.SetValue(itemName, writeValue, RegistryValueKind.MultiString);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating MultiString registry value: {ex.Message}");
            }

            return false;
        }
    }
}
