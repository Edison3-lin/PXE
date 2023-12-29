using Microsoft.Win32;
using System;
using System.Management;

namespace MyT {
    internal class Program {
        static void Main(string[] args)
        {


            // 建立 RegistryKey 物件
            RegistryKey key = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run");

            // 列出所有子鍵
            foreach (string subKey in key.GetSubKeyNames())
            {
                // 列出子鍵的名稱
                Console.WriteLine(subKey);

                // 列出子鍵的值
                foreach (RegistryValueEntry valueEntry in subKey.GetValues())
                {
                    // 列出值名稱和值
                    Console.WriteLine("{0} = {1}", valueEntry.Name, valueEntry.Value);
                }
            }

            // 關閉 RegistryKey 物件
            key.Close();

        }


    }
}
