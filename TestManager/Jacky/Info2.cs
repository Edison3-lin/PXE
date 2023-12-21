using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Management;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;

namespace Jacky
{
    public class SystemInfo
    {
        // 1.操作系统版本
		public static string GetOSVersion()
        {
            string os = string.Empty;
            try
            {
                if (Environment.OSVersion.Platform == PlatformID.Win32NT && Environment.OSVersion.Version.Major == 6 && Environment.OSVersion.Version.Minor == 1)
                {
                    os = "Windows 7";
                }
                else if (Environment.OSVersion.Version.CompareTo(new Version("10.0")) >= 0)
                {
                    os = "Windows 10";
                }
                else if (Environment.OSVersion.Platform == PlatformID.Win32NT && Environment.OSVersion.Version.Major == 5 && Environment.OSVersion.Version.Minor == 1)
                {
                    os = "Windows XP";
                }
                else if (Environment.OSVersion.Version.CompareTo(new Version("6.2")) >= 0)
                {
                    if (Environment.OSVersion.Version.CompareTo(new Version("6.3")) >= 0)
                    {
                        os = "Windows 8.1";
                    }
                    else
                    {
                        os = "Windows 8";
                    }
                }
                else if (Environment.OSVersion.Platform == PlatformID.Win32NT && Environment.OSVersion.Version.Major == 6 && Environment.OSVersion.Version.Minor == 0)
                {
                    os = "Windows Vista";
                }
                else if (Environment.OSVersion.Platform == PlatformID.Win32NT && Environment.OSVersion.Version.Major == 5 && Environment.OSVersion.Version.Minor == 2)
                {
                    os = "Windows 2003";
                }
                else if (Environment.OSVersion.Platform == PlatformID.Win32NT && Environment.OSVersion.Version.Major == 5 && Environment.OSVersion.Version.Minor == 0)
                {
                    os = "Windows 2000";
                }
                else if (Environment.OSVersion.Platform == PlatformID.Win32Windows && Environment.OSVersion.Version.Minor == 10 && Environment.OSVersion.Version.Revision.ToString() == "2222A")
                {
                    os = "Windows 98 第二版";
                }
                else if (Environment.OSVersion.Platform == PlatformID.Win32Windows && Environment.OSVersion.Version.Minor == 10 && Environment.OSVersion.Version.Revision.ToString() != "2222A")
                {
                    os = "Windows 98";
                }
                else if (Environment.OSVersion.Platform == PlatformID.Unix)
                {
                    os = "Unix";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(new StringBuilder("获取操作系统版本出错：").Append(ex.Message).ToString());
            }

            return os;
        }

        // 2.系统类型（32位 or 64位）
		public static string GetSystemType()
        {
            if (Environment.Is64BitOperatingSystem)
            {
                return "64 位操作系统";
            }
            else
            {
                return "32 位操作系统";
            }
        }

		public static int GetCPUCount()
        {
            return Environment.ProcessorCount;
        }

		public static string GetProcessorName()
        {
            //Win32_PhysicalMemory;Win32_Keyboard;Win32_ComputerSystem;Win32_OperatingSystem
            try
            {
                ManagementClass mos = new ManagementClass("Win32_Processor");
                foreach (ManagementObject mo in mos.GetInstances())
                {
                    if (mo["Name"] != null)
                    {
                        return mo["Name"].ToString();
                    }

                    //PropertyDataCollection pdc = mo.Properties;
                    //foreach (PropertyData pd in pdc)
                    //{
                    //    if ("Name" == pd.Name)
                    //    {
                    //        return pd.Value.ToString();
                    //    }
                    //}
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(new StringBuilder("获取处理器名称失败：").Append(ex.Message).ToString());
            }
            
            return string.Empty;
        }

		public static string GetRunningTime()
        {
            string result = string.Empty;
            try
            {
                int uptime = Environment.TickCount & Int32.MaxValue;
                TimeSpan ts = new TimeSpan(Convert.ToInt64(Convert.ToInt64(uptime) * 10000));
                result = new StringBuilder(ts.Days.ToString()).Append("天 ").Append(ts.Hours).Append(":").
                    Append(ts.Minutes).Append(":").Append(ts.Seconds).ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine(new StringBuilder("获取开机时间失败：").Append(ex.Message).ToString());
            }

            return result;
        }

		// public static float GetUsageOfCPU()
        // {
        // 	this.pcCpuLoad = new PerformanceCounter("Processor", "% Processor Time", "_Total");
        //     this.pcCpuLoad.MachineName = ".";
        //     this.pcCpuLoad.NextValue();
 
        //     return this.pcCpuLoad.NextValue();
        // }

		public static long GetPhysicalMemory()
        {
            ManagementClass mc = new ManagementClass("Win32_ComputerSystem");
            ManagementObjectCollection moc = mc.GetInstances();
            foreach (ManagementObject mo in moc)
            {
                if (mo["TotalPhysicalMemory"] != null)
                {
                    return long.Parse(mo["TotalPhysicalMemory"].ToString());
                }
            }
            
			return 0;
        }

        // 8.可用内存
		public static long GetAvailableMemory()
        {
            long availablebytes = 0;
            try
            {
                ManagementClass mos = new ManagementClass("Win32_OperatingSystem");
                foreach (ManagementObject mo in mos.GetInstances())
                {
                    if (mo["FreePhysicalMemory"] != null)
                    {
                        availablebytes = 1024 * long.Parse(mo["FreePhysicalMemory"].ToString());
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(new StringBuilder("获取可用内存失败：").Append(ex.Message).ToString());
            }

            return availablebytes;
        }

        // 9.进程数量、线程数量、句柄数量
		// public static void GetProcessCount(out int processCount, out int threadCount, out int handleCount) 
        // {
        //     Process[] processes = Process.GetProcesses();
        //     processCount = processes.Count();
        //     foreach (Process pro in processes)
        //     {
        //         threadCount += pro.Threads.Count;
        //         handleCount += pro.HandleCount;
        //     }
        // }



    }
}
