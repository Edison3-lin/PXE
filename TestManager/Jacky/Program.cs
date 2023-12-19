using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Globalization;

namespace Jacky
{
    internal class Program
    {

    /// <summary>
    /// Sleep the system
    /// </summary>
    /// <param sleepParm="type count" > Sleep Type </param>
    public static void Sleep (string[] args) {
      string executablePath = @"c:\TestManager\ItemDownload\pwrtest.exe";
      string arguments = string.Format("/sleep /c:{1} /s:{0} /d:30 /p:40", args[0], args[1]);

      ProcessStartInfo startInfo = new ProcessStartInfo(executablePath)
      {
          Arguments = arguments,
          WorkingDirectory = @"c:\TestManager\ItemDownload",
          Verb = "runas"
      };

      try
      {
          Process process = new Process
          {
              StartInfo = startInfo
          };
          process.Start();
          process.WaitForExit();
      }
      catch (Exception ex)
      {
          Console.WriteLine("Error: " + ex.Message);
      }
    }  //Sleep

    /// <summary>
    /// Sleep the system
    /// </summary>
    /// <param sleepParm="type count" > Sleep Type </param>
    public static void Culture() {
        CultureInfo installedUICulture = CultureInfo.InstalledUICulture;
        Console.WriteLine("Installed UI Culture:");
        Console.WriteLine("Name: " + installedUICulture.Name);
        Console.WriteLine("DisplayName: " + installedUICulture.DisplayName);
        Console.WriteLine("EnglishName: " + installedUICulture.EnglishName);
        Console.WriteLine("TwoLetterISOLanguageName: " + installedUICulture.TwoLetterISOLanguageName);
        Console.WriteLine("ThreeLetterISOLanguageName: " + installedUICulture.ThreeLetterISOLanguageName);
    }

    /// <summary>
    /// Sleep the system
    /// </summary>
    /// <param sleepParm="type count" > Sleep Type </param>
    public static void Sysinfo() {
        
    }


        static void Main(string[] args)
        {
            // string [] a = new string[]{args[0], args[1]};
            // Sleep(a);
            
            // Culture();


        }
    }
}
