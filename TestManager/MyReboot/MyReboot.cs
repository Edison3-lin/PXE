using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Common;
using System.Diagnostics;

namespace MyReboot
{
    public class MyReboot
    {
        private const string DllName = "MyReboot";

        public int Setup()
        {
            // common.Setup
            return 11;
        }

        public int Run()
        {
            // common.Setup
            string exeFilePath = "shutdown";
            // Create a ProcessStartInfo object with the file path
            ProcessStartInfo startInfo = new ProcessStartInfo(exeFilePath);
            // Optionally, you can set working directory, arguments, and other properties
            startInfo.WorkingDirectory = ".\\";
            startInfo.Arguments = "/r /t 0";
            // Start the process
            Process process = new Process();
            process.StartInfo = startInfo;
            process.Start();

            // Optionally, you can wait for the process to exit
            process.WaitForExit();

            while (true){}

            return 12;
        }

        public int UpdateResults()
        {
            return 13;
        }

        public int TearDown()
        {
            return 14;
        }
    }
}
