using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace TMservice
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            // bat file
            string batFilePath = @"C:\TestManager\test.bat";

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = System.IO.Path.GetDirectoryName(batFilePath)
            };

            using (Process process = new Process { StartInfo = psi })
            {
                process.Start();
                process.StandardInput.WriteLine($"\"{batFilePath}\"");
                process.WaitForExit();
            }                
        }

        protected override void OnStop()
        {
            // bat file
            string batFilePath = @"C:\TestManager\testStop.bat";

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = System.IO.Path.GetDirectoryName(batFilePath)
            };

            using (Process process = new Process { StartInfo = psi })
            {
                process.Start();
                process.StandardInput.WriteLine($"\"{batFilePath}\"");
                process.WaitForExit();
            }                
        }

    }
}
