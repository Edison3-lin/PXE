using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LoadDll;
using System.Management.Automation;             // *.csproj 手動加入 <Reference Include="System.Management.Automation" />
using System.Management.Automation.Runspaces;   // *.csproj 手動加入 <Reference Include="System.Management.Automation" />

namespace MyPowerShell
{
    public class MyPowerShell
    {
        private const string DllName = "MyPowerShell";
        private static string ItemDownload = "C:\\TestManager\\ItemDownload\\";

        public int Setup()
        {
            Runnner.WriteLog("Setup MyPowershell.dll...");

            return 21;
        }

        public int Run()
        {

            Runnner.WriteLog("Start MyPowershell.dll...");

            try
            {
             Runspace runspace = RunspaceFactory.CreateRunspace();
             runspace.Open();
             Pipeline pipeline = runspace.CreatePipeline();
             pipeline.Commands.AddScript(ItemDownload+"Abt1.ps1");
             pipeline.Invoke();
             runspace.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error!!! " + ex.Message);
            }

            Runnner.WriteLog("End MyPowershell.dll...");

            return 22;
        }

        public int UpdateResults()
        {
            Runnner.WriteLog("UpdateResults MyPowershell.dll...");
            return 23;
        }

        public int TearDown()
        {
            Runnner.WriteLog("TearDown MyPowershell.dll...");
            return 24;
        }

    }
}
