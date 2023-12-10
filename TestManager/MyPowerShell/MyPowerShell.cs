using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Common;
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
            Testflow.General.WriteLog("MyPowershell", "Setup MyPowershell.dll...");

            return 21;
        }

        public int Run()
        {

            Testflow.General.WriteLog("MyPowershell", "Start MyPowershell.dll...");

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

            Testflow.General.WriteLog("MyPowershell", "End MyPowershell.dll...");

            return 22;
        }

        public int UpdateResults()
        {
            Testflow.General.WriteLog("MyPowershell", "UpdateResults MyPowershell.dll...");
            return 23;
        }

        public int TearDown()
        {
            Testflow.General.WriteLog("MyPowershell", "TearDown MyPowershell.dll...");
            return 24;
        }

    }
}
