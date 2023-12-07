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
            // common.Setup
            Testflow.Setup(DllName);

            return 21;
        }

        public int Run()
        {
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

            // common.Setup
            Testflow.Run("MyPowerShell");
            return 22;
        }

        public int UpdateResults()
        {
            Testflow.UpdateResults(DllName, true);
            return 23;
        }

        public int TearDown()
        {
            Testflow.TearDown(DllName);
            return 24;
        }

    }
}
