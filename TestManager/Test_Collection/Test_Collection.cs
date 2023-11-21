using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Common;

namespace Test_Collection
{
    public class Test_Collection
    {
        private static string currentDirectory = Directory.GetCurrentDirectory() + '\\';
        private const string DllName = "Test_Collection";
        static string LibsPath = currentDirectory+"ItemDownload\\";

        public int Setup()
        {
            // common.Setup
            Testflow.Setup(DllName);

            return 81;
        }

        public int Run()
        {
            // Testflow.Run(DllName);

            try
            {
                Common.Runnner.RunTestItem(LibsPath+"TestItem1.dll");
                Common.Runnner.RunTestItem(LibsPath+"Test3.dll");
                Common.Runnner.RunTestItem(LibsPath+"TestItem2.dll");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Test all " + ex.Message);
            }
            return 82;
        }

        public int UpdateResults()
        {
            Testflow.UpdateResults(DllName, true);

            // Return the test results from 'Run'
            return 83;
        }

        public int TearDown()
        {
            Testflow.TearDown(DllName);
            return 84;
        }
    }
}
