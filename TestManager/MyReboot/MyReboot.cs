using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LoadDll;
using System.Diagnostics;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MyReboot
{
    public class MyReboot
    {
        private const string DllName = "MyReboot";
        private const string TR = "C:\\TestManager\\TR_Result.json";

        public int Setup()
        {
            // LoadDll.Setup
            return 11;
        }

        public int Run()
        {

           // Read TR_Result.json Reboot
           string jsonString = System.IO.File.ReadAllText(TR);
           JObject json = JObject.Parse(jsonString);
           string MyReboot = (string) json["Reboot"];
           Console.WriteLine(MyReboot);

           if( MyReboot == "Reboot") {
                json["Reboot"] = "After";
                string ModJson = json.ToString();
                System.IO.File.WriteAllText(TR, ModJson);
                Console.WriteLine("Had reboot...");
                return 0;
           }
           else {
                json["Reboot"] = "Before";
                string ModJson = json.ToString();
                System.IO.File.WriteAllText(TR, ModJson);
           }


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
