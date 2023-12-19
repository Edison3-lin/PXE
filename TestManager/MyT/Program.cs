using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;
using System.Xml.Linq;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading;


namespace MyT {
    internal class Program {
        static void Main(string[] args)
        {
                string TR = @"c:\\TestManager\\TR_Result.json"; // 將路徑替換為你的JSON文件的實際路徑

                string jsonString = File.ReadAllText(TR);
                Console.WriteLine(jsonString);
                JObject json = JObject.Parse(jsonString);
                int timeout = (int)json["Test_TimeOut"];
                Console.WriteLine(timeout);

        }
    }
}
