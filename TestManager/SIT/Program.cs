using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;             //手動加入參考
using System.Management.Automation.Runspaces;   //手動加入參考
using System.Threading;
using System.Reflection;
using System.Diagnostics;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace SIT
{
    internal class Program
    {
        static void Main(string[] args)
        {

            string strNumber = args[0];
            int timeout = 0;
            Console.WriteLine("strNumber " +　strNumber);
            try
            {
                timeout = int.Parse(strNumber);
                Console.WriteLine("timeout:" +　timeout.ToString());

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                timeout = 999;
            }

        }
    }
}
