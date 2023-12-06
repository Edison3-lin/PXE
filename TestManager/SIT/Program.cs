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
            // 检查是否有命令行参数
            if (args.Length == 0)
            {
                Console.WriteLine("没有提供命令行参数。");
            }
            else
            {
                Console.WriteLine("命令行参数:");

                // 使用循环显示所有命令行参数
                for (int i = 0; i < args.Length; i++)
                {
                    Console.WriteLine($"参数 {i + 1}: {args[i]}");
                }
            }        
        }
    }
}
