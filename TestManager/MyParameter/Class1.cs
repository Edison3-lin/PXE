using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LoadDll;

namespace MyParameter
{
    public class Class1
    {

        private const string DllName = "MyParameter";

        public int Setup(int a, string b)
        {
            Runnner.WriteLog($"{a.ToString()} xxx {b}");
            return 0;
        }

        public int Run(int a, int b, string c)
        {
            Runnner.WriteLog($"{a.ToString()} man {b.ToString()} woman {c}");
            return 0;
        }

        public int UpdateResults(string a, int b)
        {
            Runnner.WriteLog($"{a} Chinese {b.ToString()} hands");
            return 0;
        }

        public int TearDown(char a, string b)
        {
            Runnner.WriteLog($"{a} char {b} string");
            return 0;
        }

    }
}
