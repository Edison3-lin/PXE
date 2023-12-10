using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Common;

namespace MyParameter
{
    public class Class1
    {

        private const string DllName = "MyParameter";

        public int Setup(int a, string b)
        {
            Testflow.General.WriteLog(DllName, $"{a.ToString()} xxx {b}");
            return 0;
        }

        public int Run(int a, int b, string c)
        {
            Testflow.General.WriteLog(DllName, $"{a.ToString()} man {b.ToString()} woman {c}");
            return 0;
        }

        public int UpdateResults(string a, int b)
        {
            Testflow.General.WriteLog(DllName, $"{a} Chinese {b.ToString()} hands");
            return 0;
        }

        public int TearDown(char a, string b)
        {
            Testflow.General.WriteLog(DllName, $"{a} char {b} string");
            return 0;
        }

    }
}
