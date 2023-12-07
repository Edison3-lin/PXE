using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyParameter
{
    public class Class1
    {

        private const string DllName = "MyParameter";

        public int Setup(int a, string b)
        {
            // common.Setup
            Console.WriteLine("Setup().. "+a.ToString() +"歲的女人 "+ b);
            return 11;
        }

        public int Run(int a, int b, string c)
        {
            Console.WriteLine("Run().. "+ (a+b).ToString() +"歲的男人 "+ c);
            return 12;
        }

        public int UpdateResults(string a, int b)
        {
            Console.WriteLine("UpdateResults() .." + a + " 測試中文 "+ b.ToString() );
            return 0;
        }

        public int TearDown(char a, string b)
        {
            Console.WriteLine("TearDwon().." + a + " 字母 "+ b );
            Console.WriteLine("any key to continue...");
            Console.ReadKey();
            return 0;
        }

    }
}
