using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace C1
{
    public class Class1
    {

        private const string DllName = "C1";

        public int Setup(int a, string b)
        {
            // common.Setup
            Console.WriteLine(a.ToString() +"歲的女人 "+ b);
            Console.ReadKey();
            return 11;
        }

        public int Run(int a, int b, string c)
        {
            Console.WriteLine((a+b).ToString() +"歲的男人 "+ c);
            Console.ReadKey();
            return 12;
        }

        public int UpdateResults(string a, int b)
        {
            Console.WriteLine(a + " 是中文名字 "+ b.ToString() );
            Console.ReadKey();
            return 13;
        }

        public int TearDown(char a, string b)
        {
            Console.WriteLine(a + " 字母 "+ b );
            Console.ReadKey();
            return 14;
        }

    }
}
