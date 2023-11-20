using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace TestCase
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string[] baseinfo = {"",""};
            Setup.Setup.setup(baseinfo);
            Console.WriteLine(baseinfo[0]);
            Console.WriteLine(baseinfo[1]);
            //image_installation_driver_default.image_installation_driver_default.Run();
            image_installation_application_default.image_installation_application_default.Run();
            //while (true) 
            //{
            //    Timer timer = new Timer(2000);
            //    timer.Elapsed += TimerElapsed;
            //    timer.Start();
                
            //}
            
            //public static void click(int X, int Y, int clicks, int interval, int button)
            //image_installation_application_default.image_installation_application_default.Run();
            
        }
        //private static void TimerElapsed(object sender, ElapsedEventArgs e)
        //{
        //    Console.WriteLine("After delay");
        //    mouseByCS.MouseSimulator.click(100, 300, 1, 200, 0);
        //    ((Timer)sender).Stop();
        //}
    }
}
