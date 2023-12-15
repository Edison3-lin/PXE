using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace MyT {
    internal class Program {
        public static void CallToChildThread() {
            try{
                Console.WriteLine("执行子线程");
                // 计数到 10
                for (int counter = 0; counter <= 10; counter++)
                {
                    Thread.Sleep(500);
                    Console.WriteLine(counter);
                }
                Console.WriteLine("子线程执行完成");
            }catch (ThreadAbortException e){
                Console.WriteLine("线程终止：{0}", e);
            }finally{
                Console.WriteLine("无法捕获线程异常");
            }
        }
        static void Main(string[] args)
        {
            ThreadStart childref = new ThreadStart(CallToChildThread);
            Console.WriteLine("在 Main 函数中创建子线程");
            Thread childThread = new Thread(childref);
            childThread.Start();
            // 停止主线程一段时间
            Thread.Sleep(2000);
            // 现在中止子线程
            Console.WriteLine("在 Main 函数中终止子线程");
            childThread.Abort();
            Console.ReadKey();
        }
    }
}
