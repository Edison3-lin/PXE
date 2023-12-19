using System;
using System.Threading;

class Program
{
    static void Main()
    {
        // 建立一個新的 AppDomain
        AppDomain yourAppDomain = AppDomain.CreateDomain("YourAppDomain");

        // 開始一個新的執行緒，在這個執行緒中執行卸載操作
        Thread unloadThread = new Thread(() =>
        {
            // 在新的執行緒中卸載 yourAppDomain
            AppDomain.Unload(yourAppDomain);
        });

        // 開始新的執行緒
        unloadThread.Start();

        // 在此等待新的執行緒完成
        unloadThread.Join();

        Console.WriteLine("AppDomain Unload completed.");
    }
}
