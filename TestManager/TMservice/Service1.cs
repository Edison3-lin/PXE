using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace TMservice
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            // 設定批處理文件的路徑
            string batFilePath = @"C:\TestManager\test.bat";

            // 創建進程啟動信息
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = System.IO.Path.GetDirectoryName(batFilePath)
            };

            // 創建進程對象
            using (Process process = new Process { StartInfo = psi })
            {
                // 啟動進程
                process.Start();

                // 向 cmd 輸入執行的命令（這裡是執行批處理文件）
                process.StandardInput.WriteLine($"\"{batFilePath}\"");

                // 等待命令執行完成
                process.WaitForExit();
            }                
        }

        protected override void OnStop()
        {
            // 停止服務時執行的程式碼
        }

        private void OnTimer(object sender, ElapsedEventArgs e)
        {
            // 定期執行的程式碼
            Console.WriteLine("xxxx");
            // 這裡可以執行你的任務或工作
        }        
    }
}
