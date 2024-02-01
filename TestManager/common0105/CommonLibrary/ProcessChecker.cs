/*
* ProcessChecker.cs encapsulate Acer BIOS Setting Tool functions.
* IsRunning - Return true if number of processes with processName is greater than 0.
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Chris Huang   <Chris_Huang@quantatw.com>
*/
using System.Diagnostics;
using System.Threading.Tasks;
using System.Threading;


namespace CaptainWin.CommonAPI {
    /// <summary>
    /// Thie class is used to check if a process is running by its name
    /// </summary>
    public class ProcessChecker {
        /// <summary>
        /// Check if a process of which name is processName is running
        /// </summary>
        /// <param name="processName">The name of the process</param>
        /// <param name="sleepTime">The time lag before get process array</param>
        /// <returns>true, if the number of processses >= 1</returns>
        public static bool IsRunning(string processName, int sleepTime = 5) {

            bool result = Task.Factory.StartNew(() => {
                Thread.Sleep(sleepTime * 1000);
                Process[] processes = Process.GetProcessesByName(processName);
                if ( processes.Length > 0 )
                    return true;
                else
                    return false;
            }).Result;

            return result;
        }
    }
}
