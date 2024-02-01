/*
* Abst.cs encapsulate Acer BIOS Setting Tool functions.
* Set - Set value into BIOS
* Get - Get value from BIOS
* ImportCustomSettings - Import custom settings from XML file
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Chris Huang   <Chris_Huang@quantatw.com>
*/

using System.Diagnostics;
using System.IO;
using System.Text;

namespace CaptainWin.CommonAPI {
    /// <summary>
    /// Thie class encapsulate Acer BIOS Setting Tool, Abst64_unsign.exe, functions
    /// </summary>
    /// 
    public class Abst {
        private static string run( string filePath, string arguments = "", int timeOut = 60 ) {
            string output;

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = filePath;
            //startInfo.Verb = "runas";
            startInfo.Arguments = arguments;
            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
			
            Process p = Process.Start(startInfo);
			
            // read the output of the process asynchronously
            StringBuilder outputBuilder = new StringBuilder();
            p.OutputDataReceived += (s, e) => {
                if (!string.IsNullOrEmpty(e.Data)) {
                    outputBuilder.AppendLine(e.Data); // append each line of output to the StringBuilder
                }
            };
            p.BeginOutputReadLine();

            // wait for the process to complete, with timeout
            int timeoutMilliseconds = timeOut * 1000;
            bool processCompleted = p.WaitForExit(timeoutMilliseconds);

            if ( processCompleted ) {
                output = outputBuilder.ToString(); // get the entire output as a string
            }
            else {
                output = "Process Time Out!";
                p.Kill(); // kill the process if it has not completed within the timeout
            }
            return output;
        }
        /// <summary>
        /// Set a value into BIOS and check if vlaue is set successfully
        /// </summary>
        /// <param name="abstPath">The absolute path of abst64_unsigned.exe</param>
        /// <param name="password">Superviser password</param>
        /// <param name="item">The BIOS item to be set</param>
        /// <param name="value">The value to be set</param>
        /// <returns>success if value is set correctly in BIOS, fail if not</returns>
        public static string Set( string abstPath, string password, string item, string value ) {
            string result = "fail";
            string output;
            string aLine;

            string arguments = " /password " + password + " /set \"" + item + "=" + value + "\"";
            output = Abst.run(abstPath, arguments);

            bool setError = output.Contains("Error:");
            bool setBiosOptionSuccess = output.Contains("Comparison result : The same");
            if (setError) {

                StringReader strReader = new StringReader(output);
                aLine = strReader.ReadLine();

                while ( aLine != null ) {
                    if ( aLine.Contains("Error:") ) {
                        result = aLine;
                        break;
                    }
                    aLine = strReader.ReadLine();
                }
            }
            else {
                if ( setBiosOptionSuccess ) {
                    result = "success";
                }
            }
            return result;
        }
        /// <summary>
        /// Get the value of a BIOS item
        /// </summary>
        /// <param name="abstPath">The absolute path of abst64_unsigned.exe</param>
        /// <param name="password">Superviser password</param>
        /// <param name="item">The BIOS item to be set</param>
        /// <returns>the value of the item in BIOS, fail if not</returns>
        public static string Get( string abstPath, string password, string item ) {
            string result = "fail";
            string output;
            string aLine;
            string[] splitStrings;

            string arguments = " /password " + password + " /get \"" + item;
            output = Abst.run(abstPath, arguments);

            bool getBiosOptionSuccess = output.Contains("Get BIOS options success");
            if ( getBiosOptionSuccess ) {
                StringReader strReader = new StringReader(output);
                aLine = strReader.ReadLine();
                while ( aLine != null ) {

                    if ( aLine.Contains(item.ToLower() + ":") ) {
                        splitStrings = aLine.TrimEnd().Split(':');
                        result = splitStrings[1].TrimStart();
                    }
                    aLine = strReader.ReadLine();
                }
            }
            return result;
        }
        /// <summary>
        /// Import BIOS custom settings XML file
        /// </summary>
        /// <param name="abstPath">The absolute path of abst64_unsigned.exe</param>
        /// <param name="password">Superviser password</param>
        /// <param name="fileName">The bios settings xml file name</param>
        /// <returns>the value of the item in BIOS, fail if not</returns>
        public static string ImportCustomSettings( string abstPath, string password, string fileName ) {
            string result = "fail";
            string output;
            string exePath = abstPath;
            string arguments = " /password " + password + " /import-custom-settings \"" + fileName + "\"";

            output = Abst.run(exePath, arguments);
            if ( output.Contains(fileName) ) {
                result = "success";
            }
            return result;
        }
    }
}