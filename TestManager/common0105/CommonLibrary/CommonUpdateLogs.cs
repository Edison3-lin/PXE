using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace CaptainWin.CommonAPI{
    public class CommonUpdateLogs{
        /// <summary>
        /// To check test log json file path exist or lost
        /// </summary>
        /// <returns></returns>
        public string GetTestLogPath(){
            //Local log file path in DUT 
            string logPath = @"c:\\TestManager\\TR_Result.json";
            try{
                if (File.Exists(logPath)){
                    return logPath;
                }else{
                    Console.WriteLine("File not exist in C:\\TestManager, please help to check");
                }
            }catch (Exception ex){
                Console.WriteLine(ex.Message);
            }
            return "";
        }
        /// <summary>
        /// Read test status from json file
        /// </summary>
        public void ReadTestLogStatus(){
            //string filePath = @"c:\\TestManager\\TR_Result.json"; // 將路徑替換為你的JSON文件的實際路徑
            string filePath = GetTestLogPath();
            // 讀取JSON文件內容
            string jsonContent = File.ReadAllText(filePath);
            // 將JSON字串解析為JObject
            JObject jsonObject = JObject.Parse(jsonContent);
            // 讀取"TestStatus"的值
            string test_status = (string)jsonObject["TestStatus"];
            string test_result = (string)jsonObject["TestResult"];
            Console.WriteLine("Test[Status] is: " + test_status);
            Console.WriteLine("Test[Result] is: " + test_result);
        }
        /// <summary>
        /// Write test log status to json file
        /// </summary>
        /// <param name="writetag"></param>
        /// <param name="result"></param>
        public void WriteTestLogStatus(string writetag, string result){
            string filePath = GetTestLogPath();
            // 讀取JSON文件內容
            string jsonContent1 = File.ReadAllText(filePath);
            // 將JSON字串解析為JObject
            JObject jsonObject1 = JObject.Parse(jsonContent1);
            // 修改 "site" 內容
            jsonObject1[writetag] = result; // 在這裡將新的值賦給 "site" 屬性
                                            //jsonObject1["TestResult"] = "SWQuantaBU4SW";
                                            // 將修改後的JObject轉換回JSON字符串
            string modifiedJson1 = jsonObject1.ToString();
            // 將修改後的JSON字串保存回文件
            File.WriteAllText(filePath, modifiedJson1);
        }
        /// <summary>
        /// To dump all context in json file
        /// </summary>
        /// <param name="filePath"></param>
        public void DumpTestLogJsonFile(string filePath){
            TRlog trlog = new TRlog();
            trlog = ReadJsonFile<TRlog>(filePath);
            DisplayTRlog(trlog);
        }
        /// <summary>
        /// Read and deserialize data in json file, will return a format of deserialize json string
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <returns></returns>
        static T ReadJsonFile<T>(string filePath){
            try{
                // Read the JSON file into a string
                string jsonString = File.ReadAllText(filePath);

                // Deserialize the JSON into an object
                T jsonObject = JsonConvert.DeserializeObject<T>(jsonString);

                return jsonObject;
            }catch (Exception ex){
                Console.WriteLine($"Error reading JSON file: {ex.Message}");
                return default(T);
            }
        }
        /// <summary>
        /// Display context of json file
        /// </summary>
        /// <param name="trlog"></param>
        public void DisplayTRlog(TRlog trlog){
            if (trlog != null){
                Console.WriteLine($"TCM_ID: {trlog.TCM_ID}");
                Console.WriteLine($"TR_ID: {trlog.TR_ID}");
                Console.WriteLine($"TestResult: {trlog.TestResult}");
                Console.WriteLine($"TestStatus: {trlog.TestStatus}");
                Console.WriteLine($"Test_TimeOut: {trlog.Test_TimeOut}");
                Console.WriteLine($"TestFail_Dercription: {trlog.TestFail_Dercription}");

                foreach (var item in trlog.printlog){
                    Console.WriteLine($"time: {item.time}");
                    Console.WriteLine($"LogType: {item.LogType}");
                    Console.WriteLine($"log: {item.log}");
                }

                Console.WriteLine("Text_Log_File_Path:");
                foreach (var text_log in trlog.Text_Log_File_Path){
                    Console.WriteLine($"{text_log}");
                }
                Console.WriteLine("Image_Log_File_Path");
                foreach (var image_log in trlog.Image_Log_File_Path){
                    Console.WriteLine($"{image_log}");
                }
                Console.WriteLine($"Pass: {trlog.Pass}");
            }
        }
    }
    /// <summary>
    /// Class for mapping the data struct of TR_result.json
    /// </summary>
    public class TRlog{
        public TRlog(){
            //Console.WriteLine("Construct TRlog");
            printlog = new List<PrintLog>();
            Text_Log_File_Path = new List<string>();
            Image_Log_File_Path = new List<string>();
        }
        public int TCM_ID { get; set; }
        public int TR_ID { get; set; }
        public string TestResult { get; set; }
        public string TestStatus { get; set; }

        public int Test_TimeOut { get; set; }

        public string TestFail_Dercription { get; set; }
        public List<PrintLog> printlog { get; set; }

        public List<string> Text_Log_File_Path { get; set; }

        public List<string> Image_Log_File_Path { get; set; }

        public string Pass { get; set; }

        public class PrintLog{
            public string time { get; set; }
            public string LogType { get; set; }
            public string log { get; set; }
        }
    }
}
