/*
* EventLogHelper.cs 
* 
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Bencool   <Bencool.lin@quantatw.com>
*/
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace CaptainWin.CommonAPI
{
    /// <summary>
    /// help to get windws event log. Can pass in parameter to do the filter, like logs for windws or application, create times.
    /// </summary>
    public static  class EventLogHelper
    {
        public const string EventLogType_System = "System";
        public const string EventLogType_Application = "Application";


        /// <summary>
        /// Use C# EventLog library to query windows event logs
        /// </summary>
        /// <param name="fromDate">Datetime type for Query from datetime</param>
        /// <param name="toDate">Datetime type for Query to datetime</param>
        /// <param name="logType">pass in Const string EventLogType_System or EventLogType_Application</param>
        /// <param name="saveFolderPath">null then won't save file. If not null then save data to the pass in folder with file name:QueryEventLog_{DateTime.Now:yyyyMMddHHmmss}.json</param>
        /// <returns>List of EventLogEntryDetails for query data</returns>
        public static List<EventLogEntryDetails> QueryEventLog(DateTime fromDate, DateTime toDate, string logType, string saveFolderPath = null)
        {
            List<EventLogEntryDetails> eventLogEntries = new List<EventLogEntryDetails>();

            try
            {
                EventLog eventLog = new EventLog(logType);

                var query = from EventLogEntry entry in eventLog.Entries
                            where entry.TimeGenerated >= fromDate && entry.TimeGenerated <= toDate
                            select new EventLogEntryDetails
                            {
                                TimeGenerated = entry.TimeGenerated.ToString(),
                                EntryType = entry.EntryType.ToString(),
                                Source = entry.Source,
                                Message = entry.Message
                            };

                eventLogEntries = query.ToList();

                if (!string.IsNullOrEmpty(saveFolderPath))
                {
                    SaveEventLogToJson(eventLogEntries, saveFolderPath);
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions, e.g., log or display an error message
                Console.WriteLine($"Error querying Event Log: {ex.Message}");
            }

            return eventLogEntries;
        }

        private static void SaveEventLogToJson(List<EventLogEntryDetails> eventLogEntries, string saveFolderPath)
        {
            try
            {
                string fileName = $"QueryEventLog_{DateTime.Now:yyyyMMddHHmmss}.json";
                string filePath = Path.Combine(saveFolderPath, fileName);

                string jsonContent = JsonConvert.SerializeObject(eventLogEntries, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(filePath, jsonContent);

                Console.WriteLine($"Event Log saved to: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving Event Log to JSON: {ex.Message}");
            }
        }

    }

    public class EventLogEntryDetails
    {
        public string TimeGenerated { get; set; }
        public string EntryType { get; set; }
        public string Source { get; set; }
        public string Message { get; set; }
    }


}
