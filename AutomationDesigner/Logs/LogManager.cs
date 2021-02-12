using AutomationDesinger.Build;
using AutomationDesinger.Helpers;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutomationDesinger.Logs
{
    public class LogManager
    {
        private static List<LogData> _logData = new List<LogData>();

        public static void WriteLogs(Workbook workBook)
        {
            Worksheet logSheet = null;

            var sheets = workBook.GetWorksheets();

            var existingLog = sheets.Exists(sheet => sheet.Name.ToUpper().Contains("LOGS"));

            if (existingLog)
            {
                logSheet = sheets.FirstOrDefault(sheet => sheet.Name.ToUpper().Contains("LOGS"));
            }
            else
            {
                logSheet = workBook.Worksheets.Add();

                logSheet.Name = "Logs";
            }

            var logWriter = new LogWriter(logSheet, !existingLog, _logData); 
        }

        public static bool HasData => _logData.Count > 0;

        public static void Add(string message, string type = LogTypes.Error)
        {
            _logData.Add(new LogData(message, type));
        }

        public static void Clear()
        {
            _logData.Clear();
        }
    }

    
}
