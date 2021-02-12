using AutomationDesinger.Build;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationDesinger.Logs
{
    public class LogWriter : ExcelBaseParse
    {
        public LogWriter(Worksheet worksheet, bool createColumns, List<LogData> logs) : base(worksheet)
        {
            var startIndex = 1;

            if (createColumns)
            {
                SetValue(1, 1, "Id");
                SetValue(1, 2, "Type");
                SetValue(1, 3, "Message");

                startIndex = 2;
            }
            else
            {
                while (!string.IsNullOrWhiteSpace(GetString(startIndex, 1)))
                {
                    startIndex++;

                    if (startIndex > 10000)
                    {
                        throw new Exception("Could not find start index for writing log");
                    }
                }
            }

            foreach (var l in logs)
            {
                SetValue(startIndex, 1, (startIndex - 1).ToString());
                SetValue(startIndex, 2, l.Type);
                SetValue(startIndex, 3, l.Message);

                startIndex++;
            }    
        }
    }
}
