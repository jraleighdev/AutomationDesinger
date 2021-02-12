using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationDesinger.Logs
{
    public class LogData
    {
        public string Type { get; set; }

        public string Message { get; set; }

        public LogData(string message, string type = LogTypes.Info)
        {
            Type = type;
            Message = message;
        }
    }
}
