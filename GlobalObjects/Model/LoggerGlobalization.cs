using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.Model
{
    public class LoggerGlobalization
    {
        public string languageName;
        public string logCode;
        public string logText;

        public LoggerGlobalization(string languageName, string logCode, string logText)
        {
            this.languageName = languageName;
            this.logCode = logCode;
            this.logText = logText;
        }
    }

    public class LoggerGlobalizationSetter : List<LoggerGlobalization>
    {
        public string currentLanguageName;
        public string defaultLanguageName;
        public void Add(string languageName, string logCode, string logText)
        {
            this.Add(new LoggerGlobalization(languageName, logCode, logText));
        }

        public string Find(string languageName, string logCode)
        {
            LoggerGlobalizationSetter sameCodeList = new();
            foreach (LoggerGlobalization loggerGlobalization in this)
            {
                if (loggerGlobalization.logCode == logCode)
                {
                    if (loggerGlobalization.languageName == languageName)
                    {
                        return loggerGlobalization.logText;
                    }
                    sameCodeList.Add(loggerGlobalization);
                }
            }
            foreach (LoggerGlobalization loggerGlobalization in sameCodeList)
            {
                if (loggerGlobalization.languageName == defaultLanguageName && loggerGlobalization.logCode == logCode)
                {
                    return loggerGlobalization.logText;
                }
            }
            return $"Log text not found. Code: {logCode}, Language: {languageName}";
        }
    }
}
