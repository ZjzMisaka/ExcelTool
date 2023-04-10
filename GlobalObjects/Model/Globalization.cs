using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GlobalObjects.Model
{
    public class Globalization
    {
        public string languageName;
        public string logCode;
        public string logText;

        public Globalization(string languageName, string logCode, string logText)
        {
            this.languageName = languageName;
            this.logCode = logCode;
            this.logText = logText;
        }
    }

    public class GlobalizationSetter
    {
        public string currentLanguageName;
        public string defaultLanguageName;
        public bool enableGlobalizationForParamSetter;
        public List<Globalization> globalizationList;

        public GlobalizationSetter()
        {
            enableGlobalizationForParamSetter = false;
            globalizationList = new List<Globalization>();
        }

        public GlobalizationSetter(string currentLanguageName, string defaultLanguageName, bool enableGlobalizationForParamSetter, List<Globalization> globalizationList)
        {
            this.currentLanguageName = currentLanguageName;
            this.defaultLanguageName = defaultLanguageName;
            this.enableGlobalizationForParamSetter = enableGlobalizationForParamSetter;
            this.globalizationList = globalizationList;
        }

        public void Add(string languageName, string logCode, string logText)
        {
            this.globalizationList.Add(new Globalization(languageName, logCode, logText));
        }

        public string Find(string logCode)
        {
            return Find(currentLanguageName, logCode);
        }

        public string Find(string languageName, string logCode)
        {
            List<Globalization> sameCodeList = new();
            foreach (Globalization loggerGlobalization in this.globalizationList)
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
            foreach (Globalization loggerGlobalization in sameCodeList)
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
