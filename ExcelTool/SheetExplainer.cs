using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{
    public enum FindingMethod { SAME, CONTAIN, REGEX }
    public class SheetExplainer
    {
        public List<String> relativePathes = new List<String>();
        public KeyValuePair<FindingMethod, List<String>> fileNames = new KeyValuePair<FindingMethod, List<String>>();
        public KeyValuePair<FindingMethod, List<String>> sheetNames = new KeyValuePair<FindingMethod, List<String>>();
    }
}
