using System;
using System.Collections.Generic;

namespace ExcelTool
{
    public enum ParamType { Default, Single, Multiple };
    public class Analyzer
    {
        public String name;
        public String code;
        public Dictionary<string, ParamInfo> paramDic;
    }

    public class ParamInfo
    {
        public ParamInfo(String describe, List<String> possibleValues, ParamType type)
        {
            this.describe = describe;
            this.possibleValues = possibleValues;
            this.type = type;
        }
        public String describe;
        public List<String> possibleValues;
        public ParamType type;
    }
}
