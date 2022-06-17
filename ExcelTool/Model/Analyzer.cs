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
        public ParamInfo()
        { 
        
        }
        public ParamInfo(List<PossibleValue> possibleValues, ParamType type)
        {
            this.possibleValues = possibleValues;
            this.type = type;
        }
        public ParamInfo(String describe, List<PossibleValue> possibleValues, ParamType type)
        {
            this.describe = describe;
            this.possibleValues = possibleValues;
            this.type = type;
        }
        public String describe;
        public List<PossibleValue> possibleValues;
        public ParamType type;
    }

    public class PossibleValue
    {
        public string value;
        public string describe;
        public PossibleValue()
        {

        }
        public PossibleValue(string value)
        {
            this.value = value;
        }
        public PossibleValue(string value, string describe)
        {
            this.value = value;
            this.describe = describe;
        }
    }
}
