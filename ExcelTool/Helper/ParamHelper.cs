using GlobalObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.Helper
{
    public static class ParamHelper
    {
        public static string GetParamStr(Dictionary<string, Dictionary<string, string>> paramDicEachAnalyzer)
        {
            string paramStr = "";
            foreach (string analyzerName in paramDicEachAnalyzer.Keys)
            {
                Dictionary<string, string> paramDic = paramDicEachAnalyzer[analyzerName];
                if (analyzerName == "public")
                {
                    foreach (string key in paramDic.Keys)
                    {
                        paramStr = $"{ paramStr }|{key}:{Decode1(paramDic[key])}";
                    }
                }
                else
                {
                    string paramStrTemp = "";
                    foreach (string key in paramDic.Keys)
                    {
                        paramStrTemp = $"{ paramStrTemp }|{key}:{Decode1(paramDic[key])}";
                    }
                    paramStrTemp = paramStrTemp.Substring(1);

                    paramStr = $"{ paramStr }|{analyzerName}{{{paramStrTemp}}}";
                }
            }
            if (paramStr.Length >= 1)
            {
                paramStr = paramStr.Substring(1);
            }
            return paramStr;
        }

        public static Dictionary<string, Dictionary<string, string>> GetParamDicEachAnalyzer(string paramStr)
        {
            Dictionary<string, Dictionary<string, string>> paramDicEachAnalyzer = new Dictionary<string, Dictionary<string, string>>();

            int startIndex = 0;
            int index = 0;
            while (index <= paramStr.Length - 1)
            {
                string tempStr = paramStr.Substring(startIndex);
                int tempIndex = tempStr.IndexOf("|");
                string tempParam = "";
                if (tempIndex == -1)
                {
                    tempParam = tempStr;

                    index = paramStr.Length + 1;
                }
                else
                {
                    tempParam = tempStr.Substring(0, tempIndex);
                }
                if (tempParam.Contains("{"))
                {
                    tempIndex = tempStr.IndexOf("}");
                    tempParam = tempStr.Substring(0, tempIndex + 1);

                    int lIndex = tempParam.IndexOf("{");
                    int rIndex = tempParam.IndexOf("}");
                    string analyzerName = tempParam.Substring(0, lIndex);
                    string tempParamInTempParam = tempParam.Substring(lIndex + 1, rIndex - lIndex - 1);
                    string[] splitedParams = tempParamInTempParam.Split('|');
                    Dictionary<string, string> paramDic = new Dictionary<string, string>();
                    foreach (string param in splitedParams)
                    {
                        string[] kv = param.Split(':');
                        if (kv.Length == 2)
                        {
                            paramDic.Add(kv[0], kv[1]);
                        }
                    }
                    if (paramDicEachAnalyzer.ContainsKey(analyzerName))
                    {
                        foreach (string key in paramDic.Keys)
                        {
                            if (!paramDicEachAnalyzer[analyzerName].ContainsKey(key))
                            {
                                paramDicEachAnalyzer[analyzerName].Add(key, paramDic[key]);
                            }
                        }
                    }
                    else
                    {
                        paramDicEachAnalyzer.Add(analyzerName, paramDic);
                    }
                }
                else
                {
                    string[] kv = tempParam.Split(':');
                    if (kv.Length == 2)
                    {
                        if (!paramDicEachAnalyzer.ContainsKey("public"))
                        {
                            paramDicEachAnalyzer.Add("public", new Dictionary<string, string>());
                        }
                        paramDicEachAnalyzer["public"].Add(kv[0], kv[1]);
                    }
                }
                if (tempIndex != -1)
                {
                    index = startIndex + tempIndex + 1;
                }
                startIndex = index;
            }

            return paramDicEachAnalyzer;
        }

        public static Param MergePublicParam(Dictionary<string, Dictionary<string, string>> paramDicEachAnalyzer, string analyzerName)
        {
            Dictionary<string, string> aimDic = new Dictionary<string, string>();
            if (paramDicEachAnalyzer.ContainsKey(analyzerName))
            {
                aimDic = paramDicEachAnalyzer[analyzerName];
            }

            Dictionary<string, string> dic = new Dictionary<string, string>();
            if (paramDicEachAnalyzer.ContainsKey("public"))
            {
                dic = paramDicEachAnalyzer["public"];
            }
            
            foreach (string key in dic.Keys)
            {
                if (!aimDic.ContainsKey(key))
                {
                    aimDic.Add(key, dic[key]);
                }
            }
            return ConvertToParam(aimDic);
        }

        public static string Encode(string paramStr)
        {
            return paramStr.Replace("\\\\", "*Backslash*").Replace("\\{", "*LeftCurlyBrace*").Replace("\\}", "*RightCurlyBrace*").Replace("\\:", "*Colon*").Replace("\\|", "*VerticalLine*").Replace("\\+", "*PlusSign*");
        }

        public static string Decode(string paramStr)
        {
            return paramStr.Replace("*Backslash*", "\\").Replace("*LeftCurlyBrace*", "{").Replace("*RightCurlyBrace*", "}").Replace("*Colon*", ":").Replace("*VerticalLine*", "|").Replace("*PlusSign*", "+");
        }

        public static string Encode1(string paramStr)
        {
            return paramStr.Replace("\\", "*Backslash*").Replace("{", "*LeftCurlyBrace*").Replace("}", "*RightCurlyBrace*").Replace(":", "*Colon*").Replace("|", "*VerticalLine*").Replace("+", "*PlusSign*");
        }

        public static string Decode1(string paramStr)
        {
            return paramStr.Replace("*Backslash*", "\\\\").Replace("*LeftCurlyBrace*", "\\{").Replace("*RightCurlyBrace*", "\\}").Replace("*Colon*", "\\:").Replace("*VerticalLine*", "\\|").Replace("*PlusSign*", "\\+");
        }

        public static Param ConvertToParam(Dictionary<string, string> dic)
        {
            Dictionary<string, List<string>> resParamDic = new Dictionary<string, List<string>>();
            foreach (string key in dic.Keys)
            {
                string[] values = dic[key].Split('+');
                List<string> list = new List<string>();
                foreach (string value in values)
                {
                    list.Add(ParamHelper.Decode(value));
                }
                resParamDic.Add(key, list);
            }
            return new Param(resParamDic);
        }
    }
}
