using CustomizableMessageBox;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace GlobalObjects
{
    public enum ResultType { FILEPATH, FILENAME, MESSAGE, RESULTOBJECT };

    public static class GlobalObjects
    {
        private static Object globalParam;
        private static PropertiesSetter ps = null;
        private static Style style = null;

        public static Object GetGlobalParam()
        {
            return globalParam;
        }
        public static void SetGlobalParam(Object _globalParam)
        {
            globalParam = _globalParam;
        }
        public static Style GetBtnStyle()
        {
            if (style == null)
            {
                style = new Style();
                style.TargetType = typeof(Button);
                style.Setters.Add(new Setter(Button.BackgroundProperty, new SolidColorBrush(Colors.White)));
                style.Setters.Add(new Setter(Button.BorderBrushProperty, new SolidColorBrush(Colors.LightGray)));
                style.Setters.Add(new Setter(Button.ForegroundProperty, new SolidColorBrush(Colors.Black)));
                style.Setters.Add(new Setter(Button.FontSizeProperty, (double)12));
                style.Setters.Add(new Setter(Button.BorderThicknessProperty, new Thickness(1)));
                style.Setters.Add(new Setter(Button.MarginProperty, new Thickness(5)));
                style.Setters.Add(new Setter(Button.HeightProperty, 25d));
            }
            return style;
        }
        public static PropertiesSetter GetPropertiesSetter()
        {
            if (ps == null)
            {
                ps = new PropertiesSetter();
                ps.WndBorderThickness = new Thickness(1);
                ps.WndBorderColor = new MessageBoxColor(Colors.LightGray);
                ps.ButtonPanelColor = new MessageBoxColor("White");
                ps.MessagePanelColor = new MessageBoxColor("White");
                ps.TitlePanelColor = new MessageBoxColor(Colors.LightGray);
                ps.TitlePanelBorderThickness = new Thickness(0, 0, 0, 2);
                ps.TitlePanelBorderColor = new MessageBoxColor("#FFEFE2E2");
                ps.MessagePanelBorderThickness = new Thickness(0);
                ps.ButtonPanelBorderThickness = new Thickness(0);
                ps.TitleFontSize = 14;
                ps.TitleFontColor = new MessageBoxColor(Colors.Black);
                ps.MessageFontColor = new MessageBoxColor(Colors.Black);
                ps.MessageFontSize = 14;
                ps.ButtonFontSize = 16;
                ps.ButtonBorderThickness = new Thickness(1);
                ps.ButtonBorderColor = new MessageBoxColor(Colors.LightGray);
                ps.WindowMinHeight = 200;
                ps.LockHeight = false;
                ps.WindowWidth = 450;
                ps.WindowShowDuration = new Duration(new TimeSpan(0, 0, 0, 0, 300));
                ps.CloseTimer = new MessageBoxCloseTimer(999999, -100);
                List<Style> buttonStyleList = new List<Style>();
                buttonStyleList.Add(GetBtnStyle());
                ps.ButtonStyleList = buttonStyleList;
            }

            return ps;
        }
        public static String GetDefaultCode()
        {
            return @"using ClosedXML.Excel;
using GlobalObjects;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;

namespace AnalyzeCode
{
    class Analyze
    {
        /// <summary>
        /// 在所有分析前调用
        /// </summary>
        /// <param name=""paramDic"">传入的参数</param>
        /// <param name=""globalObject"">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name=""allFilePathList"">将会分析的所有文件路径列表</param>
        public void RunBeforeAnalyzeSheet(Dictionary<string, string> paramDic, ref Object globalObject, List<string> allFilePathList)
        {
            
        }

        /// <summary>
        /// 分析一个sheet
        /// </summary>
        /// <param name=""paramDic"">传入的参数</param>
        /// <param name=""sheet"">被分析的sheet</param>
        /// <param name=""result"">存储当前文件的信息 ResultType { (String) FILEPATH [文件路径], (String) FILENAME [文件名], (String) MESSAGE [当查找时出现问题时输出的消息], (Object) RESULTOBJECT [用户自定的分析结果] }</param>
        /// <param name=""globalObject"">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name=""invokeCount"">此分析函数被调用的次数</param>
        public void AnalyzeSheet(Dictionary<string, string> paramDic, IXLWorksheet sheet, ConcurrentDictionary<ResultType, Object> result, ref Object globalObject, int invokeCount)
        {
            
        }

        /// <summary>
        /// 在所有输出前调用
        /// </summary>
        /// <param name=""paramDic"">传入的参数</param>
        /// <param name=""workbook"">用于输出的excel文件</param>
        /// <param name=""globalObject"">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name=""resultList"">所有文件的信息</param>
        /// <param name=""allFilePathList"">分析的所有文件路径列表</param>
        public void RunBeforeSetResult(Dictionary<string, string> paramDic, XLWorkbook workbook, ref Object globalObject, ICollection<ConcurrentDictionary<ResultType, Object>> resultList, List<string> allFilePathList)
        {
            
        }

        /// <summary>
        /// 根据分析结果输出到excel中
        /// </summary>
        /// <param name=""paramDic"">传入的参数</param>
        /// <param name=""workbook"">用于输出的excel文件</param>
        /// <param name=""result"">存储当前文件的信息 ResultType { (String) FILEPATH [文件路径], (String) FILENAME [文件名], (String) MESSAGE [当查找时出现问题时输出的消息], (Object) RESULTOBJECT [用户自定的分析结果] }</param>
        /// <param name=""globalObject"">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name=""invokeCount"">此输出函数被调用的次数</param>
        /// <param name=""totalCount"">总共需要调用的输出函数的次数</param>
        public void SetResult(Dictionary<string, string> paramDic, XLWorkbook workbook, ConcurrentDictionary<ResultType, Object> result, ref Object globalObject, int invokeCount, int totalCount)
        {
            
        }

        /// <summary>
        /// 所有调用结束后调用
        /// </summary>
        /// <param name=""paramDic"">传入的参数</param>
        /// <param name=""workbook"">用于输出的excel文件</param>
        /// <param name=""globalObject"">全局存在, 可以保存需要在其他调用时使用的数据, 如当前行号等</param>
        /// <param name=""resultList"">所有文件的信息</param>
        /// <param name=""allFilePathList"">分析的所有文件路径列表</param>
        public void RunEnd(Dictionary<string, string> paramDic, XLWorkbook workbook, ref Object globalObject, ICollection<ConcurrentDictionary<ResultType, Object>> resultList, List<string> allFilePathList)
        {
            
        }
    }
}
";
        }
    }

    public static class Logger
    {
        private static string log = "";

        public static void Info(string info)
        {
            log = $"{log}[Info] {info}\n";
        }

        public static void Error(string error)
        {
            log = $"{log}[Error] {error}\n";
        }

        public static void Warn(string warn)
        {
            log = $"{log}[Warn] {warn}\n";
        }

        public static void Print(string str)
        {
            log = $"{log}{str}\n";
        }

        public static string Get()
        {
            string logTemp = log;
            Clear();
            return logTemp;
        }

        public static void Clear()
        {
            log = "";
        }
    }
}
