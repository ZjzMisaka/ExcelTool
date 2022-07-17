using ClosedXML.Excel;
using CustomizableMessageBox;
using System;
using System.CodeDom.Compiler;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace GlobalObjects
{
    public enum ResultType { FILEPATH, FILENAME, MESSAGE, RESULTOBJECT };
    public enum ProgramStatus { Default, Shutdown, Restart };

    public static class GlobalObjects
    {
        private static ProgramStatus programCurrentStatus;

        private static ConcurrentDictionary<CompilerResults, Object> globalParamDic;
        private static PropertiesSetter ps = null;
        private static PropertiesSetter psWithTimmer = null;
        private static Style style = null;

        public static ProgramStatus ProgramCurrentStatus { get => programCurrentStatus; set => programCurrentStatus = value; }

        public static Object GetGlobalParam(CompilerResults compilerResults)
        {
            Object res;
            globalParamDic.TryGetValue(compilerResults, out res);
            return res;
        }
        public static void SetGlobalParam(CompilerResults compilerResults, Object globalParam)
        {
            globalParamDic.AddOrUpdate(compilerResults, globalParam, (key, oldValue) => { return globalParam; });
        }
        public static void ClearGlobalParamDic()
        {
            globalParamDic = new ConcurrentDictionary<CompilerResults, object>();
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
                ps.ButtonMarginList = new List<Thickness>() { new Thickness(5) };
                ps.ButtonHeightList = new List<double>() { 30 };
            }

            return ps;
        }
        public static PropertiesSetter GetPropertiesSetterWithTimmer()
        {
            if (psWithTimmer == null)
            {
                psWithTimmer = new PropertiesSetter();
                psWithTimmer.WndBorderThickness = new Thickness(1);
                psWithTimmer.WndBorderColor = new MessageBoxColor(Colors.LightGray);
                psWithTimmer.ButtonPanelColor = new MessageBoxColor("White");
                psWithTimmer.MessagePanelColor = new MessageBoxColor("White");
                psWithTimmer.TitlePanelColor = new MessageBoxColor(Colors.LightGray);
                psWithTimmer.TitlePanelBorderThickness = new Thickness(0, 0, 0, 2);
                psWithTimmer.TitlePanelBorderColor = new MessageBoxColor("#FFEFE2E2");
                psWithTimmer.MessagePanelBorderThickness = new Thickness(0);
                psWithTimmer.ButtonPanelBorderThickness = new Thickness(0);
                psWithTimmer.TitleFontSize = 14;
                psWithTimmer.TitleFontColor = new MessageBoxColor(Colors.Black);
                psWithTimmer.MessageFontColor = new MessageBoxColor(Colors.Black);
                psWithTimmer.MessageFontSize = 14;
                psWithTimmer.ButtonFontSize = 16;
                psWithTimmer.ButtonBorderThickness = new Thickness(1);
                psWithTimmer.ButtonBorderColor = new MessageBoxColor(Colors.LightGray);
                psWithTimmer.WindowMinHeight = 200;
                psWithTimmer.LockHeight = false;
                psWithTimmer.WindowWidth = 450;
                psWithTimmer.WindowShowDuration = new Duration(new TimeSpan(0, 0, 0, 0, 300));
                psWithTimmer.CloseTimer = new MessageBoxCloseTimer(1, 0);
                psWithTimmer.ButtonMarginList = new List<Thickness>() { new Thickness(5) };
                psWithTimmer.ButtonHeightList = new List<double>() { 30 };
            }

            return psWithTimmer;
        }
        public static String GetDefaultCode()
        {
            return Application.Current.FindResource("Code").ToString();
        }
    }

    public class Param
    {
        private Dictionary<string, List<string>> paramDic;

        public Param(Dictionary<string, List<string>> paramDic)
        {
            this.paramDic = paramDic;
        }

        public List<string> Get(string key)
        {
            if (!paramDic.ContainsKey(key))
            {
                return new List<string>();
            }
            return paramDic[key];
        }

        public string GetOne(string key)
        {
            if (!paramDic.ContainsKey(key))
            {
                return null;
            }
            return paramDic[key].First();
        }

        public IEnumerable<String> GetKeys()
        {
            return paramDic.Keys;
        }

        public bool ContainsKey(string key)
        {
            return paramDic.ContainsKey(key);
        }
    }

    public static class Logger
    {
        private static string log = "";

        private static bool isOutputMethodNotFoundWarning = true;

        public static bool IsOutputMethodNotFoundWarning { get => isOutputMethodNotFoundWarning; set => isOutputMethodNotFoundWarning = value; }

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

    public class Scanner
    {
        private static Object lockObj;

        private static bool inputLock = false;

        private static string value;
        private static string currentInputMessage;
        private static string lastInputValue;

        public delegate void InputEventHandler(Object sender, InputEventArgs e);

        public event InputEventHandler Input;

        public static bool InputLock { get => inputLock; set => inputLock = value; }
        public static string CurrentInputMessage { get => currentInputMessage; set => currentInputMessage = value; }
        public static string LastInputValue { get => lastInputValue; set => lastInputValue = value; }

        public class InputEventArgs : EventArgs
        {
            public readonly string value;
            public InputEventArgs(string value)
            {
                this.value = value;
            }
        }
        protected virtual void OnInput(InputEventArgs e)
        {
            if (Input != null)
            {
                Input(this, e); // 调用所有注册对象的方法
            }
        }

        public void UpdateInput(string value)
        {
            InputEventArgs e = new InputEventArgs(value);
            OnInput(e);
        }

        public static void UserInput(Object sender, InputEventArgs e)
        {
            Scanner.value = e.value;
        }

        public static string WaitInput()
        {
            return GetInput("", true);
        }

        public static string GetInput()
        {
            return GetInput("", false);
        }

        public static string GetInput(string value)
        {
            return GetInput(value, false);
        }

        private static string GetInput(string value, bool isWait)
        {
            lock (lockObj)
            {
                if (!isWait)
                {
                    InputLock = true;

                    CurrentInputMessage = $"[Input] {value} > ".TrimStart();

                    while (Scanner.value == null)
                    {
                        Thread.Sleep(100);
                    }

                    InputLock = false;

                    LastInputValue = Scanner.value;

                    Scanner.value = null;
                }

                return LastInputValue;
            }
        }

        public static void ResetAll()
        {
            lockObj = new Object();
            CurrentInputMessage = "";
            InputLock = false;
            LastInputValue = "";
        }
    }

    public static class Output
    {
        private static ConcurrentDictionary<string, XLWorkbook> workbookDic;

        private static bool isSaveDefaultWorkBook = true;

        public static bool IsSaveDefaultWorkBook { get => isSaveDefaultWorkBook; set => isSaveDefaultWorkBook = value; }

        public static XLWorkbook CreateWorkbook(string name)
        {
            if (workbookDic.ContainsKey(name))
            {
                throw new Exception("The excel name is repeated multiple times. ");
            }

            XLWorkbook workbook = new XLWorkbook();

            if (!workbookDic.TryAdd(name, workbook))
            {
                return null;
            }

            return workbook;
        }

        public static XLWorkbook GetWorkbook(string name)
        {
            if (!workbookDic.ContainsKey(name))
            {
                throw new Exception("Can't find the book. ");
            }

            XLWorkbook workbook = null;
            workbookDic.TryGetValue(name, out workbook);
            return workbook;
        }

        public static IXLWorksheet GetSheet(string workbookName, string sheetName)
        {
            if (!workbookDic.ContainsKey(workbookName))
            {
                throw new Exception("Can't find the workbook. ");
            }

            XLWorkbook workbook = null;
            workbookDic.TryGetValue(workbookName, out workbook);
            if (workbook == null)
            {
                throw new Exception("Can't get the worksheet. get workbook failed. ");
            }

            return GetSheet(workbook, sheetName);
        }

        public static IXLWorksheet GetSheet(XLWorkbook workbook, string sheetName)
        {
            if (!workbook.Worksheets.Contains(sheetName))
            {
                throw new Exception("Can't find the worksheet. ");
            }

            return workbook.Worksheet(sheetName);
        }

        public static Dictionary<string, XLWorkbook> GetAllWorkbooks()
        {
            return workbookDic.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
        }

        public static void ClearWorkbooks()
        {
            workbookDic = new ConcurrentDictionary<string, XLWorkbook>();
        }
    }
}