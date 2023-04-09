using ClosedXML.Excel;
using CustomizableMessageBox;
using Microsoft.CodeAnalysis.CSharp;
using ModernWpf;
using System;
using System.CodeDom.Compiler;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace GlobalObjects
{
    public enum ProgramStatus { Default, Shutdown, Restart };

    public static class GlobalDic
    {
        private static ConcurrentDictionary<string, object> globalDic;

        public static void SetObj(string key, object value)
        {
            globalDic.AddOrUpdate(key, value, (k, v) => { return v; });
        }

        public static bool ContainsKey(string key)
        {
            return globalDic.ContainsKey(key);
        }

        public static object GetObj(string key)
        {
            globalDic.TryGetValue(key, out object value);
            return value;
        }

        public static void Reset()
        {
            globalDic = new ConcurrentDictionary<string, object>();
        }
    }

    [SupportedOSPlatform("windows7.0")]
    public static class Theme
    {
        private static Brush themeBackground;
        public static Brush ThemeBackground
        {
            get { return themeBackground; }
            set { themeBackground = value; }
        }

        private static Brush themeControlBackground;
        public static Brush ThemeControlBackground
        {
            get { return themeControlBackground; }
            set { themeControlBackground = value; }
        }

        private static Brush themeControlFocusBackground;
        public static Brush ThemeControlFocusBackground
        {
            get { return themeControlFocusBackground; }
            set { themeControlFocusBackground = value; }
        }

        private static Brush themeControlForeground;
        public static Brush ThemeControlForeground
        {
            get { return themeControlForeground; }
            set { themeControlForeground = value; }
        }

        private static Brush themeControlBorderBrush;
        public static Brush ThemeControlBorderBrush
        {
            get { return themeControlBorderBrush; }
            set { themeControlBorderBrush = value; }
        }

        private static Brush themeControlHoverBorderBrush;
        public static Brush ThemeControlHoverBorderBrush
        {
            get { return themeControlHoverBorderBrush; }
            set { themeControlHoverBorderBrush = value; }
        }

        private static Brush themeMessageBoxTitlePanelBackgroundBrush;
        public static Brush ThemeMessageBoxTitlePanelBackgroundBrush
        {
            get { return themeMessageBoxTitlePanelBackgroundBrush; }
            set { themeMessageBoxTitlePanelBackgroundBrush = value; }
        }

        private static Brush themeMessageBoxTitlePanelBorderBrush;
        public static Brush ThemeMessageBoxTitlePanelBorderBrush
        {
            get { return themeMessageBoxTitlePanelBorderBrush; }
            set { themeMessageBoxTitlePanelBorderBrush = value; }
        }

        private static Brush themeSelectionBrush;
        public static Brush ThemeSelectionBrush
        {
            get { return themeSelectionBrush; }
            set { themeSelectionBrush = value; }
        }

        private static Brush themeSelectionForeground;
        public static Brush ThemeSelectionForeground
        {
            get { return themeSelectionForeground; }
            set { themeSelectionForeground = value; }
        }

        public static void SetTheme()
        {
            if (ThemeManager.Current.ActualApplicationTheme == ApplicationTheme.Dark)
            {
                ThemeBackground = new MessageBoxColor("#2F2F31").solidColorBrush;
                ThemeControlBackground = new MessageBoxColor("#1C1C1D").solidColorBrush;
                ThemeControlFocusBackground = new MessageBoxColor("#2F2F31").solidColorBrush;
                ThemeControlForeground = new SolidColorBrush(Color.FromRgb(220, 220, 220));
                ThemeControlBorderBrush = Brushes.Gainsboro;
                ThemeControlHoverBorderBrush = Brushes.Gainsboro;
                ThemeMessageBoxTitlePanelBackgroundBrush = Brushes.DimGray;
                ThemeMessageBoxTitlePanelBorderBrush = Brushes.Gray;
                ThemeSelectionBrush = new SolidColorBrush(Color.FromArgb(100, 51, 153, 255));
                ThemeSelectionForeground = new SolidColorBrush(Color.FromRgb(220, 220, 220));
            }
            else
            {
                ThemeBackground = Brushes.White;
                ThemeControlBackground = Brushes.White;
                ThemeControlFocusBackground = Brushes.White;
                ThemeControlForeground = Brushes.Black;
                ThemeControlBorderBrush = new MessageBoxColor(ThemeManager.Current.ActualAccentColor).solidColorBrush;
                ThemeControlHoverBorderBrush = Brushes.Black;
                ThemeMessageBoxTitlePanelBackgroundBrush = Brushes.LightGray;
                ThemeMessageBoxTitlePanelBorderBrush = new MessageBoxColor("#FFEFE2E2").solidColorBrush;
                ThemeSelectionBrush = new SolidColorBrush(Colors.SkyBlue);
                ThemeSelectionForeground = new SolidColorBrush(Colors.White);
            }

            GlobalObjects.ps = new PropertiesSetter
            {
                WndBorderThickness = new Thickness(1),
                WndBorderColor = new MessageBoxColor(((SolidColorBrush)Theme.ThemeMessageBoxTitlePanelBackgroundBrush).Color),
                ButtonPanelColor = new MessageBoxColor(((SolidColorBrush)Theme.ThemeBackground).Color),
                MessagePanelColor = new MessageBoxColor(((SolidColorBrush)Theme.ThemeBackground).Color),
                ButtonFontColor = new MessageBoxColor(((SolidColorBrush)Theme.ThemeControlForeground).Color),
                TitlePanelColor = new MessageBoxColor(((SolidColorBrush)Theme.ThemeMessageBoxTitlePanelBackgroundBrush).Color),
                TitlePanelBorderThickness = new Thickness(0, 0, 0, 2),
                TitlePanelBorderColor = new MessageBoxColor(((SolidColorBrush)Theme.ThemeMessageBoxTitlePanelBorderBrush).Color),
                MessagePanelBorderThickness = new Thickness(0),
                ButtonPanelBorderThickness = new Thickness(0),
                TitleFontSize = 14,
                TitleFontColor = new MessageBoxColor(((SolidColorBrush)Theme.ThemeControlForeground).Color),
                MessageFontColor = new MessageBoxColor(((SolidColorBrush)Theme.ThemeControlForeground).Color),
                MessageFontSize = 14,
                ButtonFontSize = 16,
                ButtonBorderBrushList = new List<Brush>() { Theme.ThemeBackground },
                ButtonBorderThicknessList = new List<Thickness>() { new Thickness(0) },
                WindowMinHeight = 200,
                LockHeight = false,
                WindowWidth = 450,
                WindowShowDuration = new Duration(new TimeSpan(0, 0, 0, 0, 300)),
                ButtonMarginList = new List<Thickness>() { new Thickness(5) },
                ButtonHeightList = new List<double>() { 30 }
            };

            CustomizableMessageBox.MessageBox.DefaultProperties = GlobalObjects.ps;
        }
    }

    public static class GlobalObjects
    {
        private static ProgramStatus programCurrentStatus;

        private static ConcurrentDictionary<object, object> globalParamDic;
        private static PropertiesSetter psWithTimmer = null;
        public static PropertiesSetter ps = null;

        public static ProgramStatus ProgramCurrentStatus { get => programCurrentStatus; set => programCurrentStatus = value; }

        public static Object GetGlobalParam(object instanceObj)
        {
            globalParamDic.TryGetValue(instanceObj, out object res);
            return res;
        }
        public static void SetGlobalParam(object instanceObj, object globalParam)
        {
            globalParamDic.AddOrUpdate(instanceObj, globalParam, (key, oldValue) => { return globalParam; });
        }
        public static void ClearGlobalParamDic()
        {
            globalParamDic = new ConcurrentDictionary<object, object>();
        }

        public static void ClearPropertiesSetter()
        {
            ps = null;
            psWithTimmer = null;
        }

        public static PropertiesSetter GetPropertiesSetterWithTimmer()
        {
            psWithTimmer ??= new PropertiesSetter(ps)
            {
                CloseTimer = new MessageBoxCloseTimer(1, 0)
            };

            return psWithTimmer;
        }
        public static String GetDefaultCode()
        {
            return Application.Current.FindResource("Code").ToString();
        }
    }

    public class Param
    {
        private readonly Dictionary<string, List<string>> paramDic;

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
            Input?.Invoke(this, e); // 调用所有注册对象的方法
        }

        public void UpdateInput(string value)
        {
            InputEventArgs e = new(value);
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
        private static string outputPath = "";
        public static string OutputPath { get => outputPath; set => outputPath = value; }

        public static XLWorkbook CreateWorkbook(string name)
        {
            if (workbookDic.ContainsKey(name))
            {
                throw new Exception("The excel name is repeated multiple times. ");
            }

            XLWorkbook workbook = new();

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

            workbookDic.TryGetValue(name, out XLWorkbook workbook);
            return workbook;
        }

        public static IXLWorksheet GetSheet(string workbookName, string sheetName)
        {
            if (!workbookDic.ContainsKey(workbookName))
            {
                throw new Exception("Can't find the workbook. ");
            }

            workbookDic.TryGetValue(workbookName, out XLWorkbook workbook);
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

    public static class Running {
        private static bool userStop = false;
        public static bool UserStop { get => userStop; set => userStop = value; }
        
        private static bool nowRunning = false;
        public static bool NowRunning { get => nowRunning; set => nowRunning = value; }
    }
}