using DocumentFormat.OpenXml.Wordprocessing;
using IniParser;
using IniParser.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelTool
{
    public static class IniHelper
    {
        public static void CheckAndCreateIniFile()
        {
            FileIniDataParser parser = new FileIniDataParser();

            if (File.Exists("Setting.ini"))
            {
                return;
            }

            IniData parsedINIDataToBeSaved = new IniData();
            parsedINIDataToBeSaved.Sections.AddSection("Thread");
            parsedINIDataToBeSaved["Thread"].AddKey("MaxThreadCount", "25");
            parsedINIDataToBeSaved["Thread"].AddKey("TotalTimeoutLimitAnalyze", "120000");
            parsedINIDataToBeSaved["Thread"].AddKey("PerTimeoutLimitAnalyze", "60000");
            parsedINIDataToBeSaved["Thread"].AddKey("TotalTimeoutLimitOutput", "120000");
            parsedINIDataToBeSaved["Thread"].AddKey("PerTimeoutLimitOutput", "60000");
            parsedINIDataToBeSaved["Thread"].AddKey("FileSystemWatcherInvokeDalay", "1000");
            parsedINIDataToBeSaved["Thread"].AddKey("FreshInterval", "100");
            parsedINIDataToBeSaved.Sections.AddSection("Window");
            parsedINIDataToBeSaved["Window"].AddKey("MainWindowWidth", "800");
            parsedINIDataToBeSaved["Window"].AddKey("MainWindowHeight", "900");
            parsedINIDataToBeSaved["Window"].AddKey("AnalyzerEditorWidth", "800");
            parsedINIDataToBeSaved["Window"].AddKey("AnalyzerEditorHeight", "450");
            parsedINIDataToBeSaved["Window"].AddKey("SheetExplainerEditorWidth", "800");
            parsedINIDataToBeSaved["Window"].AddKey("SheetExplainerEditorHeight", "450");
            parsedINIDataToBeSaved["Window"].AddKey("IsExecuteInSequence", "False");
            parsedINIDataToBeSaved["Window"].AddKey("IsAutoOpen", "False");
            parsedINIDataToBeSaved["Window"].AddKey("Theme", "Light");
            parsedINIDataToBeSaved["Window"].AddKey("Language", Thread.CurrentThread.CurrentUICulture.Name);
            parsedINIDataToBeSaved.Sections.AddSection("Value");
            parsedINIDataToBeSaved["Value"].AddKey("DefaultBasePath", "");
            parsedINIDataToBeSaved["Value"].AddKey("DefaultOutputPath", Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));
            parsedINIDataToBeSaved["Value"].AddKey("DefaultOutputFileName", "Output");
            parsedINIDataToBeSaved.Sections.AddSection("Check");
            parsedINIDataToBeSaved["Check"].AddKey("SecurityCheck", "True");
            parsedINIDataToBeSaved.Sections.AddSection("Editor");
            parsedINIDataToBeSaved["Editor"].AddKey("ShowSpaces", "False");
            parsedINIDataToBeSaved["Editor"].AddKey("ShowTabs", "False");
            parsedINIDataToBeSaved["Editor"].AddKey("ShowEndOfLine", "False");
            parsedINIDataToBeSaved["Editor"].AddKey("ShowBoxForControlCharacters", "False");
            parsedINIDataToBeSaved["Editor"].AddKey("EnableHyperlinks", "True");
            parsedINIDataToBeSaved["Editor"].AddKey("IndentationSize", "4");
            parsedINIDataToBeSaved["Editor"].AddKey("ConvertTabsToSpaces", "True");
            parsedINIDataToBeSaved["Editor"].AddKey("HighlightCurrentLine", "True");
            parsedINIDataToBeSaved["Editor"].AddKey("HideCursorWhileTyping", "False");
            parsedINIDataToBeSaved["Editor"].AddKey("WordWrap", "True");
            parsedINIDataToBeSaved["Editor"].AddKey("ShowLineNumbers", "True");


            //保存文件
            parser.WriteFile("Setting.ini", parsedINIDataToBeSaved);
        }

        public static void SetMaxThreadCount(int count)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Thread"]["MaxThreadCount"] = count.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static int GetMaxThreadCount()
        {
            if (!File.Exists("Setting.ini"))
            {
                return -1;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return Int32.Parse(data["Thread"]["MaxThreadCount"]);
        }

        public static void SetTotalTimeoutLimitAnalyze(int count)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Thread"]["TotalTimeoutLimitAnalyze"] = count.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static int GetTotalTimeoutLimitAnalyze()
        {
            if (!File.Exists("Setting.ini"))
            {
                return -1;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return Int32.Parse(data["Thread"]["TotalTimeoutLimitAnalyze"]);
        }
        public static void SetPerTimeoutLimitAnalyze(int count)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Thread"]["PerTimeoutLimitAnalyze"] = count.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static int GetPerTimeoutLimitAnalyze()
        {
            if (!File.Exists("Setting.ini"))
            {
                return -1;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return Int32.Parse(data["Thread"]["PerTimeoutLimitAnalyze"]);
        }

        public static void SetTotalTimeoutLimitOutput(int count)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Thread"]["TotalTimeoutLimitOutput"] = count.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static int GetTotalTimeoutLimitOutput()
        {
            if (!File.Exists("Setting.ini"))
            {
                return -1;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return Int32.Parse(data["Thread"]["TotalTimeoutLimitOutput"]);
        }
        public static void SetPerTimeoutLimitOutput(int count)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Thread"]["PerTimeoutLimitOutput"] = count.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static int GetPerTimeoutLimitOutput()
        {
            if (!File.Exists("Setting.ini"))
            {
                return -1;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return Int32.Parse(data["Thread"]["PerTimeoutLimitOutput"]);
        }

        public static void SetFileSystemWatcherInvokeDalay(int time)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Thread"]["FileSystemWatcherInvokeDalay"] = time.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static int GetFileSystemWatcherInvokeDalay()
        {
            if (!File.Exists("Setting.ini"))
            {
                return -1;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return Int32.Parse(data["Thread"]["FileSystemWatcherInvokeDalay"]);
        }

        public static void SetFreshInterval(int time)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Thread"]["FreshInterval"] = time.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static int GetFreshInterval()
        {
            if (!File.Exists("Setting.ini"))
            {
                return -1;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return Int32.Parse(data["Thread"]["FreshInterval"]);
        }

        public static void SetWindowSize(String wndName, Point point)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Window"][$"{wndName}Width"] = point.X.ToString();
            data["Window"][$"{wndName}Height"] = point.Y.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static Point GetWindowSize(String wndName)
        {
            if (!File.Exists("Setting.ini"))
            {
                return new Point(-1, -1);
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return new Point(Double.Parse(data["Window"][$"{wndName}Width"]), Double.Parse(data["Window"][$"{wndName}Height"]));
        }

        public static void SetIsExecuteInSequence(bool isExecuteInSequence)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Window"]["IsExecuteInSequence"] = isExecuteInSequence.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static bool GetIsExecuteInSequence()
        {
            if (!File.Exists("Setting.ini"))
            {
                return false;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return bool.Parse(data["Window"]["IsExecuteInSequence"]);
        }

        public static void SetIsAutoOpen(bool isAutoOpen)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Window"]["IsAutoOpen"] = isAutoOpen.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static bool GetIsAutoOpen()
        {
            if (!File.Exists("Setting.ini"))
            {
                return false;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return bool.Parse(data["Window"]["IsAutoOpen"]);
        }

        public static void SetTheme(string theme)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Window"]["Theme"] = theme;
            parser.WriteFile("Setting.ini", data);
        }

        public static string GetTheme()
        {
            if (!File.Exists("Setting.ini"))
            {
                return "Light";
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return data["Window"]["Theme"];
        }

        public static void SetLanguage(String language)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Window"]["Language"] = language;
            parser.WriteFile("Setting.ini", data);
        }

        public static String GetLanguage()
        {
            if (!File.Exists("Setting.ini"))
            {
                return null;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return data["Window"]["Language"];
        }

        public static void SetBasePath(String basePath)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Value"]["DefaultBasePath"] = basePath;
            parser.WriteFile("Setting.ini", data);
        }

        public static String GetBasePath()
        {
            if (!File.Exists("Setting.ini"))
            {
                return null;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return data["Value"]["DefaultBasePath"];
        }

        public static void SetOutputPath(String output)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Value"]["DefaultOutputPath"] = output;
            parser.WriteFile("Setting.ini", data);
        }

        public static String GetOutputPath()
        {
            if (!File.Exists("Setting.ini"))
            {
                return null;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            string defaultOutputPath = data["Value"]["DefaultOutputPath"];
            if (String.IsNullOrWhiteSpace(defaultOutputPath))
            {
                defaultOutputPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            }
            return defaultOutputPath;
        }

        public static void SetOutputFileName(String outputFileName)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Value"]["DefaultOutputFileName"] = outputFileName;
            parser.WriteFile("Setting.ini", data);
        }

        public static String GetOutputFileName()
        {
            if (!File.Exists("Setting.ini"))
            {
                return null;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return data["Value"]["DefaultOutputFileName"];
        }

        public static void SetShowSpaces(bool showSpaces)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Editor"]["ShowSpaces"] = showSpaces.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static bool GetShowSpaces()
        {
            if (!File.Exists("Setting.ini"))
            {
                return true;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return bool.Parse(data["Editor"]["ShowSpaces"]);
        }

        public static void SetShowTabs(bool showTabs)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Editor"]["ShowTabs"] = showTabs.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static bool GetShowTabs()
        {
            if (!File.Exists("Setting.ini"))
            {
                return true;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return bool.Parse(data["Editor"]["ShowTabs"]);
        }

        public static void SetShowEndOfLine(bool showEndOfLine)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Editor"]["ShowEndOfLine"] = showEndOfLine.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static bool GetShowEndOfLine()
        {
            if (!File.Exists("Setting.ini"))
            {
                return true;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return bool.Parse(data["Editor"]["ShowEndOfLine"]);
        }

        public static void SetShowBoxForControlCharacters(bool showBoxForControlCharacters)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Editor"]["ShowBoxForControlCharacters"] = showBoxForControlCharacters.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static bool GetShowBoxForControlCharacters()
        {
            if (!File.Exists("Setting.ini"))
            {
                return true;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return bool.Parse(data["Editor"]["ShowBoxForControlCharacters"]);
        }

        public static void SetEnableHyperlinks(bool enableHyperlinks)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Editor"]["EnableHyperlinks"] = enableHyperlinks.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static bool GetEnableHyperlinks()
        {
            if (!File.Exists("Setting.ini"))
            {
                return true;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return bool.Parse(data["Editor"]["EnableHyperlinks"]);
        }

        public static void SetIndentationSize(int indentationSize)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Editor"]["IndentationSize"] = indentationSize.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static int GetIndentationSize()
        {
            if (!File.Exists("Setting.ini"))
            {
                return -1;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return Int32.Parse(data["Editor"]["IndentationSize"]);
        }

        public static void SetConvertTabsToSpaces(bool convertTabsToSpaces)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Editor"]["ConvertTabsToSpaces"] = convertTabsToSpaces.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static bool GetConvertTabsToSpaces()
        {
            if (!File.Exists("Setting.ini"))
            {
                return true;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return bool.Parse(data["Editor"]["ConvertTabsToSpaces"]);
        }

        public static void SetHighlightCurrentLine(bool highlightCurrentLine)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Editor"]["HighlightCurrentLine"] = highlightCurrentLine.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static bool GetHighlightCurrentLine()
        {
            if (!File.Exists("Setting.ini"))
            {
                return true;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return bool.Parse(data["Editor"]["HighlightCurrentLine"]);
        }

        public static void SetHideCursorWhileTyping(bool hideCursorWhileTyping)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Editor"]["HideCursorWhileTyping"] = hideCursorWhileTyping.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static bool GetHideCursorWhileTyping()
        {
            if (!File.Exists("Setting.ini"))
            {
                return true;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return bool.Parse(data["Editor"]["HideCursorWhileTyping"]);
        }

        public static void SetWordWrap(bool wordWrap)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Editor"]["WordWrap"] = wordWrap.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static bool GetWordWrap()
        {
            if (!File.Exists("Setting.ini"))
            {
                return true;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return bool.Parse(data["Editor"]["WordWrap"]);
        }

        public static void SetShowLineNumbers(bool showLineNumbers)
        {
            if (!File.Exists("Setting.ini"))
            {
                return;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            data["Editor"]["ShowLineNumbers"] = showLineNumbers.ToString();
            parser.WriteFile("Setting.ini", data);
        }

        public static bool GetShowLineNumbers()
        {
            if (!File.Exists("Setting.ini"))
            {
                return true;
            }

            FileIniDataParser parser = new FileIniDataParser();
            IniData data = parser.ReadFile("Setting.ini");
            return bool.Parse(data["Editor"]["ShowLineNumbers"]);
        }
    }
}
