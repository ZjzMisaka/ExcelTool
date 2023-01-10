using ClosedXML.Excel;
using CustomizableMessageBox;
using GlobalObjects;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using static CustomizableMessageBox.MessageBox;

namespace ExcelTool.Helper
{
    public static class FileHelper
    {
        public static void FileTraverse(bool isAuto, string folderPath, SheetExplainer sheetExplainer, List<string> filePathList)
        {
            if (string.IsNullOrEmpty(folderPath))
            {
                return;
            }

            DirectoryInfo dir = new DirectoryInfo(folderPath);
            try
            {
                if (!dir.Exists)
                    return;
                DirectoryInfo dirD = dir as DirectoryInfo;
                FileSystemInfo[] files = dirD.GetFileSystemInfos();
                foreach (FileSystemInfo FSys in files)
                {
                    System.IO.FileInfo fileInfo = FSys as System.IO.FileInfo;

                    if (fileInfo != null && (System.IO.Path.GetExtension(fileInfo.Name).Equals(".xlsx")))
                    {
                        if (sheetExplainer.fileNames.Key == FindingMethod.SAME)
                        {
                            foreach (string str in sheetExplainer.fileNames.Value)
                            {
                                if (fileInfo.Name.Equals(str))
                                {
                                    string fileName = System.IO.Path.Combine(fileInfo.DirectoryName, fileInfo.Name);
                                    if (File.Exists(fileName))
                                    {
                                        filePathList.Add(fileName);
                                    }
                                }
                            }
                        }
                        else if (sheetExplainer.fileNames.Key == FindingMethod.CONTAIN)
                        {
                            foreach (string str in sheetExplainer.fileNames.Value)
                            {
                                if (fileInfo.Name.Contains(str))
                                {
                                    string fileName = System.IO.Path.Combine(fileInfo.DirectoryName, fileInfo.Name);
                                    if (File.Exists(fileName))
                                    {
                                        filePathList.Add(fileName);
                                    }
                                }
                            }
                        }
                        else if (sheetExplainer.fileNames.Key == FindingMethod.REGEX)
                        {
                            foreach (string str in sheetExplainer.fileNames.Value)
                            {
                                Regex rgx = new Regex(str);
                                if (rgx.IsMatch(fileInfo.Name))
                                {
                                    string fileName = System.IO.Path.Combine(fileInfo.DirectoryName, fileInfo.Name);
                                    if (File.Exists(fileName))
                                    {
                                        filePathList.Add(fileName);
                                    }
                                }
                            }
                        }
                        else if (sheetExplainer.fileNames.Key == FindingMethod.ALL)
                        {
                            Regex rgx = new Regex("[\\s\\S]*.xls[xm]", RegexOptions.IgnoreCase);
                            if (rgx.IsMatch(fileInfo.Name))
                            {
                                string fileName = System.IO.Path.Combine(fileInfo.DirectoryName, fileInfo.Name);
                                if (File.Exists(fileName))
                                {
                                    filePathList.Add(fileName);
                                }
                            }
                        }
                    }
                    else
                    {
                        string pp = FSys.Name;
                        FileTraverse(isAuto, System.IO.Path.Combine(folderPath, FSys.ToString()), sheetExplainer, filePathList);
                    }
                }
            }
            catch (Exception ex)
            {
                if (!isAuto)
                {
                    CustomizableMessageBox.MessageBox.Show(new RefreshList { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, ex.Message, "ERROR", MessageBoxImage.Error);
                }
                else
                {
                    Logger.Error(ex.Message);
                }
                return;
            }
        }

        public static List<string> GetSheetExplainersList()
        {
            List<string> sheetExplainersList = Directory.GetFiles(".\\SheetExplainers", "*.json").ToList();
            sheetExplainersList.Insert(0, "");
            for (int i = 0; i < sheetExplainersList.Count; ++i)
            {
                string str = sheetExplainersList[i];
                sheetExplainersList[i] = str.Substring(str.LastIndexOf('\\') + 1);
                if (sheetExplainersList[i].Contains('.'))
                {
                    sheetExplainersList[i] = sheetExplainersList[i].Substring(0, sheetExplainersList[i].LastIndexOf('.'));
                }
            }
            return sheetExplainersList;
        }

        public static List<string> GetAnalyzersList()
        {
            List<string> analyzersList = Directory.GetFiles(".\\Analyzers", "*.json").ToList();
            analyzersList.Insert(0, "");
            for (int i = 0; i < analyzersList.Count; ++i)
            {
                string str = analyzersList[i];
                analyzersList[i] = str.Substring(str.LastIndexOf('\\') + 1);
                if (analyzersList[i].Contains('.'))
                {
                    analyzersList[i] = analyzersList[i].Substring(0, analyzersList[i].LastIndexOf('.'));
                }
            }
            return analyzersList;
        }

        public static List<string> GetRulesList()
        {
            List<string> rulesList = Directory.GetFiles(".\\Rules", "*.json").ToList();
            rulesList.Insert(0, "");
            for (int i = 0; i < rulesList.Count; ++i)
            {
                string str = rulesList[i];
                rulesList[i] = str.Substring(str.LastIndexOf('\\') + 1);
                if (rulesList[i].Contains('.'))
                {
                    rulesList[i] = rulesList[i].Substring(0, rulesList[i].LastIndexOf('.'));
                }
            }
            return rulesList;
        }

        public static List<string> GetParamsList()
        {
            if (!File.Exists(".\\Params.txt"))
            {
                FileStream fs = null;
                try
                {
                    fs = File.Create(".\\Params.txt");
                    fs.Close();
                }
                catch (Exception ex)
                {
                    CustomizableMessageBox.MessageBox.Show(new RefreshList { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, ex.Message, "ERROR", MessageBoxImage.Error);
                }
            }
            string paramStr = ParamHelper.DecodeToEscaped(File.ReadAllText(".\\Params.txt"));
            List<string> paramsList = paramStr.Split('\n').ToList<string>();
            List<string> newParamsList = new List<string>();
            foreach (string param in paramsList)
            {
                if (param.Trim() != "")
                {
                    newParamsList.Add(param.Trim());
                }
            }
            newParamsList.Insert(0, "");

            return newParamsList;
        }

        public static void DeleteCopiedDlls(List<string> copiedDllsList)
        {
            string arguments = "";
            foreach (string path in copiedDllsList)
            {
                arguments += path.Replace(" ", "|SPACE|") + " ";
            }

            if (arguments.Length > 0)
            {
                Process.Start("ExcelToolAfterClosed.exe", arguments);
            }
        }

        public static void SaveWorkbook(bool isAuto, string filePath, XLWorkbook workbook, bool isAutoOpen, bool isExecuteInSequence)
        {
            bool saveResult = false;
            SaveFile(isAuto, isExecuteInSequence, filePath, workbook, out saveResult);
            if (!saveResult)
            {
                string fileNotSavedStr = $"{Application.Current.FindResource("FileNotSaved")}: \n{filePath}";
                if (!isAuto && !isExecuteInSequence)
                {
                    Logger.Error(fileNotSavedStr);
                    CustomizableMessageBox.MessageBox.Show(new RefreshList { new ButtonSpacer(), Application.Current.FindResource(Application.Current.FindResource("Ok").ToString()).ToString() }, fileNotSavedStr, Application.Current.FindResource("Info").ToString());
                }
                else
                {
                    Logger.Error(fileNotSavedStr);
                }

                return;
            }

            string fileSavedStr = $"{Application.Current.FindResource("FileSaved").ToString()}: \n{filePath}";
            Logger.Info(fileSavedStr);
            if (!isAuto && isAutoOpen == false && !isExecuteInSequence)
            {
                Button btnOpenFile = new Button();
                btnOpenFile.Height = 30;
                btnOpenFile.HorizontalAlignment = HorizontalAlignment.Stretch;
                btnOpenFile.Margin = new Thickness(5);
                btnOpenFile.Content = Application.Current.FindResource("OpenFile").ToString();
                btnOpenFile.Click += (s, ee) =>
                {
                    System.Diagnostics.Process.Start(filePath);
                    CustomizableMessageBox.MessageBox.CloseNow();
                };

                Button btnOpenPath = new Button();
                btnOpenPath.Height = 30;
                btnOpenPath.HorizontalAlignment = HorizontalAlignment.Stretch;
                btnOpenPath.Margin = new Thickness(5);
                btnOpenPath.Content = Application.Current.FindResource("OpenPath").ToString();
                btnOpenPath.Click += (s, ee) =>
                {
                    System.Diagnostics.Process.Start("Explorer", $"/e,/select,{filePath.Replace("/", "\\")}");
                    CustomizableMessageBox.MessageBox.CloseNow();
                };

                Button btnClose = new Button();
                btnClose.Height = 30;
                btnClose.HorizontalAlignment = HorizontalAlignment.Stretch;
                btnClose.Margin = new Thickness(5);
                btnClose.Content = Application.Current.FindResource("Close").ToString();
                btnClose.Click += (s, ee) =>
                {
                    CustomizableMessageBox.MessageBox.CloseNow();
                };
                CustomizableMessageBox.MessageBox.Show(new RefreshList { btnClose, new ButtonSpacer(40), btnOpenFile, btnOpenPath }, fileSavedStr, Application.Current.FindResource("Info").ToString());
            }
            else
            {
                if (isAutoOpen == true)
                {
                    Logger.Info(Application.Current.FindResource("AutoOpened").ToString());
                    Process.Start(filePath);
                }
            }
        }

        private static void SaveFile(bool isAuto, bool isExecuteInSequence, string filePath, XLWorkbook workbook, out bool result)
        {
            try
            {
                workbook.SaveAs(filePath);
                result = true;
            }
            catch (Exception e)
            {
                int res = 2;
                if (!isAuto && !isExecuteInSequence)
                {
                    res = CustomizableMessageBox.MessageBox.Show(new RefreshList { new ButtonSpacer(), Application.Current.FindResource("Yes").ToString(), Application.Current.FindResource("No").ToString() }, $"{Application.Current.FindResource("FailedToSaveFile").ToString()} \n{e.Message}", Application.Current.FindResource("Error").ToString(), MessageBoxImage.Question);
                }
                if (res == 2)
                {
                    result = false;
                }
                else
                {
                    SaveFile(isAuto, isExecuteInSequence, filePath, workbook, out result);
                }
            }
        }

        public static void CheckAndCreateFolders()
        {
            try
            {
                if (!Directory.Exists(".\\Analyzers"))
                {
                    Directory.CreateDirectory(".\\Analyzers");
                }
                if (!Directory.Exists(".\\SheetExplainers"))
                {
                    Directory.CreateDirectory(".\\SheetExplainers");
                }
                if (!Directory.Exists(".\\Rules"))
                {
                    Directory.CreateDirectory(".\\Rules");
                }
                if (!Directory.Exists(".\\Dlls"))
                {
                    Directory.CreateDirectory(".\\Dlls");
                }
            }
            catch (Exception ex)
            {
                CustomizableMessageBox.MessageBox.Show(new RefreshList { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, $"{Application.Current.FindResource("FailedToCreateANewFolder").ToString()}\n{ex.Message}", Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
                return;
            }
        }

        public static string SavaSheetExplainerJson(string fileName, SheetExplainer sheetExplainer)
        {
            string json = JsonConvert.SerializeObject(sheetExplainer);

            string fullFileName = $".\\SheetExplainers\\{fileName}.json";
            string tempFileName = fileName;

            int i = 1;
            while (File.Exists(fullFileName))
            {
                tempFileName = $"{fileName}_{i}";
                fullFileName = $".\\SheetExplainers\\{tempFileName}.json";
                ++i;
            }

            FileStream fs;
            try
            {
                fs = File.Create(fullFileName);
                fs.Close();
                StreamWriter sw = File.CreateText(fullFileName);
                sw.Write(json);
                sw.Flush();
                sw.Close();
            }
            catch (Exception ex)
            {
                CustomizableMessageBox.MessageBox.Show(new RefreshList { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, ex.Message, Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
            }

            return tempFileName;
        }

        public static string SavaAnalyzerJson(string fileName, Analyzer analyzer)
        {
            string json = JsonConvert.SerializeObject(analyzer);

            string fullFileName = $".\\Analyzers\\{fileName}.json";
            string tempFileName = fileName;

            int i = 1;
            while (File.Exists(fullFileName))
            {
                tempFileName = $"{fileName}_{i}";
                fullFileName = $".\\Analyzers\\{tempFileName}.json";
                ++i;
            }

            FileStream fs;
            try
            {
                fs = File.Create(fullFileName);
                fs.Close();
                StreamWriter sw = File.CreateText(fullFileName);
                sw.Write(json);
                sw.Flush();
                sw.Close();
            }
            catch (Exception ex)
            {
                CustomizableMessageBox.MessageBox.Show(new RefreshList { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, ex.Message, Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
            }

            return tempFileName;
        }

        public static string SavaRunningRuleJson(string fileName, RunningRule runningRule)
        {
            string json = JsonConvert.SerializeObject(runningRule);

            string fullFileName = $".\\Rules\\{fileName}.json";
            string tempFileName = fileName;

            int i = 1;
            while (File.Exists(fullFileName))
            {
                tempFileName = $"{fileName}_{i}";
                fullFileName = $".\\Rules\\{tempFileName}.json";
                ++i;
            }

            FileStream fs;
            try
            {
                fs = File.Create(fullFileName);
                fs.Close();
                StreamWriter sw = File.CreateText(fullFileName);
                sw.Write(json);
                sw.Flush();
                sw.Close();
            }
            catch (Exception ex)
            {
                CustomizableMessageBox.MessageBox.Show(new RefreshList { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, ex.Message, Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
            }

            return tempFileName;
        }
    }
}
