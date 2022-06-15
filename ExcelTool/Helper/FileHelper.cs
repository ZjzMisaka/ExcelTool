using ClosedXML.Excel;
using CustomizableMessageBox;
using GlobalObjects;
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

namespace ExcelTool.Helper
{
    public static class FileHelper
    {
        public static void FileTraverse(bool isAuto, String folderPath, SheetExplainer sheetExplainer, List<String> filePathList)
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
                            foreach (String str in sheetExplainer.fileNames.Value)
                            {
                                if (fileInfo.Name.Equals(str))
                                {
                                    string fileName = System.IO.Path.Combine(fileInfo.DirectoryName, fileInfo.Name);
                                    filePathList.Add(fileName);
                                }
                            }
                        }
                        else if (sheetExplainer.fileNames.Key == FindingMethod.CONTAIN)
                        {
                            foreach (String str in sheetExplainer.fileNames.Value)
                            {
                                if (fileInfo.Name.Contains(str))
                                {
                                    string fileName = System.IO.Path.Combine(fileInfo.DirectoryName, fileInfo.Name);
                                    filePathList.Add(fileName);
                                }
                            }
                        }
                        else if (sheetExplainer.fileNames.Key == FindingMethod.REGEX)
                        {
                            foreach (String str in sheetExplainer.fileNames.Value)
                            {
                                Regex rgx = new Regex(str);
                                if (rgx.IsMatch(fileInfo.Name))
                                {
                                    string fileName = System.IO.Path.Combine(fileInfo.DirectoryName, fileInfo.Name);
                                    filePathList.Add(fileName);
                                }
                            }
                        }
                        else if (sheetExplainer.fileNames.Key == FindingMethod.ALL)
                        {
                            Regex rgx = new Regex("[\\s\\S]*.xls[xm]", RegexOptions.IgnoreCase);
                            if (rgx.IsMatch(fileInfo.Name))
                            {
                                string fileName = System.IO.Path.Combine(fileInfo.DirectoryName, fileInfo.Name);
                                filePathList.Add(fileName);
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
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
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
            List<String> sheetExplainersList = Directory.GetFiles(".\\SheetExplainers", "*.json").ToList();
            sheetExplainersList.Insert(0, "");
            for (int i = 0; i < sheetExplainersList.Count; ++i)
            {
                String str = sheetExplainersList[i];
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
            List<String> analyzersList = Directory.GetFiles(".\\Analyzers", "*.json").ToList();
            analyzersList.Insert(0, "");
            for (int i = 0; i < analyzersList.Count; ++i)
            {
                String str = analyzersList[i];
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
            List<String> rulesList = Directory.GetFiles(".\\Rules", "*.json").ToList();
            rulesList.Insert(0, "");
            for (int i = 0; i < rulesList.Count; ++i)
            {
                String str = rulesList[i];
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
                    StreamWriter sw = File.CreateText(".\\Params.txt");
                    sw.Write("key1:value1|key2:value2|key3:value3");
                    sw.Flush();
                    sw.Close();
                }
                catch (Exception ex)
                {
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            String paramStr = ParamHelper.Decode1(File.ReadAllText(".\\Params.txt"));
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

        public static void SaveWorkbook(bool isAuto, string filePath, XLWorkbook workbook, bool isAutoOpen)
        {
            bool saveResult = false;
            SaveFile(isAuto, filePath, workbook, out saveResult);
            if (!saveResult)
            {
                string fileNotSavedStr = $"{Application.Current.FindResource("FileNotSaved").ToString()}: \n{filePath}";
                if (!isAuto)
                {
                    Logger.Error(fileNotSavedStr);
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), Application.Current.FindResource(Application.Current.FindResource("Ok").ToString()).ToString() }, fileNotSavedStr, Application.Current.FindResource("Info").ToString());
                }
                else
                {
                    Logger.Error(fileNotSavedStr);
                }

                return;
            }

            string fileSavedStr = $"{Application.Current.FindResource("FileSaved").ToString()}: \n{filePath}";
            Logger.Info(fileSavedStr);
            if (!isAuto && isAutoOpen == false)
            {
                Button btnOpenFile = new Button();
                btnOpenFile.Style = GlobalObjects.GlobalObjects.GetBtnStyle();
                btnOpenFile.Content = Application.Current.FindResource("OpenFile").ToString();
                btnOpenFile.Click += (s, ee) =>
                {
                    System.Diagnostics.Process.Start(filePath);
                    CustomizableMessageBox.MessageBox.CloseNow();
                };

                Button btnOpenPath = new Button();
                btnOpenPath.Style = GlobalObjects.GlobalObjects.GetBtnStyle();
                btnOpenPath.Content = Application.Current.FindResource("OpenPath").ToString();
                btnOpenPath.Click += (s, ee) =>
                {
                    System.Diagnostics.Process.Start("Explorer", $"/e,/select,{filePath.Replace("/", "\\")}");
                    CustomizableMessageBox.MessageBox.CloseNow();
                };

                Button btnClose = new Button();
                btnClose.Style = GlobalObjects.GlobalObjects.GetBtnStyle();
                btnClose.Content = Application.Current.FindResource("Close").ToString();
                btnClose.Click += (s, ee) =>
                {
                    CustomizableMessageBox.MessageBox.CloseNow();
                };
                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { btnClose, new ButtonSpacer(40), btnOpenFile, btnOpenPath }, fileSavedStr, "OK");
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

        private static void SaveFile(bool isAuto, string filePath, XLWorkbook workbook, out bool result)
        {
            try
            {
                workbook.SaveAs(filePath);
                result = true;
            }
            catch (Exception e)
            {
                int res = 2;
                if (!isAuto)
                {
                    res = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), Application.Current.FindResource("Yes").ToString(), Application.Current.FindResource("No").ToString() }, $"{Application.Current.FindResource("FailedToSaveFile").ToString()} \n{e.Message}", Application.Current.FindResource("Error").ToString(), MessageBoxImage.Question);
                }
                if (res == 2)
                {
                    result = false;
                }
                else
                {
                    SaveFile(isAuto, filePath, workbook, out result);
                }
            }
        }
    }
}
