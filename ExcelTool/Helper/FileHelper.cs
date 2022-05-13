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
            String paramStr = File.ReadAllText(".\\Params.txt");
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
    }
}
