using Amib.Threading;
using ClosedXML.Excel;
using GlobalObjects;
using Microsoft.CSharp;
using Newtonsoft.Json;
using System;
using System.CodeDom.Compiler;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using CustomizableMessageBox;
using System.Xml;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Document;

namespace ExcelTool
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private enum ReadFileReturnType { ANALYZER, RESULT };
        private Boolean isStop;
        private SmartThreadPool smartThreadPoolAnalyze = null;
        private SmartThreadPool smartThreadPoolOutput = null;
        private ConcurrentDictionary<string, long> currentAnalizingDictionary;
        private ConcurrentDictionary<string, long> currentOutputtingDictionary;
        private ConcurrentDictionary<ConcurrentDictionary<ResultType, Object>, Analyzer> resultList;
        private int analyzeSheetInvokeCount;
        private int setResultInvokeCount;
        private int totalTimeoutLimitAnalyze;
        private int perTimeoutLimitAnalyze;
        private int totalTimeoutLimitOutput;
        private int perTimeoutLimitOutput;
        private String errorStr;
        public MainWindow()
        {
            InitializeComponent();

            currentAnalizingDictionary = new ConcurrentDictionary<string, long>();
            currentOutputtingDictionary = new ConcurrentDictionary<string, long>();
            resultList = new ConcurrentDictionary<ConcurrentDictionary<ResultType, Object>, Analyzer>();
            analyzeSheetInvokeCount = 0;
            setResultInvokeCount = 0;
            errorStr = "";

            isStop = false;
        }

        private void WindowLoaded(object sender, RoutedEventArgs e)
        {
            IniHelper.CheckAndCreateIniFile();

            double width = IniHelper.GetWindowSize(this.Name.Substring(2)).X;
            double height = IniHelper.GetWindowSize(this.Name.Substring(2)).Y;
            if (width > 0)
            {
                this.Width = width;
            }

            if (height > 0)
            {
                this.Height = height;
            }

            tb_base_path.Text = IniHelper.GetBasePath();
            tb_output_path.Text = IniHelper.GetOutputPath();
            tb_output_name.Text = IniHelper.GetOutputFileName();

            totalTimeoutLimitAnalyze = IniHelper.GetTotalTimeoutLimitAnalyze();
            perTimeoutLimitAnalyze = IniHelper.GetPerTimeoutLimitAnalyze();
            totalTimeoutLimitOutput = IniHelper.GetTotalTimeoutLimitOutput();
            perTimeoutLimitOutput = IniHelper.GetPerTimeoutLimitOutput();

            LoadFiles();


            IHighlightingDefinition ch;
            using (Stream s = new FileStream(@"ParamHighlighting.xshd", FileMode.Open))
            {
                if (s == null)
                    throw new InvalidOperationException("Could not find embedded resource");
                using (XmlReader reader = new XmlTextReader(s))
                {
                    //解析xshd
                    ch = ICSharpCode.AvalonEdit.Highlighting.Xshd.HighlightingLoader.Load(reader, HighlightingManager.Instance);
                }
            }
            //注册数据
            HighlightingManager.Instance.RegisterHighlighting("param", new string[] { }, ch);
            te_params.SyntaxHighlighting = HighlightingManager.Instance.GetDefinition("param");
        }

        private void CbSheetExplainersSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cb_sheetexplainers.SelectedIndex == 0) 
            {
                return;
            }
            te_sheetexplainers.Text += $"{cb_sheetexplainers.SelectedItem}\n";
            cb_sheetexplainers.SelectedIndex = 0;
        }
        private void CbAnalyzersSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cb_analyzers.SelectedIndex == 0)
            {
                return;
            }
            te_analyzers.Text += $"{cb_analyzers.SelectedItem}\n";
            cb_analyzers.SelectedIndex = 0;
        }

        private void BtnOpenSheetExplainerEditorClick(object sender, RoutedEventArgs e)
        {
            SheetExplainerEditor sheetExplainerEditor = new SheetExplainerEditor();
            sheetExplainerEditor.Show();
        }

        private void BtnOpenAnalyzerEditorClick(object sender, RoutedEventArgs e)
        {
            AnalyzerEditor analyzerEditor = new AnalyzerEditor();
            analyzerEditor.Show();
        }

        private void CbSheetExplainersPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            LoadFiles();
        }

        private void CbAnalyzersPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            LoadFiles();
        }

        private void DropPath(object sender, DragEventArgs e)
        {
            (sender as TextBox).Text = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
        }

        private void DragOverPath(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }

        private void DropFile(object sender, DragEventArgs e)
        {
            (sender as TextBox).Text = System.IO.Path.GetFileNameWithoutExtension(((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString());
        }

        private void DragOverFile(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }

        private void SelectPath(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog openFileDialog = new System.Windows.Forms.FolderBrowserDialog();

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if ((sender as Button).Name == "btn_select_output_path")
                {
                    tb_base_path.Text = openFileDialog.SelectedPath;
                }
                else
                {
                    tb_output_path.Text = openFileDialog.SelectedPath;
                }
            }
        }

        private void SelectName(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog fileDialog = new System.Windows.Forms.OpenFileDialog();
            fileDialog.Multiselect = false;
            fileDialog.Title = "请选择文件";
            fileDialog.Filter = "Excel文件|*.xlsx;*.xlsm;*.xls|All files(*.*)|*.*";
            if (fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string[] names = fileDialog.FileNames;
                tb_output_name.Text = System.IO.Path.GetFileNameWithoutExtension(names[0]);

            }
        }

        private void OpenPath(object sender, RoutedEventArgs e)
        {
            string resPath = tb_output_path.Text.Replace("\\", "/");
            string filePath;
            if (tb_output_path.Text.EndsWith("/"))
            {
                filePath = $"{resPath}{tb_output_name.Text}.xlsx";
            }
            else
            {
                filePath = $"{resPath}/{tb_output_name.Text}.xlsx";
            }
            if (File.Exists(filePath))
            {
                System.Diagnostics.Process.Start("Explorer", $"/e,/select,{filePath.Replace("/", "\\")}");
            }
            else
            {
                filePath = filePath.Replace("/", "\\");
                System.Diagnostics.Process.Start("Explorer", $"{filePath.Substring(0, filePath.LastIndexOf('\\'))}");
            }
        }
        private void Stop(object sender, RoutedEventArgs e)
        {
            Stop();
        }
        private void Stop()
        {
            isStop = true;
        }
        private async void Start(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.Invoke(() =>
            {
                btn_start.IsEnabled = false;
                btn_stop.IsEnabled = true;
            });

            isStop = false;

            analyzeSheetInvokeCount = 0;
            setResultInvokeCount = 0;
            errorStr = "";
            resultList = new ConcurrentDictionary<ConcurrentDictionary<ResultType, Object>, Analyzer>();
            GlobalObjects.GlobalObjects.SetGlobalParam(null);

            List<String> sheetExplainersList = te_sheetexplainers.Text.Split('\n').Where(str => str.Trim() != "").ToList();
            List<String> analyzersList = te_analyzers.Text.Split('\n').Where(str => str.Trim() != "").ToList();
            if(sheetExplainersList.Count != analyzersList.Count || sheetExplainersList.Count == 0)
            {
                this.Dispatcher.Invoke(() =>
                {
                    btn_start.IsEnabled = true;
                    btn_stop.IsEnabled = false;
                });

                isStop = false;
                return;
            }

            te_params.Text = $"{te_params.Text.Replace("\n\r", "").Replace("\r\n", "").Replace("\n", "").Replace("\r", "")}";
            string paramStr = te_params.Text;
            string[] splitedParams = paramStr.Split('|');
            Dictionary<string, string> paramDic = new Dictionary<string, string>();
            foreach (string param in splitedParams)
            {
                string[] kv = param.Split(':');
                if (kv.Length == 2)
                {
                    paramDic.Add(kv[0], kv[1]);
                }
            }


            STPStartInfo stpAnalyze = new STPStartInfo();
            stpAnalyze.CallToPostExecute = CallToPostExecute.WhenWorkItemNotCanceled;
            stpAnalyze.PostExecuteWorkItemCallback = delegate (IWorkItemResult wir)
            {
                ConcurrentDictionary<ReadFileReturnType, Object> methodResult = null;
                try
                {
                    methodResult = (ConcurrentDictionary<ReadFileReturnType, Object>)wir.GetResult();
                }
                catch (Exception ex)
                {
                    methodResult[ReadFileReturnType.RESULT] = new ConcurrentDictionary<ResultType, Object>();
                    ((ConcurrentDictionary<ResultType, Object>)methodResult[ReadFileReturnType.RESULT]).AddOrUpdate(ResultType.MESSAGE, ex.Message, (key, oldValue) => null);
                }

                ConcurrentDictionary<ResultType, Object> result = ((ConcurrentDictionary<ResultType, Object>)methodResult[ReadFileReturnType.RESULT]);

                long value;
                currentAnalizingDictionary.TryRemove((String)result[ResultType.FILEPATH], out value);

                resultList.AddOrUpdate(result, (Analyzer)methodResult[ReadFileReturnType.ANALYZER], (key, oldValue) => null);
            };
            RenewSmartThreadPoolAnalyze(stpAnalyze);

            STPStartInfo stpOutput = new STPStartInfo();
            stpOutput.CallToPostExecute = CallToPostExecute.WhenWorkItemNotCanceled;
            stpOutput.PostExecuteWorkItemCallback = delegate (IWorkItemResult wir)
            {
                Object[] methodResult = null;
                try
                {
                    methodResult = (Object[])wir.GetResult();
                }
                catch (Exception ex)
                {
                    methodResult = null;
                }

                long value;
                currentOutputtingDictionary.TryRemove((String)((ConcurrentDictionary<ResultType, Object>)methodResult[1])[ResultType.FILEPATH], out value);
            };
            RenewSmartThreadPoolOutput(stpOutput);


            List<SheetExplainer> sheetExplainers = new List<SheetExplainer>();
            List<Analyzer> analyzers = new List<Analyzer>();
            foreach (string name in sheetExplainersList)
            {
                SheetExplainer sheetExplainer = null;
                try
                {
                    sheetExplainer = JsonConvert.DeserializeObject<SheetExplainer>(File.ReadAllText($".\\SheetExplainers\\{name}.json"));
                    if (sheetExplainer.relativePathes == null || sheetExplainer.relativePathes.Count == 0)
                    {
                        sheetExplainer.relativePathes = new List<string>();
                        sheetExplainer.relativePathes.Add("");
                    }
                }
                catch (Exception ex)
                {
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                    this.Dispatcher.Invoke(() =>
                    {
                        btn_start.IsEnabled = true;
                        btn_stop.IsEnabled = false;
                    });
                    isStop = false;
                    return;
                }
                sheetExplainers.Add(sheetExplainer);
            }
            foreach (string name in analyzersList)
            {
                Analyzer analyzer = null;
                try
                {
                    analyzer = JsonConvert.DeserializeObject<Analyzer>(File.ReadAllText($".\\Analyzers\\{name}.json"));
                }
                catch (Exception ex)
                {
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                    this.Dispatcher.Invoke(() =>
                    {
                        btn_start.IsEnabled = true;
                        btn_stop.IsEnabled = false;
                    });
                    isStop = false;
                    return;
                }
                analyzers.Add(analyzer);
            }

            int totalCount = 0;
            for(int i = 0; i < sheetExplainersList.Count; ++i)
            {
                SheetExplainer sheetExplainer = sheetExplainers[i];
                Analyzer analyzer = analyzers[i];

                CompilerResults cresult = GetCresult(analyzer);

                foreach (String str in sheetExplainer.relativePathes) 
                {
                    List<String> filePathList = new List<String>();
                    String basePath = Path.Combine(tb_base_path.Text.Trim(), str.Trim());

                    FileTraverse(basePath, sheetExplainer, filePathList);
                    totalCount += filePathList.Count;

                    int filesCount = filePathList.Count;
                    foreach (String filePath in filePathList)
                    { 
                        List<String> pathSplit = filePath.Split('\\').ToList<string>();
                        String fileName = pathSplit[pathSplit.Count - 1];
                        fileName = fileName.Substring(0, fileName.LastIndexOf('.'));
                        List<object> readFileParams = new List<object>();
                        readFileParams.Add(filePath);
                        readFileParams.Add(fileName);
                        readFileParams.Add(sheetExplainer);
                        readFileParams.Add(analyzer);
                        readFileParams.Add(paramDic);
                        readFileParams.Add(cresult);
                        smartThreadPoolAnalyze.QueueWorkItem(new Func<List<object>, Object>(ReadFile), readFileParams);
                    }
                }
            }

            long startSs = GetNowSs();
            smartThreadPoolAnalyze.Join();
            while (smartThreadPoolAnalyze.CurrentWorkItemsCount > 0)
            {
                long nowSs = GetNowSs();
                long totalTimeCostSs = nowSs - startSs;
                if (isStop || totalTimeCostSs >= totalTimeoutLimitAnalyze)
                {
                    smartThreadPoolAnalyze.Dispose();
                    this.Dispatcher.Invoke(() =>
                    {
                        btn_start.IsEnabled = true;
                        btn_stop.IsEnabled = false;
                    });
                    if (totalTimeCostSs >= totalTimeoutLimitAnalyze)
                    {
                        CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "OK" }, $"Total time out. \n{totalTimeoutLimitAnalyze / 1000.0}(s)", "error", MessageBoxImage.Error);
                    }
                    SetErrorLog();
                    return;
                }

                StringBuilder sb = new StringBuilder();
                sb.Append("Analyzing...");

                ICollection<string> keys = currentAnalizingDictionary.Keys;
                foreach (string key in keys)
                {
                    long value = new long();
                    if (currentAnalizingDictionary.TryGetValue(key, out value))
                    {
                        long timeCostSs = GetNowSs() - currentAnalizingDictionary[key];
                        if (timeCostSs >= perTimeoutLimitAnalyze)
                        {
                            smartThreadPoolAnalyze.Dispose();
                            this.Dispatcher.Invoke(() =>
                            {
                                btn_start.IsEnabled = true;
                                btn_stop.IsEnabled = false;
                            });
                            CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "OK" }, $"{key}\nTime out. \n{perTimeoutLimitAnalyze / 1000.0}(s)", "error", MessageBoxImage.Error);
                            return;
                        }
                        sb.Append(key.Substring(key.LastIndexOf('\\') + 1)).Append("(").Append((timeCostSs / 1000.0).ToString("0.0")).Append(")");
                    }
                }

                await Task.Run(() =>
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        l_process.Content = $"{smartThreadPoolAnalyze.CurrentWorkItemsCount}/{totalCount} -- ActiveThreads: {smartThreadPoolAnalyze.ActiveThreads} -- InUseThreads: {smartThreadPoolAnalyze.InUseThreads}";
                        tb_log.Text = $"{sb.ToString()}";
                    });
                });
                Thread.Sleep(100);
            }
            smartThreadPoolAnalyze.Dispose();

            // 输出结果
            using (var workbook = new XLWorkbook())
            {
                string resPath = tb_output_path.Text.Replace("\\", "/");
                string filePath;
                if (tb_output_path.Text.EndsWith("/"))
                {
                    filePath = $"{resPath}{tb_output_name.Text}.xlsx";
                }
                else
                {
                    filePath = $"{resPath}/{tb_output_name.Text}.xlsx";
                }

                foreach (ConcurrentDictionary<ResultType, Object> result in resultList.Keys)
                {
                    List<object> setResultParams = new List<object>();
                    setResultParams.Add(workbook);
                    setResultParams.Add(result);
                    setResultParams.Add(resultList[result]);
                    setResultParams.Add(resultList.Count);
                    setResultParams.Add(paramDic);
                    smartThreadPoolOutput.QueueWorkItem(new Func<List<object>, Object[]>(SetResult), setResultParams);
                }
                startSs = GetNowSs();
                smartThreadPoolOutput.Join();
                while (smartThreadPoolOutput.CurrentWorkItemsCount > 0)
                {
                    long nowSs = GetNowSs();
                    long totalTimeCostSs = nowSs - startSs;
                    if (isStop || totalTimeCostSs >= totalTimeoutLimitOutput)
                    {
                        smartThreadPoolOutput.Dispose();
                        this.Dispatcher.Invoke(() =>
                        {
                            btn_start.IsEnabled = true;
                            btn_stop.IsEnabled = false;
                        });
                        if (totalTimeCostSs >= totalTimeoutLimitOutput)
                        {
                            CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "OK" }, $"Total time out. \n{totalTimeoutLimitOutput / 1000.0}(s)", "error", MessageBoxImage.Error);
                        }
                        SetErrorLog();
                        return;
                    }

                    StringBuilder sb = new StringBuilder();
                    sb.Append("Outputting...");

                    ICollection<string> keys = currentOutputtingDictionary.Keys;
                    foreach (string key in keys)
                    {
                        long value = new long();
                        if (currentOutputtingDictionary.TryGetValue(key, out value))
                        {
                            long timeCostSs = GetNowSs() - currentOutputtingDictionary[key];
                            if (timeCostSs >= perTimeoutLimitAnalyze)
                            {
                                smartThreadPoolOutput.Dispose();
                                this.Dispatcher.Invoke(() =>
                                {
                                    btn_start.IsEnabled = true;
                                    btn_stop.IsEnabled = false;
                                });
                                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "OK" }, $"{key}\nTime out. \n{perTimeoutLimitAnalyze / 1000.0}(s)", "error", MessageBoxImage.Error);
                                return;
                            }
                            sb.Append(key.Substring(key.LastIndexOf('\\') + 1)).Append("(").Append((timeCostSs / 1000.0).ToString("0.0")).Append(")");
                        }
                    }

                    await Task.Run(() =>
                    {
                        this.Dispatcher.Invoke(() =>
                        {
                            l_process.Content = $"{smartThreadPoolOutput.CurrentWorkItemsCount}/{totalCount} -- ActiveThreads: {smartThreadPoolOutput.ActiveThreads} -- InUseThreads: {smartThreadPoolOutput.InUseThreads}";
                            tb_log.Text = $"{sb.ToString()}";
                        });
                    });
                    Thread.Sleep(100);
                }
                smartThreadPoolOutput.Dispose();

                if (SetErrorLog())
                {
                    btn_start.IsEnabled = true;
                    btn_stop.IsEnabled = false;
                    return;
                }
                else
                {
                    tb_log.Text = "";
                }

                String paramStrFromFile = File.ReadAllText(".\\Params.txt");
                List<string> paramsList = paramStrFromFile.Split('\n').ToList<string>();
                List<string> newParamsList = new List<string>();
                foreach (string param in paramsList)
                {
                    if (param.Trim() != "")
                    {
                        newParamsList.Add(param.Trim());
                    }
                }
                paramsList = newParamsList;
                if (paramsList.Count >= 5 && !paramsList.Contains(paramStr))
                {
                    paramsList.RemoveAt(4);
                }
                if (paramsList.Contains(paramStr))
                {
                    paramsList.Remove(paramStr);
                }
                paramsList.Insert(0, paramStr);

                try
                {
                    if (!File.Exists(".\\Params.txt"))
                    {
                        FileStream fs = null;
                        fs = File.Create(".\\Params.txt");
                        fs.Close();
                    }
                    StreamWriter sw = File.CreateText(".\\Params.txt");
                    foreach (string param in paramsList)
                    {
                        sw.WriteLine(param);
                    }
                    sw.Flush();
                    sw.Close();
                }
                catch (Exception ex)
                {
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                bool saveResult = false;
                SaveFile(filePath, workbook, out saveResult);
                if (!saveResult)
                {
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "OK" }, "file not saved.", "info");

                    btn_start.IsEnabled = true;
                    btn_stop.IsEnabled = false;
                    return;
                }

                Button btnOpenFile = new Button();
                btnOpenFile.Style = GlobalObjects.GlobalObjects.GetBtnStyle();
                btnOpenFile.Content = "open file";
                btnOpenFile.Click += (s, ee) =>
                {
                    System.Diagnostics.Process.Start(filePath);
                    CustomizableMessageBox.MessageBox.CloseTimer.CloseNow();
                };

                Button btnOpenPath = new Button();
                btnOpenPath.Style = GlobalObjects.GlobalObjects.GetBtnStyle();
                btnOpenPath.Content = "open path";
                btnOpenPath.Click += (s, ee) =>
                {
                    System.Diagnostics.Process.Start("Explorer", $"/e,/select,{filePath.Replace("/", "\\")}");
                    CustomizableMessageBox.MessageBox.CloseTimer.CloseNow();
                };

                Button btnClose = new Button();
                btnClose.Style = GlobalObjects.GlobalObjects.GetBtnStyle();
                btnClose.Content = "close";
                btnClose.Click += (s, ee) =>
                {
                    CustomizableMessageBox.MessageBox.CloseTimer.CloseNow();
                };

                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { btnClose, new ButtonSpacer(40), btnOpenFile, btnOpenPath }, "ファイルを保存しました。", "OK");
            }
                
            btn_start.IsEnabled = true;
            btn_stop.IsEnabled = false;
        }

        private void SaveFile(string filePath, XLWorkbook workbook, out bool result)
        {
            try
            {
                workbook.SaveAs(filePath);
                result = true;
            }
            catch (Exception e)
            {
                int res = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "YES", "NO" }, $"保存失败 重试? \n{e.Message}", "error", MessageBoxImage.Question);
                if (res == 2)
                {
                    result = false;
                }
                else
                {
                    SaveFile(filePath, workbook, out result);
                }
            }
        }

        private ConcurrentDictionary<ReadFileReturnType, Object> ReadFile(List<object> readFileParams)
        {
            String filePath = (String)readFileParams[0];
            String fileName = (String)readFileParams[1];
            SheetExplainer sheetExplainer = (SheetExplainer)readFileParams[2];
            Analyzer analyzer = (Analyzer)readFileParams[3];
            Dictionary<string, string> paramDic = (Dictionary<string, string>)readFileParams[4];
            CompilerResults cresult = (CompilerResults)readFileParams[5];

            ConcurrentDictionary<ReadFileReturnType, Object> methodResult = new ConcurrentDictionary<ReadFileReturnType, object>();
            methodResult.AddOrUpdate(ReadFileReturnType.ANALYZER, analyzer, (key, oldValue) => null);

            ConcurrentDictionary<ResultType, Object> result = new ConcurrentDictionary<ResultType, Object>();
            result.AddOrUpdate(ResultType.FILEPATH, filePath, (key, oldValue) => null);
            result.AddOrUpdate(ResultType.FILENAME, fileName, (key, oldValue) => null);
            currentAnalizingDictionary.AddOrUpdate(filePath, GetNowSs(), (key, oldValue) => GetNowSs());

            if (filePath.Contains("~$"))
            {
                result.AddOrUpdate(ResultType.MESSAGE, "文件正在被使用中", (key, oldValue) => null);
                methodResult.AddOrUpdate(ReadFileReturnType.RESULT, result, (key, oldValue) => null);
                return methodResult;
            }

            XLWorkbook workbook = null;

            try
            {
                workbook = new XLWorkbook(filePath);
            }
            catch
            {
                result.AddOrUpdate(ResultType.MESSAGE, "文件损坏", (key, oldValue) => null);
                methodResult.AddOrUpdate(ReadFileReturnType.RESULT, result, (key, oldValue) => null);
                return methodResult;
            }

            foreach (IXLWorksheet sheet in workbook.Worksheets)
            {
                if (sheetExplainer.sheetNames.Key == FindingMethod.SAME)
                {
                    foreach (String str in sheetExplainer.sheetNames.Value)
                    {
                        if (sheet.Name.Equals(str))
                        {
                            Analyze(sheet, result, analyzer, paramDic, cresult);
                        }
                    }
                }
                else if (sheetExplainer.sheetNames.Key == FindingMethod.CONTAIN)
                {
                    foreach (String str in sheetExplainer.sheetNames.Value)
                    {
                        if (sheet.Name.Contains(str))
                        {
                            Analyze(sheet, result, analyzer, paramDic, cresult);
                        }
                    }
                }
                else if (sheetExplainer.sheetNames.Key == FindingMethod.REGEX)
                {
                    foreach (String str in sheetExplainer.sheetNames.Value)
                    {
                        Regex rgx = new Regex(str);
                        if (rgx.IsMatch(sheet.Name))
                        {
                            Analyze(sheet, result, analyzer, paramDic, cresult);
                        }
                    }
                }
            }
            methodResult.AddOrUpdate(ReadFileReturnType.RESULT, result, (key, oldValue) => null);
            return methodResult;
        }

        private void FileTraverse(String folderPath, SheetExplainer sheetExplainer, List<String> filePathList)
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
                    }
                    else
                    {
                        string pp = FSys.Name;
                        FileTraverse(System.IO.Path.Combine(folderPath, FSys.ToString()), sheetExplainer, filePathList);
                    }
                }
            }
            catch (Exception ex)
            {
                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "OK" }, ex.Message);
                return;
            }
        }

        private void LoadFiles() 
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
            }
            catch (Exception ex)
            {
                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"新建文件夹失败\n{ex.Message}", "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

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
            cb_sheetexplainers.ItemsSource = sheetExplainersList;

            List<String> AnalyzersList = Directory.GetFiles(".\\Analyzers", "*.json").ToList();
            AnalyzersList.Insert(0, "");
            for (int i = 0; i < AnalyzersList.Count; ++i)
            {
                String str = AnalyzersList[i];
                AnalyzersList[i] = str.Substring(str.LastIndexOf('\\') + 1);
                if (AnalyzersList[i].Contains('.'))
                {
                    AnalyzersList[i] = AnalyzersList[i].Substring(0, AnalyzersList[i].LastIndexOf('.'));
                }
            }
            cb_analyzers.ItemsSource = AnalyzersList;

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

            cb_params.ItemsSource = newParamsList;
        }

        private long GetNowSs()
        {
            return (DateTime.Now.ToUniversalTime().Ticks - 621355968000000000) / 10000;
        }

        private void Analyze(IXLWorksheet sheet, ConcurrentDictionary<ResultType, Object> result, Analyzer analyzer, Dictionary<string, string> paramDic, CompilerResults cresult) 
        {
            if (cresult.Errors.HasErrors)
            {

                String str = "[ERROR: ]";
                foreach (CompilerError err in cresult.Errors)
                {
                    str = $"{str}\n    {err.ErrorText} in line {err.Line}"; 
                }

                errorStr = str;
                Stop();
            }
            else
            {
                // 通过反射执行代码
                try
                {
                    Assembly objAssembly = cresult.CompiledAssembly;
                    object obj = objAssembly.CreateInstance("AnalyzeCode.Analyze");
                    MethodInfo objMI = obj.GetType().GetMethod("AnalyzeSheet");
                    ++analyzeSheetInvokeCount;
                    object[] objList = new object[] { paramDic, sheet, result, GlobalObjects.GlobalObjects.GetGlobalParam(), analyzeSheetInvokeCount };
                    objMI.Invoke(obj, objList);
                    GlobalObjects.GlobalObjects.SetGlobalParam(objList[3]);
                }
                catch (Exception e)
                {
                    errorStr += $"[ERROR: ]\n    {e.InnerException.Message}\n    {analyzer.name}.AnalyzeSheet(): {e.InnerException.StackTrace.Substring(e.InnerException.StackTrace.LastIndexOf(':') + 1)}\n";
                    Stop();
                }
            }

        }

        private Object[] SetResult(List<object> setResultParams)
        {
            XLWorkbook workbook = (XLWorkbook)setResultParams[0];
            ConcurrentDictionary<ResultType, Object> result = (ConcurrentDictionary<ResultType, Object>)setResultParams[1];
            Analyzer analyzer = (Analyzer)setResultParams[2];
            int totalCount = (int)setResultParams[3];
            Dictionary<string, string> paramDic = (Dictionary<string, string>)setResultParams[4];


            Boolean resBoolean = true;

            Object filePath = null;
            if(result.TryGetValue(ResultType.FILEPATH, out filePath))
            {
                currentOutputtingDictionary.AddOrUpdate(filePath.ToString(), GetNowSs(), (key, oldValue) => GetNowSs());
            }

            CompilerResults cresult = GetCresult(analyzer);

            if (cresult.Errors.HasErrors)
            {
                resBoolean = false;

                String str = "[ERROR: ]";
                foreach (CompilerError err in cresult.Errors)
                {
                    str = $"{str}\n    {err.ErrorText} in line {err.Line}";
                }
                errorStr = str;
                Stop();
            }
            else
            {
                // 通过反射执行代码
                try
                {
                    Assembly objAssembly = cresult.CompiledAssembly;
                    object obj = objAssembly.CreateInstance("AnalyzeCode.Analyze");
                    MethodInfo objMI = obj.GetType().GetMethod("SetResult");
                    ++setResultInvokeCount;
                    objMI.Invoke(obj, new object[] { paramDic, workbook, result, GlobalObjects.GlobalObjects.GetGlobalParam(), setResultInvokeCount, totalCount });
                }
                catch (Exception e)
                {
                    errorStr += $"[ERROR: ]\n    {e.InnerException.Message}\n    {analyzer.name}.SetResult(): {e.InnerException.StackTrace.Substring(e.InnerException.StackTrace.LastIndexOf(':') + 1)}\n";
                    resBoolean = false;
                    Stop();
                }
            }
            return new Object[] { resBoolean, result };
        }

        private void WindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            IniHelper.SetWindowSize(this.Name.Substring(2), new Point(this.Width, this.Height));
            IniHelper.SetBasePath(tb_base_path.Text);
            IniHelper.SetOutputPath(tb_output_path.Text);
            IniHelper.SetOutputFileName(tb_output_name.Text);
        }

        private void RenewSmartThreadPoolAnalyze(STPStartInfo stp)
        {
            smartThreadPoolAnalyze = new SmartThreadPool(stp);
            int maxThreadCount = IniHelper.GetMaxThreadCount();
            if (maxThreadCount > 0)
            {
                smartThreadPoolAnalyze.MaxThreads = maxThreadCount;
            }
        }

        private void RenewSmartThreadPoolOutput(STPStartInfo stp)
        {
            smartThreadPoolOutput = new SmartThreadPool(stp);
            int maxThreadCount = IniHelper.GetMaxThreadCount();
            if (maxThreadCount > 0)
            {
                smartThreadPoolOutput.MaxThreads = maxThreadCount;
            }
        }

        private Boolean SetErrorLog()
        {
            if (errorStr != "")
            {
                tb_log.Text += errorStr;
                tb_log.Text = tb_log.Text.Substring(tb_log.Text.IndexOf("[ERROR: ]"));
                return true;
            }
            return false;
        }

        private void CbParamsSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            te_params.TextChanged -= TbParamsTextChanged;
            te_params.Text = $"{cb_params.SelectedItem}";
            te_params.TextChanged += TbParamsTextChanged;
        }

        private void CbParamsPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            LoadFiles();
        }

        private void TbParamsTextChanged(object sender, EventArgs e)
        {
            cb_params.SelectionChanged -= CbParamsSelectionChanged;
            cb_params.SelectedIndex = 0;
            cb_params.SelectionChanged += CbParamsSelectionChanged;
        }

        private CompilerResults GetCresult(Analyzer analyzer) 
        {
            CSharpCodeProvider objCSharpCodePrivoder = new CSharpCodeProvider();

            CompilerParameters objCompilerParameters = new CompilerParameters();

            string[] dlls = IniHelper.GetDlls().Split('|');
            foreach (string dll in dlls)
            {
                objCompilerParameters.ReferencedAssemblies.Add(dll);
            }

            objCompilerParameters.GenerateExecutable = false;
            objCompilerParameters.GenerateInMemory = true;
            objCompilerParameters.IncludeDebugInformation = true;

            return objCSharpCodePrivoder.CompileAssemblyFromSource(objCompilerParameters, analyzer.code);
        }

        void PasteCommandBindingPreviewCanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            // Get clipboard data and stuff
            var dataObject = Clipboard.GetDataObject();
            var text = (string)dataObject.GetData(DataFormats.UnicodeText);
            // normalize newlines so we definitely get all the newlines
            text = TextUtilities.NormalizeNewLines(text, Environment.NewLine);

            // if the text contains newlines - replace them and paste again :)
            if (text.Contains(Environment.NewLine))
            {
                e.CanExecute = false;
                e.Handled = true;
                text = text.Replace(Environment.NewLine, " ");
                Clipboard.SetText(text);
                te_params.Paste();
            }
        }

        private void TbParamsLostFocus(object sender, RoutedEventArgs e)
        {
            te_params.Text = $"{te_params.Text.Replace("\n\r", "").Replace("\r\n", "").Replace("\n", "").Replace("\r", "")}";
        }
    }
}
