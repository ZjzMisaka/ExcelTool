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
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Media;
using ExcelTool.Helper;

namespace ExcelTool
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private enum ReadFileReturnType { ANALYZER, RESULT };
        private Boolean isRunning;
        private Boolean isStopByUser;
        private SmartThreadPool smartThreadPoolAnalyze = null;
        private SmartThreadPool smartThreadPoolOutput = null;
        private ConcurrentDictionary<string, long> currentAnalizingDictionary;
        private ConcurrentDictionary<string, long> currentOutputtingDictionary;
        private ConcurrentDictionary<ConcurrentDictionary<ResultType, Object>, Analyzer> resultList;
        private Dictionary<FileSystemWatcher, string> fileSystemWatcherDic;
        private List<string> copiedDllsList;
        private int analyzeSheetInvokeCount;
        private int setResultInvokeCount;
        private int totalTimeoutLimitAnalyze;
        private int perTimeoutLimitAnalyze;
        private int totalTimeoutLimitOutput;
        private int perTimeoutLimitOutput;
        private bool runNotSuccessed;
        public MainWindow()
        {
            InitializeComponent();

            ICSharpCode.AvalonEdit.Search.SearchPanel.Install(te_log);

            currentAnalizingDictionary = new ConcurrentDictionary<string, long>();
            currentOutputtingDictionary = new ConcurrentDictionary<string, long>();
            resultList = new ConcurrentDictionary<ConcurrentDictionary<ResultType, Object>, Analyzer>();
            fileSystemWatcherDic = new Dictionary<FileSystemWatcher, string>();
            copiedDllsList = new List<string>();
            analyzeSheetInvokeCount = 0;
            setResultInvokeCount = 0;

            isRunning = false;
            isStopByUser = false;
            runNotSuccessed = false;
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

            cb_isautoopen.IsChecked = IniHelper.GetIsAutoOpen();

            LoadFiles();

            SetAutoStatusAll();


            IHighlightingDefinition chParam;
            using (Stream s = new FileStream(@"Highlighting\ParamHighlighting.xshd", FileMode.Open))
            {
                if (s == null)
                    throw new InvalidOperationException("Could not find embedded resource");
                using (XmlReader reader = new XmlTextReader(s))
                {
                    //解析xshd
                    chParam = ICSharpCode.AvalonEdit.Highlighting.Xshd.HighlightingLoader.Load(reader, HighlightingManager.Instance);
                }
            }
            //注册数据
            HighlightingManager.Instance.RegisterHighlighting("param", new string[] { }, chParam);
            te_params.SyntaxHighlighting = HighlightingManager.Instance.GetDefinition("param");

            IHighlightingDefinition chLog;
            using (Stream s = new FileStream(@"Highlighting\LogHighlighting.xshd", FileMode.Open))
            {
                if (s == null)
                    throw new InvalidOperationException("Could not find embedded resource");
                using (XmlReader reader = new XmlTextReader(s))
                {
                    //解析xshd
                    chLog = ICSharpCode.AvalonEdit.Highlighting.Xshd.HighlightingLoader.Load(reader, HighlightingManager.Instance);
                }
            }
            //注册数据
            HighlightingManager.Instance.RegisterHighlighting("log", new string[] { }, chLog);
            te_log.SyntaxHighlighting = HighlightingManager.Instance.GetDefinition("log");
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
            cb_sheetexplainers.ItemsSource = FileHelper.GetSheetExplainersList();
        }

        private void CbAnalyzersPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            cb_analyzers.ItemsSource = FileHelper.GetAnalyzersList();
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
            isStopByUser = true;

            te_log.Text += Logger.Get();
        }

        private void FileSystemWatcherInvoke(object sender, FileSystemEventArgs e)
        {
            FileSystemWatcher fileSystemWatcher = (FileSystemWatcher)sender;
            if (fileSystemWatcherDic.ContainsKey(fileSystemWatcher))
            {
                try
                {
                    fileSystemWatcher.EnableRaisingEvents = false;

                    string ruleName = fileSystemWatcherDic[fileSystemWatcher];

                    RunningRule rule = JsonConvert.DeserializeObject<RunningRule>(File.ReadAllText($".\\Rules\\{ruleName}.json"));

                    if (!CheckRule(rule))
                    {
                        cb_rules.SelectedIndex = 0;
                        return;
                    }
                    this.Dispatcher.Invoke(() =>
                    {
                        Start(rule.sheetExplainers, rule.analyzers, rule.param, rule.basePath, rule.outputPath, rule.outputName, true);
                    });
                }
                finally
                {


                    while (isRunning)
                    {
                        Thread.Sleep(100);
                    }


                    fileSystemWatcher.EnableRaisingEvents = true;
                }
            }

        }

        private void Start(object sender, RoutedEventArgs e)
        {
            Start(te_sheetexplainers.Text, te_analyzers.Text, te_params.Text, tb_base_path.Text, tb_output_path.Text, tb_output_name.Text, false);
        }

        private async void Start(string sheetExplainersStr, string analyzersStr, string paramStr, string basePath, string outputPath, string outputName, bool isAuto)
        {
            this.Dispatcher.Invoke(() =>
            {
                btn_start.IsEnabled = false;
                btn_stop.IsEnabled = true;
            });

            Dictionary<SheetExplainer, List<string>> filePathListDic = new Dictionary<SheetExplainer, List<string>>();

            isRunning = true;
            isStopByUser = false;

            analyzeSheetInvokeCount = 0;
            setResultInvokeCount = 0;
            resultList = new ConcurrentDictionary<ConcurrentDictionary<ResultType, Object>, Analyzer>();
            GlobalObjects.GlobalObjects.SetGlobalParam(null);
            if (!isAuto)
            {
                te_log.Text = "";
                Logger.Clear();
            }
            else
            {
                Logger.Print("---- AUTO ANALYZE ----");
            }
            runNotSuccessed = false;

            List<String> sheetExplainersList = sheetExplainersStr.Split('\n').Where(str => str.Trim() != "").ToList();
            List<String> analyzersList = analyzersStr.Split('\n').Where(str => str.Trim() != "").ToList();
            if (sheetExplainersList.Count != analyzersList.Count || sheetExplainersList.Count == 0)
            {
                this.Dispatcher.Invoke(() =>
                {
                    btn_start.IsEnabled = true;
                    btn_stop.IsEnabled = false;
                });

                isRunning = false;
                isStopByUser = false;
                return;
            }

            paramStr = $"{paramStr.Replace("\r\n", "").Replace("\n", "")}";
            te_params.Text = paramStr;
            Dictionary<string, Dictionary<string, string>> paramDicEachAnalyzer = ParamHelper.GetParamDicEachAnalyzer(paramStr);


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
                catch (Exception)
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
                    if (!isAuto)
                    {
                        CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    else 
                    {
                        Logger.Error(ex.Message);
                    }
                    this.Dispatcher.Invoke(() =>
                    {
                        btn_start.IsEnabled = true;
                        btn_stop.IsEnabled = false;
                    });

                    isRunning = false;
                    isStopByUser = false;
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
                    if (!isAuto)
                    {
                        CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    else
                    {
                        Logger.Error(ex.Message);
                    }
                    this.Dispatcher.Invoke(() =>
                    {
                        btn_start.IsEnabled = true;
                        btn_stop.IsEnabled = false;
                    });

                    isRunning = false;
                    isStopByUser = false;
                    return;
                }
                analyzers.Add(analyzer);
            }

            Dictionary<Analyzer, Tuple<CompilerResults, SheetExplainer>> compilerDic = new Dictionary<Analyzer, Tuple<CompilerResults, SheetExplainer>>();
            int totalCount = 0;
            for (int i = 0; i < sheetExplainersList.Count; ++i)
            {
                long startTime = GetNowSs();

                SheetExplainer sheetExplainer = sheetExplainers[i];
                Analyzer analyzer = analyzers[i];

                List<String> allFilePathList = new List<String>();
                foreach (String str in sheetExplainer.relativePathes)
                {
                    List<String> filePathList = new List<String>();
                    string basePathTemp = Path.Combine(basePath.Trim(), str.Trim());
                    FileHelper.FileTraverse(isAuto, basePathTemp, sheetExplainer, filePathList);
                    allFilePathList.AddRange(filePathList);
                }
                filePathListDic.Add(sheetExplainer, allFilePathList);

                CompilerResults cresult = GetCresult(analyzer);
                Tuple<CompilerResults, SheetExplainer> cresultTuple = new Tuple<CompilerResults, SheetExplainer>(cresult, sheetExplainer);
                compilerDic.Add(analyzer, cresultTuple);

                Thread runBeforeAnalyzeSheetThread = new Thread(() => RunBeforeAnalyzeSheet(cresult, ParamHelper.MergePublicParam(paramDicEachAnalyzer, analyzer.name), analyzer, allFilePathList));
                runBeforeAnalyzeSheetThread.Start();

                while (runBeforeAnalyzeSheetThread.ThreadState == System.Threading.ThreadState.Running)
                {
                    if (isStopByUser)
                    {
                        try
                        {
                            runBeforeAnalyzeSheetThread.Abort();
                        }
                        catch (Exception)
                        {
                            // DO NOTHING
                        }
                        this.Dispatcher.Invoke(() =>
                        {
                            btn_start.IsEnabled = true;
                            btn_stop.IsEnabled = false;
                        });
                        SetErrorLog();
                        isRunning = false;
                        return;
                    }
                    long timeCostSs = GetNowSs() - startTime;
                    if (timeCostSs >= perTimeoutLimitAnalyze)
                    {
                        try
                        {
                            runBeforeAnalyzeSheetThread.Abort();
                        }
                        catch (Exception)
                        {
                            // DO NOTHING
                        }
                        this.Dispatcher.Invoke(() =>
                        {
                            btn_start.IsEnabled = true;
                            btn_stop.IsEnabled = false;
                        });
                        if (!isAuto)
                        {
                            CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "OK" }, $"RunBeforeAnalyzeSheet\nTime out. \n{perTimeoutLimitAnalyze / 1000.0}(s)", "error", MessageBoxImage.Error);
                        }
                        else
                        {
                            Logger.Error($"RunBeforeAnalyzeSheet\nTime out. \n{perTimeoutLimitAnalyze / 1000.0}(s)");
                        }
                        SetErrorLog();
                        isRunning = false;
                        return;
                    }
                    this.Dispatcher.Invoke(() =>
                    {
                        te_log.Text += Logger.Get();
                    });
                    await Task.Delay(100);
                }

                foreach (String str in sheetExplainer.relativePathes)
                {
                    List<String> filePathList = filePathListDic[sheetExplainer];
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
                        readFileParams.Add(ParamHelper.MergePublicParam(paramDicEachAnalyzer, analyzer.name));
                        readFileParams.Add(cresult);
                        smartThreadPoolAnalyze.QueueWorkItem(new Func<List<object>, Object>(ReadFile), readFileParams);
                    }
                }
            }

            long startSs = GetNowSs();
            smartThreadPoolAnalyze.Join();
            while (smartThreadPoolAnalyze.CurrentWorkItemsCount > 0)
            {
                try
                {
                    long nowSs = GetNowSs();
                    long totalTimeCostSs = nowSs - startSs;
                    if (isStopByUser || totalTimeCostSs >= totalTimeoutLimitAnalyze)
                    {
                        smartThreadPoolAnalyze.Dispose();
                        this.Dispatcher.Invoke(() =>
                        {
                            btn_start.IsEnabled = true;
                            btn_stop.IsEnabled = false;
                        });
                        if (totalTimeCostSs >= totalTimeoutLimitAnalyze)
                        {
                            if (!isAuto)
                            {
                                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "OK" }, $"Total time out. \n{totalTimeoutLimitAnalyze / 1000.0}(s)", "error", MessageBoxImage.Error);
                            }
                            else
                            {
                                Logger.Error($"Total time out. \n{totalTimeoutLimitAnalyze / 1000.0}(s)");
                            }

                        }
                        SetErrorLog();

                        isRunning = false;
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
                                if (!isAuto)
                                {
                                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "OK" }, $"{key}\nTime out. \n{perTimeoutLimitAnalyze / 1000.0}(s)", "error", MessageBoxImage.Error);
                                }
                                else
                                {
                                    Logger.Error($"{key}\nTime out. \n{perTimeoutLimitAnalyze / 1000.0}(s)");
                                }
                                SetErrorLog();
                                isRunning = false;
                                return;
                            }
                            sb.Append(key.Substring(key.LastIndexOf('\\') + 1)).Append("(").Append((timeCostSs / 1000.0).ToString("0.0")).Append(")");
                        }
                    }

                    this.Dispatcher.Invoke(() =>
                    {
                        l_process.Content = $"{smartThreadPoolAnalyze.CurrentWorkItemsCount}/{totalCount} -- ActiveThreads: {smartThreadPoolAnalyze.ActiveThreads} -- InUseThreads: {smartThreadPoolAnalyze.InUseThreads}";
                        tb_status.Text = $"{sb.ToString()}";
                        te_log.Text += Logger.Get();
                    });
                    await Task.Delay(100);
                }
                catch (Exception e)
                {
                    smartThreadPoolAnalyze.Dispose();
                    Logger.Error($"Exception has been throwed. \n{e.Message}");
                    SetErrorLog();

                    isRunning = false;
                    return;
                }
            }
            smartThreadPoolAnalyze.Dispose();

            // 输出结果
            using (var workbook = new XLWorkbook())
            {
                string resPath = outputPath.Replace("\\", "/");
                string filePath;
                if (outputPath.EndsWith("/"))
                {
                    filePath = $"{resPath}{outputName}.xlsx";
                }
                else
                {
                    filePath = $"{resPath}/{outputName}.xlsx";
                }

                foreach (Analyzer analyzer in compilerDic.Keys)
                {
                    long startTime = GetNowSs();

                    CompilerResults cresult = compilerDic[analyzer].Item1;
                    SheetExplainer sheetExplainer = compilerDic[analyzer].Item2;
                    Thread runBeforeSetResultThread = new Thread(() => RunBeforeSetResult(cresult, workbook, ParamHelper.MergePublicParam(paramDicEachAnalyzer, analyzer.name), analyzer, filePathListDic[sheetExplainer]));
                    runBeforeSetResultThread.Start();
                    while (runBeforeSetResultThread.ThreadState == System.Threading.ThreadState.Running)
                    {
                        if (isStopByUser)
                        {
                            runBeforeSetResultThread.Abort();
                            this.Dispatcher.Invoke(() =>
                            {
                                btn_start.IsEnabled = true;
                                btn_stop.IsEnabled = false;
                            });
                            SetErrorLog();
                            isRunning = false;
                            return;
                        }
                        long timeCostSs = GetNowSs() - startTime;
                        if (timeCostSs >= perTimeoutLimitAnalyze)
                        {
                            runBeforeSetResultThread.Abort();
                            this.Dispatcher.Invoke(() =>
                            {
                                btn_start.IsEnabled = true;
                                btn_stop.IsEnabled = false;
                            });
                            if (!isAuto)
                            {
                                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "OK" }, $"RunBeforeAnalyzeSheet\nTime out. \n{perTimeoutLimitAnalyze / 1000.0}(s)", "error", MessageBoxImage.Error);
                            }
                            else
                            {
                                Logger.Error($"RunBeforeAnalyzeSheet\nTime out. \n{perTimeoutLimitAnalyze / 1000.0}(s)");
                            }
                            SetErrorLog();
                            isRunning = false;
                            return;
                        }
                        this.Dispatcher.Invoke(() =>
                        {
                            te_log.Text += Logger.Get();
                        });
                        await Task.Delay(100);
                    }
                }

                foreach (ConcurrentDictionary<ResultType, Object> result in resultList.Keys)
                {
                    Analyzer analyzer = resultList[result];
                    List<object> setResultParams = new List<object>();
                    setResultParams.Add(workbook);
                    setResultParams.Add(result);
                    setResultParams.Add(analyzer.name);
                    setResultParams.Add(resultList.Count);
                    setResultParams.Add(ParamHelper.MergePublicParam(paramDicEachAnalyzer, analyzer.name));
                    setResultParams.Add(compilerDic[analyzer].Item1);
                    smartThreadPoolOutput.QueueWorkItem(new Func<List<object>, Object[]>(SetResult), setResultParams);
                }
                startSs = GetNowSs();
                smartThreadPoolOutput.Join();
                while (smartThreadPoolOutput.CurrentWorkItemsCount > 0)
                {
                    try
                    {
                        long nowSs = GetNowSs();
                        long totalTimeCostSs = nowSs - startSs;
                        if (isStopByUser || totalTimeCostSs >= totalTimeoutLimitOutput)
                        {
                            smartThreadPoolOutput.Dispose();
                            this.Dispatcher.Invoke(() =>
                            {
                                btn_start.IsEnabled = true;
                                btn_stop.IsEnabled = false;
                            });
                            if (totalTimeCostSs >= totalTimeoutLimitOutput)
                            {
                                if (!isAuto)
                                {
                                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "OK" }, $"Total time out. \n{totalTimeoutLimitOutput / 1000.0}(s)", "error", MessageBoxImage.Error);
                                }
                                else
                                {
                                    Logger.Error($"Total time out. \n{totalTimeoutLimitOutput / 1000.0}(s)");
                                }
                            }
                            SetErrorLog();

                            isRunning = false;
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
                                if (timeCostSs >= perTimeoutLimitOutput)
                                {
                                    smartThreadPoolOutput.Dispose();
                                    this.Dispatcher.Invoke(() =>
                                    {
                                        btn_start.IsEnabled = true;
                                        btn_stop.IsEnabled = false;
                                    });
                                    if (!isAuto)
                                    {
                                        CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "OK" }, $"{key}\nTime out. \n{perTimeoutLimitAnalyze / 1000.0}(s)", "error", MessageBoxImage.Error);
                                    }
                                    else
                                    {
                                        Logger.Error($"{key}\nTime out. \n{perTimeoutLimitAnalyze / 1000.0}(s)");
                                    }
                                    SetErrorLog();
                                    isRunning = false;
                                    return;
                                }
                                sb.Append(key.Substring(key.LastIndexOf('\\') + 1)).Append("(").Append((timeCostSs / 1000.0).ToString("0.0")).Append(")");
                            }
                        }

                        this.Dispatcher.Invoke(() =>
                        {
                            l_process.Content = $"{smartThreadPoolOutput.CurrentWorkItemsCount}/{totalCount} -- ActiveThreads: {smartThreadPoolOutput.ActiveThreads} -- InUseThreads: {smartThreadPoolOutput.InUseThreads}";
                            tb_status.Text = $"{sb.ToString()}";
                            te_log.Text += Logger.Get();
                        });
                        await Task.Delay(100);
                    }
                    catch (Exception e)
                    {
                        smartThreadPoolOutput.Dispose();
                        Logger.Error($"Exception has been throwed. \n{e.Message}");
                        SetErrorLog();
                        return;
                    }
                }
                smartThreadPoolOutput.Dispose();

                foreach (Analyzer analyzer in compilerDic.Keys)
                {
                    long startTime = GetNowSs();

                    CompilerResults cresult = compilerDic[analyzer].Item1;
                    SheetExplainer sheetExplainer = compilerDic[analyzer].Item2;
                    Thread runEndThread = new Thread(() => RunEnd(cresult, workbook, ParamHelper.MergePublicParam(paramDicEachAnalyzer, analyzer.name), analyzer, filePathListDic[sheetExplainer]));
                    runEndThread.Start();
                    while (runEndThread.ThreadState == System.Threading.ThreadState.Running)
                    {
                        if (isStopByUser)
                        {
                            runEndThread.Abort();
                            this.Dispatcher.Invoke(() =>
                            {
                                btn_start.IsEnabled = true;
                                btn_stop.IsEnabled = false;
                            });
                            SetErrorLog();
                            isRunning = false;
                            return;
                        }
                        long timeCostSs = GetNowSs() - startTime;
                        if (timeCostSs >= perTimeoutLimitAnalyze)
                        {
                            runEndThread.Abort();
                            this.Dispatcher.Invoke(() =>
                            {
                                btn_start.IsEnabled = true;
                                btn_stop.IsEnabled = false;
                            });
                            if (!isAuto)
                            {
                                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "OK" }, $"RunEnd\nTime out. \n{perTimeoutLimitAnalyze / 1000.0}(s)", "error", MessageBoxImage.Error);
                            }
                            else
                            {
                                Logger.Error($"RunEnd\nTime out. \n{perTimeoutLimitAnalyze / 1000.0}(s)");
                            }
                            SetErrorLog();
                            isRunning = false;
                            return;
                        }
                        this.Dispatcher.Invoke(() =>
                        {
                            te_log.Text += Logger.Get();
                        });
                        await Task.Delay(100);
                    }
                }

                if (runNotSuccessed)
                {
                    btn_start.IsEnabled = true;
                    btn_stop.IsEnabled = false;

                    isRunning = false;
                    return;
                }
                tb_status.Text = "";

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
                    if (!isAuto)
                    {
                        CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    else
                    {
                        Logger.Error(ex.Message);
                    }
                }

                bool saveResult = false;
                SaveFile(isAuto, filePath, workbook, out saveResult);
                if (!saveResult)
                {
                    if (!isAuto)
                    {
                        CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "OK" }, "file not saved.", "info");
                    }
                    else
                    {
                        Logger.Error("文件保存失败.");
                        te_log.Text += Logger.Get();
                    }

                    btn_start.IsEnabled = true;
                    btn_stop.IsEnabled = false;

                    isRunning = false;
                    return;
                }

                Button btnOpenFile = new Button();
                btnOpenFile.Style = GlobalObjects.GlobalObjects.GetBtnStyle();
                btnOpenFile.Content = "打开文件";
                btnOpenFile.Click += (s, ee) =>
                {
                    System.Diagnostics.Process.Start(filePath);
                    CustomizableMessageBox.MessageBox.CloseNow();
                };

                Button btnOpenPath = new Button();
                btnOpenPath.Style = GlobalObjects.GlobalObjects.GetBtnStyle();
                btnOpenPath.Content = "打开路径";
                btnOpenPath.Click += (s, ee) =>
                {
                    System.Diagnostics.Process.Start("Explorer", $"/e,/select,{filePath.Replace("/", "\\")}");
                    CustomizableMessageBox.MessageBox.CloseNow();
                };

                Button btnClose = new Button();
                btnClose.Style = GlobalObjects.GlobalObjects.GetBtnStyle();
                btnClose.Content = "关闭";
                btnClose.Click += (s, ee) =>
                {
                    CustomizableMessageBox.MessageBox.CloseNow();
                };

                Logger.Info($"文件已保存。\n{filePath}");
                if (!isAuto && cb_isautoopen.IsChecked == false)
                {
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { btnClose, new ButtonSpacer(40), btnOpenFile, btnOpenPath }, "文件已保存。", "OK");
                }
                else
                {
                    if (cb_isautoopen.IsChecked == true)
                    {
                        Logger.Info("已自动打开文件");
                        System.Diagnostics.Process.Start(filePath);
                    }
                }
            }

            te_log.Text += Logger.Get();
            
            isRunning = false;
            btn_start.IsEnabled = true;
            btn_stop.IsEnabled = false;
        }

        private void RunBeforeAnalyzeSheet(CompilerResults cresult, Dictionary<string, string> paramDic, Analyzer analyzer, List<String> allFilePathList)
        {
            if (cresult.Errors.HasErrors)
            {

                String str = "";
                foreach (CompilerError err in cresult.Errors)
                {
                    str = $"{str}\n    {err.ErrorText} in line {err.Line}";
                }

                Logger.Error(str);
                runNotSuccessed = true;
                Stop();
            }
            else
            {
                // 通过反射执行代码
                try
                {
                    Assembly objAssembly = cresult.CompiledAssembly;
                    object obj = objAssembly.CreateInstance("AnalyzeCode.Analyze");
                    MethodInfo objMI = obj.GetType().GetMethod("RunBeforeAnalyzeSheet");
                    object[] objList = new object[] { paramDic, GlobalObjects.GlobalObjects.GetGlobalParam(), allFilePathList };
                    objMI.Invoke(obj, objList);
                    GlobalObjects.GlobalObjects.SetGlobalParam(objList[1]);
                }
                catch (Exception e)
                {
                    if (!(e is ThreadAbortException))
                    {
                        Logger.Error($"\n    {e.InnerException.Message}\n    {analyzer.name}.RunBeforeAnalyzeSheet(): {e.InnerException.StackTrace.Substring(e.InnerException.StackTrace.LastIndexOf(':') + 1)}");
                        runNotSuccessed = true;
                        Stop();
                    }
                }
            }
        }

        private void RunBeforeSetResult(CompilerResults cresult, XLWorkbook workbook, Dictionary<string, string> paramDic, Analyzer analyzer, List<String> allFilePathList)
        {
            // 通过反射执行代码
            try
            {
                Assembly objAssembly = cresult.CompiledAssembly;
                object obj = objAssembly.CreateInstance("AnalyzeCode.Analyze");
                MethodInfo objMI = obj.GetType().GetMethod("RunBeforeSetResult");
                object[] objList = new object[] { paramDic, workbook, GlobalObjects.GlobalObjects.GetGlobalParam(), resultList.Keys, allFilePathList };
                objMI.Invoke(obj, objList);
                GlobalObjects.GlobalObjects.SetGlobalParam(objList[2]);
            }
            catch (Exception e)
            {
                if (!(e is ThreadAbortException))
                {
                    Logger.Error($"\n    {e.InnerException.Message}\n    {analyzer.name}.RunBeforeSetResult(): {e.InnerException.StackTrace.Substring(e.InnerException.StackTrace.LastIndexOf(':') + 1)}");
                    runNotSuccessed = true;
                    Stop();
                }
            }
        }

        private void RunEnd(CompilerResults cresult, XLWorkbook workbook, Dictionary<string, string> paramDic, Analyzer analyzer, List<String> allFilePathList)
        {
            // 通过反射执行代码
            try
            {
                Assembly objAssembly = cresult.CompiledAssembly;
                object obj = objAssembly.CreateInstance("AnalyzeCode.Analyze");
                MethodInfo objMI = obj.GetType().GetMethod("RunEnd");
                object[] objList = new object[] { paramDic, workbook, GlobalObjects.GlobalObjects.GetGlobalParam(), resultList.Keys, allFilePathList };
                objMI.Invoke(obj, objList);
                GlobalObjects.GlobalObjects.SetGlobalParam(objList[2]);
            }
            catch (Exception e)
            {
                if (!(e is ThreadAbortException))
                {
                    Logger.Error($"\n    {e.InnerException.Message}\n    {analyzer.name}.RunEnd(): {e.InnerException.StackTrace.Substring(e.InnerException.StackTrace.LastIndexOf(':') + 1)}");
                    runNotSuccessed = true;
                    Stop();
                }
            }
        }

        private void SaveFile(bool isAuto, string filePath, XLWorkbook workbook, out bool result)
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
                    res = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "YES", "NO" }, $"保存失败 重试? \n{e.Message}", "error", MessageBoxImage.Question);
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

        [DllImport("kernel32.dll")]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);

        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr hObject);

        public const int OF_READWRITE = 2;
        public const int OF_SHARE_DENY_NONE = 0x40;
        public readonly IntPtr HFILE_ERROR = new IntPtr(-1);
        private bool isFileUsing(string filePath)
        {
            IntPtr vHandle = _lopen(filePath, OF_READWRITE | OF_SHARE_DENY_NONE);
            if (vHandle != HFILE_ERROR)
            {
                CloseHandle(vHandle);
            }
            return vHandle == HFILE_ERROR;
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

            if (filePath.Contains("~$") || isFileUsing(filePath))
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
                else if (sheetExplainer.sheetNames.Key == FindingMethod.ALL)
                {
                    Regex rgx = new Regex("[\\s\\S]*");
                    if (rgx.IsMatch(sheet.Name))
                    {
                        Analyze(sheet, result, analyzer, paramDic, cresult);
                    }
                }
            }
            methodResult.AddOrUpdate(ReadFileReturnType.RESULT, result, (key, oldValue) => null);
            return methodResult;
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
                if (!Directory.Exists(".\\Rules"))
                {
                    Directory.CreateDirectory(".\\Rules");
                }
            }
            catch (Exception ex)
            {
                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"新建文件夹失败\n{ex.Message}", "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            cb_sheetexplainers.ItemsSource = FileHelper.GetSheetExplainersList();
            cb_analyzers.ItemsSource = FileHelper.GetAnalyzersList();
            cb_rules.ItemsSource = FileHelper.GetRulesList();
            cb_params.ItemsSource = FileHelper.GetParamsList();
        }

        private long GetNowSs()
        {
            return (DateTime.Now.ToUniversalTime().Ticks - 621355968000000000) / 10000;
        }

        private void Analyze(IXLWorksheet sheet, ConcurrentDictionary<ResultType, Object> result, Analyzer analyzer, Dictionary<string, string> paramDic, CompilerResults cresult) 
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
                Logger.Error($"\n    {e.InnerException.Message}\n    {analyzer.name}.AnalyzeSheet(): {e.InnerException.StackTrace.Substring(e.InnerException.StackTrace.LastIndexOf(':') + 1)}");
                runNotSuccessed = true;
                Stop();
            }

        }

        private Object[] SetResult(List<object> setResultParams)
        {
            XLWorkbook workbook = (XLWorkbook)setResultParams[0];
            ConcurrentDictionary<ResultType, Object> result = (ConcurrentDictionary<ResultType, Object>)setResultParams[1];
            string analyzerName = (string)setResultParams[2];
            int totalCount = (int)setResultParams[3];
            Dictionary<string, string> paramDic = (Dictionary<string, string>)setResultParams[4];
            CompilerResults cresult = (CompilerResults)setResultParams[5];


            Boolean resBoolean = true;

            Object filePath = null;
            if(result.TryGetValue(ResultType.FILEPATH, out filePath))
            {
                currentOutputtingDictionary.AddOrUpdate(filePath.ToString(), GetNowSs(), (key, oldValue) => GetNowSs());
            }

            // 通过反射执行代码
            try
            {
                Assembly objAssembly = cresult.CompiledAssembly;
                object obj = objAssembly.CreateInstance("AnalyzeCode.Analyze");
                MethodInfo objMI = obj.GetType().GetMethod("SetResult");
                ++setResultInvokeCount;
                object[] objList = new object[] { paramDic, workbook, result, GlobalObjects.GlobalObjects.GetGlobalParam(), setResultInvokeCount, totalCount };
                objMI.Invoke(obj, objList);
                GlobalObjects.GlobalObjects.SetGlobalParam(objList[3]);
            }
            catch (Exception e)
            {
                Logger.Error($"\n    {e.InnerException.Message}\n    {analyzerName}.SetResult(): {e.InnerException.StackTrace.Substring(e.InnerException.StackTrace.LastIndexOf(':') + 1)}");
                runNotSuccessed = true;
                resBoolean = false;
                Stop();
            }
            return new Object[] { resBoolean, result };
        }

        private void WindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (Application.Current.Windows.Count > 1)
            {
                MessageBoxResult result = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"尚有未关闭的子窗口, 是否关闭程序", "Warning", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result == MessageBoxResult.No)
                {
                    e.Cancel = true;
                }
                else
                {
                    foreach (Window window in Application.Current.Windows)
                    {
                        if (Application.Current.MainWindow != window)
                        {
                            window.Close();
                        }
                    }
                }
            }

            IniHelper.SetWindowSize(this.Name.Substring(2), new Point(this.Width, this.Height));
            IniHelper.SetBasePath(tb_base_path.Text);
            IniHelper.SetOutputPath(tb_output_path.Text);
            IniHelper.SetOutputFileName(tb_output_name.Text);

            if (cb_isautoopen.IsChecked == true)
            {
                IniHelper.SetIsAutoOpen(true);
            }
            else
            {
                IniHelper.SetIsAutoOpen(false);
            }
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
            string log = Logger.Get();
            if (log != null && log != "")
            {
                te_log.Text += log;
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
            cb_params.ItemsSource = FileHelper.GetParamsList();
        }

        private void TbParamsTextChanged(object sender, EventArgs e)
        {
            cb_params.SelectionChanged -= CbParamsSelectionChanged;
            cb_params.SelectedIndex = 0;
            cb_params.SelectionChanged += CbParamsSelectionChanged;
        }

        private FileSystemInfo[] GetyDllInfos()
        {
            string folderPath = Path.Combine(Environment.CurrentDirectory, "Dlls");
            DirectoryInfo dir = new DirectoryInfo(folderPath);
            FileSystemInfo[] dllInfos = null;
            if (dir.Exists)
            {
                DirectoryInfo dirD = dir as DirectoryInfo;
                dllInfos = dirD.GetFileSystemInfos();
            }

            return dllInfos;
        }

        private CompilerResults GetCresult(Analyzer analyzer) 
        {
            CSharpCodeProvider objCSharpCodePrivoder = new CSharpCodeProvider();

            CompilerParameters objCompilerParameters = new CompilerParameters();

            FileSystemInfo[] dllInfos = GetyDllInfos();
            List<string> dlls = new List<string>();
            dlls.Add("System.dll");
            dlls.Add("System.Data.dll");
            dlls.Add("System.Xml.dll");
            dlls.Add("ClosedXML.dll");
            dlls.Add("GlobalObjects.dll");
            if (dllInfos != null && dllInfos.Count() != 0)
            {
                foreach (FileSystemInfo dllInfo in dllInfos)
                {
                    dlls.Add(dllInfo.Name);
                    string destFileName = Path.Combine(Environment.CurrentDirectory, dllInfo.Name);
                    if (!copiedDllsList.Contains(destFileName))
                    {
                        File.Copy(dllInfo.FullName, destFileName, true);
                        copiedDllsList.Add(destFileName);
                    }
                }
            }
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
            var dataObject = Clipboard.GetDataObject();
            var text = (string)dataObject.GetData(DataFormats.UnicodeText);
            text = TextUtilities.NormalizeNewLines(text, Environment.NewLine);

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
            te_params.Text = $"{te_params.Text.Replace("\r\n", "").Replace("\n", "")}";
        }

        private void TeLogTextChanged(object sender, EventArgs e)
        {
            te_log.ScrollToEnd();
        }

        private void SaveRule(object sender, RoutedEventArgs e)
        {
            TextBox tbName = new TextBox();
            tbName.Margin = new Thickness(5, 8, 5, 8);
            tbName.VerticalContentAlignment = VerticalAlignment.Center;
            if (cb_rules.SelectedIndex >= 1)
            {
                tbName.Text = $"Copy Of {cb_rules.SelectedItem}";
            }
            int result = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object>() { tbName, "OK", "CANCEL" }, "name", "saving", MessageBoxImage.Information);
            if (result == 1)
            {
                RunningRule runningRule = new RunningRule();
                runningRule.analyzers = te_analyzers.Text;
                runningRule.sheetExplainers = te_sheetexplainers.Text;
                runningRule.param = te_params.Text;
                runningRule.basePath = tb_base_path.Text;
                runningRule.outputPath = tb_output_path.Text;
                runningRule.outputName = tb_output_name.Text;
                string json = JsonConvert.SerializeObject(runningRule);

                string fileName = $".\\Rules\\{tbName.Text}.json";
                FileStream fs = null;
                try
                {
                    fs = File.Create(fileName);
                    fs.Close();
                    StreamWriter sw = File.CreateText(fileName);
                    sw.Write(json);
                    sw.Flush();
                    sw.Close();
                }
                catch (Exception ex)
                {
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void DeleteRule(object sender, RoutedEventArgs e)
        {
            String path = $"{System.Environment.CurrentDirectory}\\Rules\\{cb_rules.SelectedItem.ToString()}.json";
            MessageBoxResult result = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"删除文件\n{path}", "Warning", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
            if (result == MessageBoxResult.OK)
            {
                File.Delete(path);
                cb_analyzers.SelectedIndex = 0;
            }
        }

        private void CbRulesPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            cb_rules.ItemsSource = FileHelper.GetRulesList();
        }

        private void CbRulesChanged(object sender, SelectionChangedEventArgs e)
        {

            RunningRule rule = null;

            if (cb_rules.SelectedIndex >= 1)
            {
                rule = JsonConvert.DeserializeObject<RunningRule>(File.ReadAllText($".\\Rules\\{cb_rules.SelectedItem.ToString()}.json"));

                if (!CheckRule(rule))
                {
                    cb_rules.SelectedIndex = 0;
                    return;
                }

                btn_saverule.Visibility = Visibility.Hidden;
                btn_deleterule.Visibility = Visibility.Visible;

                if (rule.watchPath != null && rule.watchPath != "")
                {
                    btn_setauto.Visibility = Visibility.Hidden;
                    btn_unsetauto.Visibility = Visibility.Visible;
                }
                else
                {
                    btn_setauto.Visibility = Visibility.Visible;
                    btn_unsetauto.Visibility = Visibility.Hidden;
                }

                te_analyzers.Text = rule.analyzers;
                te_sheetexplainers.Text = rule.sheetExplainers;
                te_params.Text = rule.param;
                tb_base_path.Text = rule.basePath;
                tb_output_path.Text = rule.outputPath;
                tb_output_name.Text = rule.outputName;
            }
            else
            {
                btn_saverule.Visibility = Visibility.Visible;
                btn_deleterule.Visibility = Visibility.Hidden;

                btn_setauto.Visibility = Visibility.Hidden;
                btn_unsetauto.Visibility = Visibility.Hidden;
            }
        }

        private void SetAuto(object sender, RoutedEventArgs e)
        {
            string ruleName = cb_rules.SelectedItem.ToString();
            RunningRule rule = JsonConvert.DeserializeObject<RunningRule>(File.ReadAllText($".\\Rules\\{ruleName}.json"));
            TextBox tbPath = new TextBox();
            tbPath.Margin = new Thickness(5, 8, 5, 8);
            tbPath.VerticalContentAlignment = VerticalAlignment.Center;
            TextBox tbFilter = new TextBox();
            tbFilter.Margin = new Thickness(5, 8, 5, 8);
            tbFilter.VerticalContentAlignment = VerticalAlignment.Center;
            int result = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object>() { tbPath, tbFilter, "OK", "CANCEL" }, "watch path and filter", "saving", MessageBoxImage.Information);
            if (result == 2)
            {
                rule.watchPath = tbPath.Text;
                rule.filter = tbFilter.Text;
            }
            else
            {
                return;
            }

            string json = JsonConvert.SerializeObject(rule);
            string fileName = $".\\Rules\\{cb_rules.SelectedItem.ToString()}.json";
            FileStream fs = null;
            try
            {
                fs = File.Create(fileName);
                fs.Close();
                StreamWriter sw = File.CreateText(fileName);
                sw.Write(json);
                sw.Flush();
                sw.Close();
            }
            catch (Exception ex)
            {
                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            SetAuto(rule, ruleName);
        }

        private void SetAuto(RunningRule rule, string ruleName)
        {
            if (!fileSystemWatcherDic.ContainsValue(ruleName))
            {
                FileSystemWatcher fileSystemWatcher = new FileSystemWatcher();
                fileSystemWatcher.Path = rule.watchPath;
                fileSystemWatcher.Filter = rule.filter;
                fileSystemWatcher.Deleted += FileSystemWatcherInvoke;
                fileSystemWatcher.Created += FileSystemWatcherInvoke;
                fileSystemWatcher.Changed += FileSystemWatcherInvoke;
                fileSystemWatcher.Renamed += FileSystemWatcherInvoke;
                fileSystemWatcher.IncludeSubdirectories = true;
                fileSystemWatcher.EnableRaisingEvents = true;
                fileSystemWatcherDic.Add(fileSystemWatcher, ruleName);

                btn_setauto.Visibility = Visibility.Hidden;
                btn_unsetauto.Visibility = Visibility.Visible;
            }
        }

        private void SetAutoStatusAll()
        {
            List<String> rulesList = Directory.GetFiles(".\\Rules", "*.json").ToList();
            foreach (string path in rulesList)
            {
                RunningRule rule = JsonConvert.DeserializeObject<RunningRule>(File.ReadAllText(path));

                if (rule.watchPath != null && rule.watchPath != "")
                {
                    if (!CheckRule(rule))
                    {
                        continue;
                    }

                    SetAuto(rule, Path.GetFileNameWithoutExtension(path));
                }
            }
        }

        private void UnsetAuto(object sender, RoutedEventArgs e)
        {
            string fileName = $".\\Rules\\{cb_rules.SelectedItem.ToString()}.json";

            RunningRule rule = JsonConvert.DeserializeObject<RunningRule>(File.ReadAllText(fileName));
            if (rule == null)
            {
                return;
            }
            rule.watchPath = null;
            rule.filter = null;

            string json = JsonConvert.SerializeObject(rule);
            FileStream fs = null;
            try
            {
                fs = File.Create(fileName);
                fs.Close();
                StreamWriter sw = File.CreateText(fileName);
                sw.Write(json);
                sw.Flush();
                sw.Close();
            }
            catch (Exception ex)
            {
                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            FileSystemWatcher fileSystemWatcher = fileSystemWatcherDic.FirstOrDefault(q => q.Value == cb_rules.SelectedItem.ToString()).Key;
            if (fileSystemWatcher != null)
            {
                fileSystemWatcher.EnableRaisingEvents = false;
                fileSystemWatcherDic.Remove(fileSystemWatcher);
            }

            btn_setauto.Visibility = Visibility.Visible;
            btn_unsetauto.Visibility = Visibility.Hidden;
        }

        private bool CheckRule(RunningRule rule)
        {
            if (rule.watchPath != null && rule.watchPath != "" && !Directory.Exists(rule.watchPath))
            {
                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"监视路径不存在: {rule.watchPath}", "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
            foreach (string analyzer in rule.analyzers.Split('\n'))
            {
                if (!cb_analyzers.Items.Contains(analyzer))
                {
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"存在未知的Analyzer: {analyzer}", "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
            }
            foreach (string sheetExplainer in rule.sheetExplainers.Split('\n'))
            {
                if (!cb_sheetexplainers.Items.Contains(sheetExplainer))
                {
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"存在未知的SheetExplainers: {sheetExplainer}", "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }

            }
            return true;
        }

        private void WindowClosed(object sender, EventArgs e)
        {
            FileHelper.DeleteCopiedDlls(copiedDllsList);
        }

        private void EditParam(object sender, RoutedEventArgs e)
        {
            ParamEditor paramEditor = new ParamEditor();

            string paramStr = $"{te_params.Text.Replace("\r\n", "").Replace("\n", "")}";
            te_params.Text = paramStr;
            Dictionary<string, Dictionary<string, string>> paramDicEachAnalyzer = ParamHelper.GetParamDicEachAnalyzer(paramStr);

            int rowNum = -1;

            List<String> analyzersList = te_analyzers.Text.Split('\n').Where(str => str.Trim() != "").ToList();

            ColumnDefinition columnDefinitionL = new ColumnDefinition();
            columnDefinitionL.Width = new GridLength(100);
            paramEditor.g_main.ColumnDefinitions.Add(columnDefinitionL);
            ColumnDefinition columnDefinitionR = new ColumnDefinition();
            paramEditor.g_main.ColumnDefinitions.Add(columnDefinitionR);

            List<string> addedAnalyzerNameList = new List<string>();

            foreach (string analyzerName in analyzersList) 
            {
                if (addedAnalyzerNameList.Contains(analyzerName))
                {
                    continue;
                }
                else
                {
                    addedAnalyzerNameList.Add(analyzerName);
                }

                Analyzer analyzer = JsonConvert.DeserializeObject<Analyzer>(File.ReadAllText($".\\Analyzers\\{analyzerName}.json"));

                RowDefinition rowDefinition = new RowDefinition();
                rowDefinition.Height = new GridLength(40);
                paramEditor.g_main.RowDefinitions.Add(rowDefinition);
                ++rowNum;
                TextBlock textBlockAnalyzerName = new TextBlock();
                textBlockAnalyzerName.Height = 35;
                textBlockAnalyzerName.FontWeight = FontWeight.FromOpenTypeWeight(600);
                textBlockAnalyzerName.FontSize = 20;
                textBlockAnalyzerName.VerticalAlignment = VerticalAlignment.Bottom;
                textBlockAnalyzerName.Text = analyzerName;
                Grid.SetRow(textBlockAnalyzerName, rowNum);
                Grid.SetColumnSpan(textBlockAnalyzerName, 2);
                paramEditor.g_main.Children.Add(textBlockAnalyzerName);

                Dictionary<string, string> paramDic = null;
                if (paramDicEachAnalyzer.ContainsKey(analyzerName))
                {
                    paramDic = paramDicEachAnalyzer[analyzerName];
                }

                if (analyzer.paramDic == null || analyzer.paramDic.Keys == null || analyzer.paramDic.Keys.Count == 0)
                {
                    continue;
                }
                foreach (string key in analyzer.paramDic.Keys)
                {
                    RowDefinition rowDefinitionKv = new RowDefinition();
                    rowDefinitionKv.Height = new GridLength(30);
                    paramEditor.g_main.RowDefinitions.Add(rowDefinitionKv);
                    ++rowNum;

                    TextBlock textBlockKey = new TextBlock();
                    textBlockKey.Height = 25;
                    textBlockKey.Text = analyzer.paramDic[key];
                    Grid.SetRow(textBlockKey, rowNum);
                    Grid.SetColumn(textBlockKey, 0);
                    paramEditor.g_main.Children.Add(textBlockKey);

                    TextBox tbValue = new TextBox();
                    tbValue.Height = 25;
                    tbValue.Margin = new Thickness(5, 0, 0, 0);
                    tbValue.HorizontalAlignment = HorizontalAlignment.Stretch;
                    tbValue.VerticalContentAlignment = VerticalAlignment.Center;
                    if (paramDic != null && paramDic.ContainsKey(key))
                    {
                        tbValue.Text = paramDic[key];
                    }
                    else if (paramDicEachAnalyzer.ContainsKey("public") && paramDicEachAnalyzer["public"] != null && paramDicEachAnalyzer["public"].ContainsKey(key))
                    {
                        tbValue.Text = paramDicEachAnalyzer["public"][key];
                    }
                    tbValue.TextChanged += (s, ex) =>
                    {
                        if (!paramDicEachAnalyzer.ContainsKey(analyzerName))
                        {
                            paramDicEachAnalyzer.Add(analyzerName, new Dictionary<string, string>());
                        }
                        paramDicEachAnalyzer[analyzerName][key] = tbValue.Text;
                    };
                    Grid.SetRow(tbValue, rowNum);
                    Grid.SetColumn(tbValue, 1);

                    paramEditor.g_main.Children.Add(tbValue);
                }
            }

            GridSplitter gridSplitter = new GridSplitter();
            gridSplitter.HorizontalAlignment = HorizontalAlignment.Right;
            gridSplitter.VerticalAlignment = VerticalAlignment.Stretch;
            gridSplitter.Width = 4;
            gridSplitter.BorderThickness = new Thickness(1, 0, 1, 0);
            gridSplitter.BorderBrush = Brushes.Black;
            gridSplitter.Background = Brushes.AntiqueWhite;
            Grid.SetRow(gridSplitter, 0);
            Grid.SetRowSpan(gridSplitter, rowNum + 1 <= 0 ? 1 : rowNum + 1);
            Grid.SetColumn(gridSplitter, 0);
            paramEditor.g_main.Children.Add(gridSplitter);

            RowDefinition rowDefinitionBlank = new RowDefinition();
            paramEditor.g_main.RowDefinitions.Add(rowDefinitionBlank);
            ++rowNum;
            RowDefinition rowDefinitionOk = new RowDefinition();
            rowDefinitionOk.Height = new GridLength(30);
            paramEditor.g_main.RowDefinitions.Add(rowDefinitionOk);
            ++rowNum;
            Button btnOk = new Button();
            btnOk.Height = 30;
            btnOk.Content = "OK";
            btnOk.Click += (s, ex) => 
            {
                te_params.Text = ParamHelper.GetParamStr(paramDicEachAnalyzer);
                paramEditor.Close();
            };
            Grid.SetRow(btnOk, rowNum);
            Grid.SetColumnSpan(btnOk, 2);
            paramEditor.g_main.Children.Add(btnOk);

            paramEditor.ShowDialog();
        }

        
    }
}
