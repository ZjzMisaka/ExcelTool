using Amib.Threading;
using ClosedXML.Excel;
using CustomizableMessageBox;
using GlobalObjects;
using ICSharpCode.AvalonEdit.Search;
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
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelTool
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private enum ReadFileReturnType { ANALYZER, RESULT };
        private Boolean isStop;
        private SmartThreadPool smartThreadPool = null;
        private ConcurrentDictionary<string, long> currentAnalizingDictionary;
        private ConcurrentDictionary<ConcurrentDictionary<ResultType, Object>, Analyzer> resultList;
        private int analyzeSheetInvokeCount;
        private int setResultInvokeCount;
        private int totalTimeoutLimit;
        private int perTimeoutLimit;
        public MainWindow()
        {
            InitializeComponent();

            currentAnalizingDictionary = new ConcurrentDictionary<string, long>();
            resultList = new ConcurrentDictionary<ConcurrentDictionary<ResultType, Object>, Analyzer>();
            analyzeSheetInvokeCount = 0;
            setResultInvokeCount = 0;

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

            totalTimeoutLimit = IniHelper.GetTotalTimeoutLimit();
            perTimeoutLimit = IniHelper.GetPerTimeoutLimit();

            LoadFiles();
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
            resultList = new ConcurrentDictionary<ConcurrentDictionary<ResultType, Object>, Analyzer>();

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


            STPStartInfo stp = new STPStartInfo();
            stp.CallToPostExecute = CallToPostExecute.WhenWorkItemNotCanceled;
            stp.PostExecuteWorkItemCallback = delegate (IWorkItemResult wir)
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
            RenewSmartThreadPool(stp);


            List<SheetExplainer> sheetExplainers = new List<SheetExplainer>();
            List<Analyzer> analyzers = new List<Analyzer>();
            foreach (string name in sheetExplainersList)
            {
                SheetExplainer sheetExplainer = null;
                try
                {
                    sheetExplainer = JsonConvert.DeserializeObject<SheetExplainer>(File.ReadAllText($".\\SheetExplainers\\{name}.json"));
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
                
                foreach (String str in sheetExplainer.relativePathes) 
                {
                    List<String> filePathList = new List<String>();
                    String basePath = $"{tb_base_path.Text}{str}".Trim();

                    FileTraverse(basePath, sheetExplainer, filePathList);
                    totalCount += filePathList.Count;

                    int filesCount = filePathList.Count;
                    foreach (String filePath in filePathList)
                    { 
                        List<String> pathSplit = filePath.Split('\\').ToList<string>();
                        String fileName = pathSplit[pathSplit.Count - 1];
                        fileName = fileName.Substring(0, fileName.LastIndexOf('.'));
                        smartThreadPool.QueueWorkItem(new Func<String, String, SheetExplainer, Analyzer, Object>(ReadFile), filePath, fileName, sheetExplainer, analyzer);
                    }
                }
            }

            long startSs = GetNowSs();
            smartThreadPool.Join();
            while (smartThreadPool.CurrentWorkItemsCount > 0)
            {
                long nowSs = GetNowSs();
                long totalTimeCostSs = nowSs - startSs;
                if (isStop || totalTimeCostSs >= totalTimeoutLimit)
                {
                    smartThreadPool.Dispose();
                    this.Dispatcher.Invoke(() =>
                    {
                        btn_start.IsEnabled = true;
                        btn_stop.IsEnabled = false;
                    });
                    if (totalTimeCostSs >= totalTimeoutLimit)
                    {
                        CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "OK" }, $"Total time out. \n{totalTimeoutLimit / 1000.0}(s)", "error", MessageBoxImage.Error);
                    }
                    return;
                }

                StringBuilder sb = new StringBuilder();

                ICollection<string> keys = currentAnalizingDictionary.Keys;
                foreach (string key in keys)
                {
                    long value = new long();
                    if (currentAnalizingDictionary.TryGetValue(key, out value))
                    {
                        long timeCostSs = GetNowSs() - currentAnalizingDictionary[key];
                        if (timeCostSs >= perTimeoutLimit)
                        {
                            smartThreadPool.Dispose();
                            this.Dispatcher.Invoke(() =>
                            {
                                btn_start.IsEnabled = true;
                                btn_stop.IsEnabled = false;
                            });
                            CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "OK" }, $"{key}\nTime out. \n{perTimeoutLimit / 1000.0}(s)", "error", MessageBoxImage.Error);
                            return;
                        }
                        sb.Append(key.Substring(key.LastIndexOf('\\') + 1)).Append("(").Append((timeCostSs / 1000.0).ToString("0.0")).Append(")");
                    }
                }

                await Task.Run(() =>
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        l_process.Content = $"{smartThreadPool.CurrentWorkItemsCount}/{totalCount} -- ActiveThreads: {smartThreadPool.ActiveThreads} -- InUseThreads: {smartThreadPool.InUseThreads}";
                        tb_log.Text = $"{sb.ToString()}";
                    });
                });
                Thread.Sleep(100);
            }
            await Task.Run(() =>
            {
                this.Dispatcher.Invoke(() =>
                {
                    l_process.Content = $"0/{totalCount}";
                    tb_log.Text = $"";
                });
            });
            smartThreadPool.Dispose();

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

                foreach(ConcurrentDictionary< ResultType, Object > result in resultList.Keys)
                {
                    SetResult(workbook, result, resultList[result], resultList.Count);
                }
                


                bool saveResult = false;
                SaveFile(filePath, workbook, out saveResult);
                if (!saveResult)
                {
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), "OK" }, "file not saved.", "info");
                    this.Dispatcher.Invoke(() =>
                    {
                        btn_start.IsEnabled = true;
                        btn_stop.IsEnabled = false;
                    });
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
            this.Dispatcher.Invoke(() =>
            {
                btn_start.IsEnabled = true;
                btn_stop.IsEnabled = false;
            });
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

        private ConcurrentDictionary<ReadFileReturnType, Object> ReadFile(String filePath, String fileName, SheetExplainer sheetExplainer, Analyzer analyzer)
        {
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
                            Analyze(sheet, result, analyzer);
                        }
                    }
                }
                else if (sheetExplainer.sheetNames.Key == FindingMethod.CONTAIN)
                {
                    foreach (String str in sheetExplainer.sheetNames.Value)
                    {
                        if (sheet.Name.Contains(str))
                        {
                            Analyze(sheet, result, analyzer);
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
                            Analyze(sheet, result, analyzer);
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
        }

        private long GetNowSs()
        {
            return (DateTime.Now.ToUniversalTime().Ticks - 621355968000000000) / 10000;
        }

        private void Analyze(IXLWorksheet sheet, ConcurrentDictionary<ResultType, Object> result, Analyzer analyzer) 
        {
            CSharpCodeProvider objCSharpCodePrivoder = new CSharpCodeProvider();

            CompilerParameters objCompilerParameters = new CompilerParameters();

            objCompilerParameters.ReferencedAssemblies.Add("System.dll");
            objCompilerParameters.ReferencedAssemblies.Add("System.Data.dll");
            objCompilerParameters.ReferencedAssemblies.Add("System.Drawing.dll");
            objCompilerParameters.ReferencedAssemblies.Add("ClosedXML.dll");
            objCompilerParameters.ReferencedAssemblies.Add("System.Xml.dll");
            objCompilerParameters.ReferencedAssemblies.Add("GlobalObjects.dll");

            objCompilerParameters.GenerateExecutable = false;
            objCompilerParameters.GenerateInMemory = true;
            objCompilerParameters.IncludeDebugInformation = true;

            CompilerResults cresult = objCSharpCodePrivoder.CompileAssemblyFromSource(objCompilerParameters, analyzer.code);

            if (cresult.Errors.HasErrors)
            {

                String str = "[ERROR: ]";
                foreach (CompilerError err in cresult.Errors)
                {
                    str = $"{str}\n    {err.ErrorText} in line {err.Line}"; 
                }
                this.Dispatcher.Invoke(() =>
                {
                    tb_log.Text = str;
                    tb_log.Text = tb_log.Text.Substring(tb_log.Text.IndexOf("[ERROR: ]"));
                });
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
                    objMI.Invoke(obj, new object[] { sheet, result, GlobalObjects.GlobalObjects.GetGlobalParam(), analyzeSheetInvokeCount });
                }
                catch (Exception e)
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        tb_log.Text += $"[ERROR: ]\n    {e.InnerException.Message}\n    {analyzer.name}.AnalyzeSheet(): {e.InnerException.StackTrace.Substring(e.InnerException.StackTrace.LastIndexOf(':') + 1)}\n";
                        tb_log.Text = tb_log.Text.Substring(tb_log.Text.IndexOf("[ERROR: ]"));
                    });
                    Stop();
                }
            }

        }

        private void SetResult(XLWorkbook workbook, ConcurrentDictionary<ResultType, Object> result, Analyzer analyzer, int totalCount)
        {
            CSharpCodeProvider objCSharpCodePrivoder = new CSharpCodeProvider();

            CompilerParameters objCompilerParameters = new CompilerParameters();

            objCompilerParameters.ReferencedAssemblies.Add("System.dll");
            objCompilerParameters.ReferencedAssemblies.Add("System.Data.dll");
            objCompilerParameters.ReferencedAssemblies.Add("System.Drawing.dll");
            objCompilerParameters.ReferencedAssemblies.Add("ClosedXML.dll");
            objCompilerParameters.ReferencedAssemblies.Add("System.Xml.dll");
            objCompilerParameters.ReferencedAssemblies.Add("GlobalObjects.dll");

            objCompilerParameters.GenerateExecutable = false;
            objCompilerParameters.GenerateInMemory = true;
            objCompilerParameters.IncludeDebugInformation = true;

            CompilerResults cresult = objCSharpCodePrivoder.CompileAssemblyFromSource(objCompilerParameters, analyzer.code);

            if (cresult.Errors.HasErrors)
            {
                String str = "[ERROR: ]";
                foreach (CompilerError err in cresult.Errors)
                {
                    str = $"{str}\n    {err.ErrorText} in line {err.Line}";
                }
                this.Dispatcher.Invoke(() =>
                {
                    tb_log.Text = str;
                    tb_log.Text = tb_log.Text.Substring(tb_log.Text.IndexOf("[ERROR: ]"));
                });
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
                    objMI.Invoke(obj, new object[] { workbook, result, GlobalObjects.GlobalObjects.GetGlobalParam(), setResultInvokeCount, totalCount });
                }
                catch (Exception e)
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        tb_log.Text += $"[ERROR: ]\n    {e.InnerException.Message}\n    {analyzer.name}.SetResult(): {e.InnerException.StackTrace.Substring(e.InnerException.StackTrace.LastIndexOf(':') + 1)}\n";
                        tb_log.Text = tb_log.Text.Substring(tb_log.Text.IndexOf("[ERROR: ]"));
                    });
                    Stop();
                }
            }

        }

        private void WindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            IniHelper.SetWindowSize(this.Name.Substring(2), new Point(this.Width, this.Height));
            IniHelper.SetBasePath(tb_base_path.Text);
            IniHelper.SetOutputPath(tb_output_path.Text);
            IniHelper.SetOutputFileName(tb_output_name.Text);
        }

        private void RenewSmartThreadPool(STPStartInfo stp)
        {
            smartThreadPool = new SmartThreadPool(stp);
            int maxThreadCount = IniHelper.GetMaxThreadCount();
            if (maxThreadCount > 0)
            {
                smartThreadPool.MaxThreads = maxThreadCount;
            }
        }
    }
}
