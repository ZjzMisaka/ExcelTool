﻿using Amib.Threading;
using ExcelTool.Helper;
using GlobalObjects;
using GongSolutions.Wpf.DragDrop;
using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Highlighting;
using Microsoft.Toolkit.Mvvm.ComponentModel;
using Microsoft.Toolkit.Mvvm.Input;
using Newtonsoft.Json;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Xml;
using ClosedXML.Excel;
using Microsoft.CSharp;
using System.CodeDom.Compiler;
using System.Text.RegularExpressions;
using CustomizableMessageBox;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace ExcelTool.ViewModel
{
    class MainWindowViewModel : ObservableObject, IDropTarget
    {
        private enum ReadFileReturnType { ANALYZER, RESULT };
        private Boolean isRunning;
        private Boolean isStopByUser;
        private SmartThreadPool smartThreadPoolAnalyze = null;
        private SmartThreadPool smartThreadPoolOutput = null;
        private Thread runningThread;
        private Thread runBeforeAnalyzeSheetThread;
        private Thread runBeforeSetResultThread;
        private Thread runEndThread;
        private Thread FileSystemWatcherInvokeThread;
        private ConcurrentDictionary<string, long> currentAnalizingDictionary;
        private ConcurrentDictionary<string, long> currentOutputtingDictionary;
        private ConcurrentDictionary<ConcurrentDictionary<ResultType, Object>, Analyzer> resultList;
        private Dictionary<FileSystemWatcher, string> fileSystemWatcherDic;
        private Stopwatch stopwatchBeforeFileSystemWatcherInvoke;
        private List<string> copiedDllsList;
        private int analyzeSheetInvokeCount;
        private int setResultInvokeCount;
        private int totalTimeoutLimitAnalyze;
        private int perTimeoutLimitAnalyze;
        private int totalTimeoutLimitOutput;
        private int perTimeoutLimitOutput;
        private bool runNotSuccessed;
        private Scanner scanner;
        private int fileSystemWatcherInvokeDalay;
        private int freshInterval;
        private string language;

        private TextEditor teLog = new TextEditor();
        public TextEditor TeLog => teLog;

        private TextEditor teParams = new TextEditor();
        public TextEditor TeParams => teParams;

        private double windowWidth;
        public double WindowWidth
        {
            get { return windowWidth; }
            set
            {
                SetProperty<double>(ref windowWidth, value);
            }
        }

        private double windowHeight;
        public double WindowHeight
        {
            get { return windowHeight; }
            set
            {
                SetProperty<double>(ref windowHeight, value);
            }
        }

        private string windowName = "MainWindow";
        public string WindowName
        {
            get { return windowName; }
            set
            {
                SetProperty<string>(ref windowName, value);
            }
        }

        private List<string> sheetExplainersItems = new List<string>();
        public List<string> SheetExplainersItems
        {
            get { return sheetExplainersItems; }
            set
            {
                SetProperty<List<string>>(ref sheetExplainersItems, value);
            }
        }

        private string selectedSheetExplainersItem = null;
        public string SelectedSheetExplainersItem
        {
            get { return selectedSheetExplainersItem; }
            set
            {
                SetProperty<string>(ref selectedSheetExplainersItem, value);
            }
        }

        private int selectedSheetExplainersIndex = 0;
        public int SelectedSheetExplainersIndex
        {
            get { return selectedSheetExplainersIndex; }
            set
            {
                SetProperty<int>(ref selectedSheetExplainersIndex, value);
            }
        }

        private List<string> analyzersItems = new List<string>();
        public List<string> AnalyzersItems
        {
            get { return analyzersItems; }
            set
            {
                SetProperty<List<string>>(ref analyzersItems, value);
            }
        }

        private string selectedAnalyzersItem = null;
        public string SelectedAnalyzersItem
        {
            get { return selectedAnalyzersItem; }
            set
            {
                SetProperty<string>(ref selectedAnalyzersItem, value);
            }
        }

        private int selectedAnalyzersIndex = 0;
        public int SelectedAnalyzersIndex
        {
            get { return selectedAnalyzersIndex; }
            set
            {
                SetProperty<int>(ref selectedAnalyzersIndex, value);
            }
        }

        private TextDocument teSheetExplainersDocument = new TextDocument();
        public TextDocument TeSheetExplainersDocument
        {
            get => teSheetExplainersDocument;
            set
            {
                SetProperty<TextDocument>(ref teSheetExplainersDocument, value);
            }
        }

        private TextDocument teAnalyzersDocument = new TextDocument();
        public TextDocument TeAnalyzersDocument
        {
            get => teAnalyzersDocument;
            set
            {
                SetProperty<TextDocument>(ref teAnalyzersDocument, value);
            }
        }

        private List<string> paramsItems = new List<string>();
        public List<string> ParamsItems
        {
            get { return paramsItems; }
            set
            {
                SetProperty<List<string>>(ref paramsItems, value);
            }
        }

        private string selectedParamsItem = null;
        public string SelectedParamsItem
        {
            get { return selectedParamsItem; }
            set
            {
                SetProperty<string>(ref selectedParamsItem, value);
            }
        }

        private int selectedParamssIndex = 0;
        public int SelectedParamsIndex
        {
            get { return selectedParamssIndex; }
            set
            {
                SetProperty<int>(ref selectedParamssIndex, value);
            }
        }

        private string tbBasePathText = null;
        public string TbBasePathText
        {
            get => tbBasePathText;
            set
            {
                SetProperty<string>(ref tbBasePathText, value);
            }
        }

        private string tbOutputPathText = null;
        public string TbOutputPathText
        {
            get => tbOutputPathText;
            set
            {
                SetProperty<string>(ref tbOutputPathText, value);
            }
        }

        private string tbOutputNameText = null;
        public string TbOutputNameText
        {
            get => tbOutputNameText;
            set
            {
                SetProperty<string>(ref tbOutputNameText, value);
            }
        }

        private List<string> ruleItems = null;
        public List<string> RuleItems
        {
            get { return ruleItems; }
            set
            {
                SetProperty<List<string>>(ref ruleItems, value);
            }
        }

        private string selectedRulesItem = null;
        public string SelectedRulesItem
        {
            get { return selectedRulesItem; }
            set
            {
                SetProperty<string>(ref selectedRulesItem, value);
            }
        }

        private int selectedRulesIndex = 0;
        public int SelectedRulesIndex
        {
            get { return selectedRulesIndex; }
            set
            {
                SetProperty<int>(ref selectedRulesIndex, value);
            }
        }

        private Visibility btnSaveRuleVisibility = Visibility.Hidden;
        public Visibility BtnSaveRuleVisibility
        {
            get { return btnSaveRuleVisibility; }
            set
            {
                SetProperty<Visibility>(ref btnSaveRuleVisibility, value);
            }
        }

        private Visibility btnDeleteRuleVisibility = Visibility.Visible;
        public Visibility BtnDeleteRuleVisibility
        {
            get { return btnDeleteRuleVisibility; }
            set
            {
                SetProperty<Visibility>(ref btnDeleteRuleVisibility, value);
            }
        }

        private bool btnDeleteRuleIsEnabled;
        public bool BtnDeleteRuleIsEnabled
        {
            get { return btnDeleteRuleIsEnabled; }
            set
            {
                SetProperty<bool>(ref btnDeleteRuleIsEnabled, value);
            }
        }

        private Visibility btnSetAutoVisibility = Visibility.Hidden;
        public Visibility BtnSetAutoVisibility
        {
            get { return btnSetAutoVisibility; }
            set
            {
                SetProperty<Visibility>(ref btnSetAutoVisibility, value);
            }
        }

        private Visibility btnUnsetAutoVisibility = Visibility.Hidden;
        public Visibility BtnUnsetAutoVisibility
        {
            get { return btnUnsetAutoVisibility; }
            set
            {
                SetProperty<Visibility>(ref btnUnsetAutoVisibility, value);
            }
        }

        private string tbStatusText = null;
        public string TbStatusText
        {
            get { return tbStatusText; }
            set
            {
                SetProperty<string>(ref tbStatusText, value);
            }
        }

        private string lProcessContent;
        public string LProcessContent
        {
            get { return lProcessContent; }
            set
            {
                SetProperty<string>(ref lProcessContent, value);
            }
        }

        private bool cbIsAutoOpenIsChecked;
        public bool CbIsAutoOpenIsChecked
        {
            get { return cbIsAutoOpenIsChecked; }
            set
            {
                SetProperty<bool>(ref cbIsAutoOpenIsChecked, value);
            }
        }

        private bool btnStopIsEnabled = false;
        public bool BtnStopIsEnabled
        {
            get { return btnStopIsEnabled; }
            set
            {
                SetProperty<bool>(ref btnStopIsEnabled, value);
            }
        }

        private bool btnStartIsEnabled = true;
        public bool BtnStartIsEnabled
        {
            get { return btnStartIsEnabled; }
            set
            {
                SetProperty<bool>(ref btnStartIsEnabled, value);
            }
        }

        public ICommand WindowLoadedCommand { get; set; }
        public ICommand WindowClosingCommand { get; set; }
        public ICommand WindowClosedCommand { get; set; }
        public ICommand BtnOpenSheetExplainerEditorClickCommand { get; set; }
        public ICommand BtnOpenAnalyzerEditorClickCommand { get; set; }
        public ICommand CbSheetExplainersPreviewMouseLeftButtonDownCommand { get; set; }
        public ICommand CbSheetExplainersSelectionChangedCommand { get; set; }
        public ICommand CbAnalyzersPreviewMouseLeftButtonDownCommand { get; set; }
        public ICommand CbAnalyzersSelectionChangedCommand { get; set; }
        public ICommand CbParamsPreviewMouseLeftButtonDownCommand { get; set; }
        public ICommand CbParamsSelectionChangedCommand { get; set; }
        public ICommand EditParamCommand { get; set; }
        public ICommand TbParamsLostFocusCommand { get; set; }
        public ICommand SelectPathCommand { get; set; }
        public ICommand OpenPathCommand { get; set; }
        public ICommand SelectNameCommand { get; set; }
        public ICommand CbRulesChangedCommand { get; set; }
        public ICommand CbRulesPreviewMouseLeftButtonDownCommand { get; set; }
        public ICommand SaveRuleCommand { get; set; }
        public ICommand DeleteRuleCommand { get; set; }
        public ICommand SetAutoCommand { get; set; }
        public ICommand UnsetAutoCommand { get; set; }
        public ICommand StopCommand { get; set; }
        public ICommand StartCommand { get; set; }
        public MainWindowViewModel()
        {
            ICSharpCode.AvalonEdit.Search.SearchPanel.Install(TeLog);

            currentAnalizingDictionary = new ConcurrentDictionary<string, long>();
            currentOutputtingDictionary = new ConcurrentDictionary<string, long>();
            resultList = new ConcurrentDictionary<ConcurrentDictionary<ResultType, Object>, Analyzer>();
            fileSystemWatcherDic = new Dictionary<FileSystemWatcher, string>();
            copiedDllsList = new List<string>();
            scanner = new Scanner();

            analyzeSheetInvokeCount = 0;
            setResultInvokeCount = 0;

            isRunning = false;
            isStopByUser = false;
            runNotSuccessed = false;

            scanner.Input += Scanner.UserInput;

            TeLog.ShowLineNumbers = false;
            TeLog.WordWrap = true;
            TeLog.IsReadOnly = true;
            TeLog.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            TeLog.PreviewKeyDown += TeLogPreviewKeyDown;
            TeLog.TextChanged += TeLogTextChanged;
            TeParams.ShowLineNumbers = false;
            TeParams.WordWrap = false;
            TeParams.HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden;
            TeParams.VerticalScrollBarVisibility = ScrollBarVisibility.Hidden;
            TeParams.TextChanged += TeParamsTextChanged;
            TeParams.PreviewKeyDown += TeParamsPreviewKeyDown;
            TeParams.BorderBrush = Brushes.Gray;
            TeParams.BorderThickness = new Thickness(1);
            TeParams.Padding = new Thickness(3);
            WindowLoadedCommand = new RelayCommand<RoutedEventArgs>(WindowLoaded);
            WindowClosingCommand = new RelayCommand<CancelEventArgs>(WindowClosing);
            WindowClosedCommand = new RelayCommand(WindowClosed);
            BtnOpenSheetExplainerEditorClickCommand = new RelayCommand(BtnOpenSheetExplainerEditorClick);
            BtnOpenAnalyzerEditorClickCommand = new RelayCommand(BtnOpenAnalyzerEditorClick);
            CbSheetExplainersPreviewMouseLeftButtonDownCommand = new RelayCommand(CbSheetExplainersPreviewMouseLeftButtonDown);
            CbSheetExplainersSelectionChangedCommand = new RelayCommand(CbSheetExplainersSelectionChanged);
            CbAnalyzersPreviewMouseLeftButtonDownCommand = new RelayCommand(CbAnalyzersPreviewMouseLeftButtonDown);
            CbAnalyzersSelectionChangedCommand = new RelayCommand(CbAnalyzersSelectionChanged);
            CbParamsPreviewMouseLeftButtonDownCommand = new RelayCommand(CbParamsPreviewMouseLeftButtonDown);
            CbParamsSelectionChangedCommand = new RelayCommand(CbParamsSelectionChanged);
            EditParamCommand = new RelayCommand(EditParam);
            TbParamsLostFocusCommand = new RelayCommand(TbParamsLostFocus);
            SelectPathCommand = new RelayCommand<object>(SelectPath);
            OpenPathCommand = new RelayCommand(OpenPath);
            SelectNameCommand = new RelayCommand(SelectName);
            CbRulesChangedCommand = new RelayCommand(CbRulesChanged);
            CbRulesPreviewMouseLeftButtonDownCommand = new RelayCommand(CbRulesPreviewMouseLeftButtonDown);
            SaveRuleCommand = new RelayCommand(SaveRule);
            DeleteRuleCommand = new RelayCommand(DeleteRule);
            SetAutoCommand = new RelayCommand(SetAuto);
            UnsetAutoCommand = new RelayCommand(UnsetAuto);
            StopCommand = new RelayCommand(Stop);
            StartCommand = new RelayCommand(Start);
        }

        private void WindowLoaded(RoutedEventArgs eventArgs)
        {
            Window window = (Window)eventArgs.Source;

            language = IniHelper.GetLanguage();
            if (String.IsNullOrWhiteSpace(language))
            {
                language = Thread.CurrentThread.CurrentUICulture.Name;
            }
            List<ResourceDictionary> dictionaryList = new List<ResourceDictionary>();
            foreach (ResourceDictionary dictionary in Application.Current.Resources.MergedDictionaries)
            {
                dictionaryList.Add(dictionary);
            }
            string requestedCulture = string.Format(@"Resources\StringResource.{0}.xaml", language);
            ResourceDictionary resourceDictionary = dictionaryList.FirstOrDefault(d => d.Source.OriginalString.Equals(requestedCulture));
            if (resourceDictionary == null)
            {
                requestedCulture = @"Resources\StringResource.zh-CN.xaml";
                resourceDictionary = dictionaryList.FirstOrDefault(d => d.Source.OriginalString.Equals(requestedCulture));
            }
            if (resourceDictionary != null)
            {
                Application.Current.Resources.MergedDictionaries.Remove(resourceDictionary);
                Application.Current.Resources.MergedDictionaries.Add(resourceDictionary);
            }

            IniHelper.CheckAndCreateIniFile();

            double width = IniHelper.GetWindowSize(WindowName).X;
            double height = IniHelper.GetWindowSize(WindowName).Y;
            if (width > 0)
            {
                window.Dispatcher.Invoke(() => {
                    window.Width = width;
                });
            }

            if (height > 0)
            {
                window.Dispatcher.Invoke(() => {
                    window.Height = height;
                });
            }

            TbBasePathText = IniHelper.GetBasePath();
            TbOutputPathText = IniHelper.GetOutputPath();
            TbOutputNameText = IniHelper.GetOutputFileName();

            totalTimeoutLimitAnalyze = IniHelper.GetTotalTimeoutLimitAnalyze();
            perTimeoutLimitAnalyze = IniHelper.GetPerTimeoutLimitAnalyze();
            totalTimeoutLimitOutput = IniHelper.GetTotalTimeoutLimitOutput();
            perTimeoutLimitOutput = IniHelper.GetPerTimeoutLimitOutput();

            CbIsAutoOpenIsChecked = IniHelper.GetIsAutoOpen();

            fileSystemWatcherInvokeDalay = IniHelper.GetFileSystemWatcherInvokeDalay();
            freshInterval = IniHelper.GetFreshInterval();

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
            TeParams.Dispatcher.Invoke(() =>
            {
                TeParams.SyntaxHighlighting = HighlightingManager.Instance.GetDefinition("param");
            });

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
            TeLog.SyntaxHighlighting = HighlightingManager.Instance.GetDefinition("log");

            runningThread = new Thread(() => WhenRunning());
            runningThread.Start();
        }
        private void WindowClosing(CancelEventArgs eventArgs)
        {
            if (Application.Current.Windows.Count > 1)
            {
                MessageBoxResult result = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), Application.Current.FindResource("ProgramClosingCheck").ToString(), Application.Current.FindResource("Warning").ToString(), MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result == MessageBoxResult.No)
                {
                    eventArgs.Cancel = true;
                }
                else
                {
                    foreach (Window currentWindow in Application.Current.Windows)
                    {
                        if (Application.Current.MainWindow != currentWindow)
                        {
                            currentWindow.Close();
                        }
                    }
                }
            }

            IniHelper.SetWindowSize(WindowName, new Point(WindowWidth, WindowHeight));
            IniHelper.SetBasePath(TbBasePathText);
            IniHelper.SetOutputPath(TbOutputPathText);
            IniHelper.SetOutputFileName(TbOutputNameText);
            IniHelper.SetLanguage(language);

            if (CbIsAutoOpenIsChecked == true)
            {
                IniHelper.SetIsAutoOpen(true);
            }
            else
            {
                IniHelper.SetIsAutoOpen(false);
            }
        }

        private void WindowClosed()
        {
            CheckAndCloseThreads(true);

            FileHelper.DeleteCopiedDlls(copiedDllsList);
        }

        private void BtnOpenSheetExplainerEditorClick()
        {
            SheetExplainerEditor sheetExplainerEditor = new SheetExplainerEditor();
            sheetExplainerEditor.Show();
        }

        private void BtnOpenAnalyzerEditorClick()
        {
            AnalyzerEditor analyzerEditor = new AnalyzerEditor();
            analyzerEditor.Show();
        }

        private void CbSheetExplainersPreviewMouseLeftButtonDown()
        {
            SheetExplainersItems = FileHelper.GetSheetExplainersList();
        }

        private void CbSheetExplainersSelectionChanged()
        {
            if (SelectedSheetExplainersIndex == 0)
            {
                return;
            }
            TeSheetExplainersDocument = new TextDocument($"{TeSheetExplainersDocument.Text}{SelectedSheetExplainersItem}\n");
            SelectedSheetExplainersIndex = 0;
        }

        private void CbAnalyzersPreviewMouseLeftButtonDown()
        {
            AnalyzersItems = FileHelper.GetAnalyzersList();
        }
        private void CbAnalyzersSelectionChanged()
        {
            if (SelectedAnalyzersIndex == 0)
            {
                return;
            }
            TeAnalyzersDocument = new TextDocument($"{TeAnalyzersDocument.Text}{SelectedAnalyzersItem}\n");
            SelectedAnalyzersIndex = 0;
        }

        private void CbParamsPreviewMouseLeftButtonDown()
        {
            paramsItems = FileHelper.GetParamsList();
        }

        private void CbParamsSelectionChanged()
        {
            if (CbParamsSelectionChangedCommand != null)
            {
                TeParams.Dispatcher.Invoke(() =>
                {
                    TeParams.TextChanged -= TeParamsTextChanged;
                    TeParams.Document = new TextDocument(SelectedParamsItem);
                    TeParams.TextChanged += TeParamsTextChanged;
                });
            }
        }

        private void EditParam()
        {
            ParamEditor paramEditor = new ParamEditor();

            string paramStr = "";
            TeParams.Dispatcher.Invoke(() =>
            {
                paramStr = $"{TeParams.Text.Replace("\r\n", "").Replace("\n", "")}";
                TeParams.Text = paramStr;
            });
            Dictionary<string, Dictionary<string, string>> paramDicEachAnalyzer = ParamHelper.GetParamDicEachAnalyzer(paramStr);

            int rowNum = -1;

            List<String> analyzersList = TeAnalyzersDocument.Text.Split('\n').Where(str => str.Trim() != "").ToList();

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
                    textBlockKey.MouseEnter += (s, ex) =>
                    {
                        textBlockKey.Text = key;
                    };
                    textBlockKey.TouchEnter += (s, ex) =>
                    {
                        textBlockKey.Text = key;
                    };
                    textBlockKey.MouseLeave += (s, ex) =>
                    {
                        textBlockKey.Text = analyzer.paramDic[key];
                    };
                    textBlockKey.TouchLeave += (s, ex) =>
                    {
                        textBlockKey.Text = analyzer.paramDic[key];
                    };
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
            btnOk.Content = Application.Current.FindResource("Ok").ToString();
            btnOk.Click += (s, ex) =>
            {
                TeParams.Text = ParamHelper.GetParamStr(paramDicEachAnalyzer);
                paramEditor.Close();
            };
            Grid.SetRow(btnOk, rowNum);
            Grid.SetColumnSpan(btnOk, 2);
            paramEditor.g_main.Children.Add(btnOk);

            paramEditor.ShowDialog();
        }

        private void TeParamsTextChanged(object sender, EventArgs e)
        {
            CbParamsSelectionChangedCommand = null;
            SelectedParamsIndex = 0;
            CbParamsSelectionChangedCommand = new RelayCommand(CbParamsSelectionChanged);
        }

        private void TeParamsPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
            }
        }

        private void TbParamsLostFocus()
        {
            TeParams.Dispatcher.Invoke(() =>
            {
                TeParams.Text = $"{TeParams.Text.Replace("\r\n", "").Replace("\n", "")}";
            });
        }

        public void DragEnter(IDropInfo dropInfo)
        {

        }

        public void DragOver(IDropInfo dropInfo)
        {
            dropInfo.DropTargetAdorner = DropTargetAdorners.Highlight;
            dropInfo.Effects = DragDropEffects.Copy;
        }

        public void DragLeave(IDropInfo dropInfo)
        {

        }

        public void Drop(IDropInfo dropInfo)
        {
            if (dropInfo.DragInfo?.VisualSource is null
                && dropInfo.Data is DataObject dataObject
                && dataObject.GetDataPresent(DataFormats.FileDrop)
                && dataObject.ContainsFileDropList())
            {
                string parentName = ((TextBox)((Border)dropInfo.TargetScrollViewer.Parent).TemplatedParent).Name;
                if (parentName == "tb_base_path")
                {
                    TbBasePathText = dataObject.GetFileDropList()[0];
                }
                else if(parentName == "tb_output_path")
                {
                    TbOutputPathText = dataObject.GetFileDropList()[0];
                }
                else if (parentName == "tb_output_name")
                {
                    TbOutputNameText = Path.GetFileNameWithoutExtension(dataObject.GetFileDropList()[0]);
                }
            }
            else
            {
                GongSolutions.Wpf.DragDrop.DragDrop.DefaultDropHandler.Drop(dropInfo);
            }
        }

        private void SelectPath(object sender)
        {
            System.Windows.Forms.FolderBrowserDialog openFileDialog = new System.Windows.Forms.FolderBrowserDialog();

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (((Button)sender).Name == "btn_select_output_path")
                {
                    TbBasePathText = openFileDialog.SelectedPath;
                }
                else
                {
                    TbOutputPathText = openFileDialog.SelectedPath;
                }
            }
        }

        private void OpenPath()
        {
            string resPath = TbOutputPathText.Replace("\\", "/");
            string filePath;
            if (TbOutputPathText.EndsWith("/"))
            {
                filePath = $"{resPath}{TbOutputNameText}.xlsx";
            }
            else
            {
                filePath = $"{resPath}/{TbOutputNameText}.xlsx";
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

        private void SelectName()
        {
            System.Windows.Forms.OpenFileDialog fileDialog = new System.Windows.Forms.OpenFileDialog();
            fileDialog.Multiselect = false;
            fileDialog.Title = Application.Current.FindResource("SelectFile").ToString();
            fileDialog.Filter = $"Excel{Application.Current.FindResource("File").ToString()}|*.xlsx;*.xlsm;*.xls|All files(*.*)|*.*";
            if (fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string[] names = fileDialog.FileNames;
                TbOutputNameText = Path.GetFileNameWithoutExtension(names[0]);
            }
        }

        private void CbRulesChanged()
        {

            RunningRule rule = null;

            if (SelectedRulesIndex >= 1)
            {
                rule = JsonConvert.DeserializeObject<RunningRule>(File.ReadAllText($".\\Rules\\{SelectedRulesItem}.json"));

                if (!CheckRule(rule))
                {
                    SelectedRulesIndex = 0;
                    return;
                }

                BtnSaveRuleVisibility = Visibility.Hidden;
                BtnDeleteRuleVisibility = Visibility.Visible;

                if (rule.watchPath != null && rule.watchPath != "")
                {
                    BtnSetAutoVisibility = Visibility.Hidden;
                    BtnUnsetAutoVisibility = Visibility.Visible;
                    BtnDeleteRuleIsEnabled = false;
                }
                else
                {
                    BtnSetAutoVisibility = Visibility.Visible;
                    BtnUnsetAutoVisibility = Visibility.Hidden;
                    BtnDeleteRuleIsEnabled = true;
                }

                TeAnalyzersDocument = new TextDocument(rule.analyzers);
                TeSheetExplainersDocument = new TextDocument(rule.sheetExplainers);
                TeParams.Dispatcher.Invoke(() =>
                {
                    TeParams.Text = rule.param;
                });

                TbBasePathText = rule.basePath;
                TbOutputPathText = rule.outputPath;
                TbOutputNameText = rule.outputName;
            }
            else
            {
                BtnSaveRuleVisibility = Visibility.Visible;
                BtnDeleteRuleVisibility = Visibility.Hidden;

                BtnSetAutoVisibility = Visibility.Hidden;
                BtnUnsetAutoVisibility = Visibility.Hidden;
            }
        }

        private void CbRulesPreviewMouseLeftButtonDown()
        {
            RuleItems = FileHelper.GetRulesList();
        }
        private void SaveRule()
        {
            TextBox tbName = new TextBox();
            tbName.Margin = new Thickness(5, 8, 5, 8);
            tbName.VerticalContentAlignment = VerticalAlignment.Center;
            if (SelectedRulesIndex >= 1)
            {
                tbName.Text = $"Copy Of {SelectedRulesItem}";
            }
            int result = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object>() { tbName, Application.Current.FindResource("Ok").ToString(), Application.Current.FindResource("Cancel").ToString() }, Application.Current.FindResource("Name").ToString(), Application.Current.FindResource("Saving").ToString(), MessageBoxImage.Information);
            if (result == 1)
            {
                RunningRule runningRule = new RunningRule();
                runningRule.analyzers = TeAnalyzersDocument.Text;
                runningRule.sheetExplainers = TeSheetExplainersDocument.Text;
                TeParams.Dispatcher.Invoke(() =>
                {
                    runningRule.param = TeParams.Text;
                });
                runningRule.basePath = TbBasePathText;
                runningRule.outputPath = TbOutputPathText;
                runningRule.outputName = TbOutputNameText;
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
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, Application.Current.FindResource("Error").ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void DeleteRule()
        {
            String path = $"{System.Environment.CurrentDirectory}\\Rules\\{SelectedRulesItem}.json";
            MessageBoxResult result = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"{Application.Current.FindResource("Delete").ToString()}\n{path}", Application.Current.FindResource("Warning").ToString(), MessageBoxButton.OKCancel, MessageBoxImage.Warning);
            if (result == MessageBoxResult.OK)
            {
                File.Delete(path);
                SelectedAnalyzersIndex = 0;
            }
        }

        private void SetAuto()
        {
            string ruleName = SelectedRulesItem;
            RunningRule rule = JsonConvert.DeserializeObject<RunningRule>(File.ReadAllText($".\\Rules\\{ruleName}.json"));
            TextBox tbPath = new TextBox();
            tbPath.Margin = new Thickness(5, 8, 5, 8);
            tbPath.VerticalContentAlignment = VerticalAlignment.Center;
            TextBox tbFilter = new TextBox();
            tbFilter.Margin = new Thickness(5, 8, 5, 8);
            tbFilter.VerticalContentAlignment = VerticalAlignment.Center;
            int result = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object>() { tbPath, tbFilter, Application.Current.FindResource("Ok").ToString(), Application.Current.FindResource("Cancel").ToString() }, Application.Current.FindResource("WatchPathAndFilter").ToString(), Application.Current.FindResource("Saving").ToString(), MessageBoxImage.Information);
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
            string fileName = $".\\Rules\\{SelectedRulesItem}.json";
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
                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, Application.Current.FindResource("Error").ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
            }

            SetAuto(rule, ruleName);
        }

        private void UnsetAuto()
        {
            string fileName = $".\\Rules\\{SelectedRulesItem}.json";

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
                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, Application.Current.FindResource("Error").ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
            }

            FileSystemWatcher fileSystemWatcher = fileSystemWatcherDic.FirstOrDefault(q => q.Value == SelectedRulesItem).Key;
            if (fileSystemWatcher != null)
            {
                fileSystemWatcher.EnableRaisingEvents = false;
                fileSystemWatcherDic.Remove(fileSystemWatcher);
            }

            BtnSetAutoVisibility = Visibility.Visible;
            BtnUnsetAutoVisibility = Visibility.Hidden;
            BtnDeleteRuleIsEnabled = true;
        }
        private void TeLogPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && !TeLog.IsReadOnly)
            {
                string text = TeLog.Text;
                int lastChangeLineIndex = text.LastIndexOf("\n");
                int subIndex = -1;
                if (text.Length > lastChangeLineIndex + 1)
                {
                    subIndex = lastChangeLineIndex + 1;
                }
                else
                {
                    subIndex = lastChangeLineIndex;
                }
                scanner.UpdateInput(text.Substring(subIndex).Replace(Scanner.CurrentInputMessage, ""));

                TeLog.Text += " <\n";

                TeLog.IsReadOnly = true;
            }

            if (e.Key == Key.Left && !TeLog.IsReadOnly)
            {
                if (TeLog.SelectionStart == TeLog.Text.LastIndexOf('\n') + Scanner.CurrentInputMessage.Length + 1)
                {
                    e.Handled = true;
                }
            }
            else if (e.Key == Key.Up && !TeLog.IsReadOnly)
            {
                e.Handled = true;
            }
        }

        private void TeLogTextChanged(object sender, EventArgs e)
        {
            TeLog.ScrollToEnd();
        }

        private void Stop()
        {
            isStopByUser = true;
            TeLog.Dispatcher.Invoke(() =>
            {
                TeLog.Text += Logger.Get();
            });
        }

        private void Start()
        {
            string param = "";
            TeParams.Dispatcher.Invoke(() =>
            {
                param = TeParams.Text;
            });
            StartLogic(TeSheetExplainersDocument.Text, TeAnalyzersDocument.Text, param, TbBasePathText, TbOutputPathText, TbOutputNameText, false);
        }

        // ---------------------------------------------------- Common Logic

        private bool CheckRule(RunningRule rule)
        {
            if (rule.watchPath != null && rule.watchPath != "" && !Directory.Exists(rule.watchPath))
            {
                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"{Application.Current.FindResource("WatchPathNotExists").ToString()}: {rule.watchPath}", Application.Current.FindResource("Error").ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
            foreach (string analyzer in rule.analyzers.Split('\n'))
            {
                if (!AnalyzersItems.Contains(analyzer))
                {
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"{Application.Current.FindResource("Unknown").ToString()}{Application.Current.FindResource("Analyzer").ToString()}: {analyzer}", Application.Current.FindResource("Error").ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
            }
            foreach (string sheetExplainer in rule.sheetExplainers.Split('\n'))
            {
                if (!SheetExplainersItems.Contains(sheetExplainer))
                {
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"{Application.Current.FindResource("Unknown").ToString()}{Application.Current.FindResource("SheetExplainer").ToString()}: {sheetExplainer}", Application.Current.FindResource("Error").ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }

            }
            return true;
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

                BtnSetAutoVisibility = Visibility.Hidden;
                BtnUnsetAutoVisibility = Visibility.Visible;
                BtnDeleteRuleIsEnabled = false;
            }
        }

        private void FileSystemWatcherInvoke(object sender, FileSystemEventArgs e)
        {
            FileSystemWatcherInvokeThread = new Thread(() =>
            {
                if (stopwatchBeforeFileSystemWatcherInvoke == null)
                {
                    stopwatchBeforeFileSystemWatcherInvoke = new Stopwatch();
                    stopwatchBeforeFileSystemWatcherInvoke.Start();

                    while (stopwatchBeforeFileSystemWatcherInvoke.ElapsedMilliseconds <= fileSystemWatcherInvokeDalay)
                    {
                        Thread.Sleep(freshInterval);
                    }

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
                                SelectedRulesIndex = 0;
                                return;
                            }
                            StartLogic(rule.sheetExplainers, rule.analyzers, rule.param, rule.basePath, rule.outputPath, rule.outputName, true);
                        }
                        finally
                        {
                            while (isRunning)
                            {
                                Thread.Sleep(freshInterval);
                            }

                            stopwatchBeforeFileSystemWatcherInvoke = null;
                            fileSystemWatcher.EnableRaisingEvents = true;
                        }
                    }
                }
                else
                {
                    stopwatchBeforeFileSystemWatcherInvoke.Restart();
                }
            });
            FileSystemWatcherInvokeThread.Start();
        }



        private async void StartLogic(string sheetExplainersStr, string analyzersStr, string paramStr, string basePath, string outputPath, string outputName, bool isAuto)
        {
            BtnStartIsEnabled = false;
            BtnStopIsEnabled = true;

            Dictionary<SheetExplainer, List<string>> filePathListDic = new Dictionary<SheetExplainer, List<string>>();

            isRunning = true;
            isStopByUser = false;

            analyzeSheetInvokeCount = 0;
            setResultInvokeCount = 0;
            resultList = new ConcurrentDictionary<ConcurrentDictionary<ResultType, Object>, Analyzer>();
            GlobalObjects.GlobalObjects.SetGlobalParam(null);

            Scanner.ResetAll();

            if (!isAuto)
            {
                TeLog.Dispatcher.Invoke(() =>
                {
                    TeLog.Text = "";
                });
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
                FinishRunning();
                return;
            }

            paramStr = $"{paramStr.Replace("\r\n", "").Replace("\n", "")}";
            TeParams.Dispatcher.Invoke(() =>
            {
                TeParams.Text = paramStr;
            });
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
                    if (sheetExplainer.pathes == null || sheetExplainer.pathes.Count == 0)
                    {
                        sheetExplainer.pathes = new List<string>();
                        sheetExplainer.pathes.Add("");
                    }
                }
                catch (Exception ex)
                {
                    if (!isAuto)
                    {
                        CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, Application.Current.FindResource("Error").ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    else
                    {
                        Logger.Error(ex.Message);
                    }
                    FinishRunning();
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
                        CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, Application.Current.FindResource("Error").ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    else
                    {
                        Logger.Error(ex.Message);
                    }
                    FinishRunning();
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
                foreach (String str in sheetExplainer.pathes)
                {
                    List<String> filePathList = new List<String>();
                    string basePathTemp = "";
                    if (Path.IsPathRooted(str))
                    {
                        basePathTemp = str.Trim();
                    }
                    else
                    {
                        basePathTemp = Path.Combine(basePath.Trim(), str.Trim());
                    }

                    if (!Directory.Exists(basePathTemp))
                    {
                        if (!isAuto)
                        {
                            CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), Application.Current.FindResource("PathNotExists").ToString(), Application.Current.FindResource("Error").ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                        else
                        {
                            Logger.Error(Application.Current.FindResource("PathNotExists").ToString());
                        }
                        FinishRunning();
                        return;
                    }

                    FileHelper.FileTraverse(isAuto, basePathTemp, sheetExplainer, filePathList);
                    allFilePathList.AddRange(filePathList);
                }
                filePathListDic.Add(sheetExplainer, allFilePathList);

                CompilerResults cresult = GetCresult(analyzer);
                Tuple<CompilerResults, SheetExplainer> cresultTuple = new Tuple<CompilerResults, SheetExplainer>(cresult, sheetExplainer);
                compilerDic.Add(analyzer, cresultTuple);

                runBeforeAnalyzeSheetThread = new Thread(() => RunBeforeAnalyzeSheet(cresult, ParamHelper.MergePublicParam(paramDicEachAnalyzer, analyzer.name), analyzer, allFilePathList));
                runBeforeAnalyzeSheetThread.Start();

                while (runBeforeAnalyzeSheetThread.IsAlive)
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
                        FinishRunning();
                        return;
                    }
                    long timeCostSs = GetNowSs() - startTime;
                    if (perTimeoutLimitAnalyze > 0 && timeCostSs >= perTimeoutLimitAnalyze)
                    {
                        try
                        {
                            runBeforeAnalyzeSheetThread.Abort();
                        }
                        catch (Exception)
                        {
                            // DO NOTHING
                        }
                        if (!isAuto)
                        {
                            CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, $"RunBeforeAnalyzeSheet\n{Application.Current.FindResource("Timeout").ToString()}. \n{perTimeoutLimitAnalyze / 1000.0}(s)", Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
                        }
                        else
                        {
                            Logger.Error($"RunBeforeAnalyzeSheet\n{Application.Current.FindResource("Timeout").ToString()}. \n{perTimeoutLimitAnalyze / 1000.0}(s)");
                        }
                        FinishRunning();
                        return;
                    }
                    await Task.Delay(freshInterval);
                }

                List<String> filePathListFromSheetExplainer = filePathListDic[sheetExplainer];
                totalCount += filePathListFromSheetExplainer.Count;

                int filesCount = filePathListFromSheetExplainer.Count;
                foreach (String filePath in filePathListFromSheetExplainer)
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

            long startSs = GetNowSs();
            smartThreadPoolAnalyze.Join();
            while (smartThreadPoolAnalyze.CurrentWorkItemsCount > 0)
            {
                try
                {
                    long nowSs = GetNowSs();
                    long totalTimeCostSs = nowSs - startSs;
                    if (isStopByUser || (totalTimeoutLimitAnalyze > 0 && totalTimeCostSs >= totalTimeoutLimitAnalyze))
                    {
                        smartThreadPoolAnalyze.Dispose();
                        if (totalTimeoutLimitAnalyze > 0 && totalTimeCostSs >= totalTimeoutLimitAnalyze)
                        {
                            if (!isAuto)
                            {
                                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, $"{Application.Current.FindResource("TotalTimeout").ToString()}. \n{totalTimeoutLimitAnalyze / 1000.0}(s)", Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
                            }
                            else
                            {
                                Logger.Error($"{Application.Current.FindResource("TotalTimeout").ToString()}. \n{totalTimeoutLimitAnalyze / 1000.0}(s)");
                            }

                        }
                        FinishRunning();
                        return;
                    }

                    StringBuilder sb = new StringBuilder();
                    sb.Append($"{Application.Current.FindResource("Analyzing").ToString()}\n");

                    ICollection<string> keys = currentAnalizingDictionary.Keys;
                    foreach (string key in keys)
                    {
                        long value = new long();
                        if (currentAnalizingDictionary.TryGetValue(key, out value))
                        {
                            long timeCostSs = GetNowSs() - currentAnalizingDictionary[key];
                            if (perTimeoutLimitAnalyze > 0 && timeCostSs >= perTimeoutLimitAnalyze)
                            {
                                smartThreadPoolAnalyze.Dispose();
                                if (!isAuto)
                                {
                                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, $"{key}\n{Application.Current.FindResource("Timeout").ToString()}. \n{perTimeoutLimitAnalyze / 1000.0}(s)", Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
                                }
                                else
                                {
                                    Logger.Error($"{key}\n{Application.Current.FindResource("Timeout").ToString()}. \n{perTimeoutLimitAnalyze / 1000.0}(s)");
                                }
                                FinishRunning();
                                return;
                            }
                            sb.Append(key.Substring(key.LastIndexOf('\\') + 1)).Append(" [").Append((timeCostSs / 1000.0).ToString("0.0")).Append("s]\n");
                        }
                    }

                    LProcessContent = $"{smartThreadPoolAnalyze.CurrentWorkItemsCount}/{totalCount} | {Application.Current.FindResource("ActiveThreads").ToString()}: {smartThreadPoolAnalyze.ActiveThreads} | {Application.Current.FindResource("InUseThreads").ToString()}: {smartThreadPoolAnalyze.InUseThreads}";
                    TbStatusText = $"{sb.ToString()}";
                    await Task.Delay(freshInterval);
                }
                catch (Exception e)
                {
                    smartThreadPoolAnalyze.Dispose();
                    Logger.Error($"{Application.Current.FindResource("ExceptionHasBeenThrowed").ToString()} \n{e.Message}");
                    FinishRunning();
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
                    runBeforeSetResultThread = new Thread(() => RunBeforeSetResult(cresult, workbook, ParamHelper.MergePublicParam(paramDicEachAnalyzer, analyzer.name), analyzer, filePathListDic[sheetExplainer]));
                    runBeforeSetResultThread.Start();
                    while (runBeforeSetResultThread.IsAlive)
                    {
                        if (isStopByUser)
                        {
                            runBeforeSetResultThread.Abort();
                            FinishRunning();
                            return;
                        }
                        long timeCostSs = GetNowSs() - startTime;
                        if (perTimeoutLimitAnalyze > 0 && timeCostSs >= perTimeoutLimitAnalyze)
                        {
                            runBeforeSetResultThread.Abort();
                            if (!isAuto)
                            {
                                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, $"RunBeforeAnalyzeSheet\n{Application.Current.FindResource("Timeout").ToString()}. \n{perTimeoutLimitAnalyze / 1000.0}(s)", Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
                            }
                            else
                            {
                                Logger.Error($"RunBeforeAnalyzeSheet\n{Application.Current.FindResource("Timeout").ToString()}. \n{perTimeoutLimitAnalyze / 1000.0}(s)");
                            }
                            FinishRunning();
                            return;
                        }
                        await Task.Delay(freshInterval);
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
                        if (isStopByUser || (totalTimeoutLimitOutput > 0 && totalTimeCostSs >= totalTimeoutLimitOutput))
                        {
                            smartThreadPoolOutput.Dispose();
                            if (totalTimeoutLimitOutput > 0 && totalTimeCostSs >= totalTimeoutLimitOutput)
                            {
                                if (!isAuto)
                                {
                                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, $"{Application.Current.FindResource("TotalTimeout").ToString()}. \n{totalTimeoutLimitOutput / 1000.0}(s)", Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
                                }
                                else
                                {
                                    Logger.Error($"{Application.Current.FindResource("TotalTimeout").ToString()}. \n{totalTimeoutLimitOutput / 1000.0}(s)");
                                }
                            }
                            FinishRunning();
                            return;
                        }

                        StringBuilder sb = new StringBuilder();
                        sb.Append($"{Application.Current.FindResource("Outputting").ToString()} \n");

                        ICollection<string> keys = currentOutputtingDictionary.Keys;
                        foreach (string key in keys)
                        {
                            long value = new long();
                            if (currentOutputtingDictionary.TryGetValue(key, out value))
                            {
                                long timeCostSs = GetNowSs() - currentOutputtingDictionary[key];
                                if (perTimeoutLimitOutput > 0 && timeCostSs >= perTimeoutLimitOutput)
                                {
                                    smartThreadPoolOutput.Dispose();
                                    if (!isAuto)
                                    {
                                        CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, $"{key}\n{Application.Current.FindResource("Timeout").ToString()}. \n{perTimeoutLimitOutput / 1000.0}(s)", Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
                                    }
                                    else
                                    {
                                        Logger.Error($"{key}\n{Application.Current.FindResource("Timeout").ToString()}. \n{perTimeoutLimitOutput / 1000.0}(s)");
                                    }
                                    FinishRunning();
                                    return;
                                }
                                sb.Append(key.Substring(key.LastIndexOf('\\') + 1)).Append(" [").Append((timeCostSs / 1000.0).ToString("0.0")).Append("s]\n");
                            }
                        }

                        LProcessContent = $"{smartThreadPoolOutput.CurrentWorkItemsCount}/{totalCount} | {Application.Current.FindResource("ActiveThreads").ToString()}: {smartThreadPoolOutput.ActiveThreads} | {Application.Current.FindResource("InUseThreads").ToString()}: {smartThreadPoolOutput.InUseThreads}";
                        TbStatusText = $"{sb.ToString()}";
                        await Task.Delay(freshInterval);
                    }
                    catch (Exception e)
                    {
                        smartThreadPoolOutput.Dispose();
                        Logger.Error($"{Application.Current.FindResource("ExceptionHasBeenThrowed").ToString()} \n{e.Message}");
                        FinishRunning();
                        return;
                    }
                }
                smartThreadPoolOutput.Dispose();

                foreach (Analyzer analyzer in compilerDic.Keys)
                {
                    long startTime = GetNowSs();

                    CompilerResults cresult = compilerDic[analyzer].Item1;
                    SheetExplainer sheetExplainer = compilerDic[analyzer].Item2;
                    runEndThread = new Thread(() => RunEnd(cresult, workbook, ParamHelper.MergePublicParam(paramDicEachAnalyzer, analyzer.name), analyzer, filePathListDic[sheetExplainer]));
                    runEndThread.Start();
                    while (runEndThread.IsAlive)
                    {
                        if (isStopByUser)
                        {
                            runEndThread.Abort();
                            FinishRunning();
                            return;
                        }
                        long timeCostSs = GetNowSs() - startTime;
                        if (perTimeoutLimitOutput > 0 && timeCostSs >= perTimeoutLimitOutput)
                        {
                            runEndThread.Abort();

                            BtnStartIsEnabled = true;
                            BtnStopIsEnabled = false;
                            if (!isAuto)
                            {
                                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, $"RunEnd\n{Application.Current.FindResource("Timeout").ToString()}. \n{perTimeoutLimitOutput / 1000.0}(s)", Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
                            }
                            else
                            {
                                Logger.Error($"RunEnd\n{Application.Current.FindResource("Timeout").ToString()}. \n{perTimeoutLimitOutput / 1000.0}(s)");
                            }
                            FinishRunning();
                            return;
                        }
                        await Task.Delay(freshInterval);
                    }
                }

                if (runNotSuccessed)
                {
                    FinishRunning();
                    return;
                }
                TbStatusText = "";

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
                        CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, Application.Current.FindResource("Error").ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
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
                        CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), Application.Current.FindResource(Application.Current.FindResource("Ok").ToString()).ToString() }, Application.Current.FindResource("FileNotSaved").ToString(), Application.Current.FindResource("Info").ToString());
                    }
                    else
                    {
                        Logger.Error(Application.Current.FindResource("FileNotSaved").ToString());
                    }

                    FinishRunning();
                    return;
                }

                Logger.Info($"{Application.Current.FindResource("FileSaved").ToString()}\n{filePath}");
                if (!isAuto && CbIsAutoOpenIsChecked == false)
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
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { btnClose, new ButtonSpacer(40), btnOpenFile, btnOpenPath }, "文件已保存。", "OK");
                }
                else
                {
                    if (CbIsAutoOpenIsChecked == true)
                    {
                        Logger.Info(Application.Current.FindResource("AutoOpened").ToString());
                        System.Diagnostics.Process.Start(filePath);
                    }
                }
            }

            TeLog.Dispatcher.Invoke(() =>
            {
                TeLog.Text += Logger.Get();
            });

            FinishRunning();
        }

        private void WhenRunning()
        {
            while (true)
            {
                TeLog.Dispatcher.Invoke(() =>
                {
                    string logTemp = Logger.Get();
                    if (!String.IsNullOrEmpty(logTemp))
                    {
                        if (TeLog.Text.Length > TeLog.Text.LastIndexOf('\n') + 1)
                        {
                            TeLog.Text = TeLog.Text.Remove(TeLog.Text.LastIndexOf('\n') + 1) + logTemp;
                        }
                        else
                        {
                            TeLog.Text += logTemp;
                        }
                    }

                    if (isRunning && Scanner.InputLock)
                    {
                        string message = Scanner.CurrentInputMessage;
                        if (message != "")
                        {
                            string logText = TeLog.Text;
                            int index = logText.LastIndexOf(message);
                            if (index != -1 && logText.IndexOf('\n', index) != -1
                            && (index - 1 < 0 || (index - 1 >= 0 && logText[index - 1] == '\n') && (logText.Length - 1 < index + logText.Length || (logText.Length - 1 >= index + logText.Length && logText[index + logText.Length] == '\n'))))
                            {
                                TeLog.Text += message;
                                TeLog.Select(TeLog.Text.Length, 0);
                            }
                            else if (index == -1)
                            {
                                TeLog.Text += message;
                                TeLog.Select(TeLog.Text.Length, 0);
                            }
                        }

                        TeLog.IsReadOnly = false;
                        if (TeLog.SelectionStart <= TeLog.Text.LastIndexOf('\n') + Scanner.CurrentInputMessage.Length)
                        {
                            TeLog.Select(TeLog.Text.Length, 0);
                        }
                    }
                    else
                    {
                        TeLog.IsReadOnly = true;
                    }
                });

                Thread.Sleep(freshInterval);
            }
        }

        private void FinishRunning()
        {
            CheckAndCloseThreads(false);
            isRunning = false;
            BtnStartIsEnabled = true;
            BtnStopIsEnabled = false;

            TbStatusText = "";
            LProcessContent = "";
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
                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"{Application.Current.FindResource("FailedToCreateANewFolder").ToString()}\n{ex.Message}", Application.Current.FindResource("Error").ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            SheetExplainersItems = FileHelper.GetSheetExplainersList();
            AnalyzersItems = FileHelper.GetAnalyzersList();
            RuleItems = FileHelper.GetRulesList();
            ParamsItems = FileHelper.GetParamsList();
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
            if (result.TryGetValue(ResultType.FILEPATH, out filePath))
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

        private void CheckAndCloseThreads(bool isCloseWindow)
        {
            if (smartThreadPoolAnalyze != null && !smartThreadPoolAnalyze.IsShuttingdown)
            {
                try
                {
                    smartThreadPoolAnalyze.Shutdown(true);
                }
                catch
                {
                    // DO NOTHING
                }
            }
            if (smartThreadPoolOutput != null && !smartThreadPoolOutput.IsShuttingdown)
            {
                try
                {
                    smartThreadPoolOutput.Shutdown(true);
                }
                catch
                {
                    // DO NOTHING
                }
            }
            if (runBeforeAnalyzeSheetThread != null && runBeforeAnalyzeSheetThread.IsAlive)
            {
                runBeforeAnalyzeSheetThread.Abort();
            }
            if (runBeforeSetResultThread != null && runBeforeSetResultThread.IsAlive)
            {
                runBeforeSetResultThread.Abort();
            }
            if (runEndThread != null && runEndThread.IsAlive)
            {
                runEndThread.Abort();
            }
            if (isCloseWindow)
            {
                if (FileSystemWatcherInvokeThread != null && FileSystemWatcherInvokeThread.IsAlive)
                {
                    FileSystemWatcherInvokeThread.Abort();
                }
                runningThread.Abort();
            }
        }
    }

}