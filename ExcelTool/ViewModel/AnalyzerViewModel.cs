﻿using CustomizableMessageBox;
using DocumentFormat.OpenXml.Bibliography;
using ExcelTool.Helper;
using GlobalObjects;
using Highlighting;
using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Highlighting;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.CSharp.Scripting.Hosting;
using Microsoft.CodeAnalysis.Scripting;
using Microsoft.CodeAnalysis.Scripting.Hosting;
using Microsoft.Toolkit.Mvvm.ComponentModel;
using Microsoft.Toolkit.Mvvm.Input;
using ModernWpf;
using Newtonsoft.Json;
using RoslynPad.Editor;
using RoslynPad.Roslyn;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using static CustomizableMessageBox.MessageBox;

namespace ExcelTool.ViewModel
{
    class AnalyzerViewModel : ObservableObject
    {
        private readonly ObservableCollection<DocumentViewModel> _documents;
        private RoslynHost _host;
        private RoslynCodeEditor editor;
        private Dictionary<string, ParamInfo> paramDicForChange;
        private ItemsControl itemsControl;
        private bool loadFailed;
        private DocumentViewModel documentViewModel;
        private DocumentId documentId;

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

        private double windowActualWidth;
        public double WindowActualWidth
        {
            get { return windowActualWidth; }
            set
            {
                SetProperty<double>(ref windowActualWidth, value);
            }
        }

        private double windowActualHeight;
        public double WindowActualHeight
        {
            get { return windowActualHeight; }
            set
            {
                SetProperty<double>(ref windowActualHeight, value);
            }
        }

        private string windowName = "AnalyzerEditor";
        public string WindowName
        {
            get { return windowName; }
            set
            {
                SetProperty<string>(ref windowName, value);
            }
        }

        private WindowState windowNowState;
        public WindowState WindowNowState
        {
            get { return windowNowState; }
            set
            {
                SetProperty<WindowState>(ref windowNowState, value);
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

        private int selectedAnalyzersIndex = 0;
        public int SelectedAnalyzersIndex
        {
            get { return selectedAnalyzersIndex; }
            set
            {
                SetProperty<int>(ref selectedAnalyzersIndex, value);
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

        private bool btnDeleteIsEnabled = false;
        public bool BtnDeleteIsEnabled
        {
            get { return btnDeleteIsEnabled; }
            set
            {
                SetProperty<bool>(ref btnDeleteIsEnabled, value);
            }
        }

        private bool btnEditParamIsEnabled = false;
        public bool BtnEditParamIsEnabled
        {
            get { return btnEditParamIsEnabled; }
            set
            {
                SetProperty<bool>(ref btnEditParamIsEnabled, value);
            }
        }

        private ObservableCollection<DocumentViewModel> itemsSource;
        public ObservableCollection<DocumentViewModel> ItemsSource
        {
            get { return itemsSource; }
            set
            {
                SetProperty<ObservableCollection<DocumentViewModel>>(ref itemsSource, value);
            }
        }

        public ICommand WindowLoadedCommand { get; set; }
        public ICommand WindowClosingCommand { get; set; }
        public ICommand MenuOpenCommand { get; set; }
        public ICommand WindowSizeChangedCommand { get; set; }
        public ICommand WindowStateChangedCommand { get; set; }
        public ICommand WindowContentRenderedCommand { get; set; }
        public ICommand KeyBindingSaveCommand { get; set; }
        public ICommand KeyBindingRenameSaveCommand { get; set; }
        public ICommand CbAnalyzersPreviewMouseLeftButtonDownCommand { get; set; }
        public ICommand CbAnalyzersSelectionChangedCommand { get; set; }
        public ICommand BtnDeleteClickCommand { get; set; }
        public ICommand BtnEditParamClickCommand { get; set; }
        public ICommand ItemLoadedCommand { get; set; }
        public ICommand BtnSaveClickCommand { get; set; }

        public AnalyzerViewModel()
        {
            _documents = new ObservableCollection<DocumentViewModel>();
            ItemsSource = _documents;
            loadFailed = false;

            WindowLoadedCommand = new RelayCommand<RoutedEventArgs>(WindowLoaded);
            WindowClosingCommand = new RelayCommand<CancelEventArgs>(WindowClosing);
            MenuOpenCommand = new RelayCommand<object>(MenuOpen);
            WindowSizeChangedCommand = new RelayCommand(WindowSizeChanged);
            WindowStateChangedCommand = new RelayCommand(WindowStateChanged);
            WindowContentRenderedCommand = new RelayCommand(WindowContentRendered);
            KeyBindingSaveCommand = new RelayCommand(SaveByKeyDown);
            KeyBindingRenameSaveCommand = new RelayCommand(RenameSaveByKeyDown);
            CbAnalyzersPreviewMouseLeftButtonDownCommand = new RelayCommand(CbAnalyzersPreviewMouseLeftButtonDown);
            CbAnalyzersSelectionChangedCommand = new RelayCommand(CbAnalyzersSelectionChanged);
            BtnDeleteClickCommand = new RelayCommand(BtnDeleteClick);
            BtnEditParamClickCommand = new RelayCommand(BtnEditParamClick);
            ItemLoadedCommand = new RelayCommand<object>(ItemLoaded);
            BtnSaveClickCommand = new RelayCommand(BtnSaveClick);

            ModernWpf.ThemeManager.Current.ActualApplicationThemeChanged += ActualApplicationThemeChanged;
        }

        private void WindowLoaded(RoutedEventArgs e)
        {
            Window window = (Window)e.Source;

            double width = IniHelper.GetWindowSize(WindowName).X;
            double height = IniHelper.GetWindowSize(WindowName).Y;
            if (width > 0)
            {
                window.Dispatcher.Invoke(() =>
                {
                    window.Width = width;
                });
            }

            if (height > 0)
            {
                window.Dispatcher.Invoke(() =>
                {
                    window.Height = height;
                });
            }

            List<Assembly> assemblies = new List<Assembly>();
            assemblies.Add(typeof(object).Assembly);
            string folderPath = System.IO.Path.Combine(Environment.CurrentDirectory, "Dlls");
            DirectoryInfo dir = new DirectoryInfo(folderPath);
            FileSystemInfo[] dllInfos = null;
            if (dir.Exists)
            {
                DirectoryInfo dirD = dir as DirectoryInfo;
                dllInfos = dirD.GetFileSystemInfos();
            }
            List<string> dlls = new List<string>();
            dlls.Add("ClosedXML.dll");
            dlls.Add("GlobalObjects.dll");
            if (dllInfos != null && dllInfos.Count() != 0)
            {
                foreach (FileSystemInfo dllInfo in dllInfos)
                {
                    dlls.Add(dllInfo.FullName);
                }
            }
            foreach (string dll in dlls)
            {
                Assembly dllFile = null;
                if (IniHelper.GetSecurityCheck())
                {
                    try
                    {
                        dllFile = Assembly.LoadFrom(dll);
                    }
                    catch (FileLoadException ex)
                    {
                        ComboBox comboBox = new ComboBox();
                        comboBox.HorizontalAlignment = HorizontalAlignment.Stretch;
                        comboBox.Margin = new Thickness(5);
                        List<string> items = new List<string>();
                        items.Add(Application.Current.FindResource("NoAction").ToString());
                        items.Add(Application.Current.FindResource("CloseTheProgramOnly").ToString());
                        items.Add(Application.Current.FindResource("BanSecurityCheckWithoutRestart").ToString());
                        items.Add(Application.Current.FindResource("BanSecurityCheckWithRestart").ToString());
                        comboBox.ItemsSource = items;
                        comboBox.SelectedIndex = 0;
                        CustomizableMessageBox.MessageBox.Show(new RefreshList { comboBox, Application.Current.FindResource("Ok").ToString() }, $"{Application.Current.FindResource("UnblockDllsCopiedFromTheWeb").ToString().Replace("{0}", dll)}\n\n{ex.Message}", Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
                        if (comboBox.SelectedIndex == 0)
                        {
                            // Do Nothing
                        }
                        else if (comboBox.SelectedIndex == 1)
                        {
                            GlobalObjects.GlobalObjects.ProgramCurrentStatus = GlobalObjects.ProgramStatus.Shutdown;
                            Application.Current.Shutdown();
                        }
                        else if (comboBox.SelectedIndex == 2)
                        {
                            IniHelper.SetSecurityCheck(false);
                        }
                        else if (comboBox.SelectedIndex == 3)
                        {
                            IniHelper.SetSecurityCheck(false);
                            GlobalObjects.GlobalObjects.ProgramCurrentStatus = GlobalObjects.ProgramStatus.Restart;
                            Application.Current.Shutdown();
                        }
                        window.Close();
                        loadFailed = true;
                        return;
                    }
                }
                else
                {
                    try
                    {
                        dllFile = Assembly.UnsafeLoadFrom(dll);
                    }
                    catch (FileLoadException ex)
                    {
                        CustomizableMessageBox.MessageBox.Show(new RefreshList { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, $"{Application.Current.FindResource("FileNotSupported").ToString().Replace("{0}", dll)}\n\n{ex.Message}", Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
                        window.Close();
                        loadFailed = true;
                        return;
                    }
                }
                foreach (Type type in dllFile.GetExportedTypes())
                {
                    assemblies.Add(type.Assembly);
                }
            }

            _host = new RoslynHost(additionalAssemblies: new[]
            {
                Assembly.Load("RoslynPad.Roslyn.Windows"),
                Assembly.Load("RoslynPad.Editor.Windows")
            }, RoslynHostReferences.NamespaceDefault.With(assemblyReferences: assemblies));
            AddNewDocument();
        }

        private void WindowClosing(CancelEventArgs e)
        {
            IniHelper.SetWindowSize(WindowName, new Point(WindowActualWidth, WindowActualHeight));
        }

        private void MenuOpen(object sender)
        {
            if (editor == null)
            {
                return;
            }
            if (((MenuItem)sender).Name == "menu_show_spaces")
            {
                editor.Options.ShowSpaces = !editor.Options.ShowSpaces;
                IniHelper.SetShowSpaces(editor.Options.ShowSpaces);
            }
            else if (((MenuItem)sender).Name == "menu_show_tabs")
            {
                editor.Options.ShowTabs = !editor.Options.ShowTabs;
                IniHelper.SetShowTabs(editor.Options.ShowTabs);
            }
            else if (((MenuItem)sender).Name == "menu_show_end_of_line")
            {
                editor.Options.ShowEndOfLine = !editor.Options.ShowEndOfLine;
                IniHelper.SetShowEndOfLine(editor.Options.ShowEndOfLine);
            }
            else if (((MenuItem)sender).Name == "menu_show_box_for_control_characters")
            {
                editor.Options.ShowBoxForControlCharacters = !editor.Options.ShowBoxForControlCharacters;
                IniHelper.SetShowBoxForControlCharacters(editor.Options.ShowBoxForControlCharacters);
            }
            else if (((MenuItem)sender).Name == "menu_enable_hyperlinks")
            {
                editor.Options.EnableHyperlinks = !editor.Options.EnableHyperlinks;
                IniHelper.SetEnableHyperlinks(editor.Options.EnableHyperlinks);
            }
            else if (((MenuItem)sender).Name == "menu_indentation_size")
            {
                TextBox textBox = new TextBox();
                textBox.Margin = new Thickness(5);
                textBox.Height = 30;
                textBox.VerticalContentAlignment = VerticalAlignment.Center;
                textBox.Text = editor.Options.IndentationSize.ToString();
                int res = CustomizableMessageBox.MessageBox.Show(new RefreshList { textBox, new ButtonSpacer(1, GridUnitType.Star, true), Application.Current.FindResource("Ok").ToString(), Application.Current.FindResource("Cancel").ToString() }, Application.Current.FindResource("IndentationSize").ToString(), Application.Current.FindResource("Setting").ToString());
                if (res == 3)
                {
                    return;
                }
                try
                {
                    editor.Options.IndentationSize = int.Parse(textBox.Text.Trim());
                }
                catch (Exception ex)
                {
                    CustomizableMessageBox.MessageBox.Show(new RefreshList { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, ex.Message, Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
                }

                IniHelper.SetIndentationSize(editor.Options.IndentationSize);
            }
            else if (((MenuItem)sender).Name == "menu_convert_tabs_to_spaces")
            {
                editor.Options.ConvertTabsToSpaces = !editor.Options.ConvertTabsToSpaces;
                IniHelper.SetConvertTabsToSpaces(editor.Options.ConvertTabsToSpaces);
            }
            else if (((MenuItem)sender).Name == "menu_highlight_current_line")
            {
                editor.Options.HighlightCurrentLine = !editor.Options.HighlightCurrentLine;
                IniHelper.SetHighlightCurrentLine(editor.Options.HighlightCurrentLine);
            }
            else if (((MenuItem)sender).Name == "menu_hide_cursor_while_typing")
            {
                editor.Options.HideCursorWhileTyping = !editor.Options.HideCursorWhileTyping;
                IniHelper.SetHideCursorWhileTyping(editor.Options.HideCursorWhileTyping);
            }
            else if (((MenuItem)sender).Name == "menu_word_wrap")
            {
                editor.WordWrap = !editor.WordWrap;
                IniHelper.SetWordWrap(editor.WordWrap);
            }
            else if (((MenuItem)sender).Name == "menu_show_line_numbers")
            {
                editor.ShowLineNumbers = !editor.ShowLineNumbers;
                IniHelper.SetShowLineNumbers(editor.ShowLineNumbers);
            }
        }

        private void WindowSizeChanged()
        {
            ResetEditorSize(editor);
        }

        private void WindowStateChanged()
        {
            ResetEditorSize(editor);
        }

        private void WindowContentRendered()
        {
            if (loadFailed)
            {
                return;
            }

            foreach (RoslynCodeEditor editorTemp in FindVisualChildren<RoslynCodeEditor>(itemsControl))
            {
                if (editorTemp.Name == "rce_editor")
                {
                    editor = editorTemp;
                }
            }

            ItemLoadedCommand = null;
            editor.Focus();

            ResetEditorSize(editor);

            documentViewModel = (DocumentViewModel)editor.DataContext;

            ActualApplicationThemeChanged(null, null);

            editor.TextArea.SelectionCornerRadius = 0;
            editor.TextArea.SelectionBorder = new Pen(new SolidColorBrush(Colors.DeepSkyBlue), 1);

            InitEditor();

            editor.Text = GlobalObjects.GlobalObjects.GetDefaultCode();
        }

        private void InitEditor()
        {
            editor.Options.ShowSpaces = IniHelper.GetShowSpaces();
            editor.Options.ShowTabs = IniHelper.GetShowTabs();
            editor.Options.ShowEndOfLine = IniHelper.GetShowEndOfLine();
            editor.Options.ShowBoxForControlCharacters = IniHelper.GetShowBoxForControlCharacters();
            editor.Options.EnableHyperlinks = IniHelper.GetEnableHyperlinks();
            editor.Options.IndentationSize = IniHelper.GetIndentationSize();
            editor.Options.ConvertTabsToSpaces = IniHelper.GetConvertTabsToSpaces();
            editor.Options.HighlightCurrentLine = IniHelper.GetHighlightCurrentLine();
            editor.Options.HideCursorWhileTyping = IniHelper.GetHideCursorWhileTyping();
            editor.WordWrap = IniHelper.GetWordWrap();
            editor.ShowLineNumbers = IniHelper.GetShowLineNumbers();
        }

        private void SaveByKeyDown()
        {
            if (SelectedAnalyzersIndex >= 1)
            {
                Save(false);
            }
            else
            {
                Save(true);
            }
        }

        private void RenameSaveByKeyDown()
        {
            Save(true);
        }


        private void CbAnalyzersPreviewMouseLeftButtonDown()
        {
            try
            {
                if (!Directory.Exists(".\\Analyzers"))
                {
                    Directory.CreateDirectory(".\\Analyzers");
                }
            }
            catch (Exception ex)
            {
                CustomizableMessageBox.MessageBox.Show(new RefreshList { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, $"{Application.Current.FindResource("FailedToCreateANewFolder").ToString()}\n{ex.Message}", Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
                return;
            }

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

            AnalyzersItems = AnalyzersList;
        }

        private void CbAnalyzersSelectionChanged()
        {
            if (editor == null)
            {
                return;
            }

            BtnDeleteIsEnabled = SelectedAnalyzersIndex >= 1 ? true : false;
            BtnEditParamIsEnabled = SelectedAnalyzersIndex >= 1 ? true : false;

            if (SelectedAnalyzersIndex == 0)
            {
                editor.Text = GlobalObjects.GlobalObjects.GetDefaultCode();
                paramDicForChange = new Dictionary<string, ParamInfo>();
                return;
            }
            Analyzer analyzer = JsonHelper.TryDeserializeObject<Analyzer>($".\\Analyzers\\{SelectedAnalyzersItem}.json");
            if (analyzer == null)
            {
                SelectedAnalyzersIndex = 0;
                return;
            }
            editor.Text = analyzer.code;

            paramDicForChange = analyzer.paramDic;
        }

        private void BtnDeleteClick()
        {
            String path = $"{System.Environment.CurrentDirectory}\\Analyzers\\{SelectedAnalyzersItem}.json";
            MessageBoxResult result = CustomizableMessageBox.MessageBox.Show($"{Application.Current.FindResource("Delete").ToString()}\n{path}", Application.Current.FindResource("Warning").ToString(), MessageBoxButton.OKCancel, MessageBoxImage.Warning);
            if (result == MessageBoxResult.OK)
            {
                File.Delete(path);
                SelectedAnalyzersIndex = 0;
            }
        }

        private void BtnEditParamClick()
        {
            ParamEditor paramEditor = new ParamEditor();

            RowDefinition rowDefinition = new RowDefinition();
            paramEditor.g_main.RowDefinitions.Add(rowDefinition);

            RowDefinition rowDefinitionOk = new RowDefinition();
            rowDefinitionOk.Height = new GridLength(40);
            paramEditor.g_btn.RowDefinitions.Add(rowDefinitionOk);

            ColumnDefinition columnDefinitionL = new ColumnDefinition();
            paramEditor.g_main.ColumnDefinitions.Add(columnDefinitionL);

            ColumnDefinition columnDefinitionML = new ColumnDefinition();
            columnDefinitionML.Width = new GridLength(20);
            paramEditor.g_main.ColumnDefinitions.Add(columnDefinitionML);

            ColumnDefinition columnDefinitionM = new ColumnDefinition();
            paramEditor.g_main.ColumnDefinitions.Add(columnDefinitionM);

            ColumnDefinition columnDefinitionMR = new ColumnDefinition();
            columnDefinitionMR.Width = new GridLength(20);
            paramEditor.g_main.ColumnDefinitions.Add(columnDefinitionMR);

            ColumnDefinition columnDefinitionR = new ColumnDefinition();
            paramEditor.g_main.ColumnDefinitions.Add(columnDefinitionR);

            TextBlock textBlockTitleL = new TextBlock();
            textBlockTitleL.Text = Application.Current.FindResource("ParamKey").ToString();
            textBlockTitleL.Margin = new Thickness(15, 0, 0, 0);
            textBlockTitleL.Height = 25;
            textBlockTitleL.HorizontalAlignment = HorizontalAlignment.Left;
            textBlockTitleL.VerticalAlignment = VerticalAlignment.Top;
            Grid.SetRow(textBlockTitleL, 0);
            Grid.SetColumn(textBlockTitleL, 0);
            paramEditor.g_main.Children.Add(textBlockTitleL);

            TextBlock textBlockTitleM = new TextBlock();
            textBlockTitleM.Text = Application.Current.FindResource("ParamDescription").ToString();
            textBlockTitleM.Margin = new Thickness(15, 0, 0, 0);
            textBlockTitleM.Height = 25;
            textBlockTitleM.HorizontalAlignment = HorizontalAlignment.Left;
            textBlockTitleM.VerticalAlignment = VerticalAlignment.Top;
            Grid.SetRow(textBlockTitleM, 0);
            Grid.SetColumn(textBlockTitleM, 2);
            paramEditor.g_main.Children.Add(textBlockTitleM);

            TextBlock textBlockTitleR = new TextBlock();
            textBlockTitleR.Text = Application.Current.FindResource("ParamPossibleValues").ToString();
            textBlockTitleR.Margin = new Thickness(15, 0, 0, 0);
            textBlockTitleR.Height = 25;
            textBlockTitleR.HorizontalAlignment = HorizontalAlignment.Left;
            textBlockTitleR.VerticalAlignment = VerticalAlignment.Top;
            Grid.SetRow(textBlockTitleR, 0);
            Grid.SetColumn(textBlockTitleR, 4);
            paramEditor.g_main.Children.Add(textBlockTitleR);

            TextEditor textEditorL = new TextEditor();
            textEditorL.Dispatcher.Invoke(() =>
            {
                textEditorL.SyntaxHighlighting = HighlightingManager.Instance.GetDefinition("param");
            });
            textEditorL.Margin = new Thickness(0, 25, 0, 0);
            textEditorL.ShowLineNumbers = true;
            textEditorL.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
            textEditorL.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            textEditorL.Background = Theme.ThemeControlBackground;
            textEditorL.Foreground = Theme.ThemeControlForeground;
            if (paramDicForChange != null && paramDicForChange.Keys != null && paramDicForChange.Keys.Count > 0)
            {
                foreach (string str in paramDicForChange.Keys)
                {
                    textEditorL.Text += $"{ParamHelper.DecodeToEscaped(str)}\n";
                }
            }
            Grid.SetRow(textEditorL, 0);
            Grid.SetColumn(textEditorL, 0);
            paramEditor.g_main.Children.Add(textEditorL);

            TextBlock textBlockML = new TextBlock();
            textBlockML.Margin = new Thickness(0, 25, 0, 0);
            textBlockML.Text = "=";
            textBlockML.HorizontalAlignment = HorizontalAlignment.Center;
            textBlockML.VerticalAlignment = VerticalAlignment.Center;
            Grid.SetRow(textBlockML, 0);
            Grid.SetColumn(textBlockML, 1);
            paramEditor.g_main.Children.Add(textBlockML);

            TextEditor textEditorM = new TextEditor();
            textEditorM.Dispatcher.Invoke(() =>
            {
                textEditorM.SyntaxHighlighting = HighlightingManager.Instance.GetDefinition("param");
            });
            textEditorM.Margin = new Thickness(0, 25, 0, 0);
            textEditorM.ShowLineNumbers = true;
            textEditorM.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
            textEditorM.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            textEditorM.Background = Theme.ThemeControlBackground;
            textEditorM.Foreground = Theme.ThemeControlForeground;
            if (paramDicForChange != null && paramDicForChange.Keys != null && paramDicForChange.Keys.Count > 0)
            {
                foreach (ParamInfo paramInfo in paramDicForChange.Values)
                {
                    textEditorM.Text += $"{ParamHelper.DecodeToEscaped(paramInfo.describe)}\n";
                }
            }
            Grid.SetRow(textEditorM, 0);
            Grid.SetColumn(textEditorM, 2);
            paramEditor.g_main.Children.Add(textEditorM);

            TextBlock textBlockMR = new TextBlock();
            textBlockMR.Margin = new Thickness(0, 25, 0, 0);
            textBlockMR.Text = "=";
            textBlockMR.HorizontalAlignment = HorizontalAlignment.Center;
            textBlockMR.VerticalAlignment = VerticalAlignment.Center;
            Grid.SetRow(textBlockMR, 0);
            Grid.SetColumn(textBlockMR, 3);
            paramEditor.g_main.Children.Add(textBlockMR);

            TextEditor textEditorR = new TextEditor();
            textEditorR.Dispatcher.Invoke(() =>
            {
                textEditorR.SyntaxHighlighting = HighlightingManager.Instance.GetDefinition("param");
            });
            textEditorR.Margin = new Thickness(0, 25, 0, 0);
            textEditorR.ShowLineNumbers = true;
            textEditorR.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
            textEditorR.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            textEditorR.Background = Theme.ThemeControlBackground;
            textEditorR.Foreground = Theme.ThemeControlForeground;
            if (paramDicForChange != null && paramDicForChange.Keys != null && paramDicForChange.Keys.Count > 0)
            {
                foreach (ParamInfo paramInfo in paramDicForChange.Values)
                {
                    string possibleValues = "";
                    if (paramInfo.possibleValues != null)
                    {
                        foreach (PossibleValue possibleValue in paramInfo.possibleValues)
                        {
                            if (paramInfo.type == ParamType.Single)
                            {
                                if (possibleValue.describe == null)
                                {
                                    possibleValues += $"{possibleValue.value}|";
                                }
                                else
                                {
                                    possibleValues += $"{possibleValue.value}:{possibleValue.describe}|";
                                }
                            }
                            else if (paramInfo.type == ParamType.Multiple)
                            {
                                if (possibleValue.describe == null)
                                {
                                    possibleValues += $"{possibleValue.value}+";
                                }
                                else
                                {
                                    possibleValues += $"{possibleValue.value}:{possibleValue.describe}+";
                                }
                            }
                        }
                    }
                    if (possibleValues != "")
                    {
                        textEditorR.Text += $"{ParamHelper.DecodeToEscaped(possibleValues.Remove(possibleValues.Length - 1))}\n";
                    }
                    else
                    {
                        textEditorR.Text += "\n";
                    }
                }
            }
            Grid.SetRow(textEditorR, 0);
            Grid.SetColumn(textEditorR, 4);
            paramEditor.g_main.Children.Add(textEditorR);

            Button btnOk = new Button();
            btnOk.HorizontalAlignment = HorizontalAlignment.Stretch;
            btnOk.Content = Application.Current.FindResource("Ok").ToString();
            btnOk.Click += (s, ex) =>
            {
                List<string> keyList = textEditorL.Text.Replace("\r", "").Split('\n').Where(str => str.Trim() != "").ToList();
                List<string> describeList = textEditorM.Text.Replace("\r", "").Split('\n').ToList();
                List<string> possibleValueList = textEditorR.Text.Replace("\r", "").Split('\n').ToList();
                if (keyList.Count <= describeList.Count && keyList.Count <= possibleValueList.Count)
                {
                    Dictionary<string, ParamInfo> changedParamDic = new Dictionary<string, ParamInfo>();
                    for (int i = 0; i < keyList.Count; ++i)
                    {
                        string possibleValue = ParamHelper.EncodeFromEscaped(possibleValueList[i]);
                        if (possibleValue.Contains('|'))
                        {
                            List<string> valueList = possibleValue.Split('|').ToList();
                            List<PossibleValue> possibleValueListOr = new List<PossibleValue>();
                            foreach (string value in valueList)
                            {
                                string valveInClass = null;
                                string describe = null;
                                if (value.Contains(':'))
                                {
                                    string[] splited = value.Split(':');
                                    valveInClass = splited[0];
                                    describe = splited[1];
                                    possibleValueListOr.Add(new PossibleValue(valveInClass, describe));
                                }
                                else
                                {
                                    valveInClass = value;
                                    possibleValueListOr.Add(new PossibleValue(valveInClass));
                                }
                            }
                            if (describeList[i].Trim() == "")
                            {
                                changedParamDic.Add(ParamHelper.EncodeFromEscaped(keyList[i]), new ParamInfo(possibleValueListOr, ParamType.Single));
                            }
                            else
                            {
                                changedParamDic.Add(ParamHelper.EncodeFromEscaped(keyList[i]), new ParamInfo(ParamHelper.EncodeFromEscaped(describeList[i]), possibleValueListOr, ParamType.Single));
                            }
                        }
                        else if (possibleValue.Contains('+'))
                        {
                            List<string> valueList = possibleValue.Split('+').ToList();
                            List<PossibleValue> possibleValueListAnd = new List<PossibleValue>();
                            foreach (string value in valueList)
                            {
                                string valveInClass = null;
                                string describe = null;
                                if (value.Contains(':'))
                                {
                                    string[] splited = value.Split(':');
                                    valveInClass = splited[0];
                                    describe = splited[1];
                                    possibleValueListAnd.Add(new PossibleValue(valveInClass, describe));
                                }
                                else
                                {
                                    valveInClass = value;
                                    possibleValueListAnd.Add(new PossibleValue(valveInClass));
                                }
                            }
                            if (describeList[i].Trim() == "")
                            {
                                changedParamDic.Add(ParamHelper.EncodeFromEscaped(keyList[i]), new ParamInfo(possibleValueListAnd, ParamType.Single));
                            }
                            else
                            {
                                changedParamDic.Add(ParamHelper.EncodeFromEscaped(keyList[i]), new ParamInfo(ParamHelper.EncodeFromEscaped(describeList[i]), possibleValueListAnd, ParamType.Multiple));
                            }
                        }
                        else
                        {
                            if (describeList[i].Trim() == "")
                            {
                                changedParamDic.Add(ParamHelper.EncodeFromEscaped(keyList[i]), new ParamInfo(null, ParamType.Single));
                            }
                            else
                            {
                                changedParamDic.Add(ParamHelper.EncodeFromEscaped(keyList[i]), new ParamInfo(ParamHelper.EncodeFromEscaped(describeList[i]), null, ParamType.Multiple));
                            }
                        }
                    }
                    paramDicForChange = changedParamDic;
                }

                paramEditor.Close();
            };
            Grid.SetRow(btnOk, 1);
            Grid.SetColumnSpan(btnOk, 5);
            paramEditor.g_btn.Children.Add(btnOk);

            paramEditor.ShowDialog();
        }

        private void ItemLoaded(object e)
        {
            itemsControl = (ItemsControl)(((RoutedEventArgs)e).Source);
        }

        private void BtnSaveClick()
        {
            Save(true);
        }

        private void ResetEditorSize(RoslynCodeEditor editor)
        {
            if (editor != null && WindowNowState != WindowState.Minimized)
            {
                editor.Height = WindowActualHeight - 155;
            }
        }

        // ---------------------------------------------------- Common Logic
        private void Save(bool isRename)
        {
            TextBox tbName = new TextBox();
            tbName.Margin = new Thickness(5);
            tbName.Height = 30;
            tbName.VerticalContentAlignment = VerticalAlignment.Center;

            string newName = "";
            if (SelectedAnalyzersIndex >= 1)
            {
                newName = SelectedAnalyzersItem.ToString();
            }
            if (isRename)
            {
                if (SelectedAnalyzersIndex >= 1)
                {
                    tbName.Text = $"Copy Of {SelectedAnalyzersItem}";
                }
                int result = CustomizableMessageBox.MessageBox.Show(new RefreshList() { tbName, Application.Current.FindResource("Ok").ToString(), Application.Current.FindResource("Cancel").ToString() }, Application.Current.FindResource("Name").ToString(), Application.Current.FindResource("Saving").ToString(), MessageBoxImage.Information);

                if (result != 1)
                {
                    return;
                }

                if (tbName.Text == "")
                {
                    CustomizableMessageBox.MessageBox.Show(new RefreshList { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, Application.Current.FindResource("FileNameEmptyError").ToString(), Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
                    return;
                }

                newName = tbName.Text;
            }

            Analyzer analyzer = new Analyzer();
            analyzer.code = editor.Text;
            analyzer.name = newName;
            analyzer.paramDic = paramDicForChange;

            FileHelper.SavaAnalyzerJson(newName, analyzer, true);

            CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetterWithTimmer(), new RefreshList { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, Application.Current.FindResource("SuccessfullySaved").ToString(), Application.Current.FindResource("Save").ToString(), MessageBoxImage.Information);

        }

        private void AddNewDocument(DocumentViewModel previous = null)
        {
            _documents.Add(new DocumentViewModel(_host, previous));
        }

        public IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj != null)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(depObj, i);

                    if (child != null && child is T)
                        yield return (T)child;

                    foreach (T childOfChild in FindVisualChildren<T>(child))
                        yield return childOfChild;
                }
            }
        }

        private void ActualApplicationThemeChanged(ThemeManager themeManager, object obj)
        {
            if (editor == null)
            {
                return;
            }

            if (themeManager == null)
            {
                if (ThemeManager.Current.ActualApplicationTheme == ApplicationTheme.Light)
                {
                    documentId = editor.Initialize(_host, new ClassificationHighlightColors(),
                    Directory.GetCurrentDirectory(), string.Empty);
                    documentViewModel.Initialize(documentId);
                }
                else if (ThemeManager.Current.ActualApplicationTheme == ApplicationTheme.Dark)
                {
                    documentId = editor.Initialize(_host, new DarkModeHighlightColors(),
                        Directory.GetCurrentDirectory(), string.Empty);
                    documentViewModel.Initialize(documentId);
                }
            }
            else
            {
                if (ThemeManager.Current.ActualApplicationTheme == ApplicationTheme.Light)
                {
                    ClassificationHighlightColors classificationHighlightColors = new ClassificationHighlightColors();

                    FieldInfo classificationHighlightColorsFieldInfo = editor.GetType().GetField("_classificationHighlightColors",
                        BindingFlags.Instance | BindingFlags.GetField | BindingFlags.NonPublic | BindingFlags.ExactBinding);
                    classificationHighlightColorsFieldInfo.SetValue(editor, classificationHighlightColors);

                    FieldInfo braceMatcherHighlighterFieldInfo = editor.GetType().GetField("_braceMatcherHighlighter",
                        BindingFlags.Instance | BindingFlags.GetField | BindingFlags.NonPublic | BindingFlags.ExactBinding);
                    braceMatcherHighlighterFieldInfo.SetValue(editor, new BraceMatcherHighlightRenderer(editor.TextArea.TextView, classificationHighlightColors));

                    editor.RefreshHighlighting();
                }
                else if (ThemeManager.Current.ActualApplicationTheme == ApplicationTheme.Dark)
                {
                    DarkModeHighlightColors darkModeHighlightColors = new DarkModeHighlightColors();

                    FieldInfo classificationHighlightColorsFieldInfo = editor.GetType().GetField("_classificationHighlightColors",
                        BindingFlags.Instance | BindingFlags.GetField | BindingFlags.NonPublic | BindingFlags.ExactBinding);
                    classificationHighlightColorsFieldInfo.SetValue(editor, darkModeHighlightColors);

                    FieldInfo braceMatcherHighlighterFieldInfo = editor.GetType().GetField("_braceMatcherHighlighter",
                        BindingFlags.Instance | BindingFlags.GetField | BindingFlags.NonPublic | BindingFlags.ExactBinding);
                    braceMatcherHighlighterFieldInfo.SetValue(editor, new BraceMatcherHighlightRenderer(editor.TextArea.TextView, darkModeHighlightColors));

                    editor.RefreshHighlighting();
                }
            }

            editor.TextArea.SelectionBrush = Theme.ThemeSelectionBrush;
            editor.TextArea.SelectionForeground = Theme.ThemeSelectionForeground;
        }

        public class DocumentViewModel : INotifyPropertyChanged
        {
            private bool _isReadOnly;
            private readonly RoslynHost _host;
            private string _result;

            public DocumentViewModel(RoslynHost host, DocumentViewModel previous)
            {
                _host = host;
                Previous = previous;
            }

            internal void Initialize(DocumentId id)
            {
                Id = id;
            }


            public DocumentId Id { get; private set; }

            public bool IsReadOnly
            {
                get { return _isReadOnly; }
                private set { SetProperty(ref _isReadOnly, value); }
            }

            public DocumentViewModel Previous { get; }

            public DocumentViewModel LastGoodPrevious
            {
                get
                {
                    var previous = Previous;

                    while (previous != null && previous.HasError)
                    {
                        previous = previous.Previous;
                    }

                    return previous;
                }
            }

            public Script<object> Script { get; private set; }

            public string Text { get; set; }

            public bool HasError { get; private set; }

            public string Result
            {
                get { return _result; }
                private set { SetProperty(ref _result, value); }
            }

            private static MethodInfo HasSubmissionResult { get; } =
                typeof(Compilation).GetMethod(nameof(HasSubmissionResult), BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);

            private static PrintOptions PrintOptions { get; } =
                new PrintOptions { MemberDisplayFormat = MemberDisplayFormat.SeparateLines };

            public async Task<bool> TrySubmit()
            {
                Result = null;

                Script = LastGoodPrevious?.Script.ContinueWith(Text) ??
                    CSharpScript.Create(Text, ScriptOptions.Default
                        .WithReferences(_host.DefaultReferences)
                        .WithImports(_host.DefaultImports));

                var compilation = Script.GetCompilation();
                var hasResult = (bool)HasSubmissionResult.Invoke(compilation, null);
                var diagnostics = Script.Compile();
                if (diagnostics.Any(t => t.Severity == DiagnosticSeverity.Error))
                {
                    Result = string.Join(Environment.NewLine, diagnostics.Select(FormatObject));
                    return false;
                }

                IsReadOnly = true;

                await Execute(hasResult);

                return true;
            }

            private async Task Execute(bool hasResult)
            {
                try
                {
                    var result = await Script.RunAsync();

                    if (result.Exception != null)
                    {
                        HasError = true;
                        Result = FormatException(result.Exception);
                    }
                    else
                    {
                        Result = hasResult ? FormatObject(result.ReturnValue) : null;
                    }
                }
                catch (Exception ex)
                {
                    HasError = true;
                    Result = FormatException(ex);
                }
            }

            private static string FormatException(Exception ex)
            {
                return CSharpObjectFormatter.Instance.FormatException(ex);
            }

            private static string FormatObject(object o)
            {
                return CSharpObjectFormatter.Instance.FormatObject(o, PrintOptions);
            }

            public event PropertyChangedEventHandler PropertyChanged;

            protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
            {
                if (!EqualityComparer<T>.Default.Equals(field, value))
                {
                    field = value;
                    OnPropertyChanged(propertyName);
                    return true;
                }
                return false;
            }

            protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
