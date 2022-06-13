﻿using ICSharpCode.AvalonEdit;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Scripting;
using Microsoft.Toolkit.Mvvm.ComponentModel;
using Microsoft.Toolkit.Mvvm.Input;
using Newtonsoft.Json;
using RoslynPad.Editor;
using RoslynPad.Roslyn;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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

namespace ExcelTool.ViewModel
{
    class AnalyzerViewModel : ObservableObject
    {
        private readonly ObservableCollection<DocumentViewModel> _documents;
        private RoslynHost _host;
        private RoslynCodeEditor editor;
        private Dictionary<string, string> paramDicForChange;
        ItemsControl itemsControl;

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

            WindowLoadedCommand = new RelayCommand<RoutedEventArgs>(WindowLoaded);
            WindowClosingCommand = new RelayCommand<CancelEventArgs>(WindowClosing);
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
            dlls.Add("System.dll");
            dlls.Add("System.Data.dll");
            dlls.Add("System.Xml.dll");
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
                var dllFile = Assembly.LoadFrom(dll);
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

            var viewModel = (DocumentViewModel)editor.DataContext;
            var workingDirectory = Directory.GetCurrentDirectory();

            var previous = viewModel.LastGoodPrevious;
            if (previous != null)
            {
                editor.CreatingDocument += (o, args) =>
                {
                    args.DocumentId = _host.AddRelatedDocument(previous.Id, new DocumentCreationArgs(
                        args.TextContainer, workingDirectory, args.ProcessDiagnostics,
                        args.TextContainer.UpdateText));
                };
            }

            var documentId = editor.Initialize(_host, new ClassificationHighlightColors(),
                workingDirectory, string.Empty);

            viewModel.Initialize(documentId);

            editor.Text = GlobalObjects.GlobalObjects.GetDefaultCode();
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
                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"{Application.Current.FindResource("FailedToCreateANewFolder").ToString()}\n{ex.Message}", Application.Current.FindResource("Error").ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
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
            BtnDeleteIsEnabled = SelectedAnalyzersIndex >= 1 ? true : false;
            BtnEditParamIsEnabled = SelectedAnalyzersIndex >= 1 ? true : false;

            editor.Text = GlobalObjects.GlobalObjects.GetDefaultCode();

            if (SelectedAnalyzersIndex == 0)
            {
                return;
            }
            Analyzer analyzer = JsonConvert.DeserializeObject<Analyzer>(File.ReadAllText($".\\Analyzers\\{SelectedAnalyzersItem}.json"));
            editor.Text = analyzer.code;

            paramDicForChange = analyzer.paramDic;
        }

        private void BtnDeleteClick()
        {
            String path = $"{System.Environment.CurrentDirectory}\\Analyzers\\{SelectedAnalyzersItem}.json";
            MessageBoxResult result = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"{Application.Current.FindResource("Delete").ToString()}\n{path}", Application.Current.FindResource("Warning").ToString(), MessageBoxButton.OKCancel, MessageBoxImage.Warning);
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
            rowDefinitionOk.Height = new GridLength(30);
            paramEditor.g_main.RowDefinitions.Add(rowDefinitionOk);

            ColumnDefinition columnDefinitionL = new ColumnDefinition();
            paramEditor.g_main.ColumnDefinitions.Add(columnDefinitionL);

            ColumnDefinition columnDefinitionM = new ColumnDefinition();
            columnDefinitionM.Width = new GridLength(20);
            paramEditor.g_main.ColumnDefinitions.Add(columnDefinitionM);

            ColumnDefinition columnDefinitionR = new ColumnDefinition();
            paramEditor.g_main.ColumnDefinitions.Add(columnDefinitionR);

            TextEditor textEditorL = new TextEditor();
            textEditorL.ShowLineNumbers = true;
            textEditorL.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
            textEditorL.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            if (paramDicForChange != null && paramDicForChange.Keys != null && paramDicForChange.Keys.Count > 0)
            {
                foreach (string str in paramDicForChange.Keys)
                {
                    textEditorL.Text += $"{str}\n";
                }
            }
            Grid.SetRow(textEditorL, 0);
            Grid.SetColumn(textEditorL, 0);
            paramEditor.g_main.Children.Add(textEditorL);

            TextBlock textBlock = new TextBlock();
            textBlock.Text = "=";
            textBlock.HorizontalAlignment = HorizontalAlignment.Center;
            textBlock.VerticalAlignment = VerticalAlignment.Center;
            Grid.SetRow(textBlock, 0);
            Grid.SetColumn(textBlock, 1);
            paramEditor.g_main.Children.Add(textBlock);

            TextEditor textEditorR = new TextEditor();
            textEditorR.ShowLineNumbers = true;
            textEditorR.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
            textEditorR.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            if (paramDicForChange != null && paramDicForChange.Keys != null && paramDicForChange.Keys.Count > 0)
            {
                foreach (string str in paramDicForChange.Values)
                {
                    textEditorR.Text += $"{str}\n";
                }
            }
            Grid.SetRow(textEditorR, 0);
            Grid.SetColumn(textEditorR, 2);
            paramEditor.g_main.Children.Add(textEditorR);

            Button btnOk = new Button();
            btnOk.Height = 30;
            btnOk.Content = Application.Current.FindResource("Ok").ToString();
            btnOk.Click += (s, ex) =>
            {
                List<string> keyList = textEditorL.Text.Replace("\r", "").Split('\n').Where(str => str.Trim() != "").ToList();
                List<string> valueList = textEditorR.Text.Replace("\r", "").Split('\n').Where(str => str.Trim() != "").ToList();
                if (keyList.Count == valueList.Count)
                {
                    Dictionary<string, string> changedParamDic = new Dictionary<string, string>();
                    for (int i = 0; i < keyList.Count; ++i)
                    {
                        changedParamDic.Add(keyList[i], valueList[i]);
                    }
                    paramDicForChange = changedParamDic;
                }

                paramEditor.Close();
            };
            Grid.SetRow(btnOk, 1);
            Grid.SetColumnSpan(btnOk, 3);
            paramEditor.g_main.Children.Add(btnOk);

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
                editor.Height = WindowActualHeight - 100;
            }
        }

        // ---------------------------------------------------- Common Logic
        private void Save(bool isRename)
        {
            TextBox tbName = new TextBox();
            tbName.Margin = new Thickness(5, 8, 5, 8);
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
                int result = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object>() { tbName, Application.Current.FindResource("Ok").ToString(), Application.Current.FindResource("Cancel").ToString() }, Application.Current.FindResource("Name").ToString(), Application.Current.FindResource("Saving").ToString(), MessageBoxImage.Information);

                if (result != 1)
                {
                    return;
                }

                newName = tbName.Text;
            }

            Analyzer analyzer = new Analyzer();
            analyzer.code = editor.Text;
            analyzer.name = newName;
            analyzer.paramDic = paramDicForChange;
            string json = JsonConvert.SerializeObject(analyzer);

            string fileName = $".\\Analyzers\\{newName}.json";
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

            CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetterWithTimmer(), Application.Current.FindResource("SuccessfullySaved").ToString(), Application.Current.FindResource("Save").ToString(), MessageBoxButton.OK, MessageBoxImage.Information);

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

        public class DocumentViewModel : INotifyPropertyChanged
        {
            private bool _isReadOnly;
            private readonly RoslynHost _host;

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

            private static MethodInfo HasSubmissionResult { get; } =
                typeof(Compilation).GetMethod(nameof(HasSubmissionResult), BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);

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