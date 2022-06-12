using CustomizableMessageBox;
using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Search;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.CSharp.Scripting.Hosting;
using Microsoft.CodeAnalysis.Scripting;
using Microsoft.CodeAnalysis.Scripting.Hosting;
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
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ExcelTool
{
    /// <summary>
    /// AnalyzerEditor.xaml の相互作用ロジック
    /// </summary>
    public partial class AnalyzerEditor : Window
    {
        private readonly ObservableCollection<DocumentViewModel> _documents;
        private RoslynHost _host;
        RoslynCodeEditor editor;
        Dictionary<string, string> paramDicForChange;
        public AnalyzerEditor()
        {
            InitializeComponent();

            _documents = new ObservableCollection<DocumentViewModel>();
            Items.ItemsSource = _documents;
        }

        private void BtnSaveClick(object sender, RoutedEventArgs e)
        {
            Save(true);
        }

        private void BtnExitClick(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void WindowLoaded(object sender, RoutedEventArgs e)
        {
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

        private void AddNewDocument(DocumentViewModel previous = null)
        {
            _documents.Add(new DocumentViewModel(_host, previous));
        }

        private void OnItemLoaded(object sender, EventArgs e)
        {
            editor = (RoslynCodeEditor)sender;
            editor.Loaded -= OnItemLoaded;
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

        private void CbAnalyzersSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            btn_delete.IsEnabled = cb_analyzers.SelectedIndex >= 1 ? true : false;
            btn_editparam.IsEnabled = cb_analyzers.SelectedIndex >= 1 ? true : false;

            editor.Text = GlobalObjects.GlobalObjects.GetDefaultCode();

            if (cb_analyzers.SelectedIndex == 0)
            {
                return;
            }
            Analyzer analyzer = JsonConvert.DeserializeObject<Analyzer>(File.ReadAllText($".\\Analyzers\\{cb_analyzers.SelectedItem.ToString()}.json"));
            editor.Text = analyzer.code;

            paramDicForChange = analyzer.paramDic;
        }

        private void CbAnalyzersPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
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

            cb_analyzers.ItemsSource = AnalyzersList;
        }

        private void BtnDeleteClick(object sender, RoutedEventArgs e)
        {
            String path = $"{System.Environment.CurrentDirectory}\\Analyzers\\{cb_analyzers.SelectedItem.ToString()}.json";
            MessageBoxResult result = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"{Application.Current.FindResource("Delete").ToString()}\n{path}", Application.Current.FindResource("Warning").ToString(), MessageBoxButton.OKCancel, MessageBoxImage.Warning);
            if (result == MessageBoxResult.OK)
            {
                File.Delete(path);
                cb_analyzers.SelectedIndex = 0;
            }
        }

        private void WindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            IniHelper.SetWindowSize(this.Name.Substring(2), new Point(this.Width, this.Height));
        }

        class DocumentViewModel : INotifyPropertyChanged
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

        private void SaveByKeyDownCanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void SaveByKeyDown(object sender, ExecutedRoutedEventArgs e)
        {
            if (cb_analyzers.SelectedIndex >= 1)
            {
                Save(false);
            }
            else
            {
                Save(true);
            }
        }

        private void RenameSaveByKeyDownCanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void RenameSaveByKeyDown(object sender, ExecutedRoutedEventArgs e)
        {
            Save(true);
        }

        private void Save(bool isRename)
        {
            TextBox tbName = new TextBox();
            tbName.Margin = new Thickness(5, 8, 5, 8);
            tbName.VerticalContentAlignment = VerticalAlignment.Center;

            string newName = "";
            if (cb_analyzers.SelectedIndex >= 1)
            {
                newName = cb_analyzers.SelectedItem.ToString();
            }
            if (isRename)
            {
                if (cb_analyzers.SelectedIndex >= 1)
                {
                    tbName.Text = $"Copy Of {cb_analyzers.SelectedItem}";
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

        private void BtnEditParamClick(object sender, RoutedEventArgs e)
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
                List<string> keyList = textEditorL.Text.Replace("\r" , "").Split('\n').Where(str => str.Trim() != "").ToList();
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

        private void AnalyzerEditorSizeChanged(object sender, SizeChangedEventArgs e)
        {
            ResetEditorSize(editor);
        }

        private void ResetEditorSize(RoslynCodeEditor editor)
        {
            if (editor != null && this.WindowState != WindowState.Minimized)
            {
                editor.Height = this.ActualHeight - 100;
            }
        }

        private void AnalyzerEditorStateChanged(object sender, EventArgs e)
        {
            ResetEditorSize(editor);
        }
    }
}
