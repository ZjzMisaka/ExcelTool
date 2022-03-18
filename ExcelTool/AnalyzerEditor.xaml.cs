using CustomizableMessageBox;
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
        public AnalyzerEditor()
        {
            InitializeComponent();

            _documents = new ObservableCollection<DocumentViewModel>();
            Items.ItemsSource = _documents;
        }

        private void BtnSaveClick(object sender, RoutedEventArgs e)
        {
            Save();
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
            string[] dlls = IniHelper.GetDlls().Split('|');
            foreach (string dll in dlls)
            {
                var dllFiles = Assembly.LoadFile(System.IO.Path.Combine(System.Environment.CurrentDirectory, dll));

                foreach (Type type in dllFiles.GetExportedTypes())
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

            editor.Text = GlobalObjects.GlobalObjects.GetDefaultCode();

            if (cb_analyzers.SelectedIndex == 0)
            {
                return;
            }
            Analyzer analyzer = JsonConvert.DeserializeObject<Analyzer>(File.ReadAllText($".\\Analyzers\\{cb_analyzers.SelectedItem.ToString()}.json"));
            editor.Text = analyzer.code;
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
                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"新建文件夹失败\n{ex.Message}", "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
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
            MessageBoxResult result = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"删除文件\n{path}", "Warning", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
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

        private void ScrollViewerPreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            var eventArg = new MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta);
            eventArg.RoutedEvent = UIElement.MouseWheelEvent;
            eventArg.Source = sender;
            sv_editor.RaiseEvent(eventArg);
        }

        private void SaveByKeyDownCanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void SaveByKeyDown(object sender, ExecutedRoutedEventArgs e)
        {
            Save();
        }

        private void Save()
        {
            TextBox tbName = new TextBox();
            tbName.Margin = new Thickness(5, 8, 5, 8);
            tbName.VerticalContentAlignment = VerticalAlignment.Center;
            if (cb_analyzers.SelectedIndex >= 1)
            {
                tbName.Text = $"Copy Of {cb_analyzers.SelectedItem}";
            }
            int result = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object>() { tbName, "OK", "CANCEL" }, "name", "saving", MessageBoxImage.Information);
            if (result == 1)
            {
                string analyzerName = tbName.Text;

                Analyzer analyzer = new Analyzer();
                analyzer.code = editor.Text;
                analyzer.name = analyzerName;
                string json = JsonConvert.SerializeObject(analyzer);

                string fileName = $".\\Analyzers\\{analyzerName}.json";
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

                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), "保存成功", "保存", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            
        }
    }
}
