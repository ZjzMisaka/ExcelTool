using ExcelTool.Helper;
using GongSolutions.Wpf.DragDrop;
using ICSharpCode.AvalonEdit.Document;
using Microsoft.Toolkit.Mvvm.ComponentModel;
using Microsoft.Toolkit.Mvvm.Input;
using Microsoft.Toolkit.Mvvm.Messaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ExcelTool.ViewModel
{
    class MainWindowViewModel : ObservableObject, IDropTarget
    {
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

        private TextDocument teParamsDocument = new TextDocument();
        public TextDocument TeParamsDocument
        {
            get => teParamsDocument;
            set
            {
                SetProperty<TextDocument>(ref teParamsDocument, value);
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
        public MainWindowViewModel()
        {
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
            //WeakReferenceMessenger.Default.Register<string, string>(this, "MainWindowViewModel",(obj, msg) => 
            //{
            //    MethodInfo method = obj.GetType().GetMethod(msg);
            //});
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
            TeParamsDocument = new TextDocument(SelectedParamsItem);
        }

        private void EditParam()
        {
            //ParamEditor paramEditor = new ParamEditor();

            //string paramStr = $"{te_params.Text.Replace("\r\n", "").Replace("\n", "")}";
            //te_params.Text = paramStr;
            //Dictionary<string, Dictionary<string, string>> paramDicEachAnalyzer = ParamHelper.GetParamDicEachAnalyzer(paramStr);

            //int rowNum = -1;

            //List<String> analyzersList = te_analyzers.Text.Split('\n').Where(str => str.Trim() != "").ToList();

            //ColumnDefinition columnDefinitionL = new ColumnDefinition();
            //columnDefinitionL.Width = new GridLength(100);
            //paramEditor.g_main.ColumnDefinitions.Add(columnDefinitionL);
            //ColumnDefinition columnDefinitionR = new ColumnDefinition();
            //paramEditor.g_main.ColumnDefinitions.Add(columnDefinitionR);

            //List<string> addedAnalyzerNameList = new List<string>();

            //foreach (string analyzerName in analyzersList)
            //{
            //    if (addedAnalyzerNameList.Contains(analyzerName))
            //    {
            //        continue;
            //    }
            //    else
            //    {
            //        addedAnalyzerNameList.Add(analyzerName);
            //    }

            //    Analyzer analyzer = JsonConvert.DeserializeObject<Analyzer>(File.ReadAllText($".\\Analyzers\\{analyzerName}.json"));

            //    RowDefinition rowDefinition = new RowDefinition();
            //    rowDefinition.Height = new GridLength(40);
            //    paramEditor.g_main.RowDefinitions.Add(rowDefinition);
            //    ++rowNum;
            //    TextBlock textBlockAnalyzerName = new TextBlock();
            //    textBlockAnalyzerName.Height = 35;
            //    textBlockAnalyzerName.FontWeight = FontWeight.FromOpenTypeWeight(600);
            //    textBlockAnalyzerName.FontSize = 20;
            //    textBlockAnalyzerName.VerticalAlignment = VerticalAlignment.Bottom;
            //    textBlockAnalyzerName.Text = analyzerName;
            //    Grid.SetRow(textBlockAnalyzerName, rowNum);
            //    Grid.SetColumnSpan(textBlockAnalyzerName, 2);
            //    paramEditor.g_main.Children.Add(textBlockAnalyzerName);

            //    Dictionary<string, string> paramDic = null;
            //    if (paramDicEachAnalyzer.ContainsKey(analyzerName))
            //    {
            //        paramDic = paramDicEachAnalyzer[analyzerName];
            //    }

            //    if (analyzer.paramDic == null || analyzer.paramDic.Keys == null || analyzer.paramDic.Keys.Count == 0)
            //    {
            //        continue;
            //    }
            //    foreach (string key in analyzer.paramDic.Keys)
            //    {
            //        RowDefinition rowDefinitionKv = new RowDefinition();
            //        rowDefinitionKv.Height = new GridLength(30);
            //        paramEditor.g_main.RowDefinitions.Add(rowDefinitionKv);
            //        ++rowNum;

            //        TextBlock textBlockKey = new TextBlock();
            //        textBlockKey.Height = 25;
            //        textBlockKey.Text = analyzer.paramDic[key];
            //        Grid.SetRow(textBlockKey, rowNum);
            //        Grid.SetColumn(textBlockKey, 0);
            //        textBlockKey.MouseEnter += (s, ex) =>
            //        {
            //            textBlockKey.Text = key;
            //        };
            //        textBlockKey.TouchEnter += (s, ex) =>
            //        {
            //            textBlockKey.Text = key;
            //        };
            //        textBlockKey.MouseLeave += (s, ex) =>
            //        {
            //            textBlockKey.Text = analyzer.paramDic[key];
            //        };
            //        textBlockKey.TouchLeave += (s, ex) =>
            //        {
            //            textBlockKey.Text = analyzer.paramDic[key];
            //        };
            //        paramEditor.g_main.Children.Add(textBlockKey);

            //        TextBox tbValue = new TextBox();
            //        tbValue.Height = 25;
            //        tbValue.Margin = new Thickness(5, 0, 0, 0);
            //        tbValue.HorizontalAlignment = HorizontalAlignment.Stretch;
            //        tbValue.VerticalContentAlignment = VerticalAlignment.Center;
            //        if (paramDic != null && paramDic.ContainsKey(key))
            //        {
            //            tbValue.Text = paramDic[key];
            //        }
            //        else if (paramDicEachAnalyzer.ContainsKey("public") && paramDicEachAnalyzer["public"] != null && paramDicEachAnalyzer["public"].ContainsKey(key))
            //        {
            //            tbValue.Text = paramDicEachAnalyzer["public"][key];
            //        }
            //        tbValue.TextChanged += (s, ex) =>
            //        {
            //            if (!paramDicEachAnalyzer.ContainsKey(analyzerName))
            //            {
            //                paramDicEachAnalyzer.Add(analyzerName, new Dictionary<string, string>());
            //            }
            //            paramDicEachAnalyzer[analyzerName][key] = tbValue.Text;
            //        };
            //        Grid.SetRow(tbValue, rowNum);
            //        Grid.SetColumn(tbValue, 1);

            //        paramEditor.g_main.Children.Add(tbValue);
            //    }
            //}

            //GridSplitter gridSplitter = new GridSplitter();
            //gridSplitter.HorizontalAlignment = HorizontalAlignment.Right;
            //gridSplitter.VerticalAlignment = VerticalAlignment.Stretch;
            //gridSplitter.Width = 4;
            //gridSplitter.BorderThickness = new Thickness(1, 0, 1, 0);
            //gridSplitter.BorderBrush = Brushes.Black;
            //gridSplitter.Background = Brushes.AntiqueWhite;
            //Grid.SetRow(gridSplitter, 0);
            //Grid.SetRowSpan(gridSplitter, rowNum + 1 <= 0 ? 1 : rowNum + 1);
            //Grid.SetColumn(gridSplitter, 0);
            //paramEditor.g_main.Children.Add(gridSplitter);

            //RowDefinition rowDefinitionBlank = new RowDefinition();
            //paramEditor.g_main.RowDefinitions.Add(rowDefinitionBlank);
            //++rowNum;
            //RowDefinition rowDefinitionOk = new RowDefinition();
            //rowDefinitionOk.Height = new GridLength(30);
            //paramEditor.g_main.RowDefinitions.Add(rowDefinitionOk);
            //++rowNum;
            //Button btnOk = new Button();
            //btnOk.Height = 30;
            //btnOk.Content = Application.Current.FindResource("Ok").ToString();
            //btnOk.Click += (s, ex) =>
            //{
            //    te_params.Text = ParamHelper.GetParamStr(paramDicEachAnalyzer);
            //    paramEditor.Close();
            //};
            //Grid.SetRow(btnOk, rowNum);
            //Grid.SetColumnSpan(btnOk, 2);
            //paramEditor.g_main.Children.Add(btnOk);

            //paramEditor.ShowDialog();
        }
        private void TbParamsLostFocus()
        {
            
        }

        private void DragOverPath(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }

        private void DropPath(object sender, DragEventArgs e)
        {
            TbBasePathText = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
        }

        public void DragEnter(IDropInfo dropInfo)
        {
            dropInfo.Effects = DragDropEffects.Copy;
        }

        public void DragOver(IDropInfo dropInfo)
        {
            dropInfo.DropTargetAdorner = DropTargetAdorners.Highlight;
            dropInfo.Effects = DragDropEffects.Copy;
        }

        public void DragLeave(IDropInfo dropInfo)
        {
            //throw new NotImplementedException();
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
            // TODO 翻译
            fileDialog.Title = "请选择文件";
            fileDialog.Filter = "Excel文件|*.xlsx;*.xlsm;*.xls|All files(*.*)|*.*";
            if (fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string[] names = fileDialog.FileNames;
                TbOutputNameText = Path.GetFileNameWithoutExtension(names[0]);
            }
        }
    }
}
