using ExcelTool.Helper;
using ExcelTool.Model;
using Microsoft.Toolkit.Mvvm.ComponentModel;
using Microsoft.Toolkit.Mvvm.Input;
using Microsoft.Toolkit.Mvvm.Messaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ExcelTool.ViewModel
{
    class MainWindowViewModel
    {
        MainWindowModel model = new MainWindowModel();
        public ICommand BtnOpenSheetExplainerEditorClickCommand { get; set; }
        public ICommand BtnOpenAnalyzerEditorClickCommand { get; set; }
        public ICommand CbSheetExplainersPreviewMouseLeftButtonDownCommand { get; set; }
        public ICommand CbSheetExplainersSelectionChangedCommand { get; set; }
        public MainWindowViewModel()
        {
            BtnOpenSheetExplainerEditorClickCommand = new RelayCommand(BtnOpenSheetExplainerEditorClick);
            BtnOpenAnalyzerEditorClickCommand = new RelayCommand(BtnOpenAnalyzerEditorClick);
            CbSheetExplainersPreviewMouseLeftButtonDownCommand = new RelayCommand(CbSheetExplainersPreviewMouseLeftButtonDown);
            CbSheetExplainersSelectionChangedCommand = new RelayCommand(CbSheetExplainersSelectionChanged);

            WeakReferenceMessenger.Default.Register<string, string>(this, "MainWindowViewModel",(obj, msg) => 
            {
                MethodInfo method = obj.GetType().GetMethod(msg);
            });
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
            model.SheetExplainersItems = FileHelper.GetSheetExplainersList();
        }

        private void CbSheetExplainersSelectionChanged()
        {
            //if (cb_sheetexplainers.SelectedIndex == 0)
            //{
            //    return;
            //}
            //te_sheetexplainers.Text += $"{cb_sheetexplainers.SelectedItem}\n";
            //cb_sheetexplainers.SelectedIndex = 0;
        }
    }
}
