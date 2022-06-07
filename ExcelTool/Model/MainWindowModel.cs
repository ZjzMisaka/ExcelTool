using Microsoft.Toolkit.Mvvm.ComponentModel;
using Microsoft.Toolkit.Mvvm.Messaging;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.Model
{
    class MainWindowModel : ObservableObject
    {
        private string teSheetExplainersValue;

        public string TeSheetExplainersValue
        {
            get => teSheetExplainersValue;
            set
            {
                SetProperty<string>(ref teSheetExplainersValue, value);
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
                WeakReferenceMessenger.Default.Send<string, string>("CbSheetExplainersSelectionChanged", "MainWindowViewModel");
                SetProperty<string>(ref selectedSheetExplainersItem, value);
            }
        }
    }
}
