using CustomizableMessageBox;
using ExcelTool.ViewModel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
    /// SheetExplainerEditor.xaml の相互作用ロジック
    /// </summary>
    public partial class SheetExplainerEditor : Window
    {
        public SheetExplainerEditor()
        {
            InitializeComponent();
            this.DataContext = new SheetExplainerViewModel();
        }
    }
}
