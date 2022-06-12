using CustomizableMessageBox;
using ExcelTool.ViewModel;
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
        public AnalyzerEditor()
        {
            InitializeComponent();
            this.DataContext = new AnalyzerViewModel();
        }
    }
}
