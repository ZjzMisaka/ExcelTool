using Amib.Threading;
using ClosedXML.Excel;
using GlobalObjects;
using Microsoft.CSharp;
using Newtonsoft.Json;
using System;
using System.CodeDom.Compiler;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using CustomizableMessageBox;
using System.Xml;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Document;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Media;
using ExcelTool.Helper;
using ExcelTool.ViewModel;

namespace ExcelTool
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = new MainWindowViewModel();
        }
    }
}
