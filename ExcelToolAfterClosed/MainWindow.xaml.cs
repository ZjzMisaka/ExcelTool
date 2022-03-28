using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelToolAfterClosed
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        String[] args = null;

        public MainWindow(String[] args)
        {
            this.args = args;
            InitializeComponent();
        }

        private void WindowLoaded(object sender, RoutedEventArgs e)
        {
            pb_process.Maximum = args.Length;
            int index = 0;

            Thread th = new Thread(WaitForClose);
            th.Start();
            th.Join();

            foreach (string path in args)
            {
                try
                {
                    File.Delete(path.Replace("|SPACE|", " "));
                    ++index;
                    pb_process.Value = index;
                }
                catch (Exception ex)
                {
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"删除Dll文件失败\n{path}\n{ex.Message}\n{ex.StackTrace}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            
            this.Close();
        }

        private void WaitForClose()
        {
            while (Process.GetProcessesByName("ExcelTool").Length > 0)
            {
                Thread.Sleep(100);
            }
        }
    }
}
