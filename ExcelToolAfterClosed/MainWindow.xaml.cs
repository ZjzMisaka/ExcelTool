using CustomizableMessageBox;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, $"Failed to delete dll file. \n{path}\n{ex.Message}\n{ex.StackTrace}", "错误", MessageBoxImage.Error);
                }
            }

            w_AfterClosedMainWindow.Close();
        }

        private void WaitForClose()
        {
            int tried = 0;
            string message = "";
            while (Process.GetProcessesByName("ExcelTool").Length > 0)
            {
                foreach (var process in Process.GetProcessesByName("ExcelTool"))
                {
                    try
                    {
                        process.Kill();
                        process.WaitForExit();
                        ++tried;
                    }
                    catch (Win32Exception ex)
                    {
                        message += $"\n{ex.Message}";
                    }
                    catch (InvalidOperationException)
                    {
                    }
                }

                if (tried == 10)
                {
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object> { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, $"Tried 10 times and cannot kill process ExcelTool. \nLog: {message}", "Error", MessageBoxImage.Error);
                }

                Thread.Sleep(100);
            }
        }
    }
}
