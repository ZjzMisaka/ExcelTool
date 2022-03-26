using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelToolAfterClosed
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        /// <summary>
        /// 重写Startup函数
        /// </summary>
        /// <param name="e"></param>
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            // 接收参数数组
            string[] args = e.Args;
            // 定义为字符数组也可以
            //string[] args = e.Args;
            // 判断参数中是否包含 TestWindows
            if (args.Length != 0)
            {
                new MainWindow(args).Show();
            }
        }
    }
}
