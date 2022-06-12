using Microsoft.Xaml.Behaviors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace ExcelTool
{
    public class EventCommandAction : TriggerAction<DependencyObject>
    {
        /// <summary>
        /// 事件要绑定的命令
        /// </summary>
        public ICommand Command
        {
            get { return (ICommand)GetValue(CommandProperty); }
            set { SetValue(CommandProperty, value); }
        }

        // Using a DependencyProperty as the backing store for MsgName.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty CommandProperty =
            DependencyProperty.Register("Command", typeof(ICommand), typeof(EventCommandAction), new PropertyMetadata(null));

        /// <summary>
        /// 绑定命令的参数，保持为空就是事件的参数
        /// </summary>
        public object CommandParameter
        {
            get { return (object)GetValue(CommandParateterProperty); }
            set { SetValue(CommandParateterProperty, value); }
        }

        // Using a DependencyProperty as the backing store for CommandParateter.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty CommandParateterProperty =
            DependencyProperty.Register("CommandParameter", typeof(object), typeof(EventCommandAction), new PropertyMetadata(null));

        //执行事件
        protected override void Invoke(object parameter)
        {
            if (CommandParameter != null)
                parameter = CommandParameter;
            var cmd = Command;
            if (cmd != null && cmd.CanExecute(parameter))
                cmd.Execute(parameter);
        }
    }
}
