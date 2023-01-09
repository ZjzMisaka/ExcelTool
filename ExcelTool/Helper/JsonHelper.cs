using CustomizableMessageBox;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static CustomizableMessageBox.MessageBox;
using System.Windows;
using GlobalObjects;
using Newtonsoft.Json.Converters;

namespace ExcelTool.Helper
{
    public static class JsonHelper
    {
        static public T TryDeserializeObject<T>(string filePath, bool useLog = false) where T : class
        {
            T res = null;
            try
            {
                JsonSerializerSettings jsonSerializerSettings = new JsonSerializerSettings();
                jsonSerializerSettings.MissingMemberHandling = MissingMemberHandling.Error;
                res = JsonConvert.DeserializeObject<T>(File.ReadAllText(filePath), jsonSerializerSettings);
            }
            catch
            {
                if (!useLog)
                {
                    CustomizableMessageBox.MessageBox.Show(new RefreshList { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, Application.Current.FindResource("ErrorWhileParsingJsonFile").ToString().Replace("{0}", $"\n{filePath}"), Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
                }
                else
                {
                    Logger.Error(Application.Current.FindResource("ErrorWhileParsingJsonFile").ToString().Replace("{0}", $"\n{filePath}"));
                }
            }
            return res;
        }
    }
}
