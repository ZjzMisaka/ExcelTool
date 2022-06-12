using CustomizableMessageBox;
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
        }

        private void BtnSaveClick(object sender, RoutedEventArgs e)
        {
            Save(true);
        }

        private List<String> StringListDeteleBlank(List<String> list)
        {
            for (int i = 0; i < list.Count; ++i)
            {
                list[i] = list[i].Trim();
                if (list[i].Length == 0)
                {
                    list.RemoveAt(i);
                    --i;
                }
            }
            return list;
        }

        private void BtnExitClick(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void CbSheetExplainersSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            btn_delete.IsEnabled = cb_sheetexplainers.SelectedIndex >= 1 ? true : false;

            tb_relativepaths.Text = "";
            cb_filenamestype.SelectedIndex = 0;
            tb_filenames.Text = "";
            cb_sheetnamestype.SelectedIndex = 0;
            tb_sheetnames.Text = "";

            if (cb_sheetexplainers.SelectedIndex == 0)
            {
                return;
            }
            SheetExplainer sheetExplainer = JsonConvert.DeserializeObject<SheetExplainer>(File.ReadAllText($".\\SheetExplainers\\{cb_sheetexplainers.SelectedItem.ToString()}.json"));
            foreach(String str in sheetExplainer.pathes)
            {
                tb_relativepaths.Text += $"{str}\n";
            }
            if (sheetExplainer.fileNames.Key == FindingMethod.SAME)
            {
                cb_filenamestype.SelectedIndex = 0;
            }
            else if (sheetExplainer.fileNames.Key == FindingMethod.CONTAIN)
            {
                cb_filenamestype.SelectedIndex = 1;
            }
            else if (sheetExplainer.fileNames.Key == FindingMethod.REGEX)
            {
                cb_filenamestype.SelectedIndex = 2;
            }
            else if (sheetExplainer.fileNames.Key == FindingMethod.ALL)
            {
                cb_filenamestype.SelectedIndex = 3;
            }
            foreach (String str in sheetExplainer.fileNames.Value)
            {
                tb_filenames.Text += $"{str}\n";
            }
            if (sheetExplainer.sheetNames.Key == FindingMethod.SAME)
            {
                cb_sheetnamestype.SelectedIndex = 0;
            }
            else if (sheetExplainer.sheetNames.Key == FindingMethod.CONTAIN)
            {
                cb_sheetnamestype.SelectedIndex = 1;
            }
            else if (sheetExplainer.sheetNames.Key == FindingMethod.REGEX)
            {
                cb_sheetnamestype.SelectedIndex = 2;
            }
            else if (sheetExplainer.sheetNames.Key == FindingMethod.ALL)
            {
                cb_sheetnamestype.SelectedIndex = 3;
            }
            foreach (String str in sheetExplainer.sheetNames.Value)
            {
                tb_sheetnames.Text += $"{str}\n";
            }
        }

        private void CbSheetExplainersPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (!Directory.Exists(".\\SheetExplainers"))
                {
                    Directory.CreateDirectory(".\\SheetExplainers");
                }
            }
            catch (Exception ex)
            {
                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"{Application.Current.FindResource("FailedToCreateANewFolder").ToString()}\n{ex.Message}", Application.Current.FindResource("Error").ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            List<String> sheetExplainersList = Directory.GetFiles(".\\SheetExplainers", "*.json").ToList();
            sheetExplainersList.Insert(0, "");
            for (int i = 0; i < sheetExplainersList.Count; ++i)
            {
                String str = sheetExplainersList[i];
                sheetExplainersList[i] = str.Substring(str.LastIndexOf('\\') + 1);
                if (sheetExplainersList[i].Contains('.'))
                {
                    sheetExplainersList[i] = sheetExplainersList[i].Substring(0, sheetExplainersList[i].LastIndexOf('.'));
                }
            }
            
            cb_sheetexplainers.ItemsSource = sheetExplainersList;
        }

        private void BtnDeleteClick(object sender, RoutedEventArgs e)
        {
            String path = $"{System.Environment.CurrentDirectory}\\SheetExplainers\\{cb_sheetexplainers.SelectedItem.ToString()}.json";
            MessageBoxResult result = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), $"{Application.Current.FindResource("Delete").ToString()}\n{path}", Application.Current.FindResource("Warning").ToString(), MessageBoxButton.OKCancel, MessageBoxImage.Warning);
            if (result == MessageBoxResult.OK)
            {
                File.Delete(path);
                cb_sheetexplainers.SelectedIndex = 0;
            }
        }

        private void WindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            IniHelper.SetWindowSize(this.Name.Substring(2), new Point(this.Width, this.Height));
        }

        private void WindowLoaded(object sender, RoutedEventArgs e)
        {
            double width = IniHelper.GetWindowSize(this.Name.Substring(2)).X;
            double height = IniHelper.GetWindowSize(this.Name.Substring(2)).Y;
            if (width > 0)
            {
                this.Width = width;
            }

            if (height > 0)
            {
                this.Height = height;
            }
        }

        private void TextBoxPreviewDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string msg = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
                Clipboard.SetText(msg);
                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetterWithTimmer(), msg, Application.Current.FindResource("Copied").ToString(), MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void TextBoxPreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void SaveByKeyDownCanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void SaveByKeyDown(object sender, ExecutedRoutedEventArgs e)
        {
            if (cb_sheetexplainers.SelectedIndex >= 1)
            {
                Save(false);
            }
            else
            {
                Save(true);
            }
        }

        private void RenameSaveByKeyDownCanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void RenameSaveByKeyDown(object sender, ExecutedRoutedEventArgs e)
        {
            Save(true);
        }

        private void Save(bool isRename)
        {
            TextBox tbName = new TextBox();
            tbName.Margin = new Thickness(5, 8, 5, 8);
            tbName.VerticalContentAlignment = VerticalAlignment.Center;

            string newName = "";
            if (cb_sheetexplainers.SelectedIndex >= 1)
            {
                newName = cb_sheetexplainers.SelectedItem.ToString();
            }
            if (isRename)
            {
                if (cb_sheetexplainers.SelectedIndex >= 1)
                {
                    tbName.Text = $"Copy Of {cb_sheetexplainers.SelectedItem}";
                }
                int result = CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), new List<Object>() { tbName, Application.Current.FindResource("Ok").ToString(), Application.Current.FindResource("Cancel").ToString() }, Application.Current.FindResource("Name").ToString(), Application.Current.FindResource("Saving").ToString(), MessageBoxImage.Information);

                if (result != 1)
                {
                    return;
                }

                newName = tbName.Text;
            }

            SheetExplainer sheetExplainer = new SheetExplainer();
            sheetExplainer.pathes = StringListDeteleBlank(tb_relativepaths.Text.Split('\n').ToList());

            FindingMethod fileNamesFindingMethod = new FindingMethod();
            if (cb_filenamestype.SelectedIndex == 0)
            {
                fileNamesFindingMethod = FindingMethod.SAME;
            }
            else if (cb_filenamestype.SelectedIndex == 1)
            {
                fileNamesFindingMethod = FindingMethod.CONTAIN;
            }
            else if (cb_filenamestype.SelectedIndex == 2)
            {
                fileNamesFindingMethod = FindingMethod.REGEX;
            }
            else if (cb_filenamestype.SelectedIndex == 3)
            {
                fileNamesFindingMethod = FindingMethod.ALL;
            }
            KeyValuePair<FindingMethod, List<string>> fileNames = new KeyValuePair<FindingMethod, List<string>>(fileNamesFindingMethod, StringListDeteleBlank(tb_filenames.Text.Split('\n').ToList()));
            sheetExplainer.fileNames = fileNames;

            FindingMethod sheetNamesFindingMethod = new FindingMethod();
            if (cb_sheetnamestype.SelectedIndex == 0)
            {
                sheetNamesFindingMethod = FindingMethod.SAME;
            }
            else if (cb_sheetnamestype.SelectedIndex == 1)
            {
                sheetNamesFindingMethod = FindingMethod.CONTAIN;
            }
            else if (cb_sheetnamestype.SelectedIndex == 2)
            {
                sheetNamesFindingMethod = FindingMethod.REGEX;
            }
            else if (cb_sheetnamestype.SelectedIndex == 3)
            {
                sheetNamesFindingMethod = FindingMethod.ALL;
            }
            KeyValuePair<FindingMethod, List<string>> sheetNames = new KeyValuePair<FindingMethod, List<string>>(sheetNamesFindingMethod, StringListDeteleBlank(tb_sheetnames.Text.Split('\n').ToList()));
            sheetExplainer.sheetNames = sheetNames;

            string json = JsonConvert.SerializeObject(sheetExplainer);

            string fileName = $".\\SheetExplainers\\{newName}.json";
            FileStream fs = null;
            try
            {
                fs = File.Create(fileName);
                fs.Close();
                StreamWriter sw = File.CreateText(fileName);
                sw.Write(json);
                sw.Flush();
                sw.Close();
            }
            catch (Exception ex)
            {
                CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetter(), ex.Message, Application.Current.FindResource("Error").ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
            }

            CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetterWithTimmer(), Application.Current.FindResource("SuccessfullySaved").ToString(), Application.Current.FindResource("Save").ToString(), MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
