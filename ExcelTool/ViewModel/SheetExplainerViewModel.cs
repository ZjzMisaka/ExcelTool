using CustomizableMessageBox;
using ExcelTool.Helper;
using GongSolutions.Wpf.DragDrop;
using Microsoft.Toolkit.Mvvm.ComponentModel;
using Microsoft.Toolkit.Mvvm.Input;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.Tracing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using static CustomizableMessageBox.MessageBox;
using static System.Net.WebRequestMethods;
using File = System.IO.File;

namespace ExcelTool.ViewModel
{
    class SheetExplainerViewModel : ObservableObject, IDropTarget
    {
        private double windowWidth;
        public double WindowWidth
        {
            get { return windowWidth; }
            set
            {
                SetProperty<double>(ref windowWidth, value);
            }
        }

        private double windowHeight;
        public double WindowHeight
        {
            get { return windowHeight; }
            set
            {
                SetProperty<double>(ref windowHeight, value);
            }
        }

        private string windowName = "SheetExplainerEditor";
        public string WindowName
        {
            get { return windowName; }
            set
            {
                SetProperty<string>(ref windowName, value);
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

        private int selectedSheetExplainersIndex = 0;
        public int SelectedSheetExplainersIndex
        {
            get { return selectedSheetExplainersIndex; }
            set
            {
                SetProperty<int>(ref selectedSheetExplainersIndex, value);
            }
        }

        private string selectedSheetExplainersItem = null;
        public string SelectedSheetExplainersItem
        {
            get { return selectedSheetExplainersItem; }
            set
            {
                SetProperty<string>(ref selectedSheetExplainersItem, value);
            }
        }

        private bool btnDeleteIsEnabled = false;
        public bool BtnDeleteIsEnabled
        {
            get { return btnDeleteIsEnabled; }
            set
            {
                SetProperty<bool>(ref btnDeleteIsEnabled, value);
            }
        }

        private string tbPathsText = "";
        public string TbPathsText
        {
            get { return tbPathsText; }
            set
            {
                SetProperty<string>(ref tbPathsText, value);
            }
        }

        private int selectedFileNamesTypeIndex = 0;
        public int SelectedFileNamesTypeIndex
        {
            get { return selectedFileNamesTypeIndex; }
            set
            {
                SetProperty<int>(ref selectedFileNamesTypeIndex, value);
            }
        }

        private string tbFileNamesText = "";
        public string TbFileNamesText
        {
            get { return tbFileNamesText; }
            set
            {
                SetProperty<string>(ref tbFileNamesText, value);
            }
        }

        private int selectedSheetNamesTypeIndex = 0;
        public int SelectedSheetNamesTypeIndex
        {
            get { return selectedSheetNamesTypeIndex; }
            set
            {
                SetProperty<int>(ref selectedSheetNamesTypeIndex, value);
            }
        }

        private string tbSheetNamesText = "";
        public string TbSheetNamesText
        {
            get { return tbSheetNamesText; }
            set
            {
                SetProperty<string>(ref tbSheetNamesText, value);
            }
        }

        public ICommand WindowLoadedCommand { get; set; }
        public ICommand WindowClosingCommand { get; set; }
        public ICommand KeyBindingSaveCommand { get; set; }
        public ICommand KeyBindingRenameSaveCommand { get; set; }
        public ICommand CbSheetExplainersPreviewMouseLeftButtonDownCommand { get; set; }
        public ICommand CbSheetExplainersSelectionChangedCommand { get; set; }
        public ICommand BtnDeleteClickCommand { get; set; }
        public ICommand BtnClearTempClickCommand { get; set; }
        public ICommand BtnSaveClickCommand { get; set; }
        public SheetExplainerViewModel()
        {
            WindowLoadedCommand = new RelayCommand<RoutedEventArgs>(WindowLoaded);
            WindowClosingCommand = new RelayCommand<CancelEventArgs>(WindowClosing);
            KeyBindingSaveCommand = new RelayCommand(SaveByKeyDown);
            KeyBindingRenameSaveCommand = new RelayCommand(RenameSaveByKeyDown);
            CbSheetExplainersPreviewMouseLeftButtonDownCommand = new RelayCommand(CbSheetExplainersPreviewMouseLeftButtonDown);
            CbSheetExplainersSelectionChangedCommand = new RelayCommand(CbSheetExplainersSelectionChanged);
            BtnDeleteClickCommand = new RelayCommand(BtnDeleteClick);
            BtnClearTempClickCommand = new RelayCommand(BtnClearTempClick);
            BtnSaveClickCommand = new RelayCommand(BtnSaveClick);
        }

        private void WindowLoaded(RoutedEventArgs e)
        {
            Window window = (Window)e.Source;

            double width = IniHelper.GetWindowSize(WindowName).X;
            double height = IniHelper.GetWindowSize(WindowName).Y;
            if (width > 0)
            {
                window.Width = width;
            }

            if (height > 0)
            {
                window.Height = height;
            }
        }

        private void WindowClosing(CancelEventArgs e)
        {
            IniHelper.SetWindowSize(WindowName, new Point(WindowWidth, WindowHeight));
        }

        private void SaveByKeyDown()
        {
            if (SelectedSheetExplainersIndex >= 1)
            {
                Save(false);
            }
            else
            {
                Save(true);
            }
        }

        private void RenameSaveByKeyDown()
        {
            Save(true);
        }

        private void CbSheetExplainersPreviewMouseLeftButtonDown()
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
                CustomizableMessageBox.MessageBox.Show(new RefreshList { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, $"{Application.Current.FindResource("FailedToCreateANewFolder").ToString()}\n{ex.Message}", Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
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

            SheetExplainersItems = sheetExplainersList;
        }

        private void CbSheetExplainersSelectionChanged()
        {
            BtnDeleteIsEnabled = SelectedSheetExplainersIndex >= 1 ? true : false;

            TbPathsText = "";
            SelectedFileNamesTypeIndex = 0;
            TbFileNamesText = "";
            SelectedSheetNamesTypeIndex = 0;
            TbSheetNamesText = "";

            if (SelectedSheetExplainersIndex == 0)
            {
                return;
            }
            SheetExplainer sheetExplainer = JsonHelper.TryDeserializeObject<SheetExplainer>($".\\SheetExplainers\\{SelectedSheetExplainersItem}.json");
            if (sheetExplainer == null)
            {
                SelectedSheetExplainersIndex = 0;
                return;
            }
            foreach (String str in sheetExplainer.pathes)
            {
                TbPathsText += $"{str}\n";
            }
            if (sheetExplainer.fileNames.Key == FindingMethod.SAME)
            {
                SelectedFileNamesTypeIndex = 0;
            }
            else if (sheetExplainer.fileNames.Key == FindingMethod.CONTAIN)
            {
                SelectedFileNamesTypeIndex = 1;
            }
            else if (sheetExplainer.fileNames.Key == FindingMethod.REGEX)
            {
                SelectedFileNamesTypeIndex = 2;
            }
            else if (sheetExplainer.fileNames.Key == FindingMethod.ALL)
            {
                SelectedFileNamesTypeIndex = 3;
            }
            foreach (String str in sheetExplainer.fileNames.Value)
            {
                TbFileNamesText += $"{str}\n";
            }
            if (sheetExplainer.sheetNames.Key == FindingMethod.SAME)
            {
                SelectedSheetNamesTypeIndex = 0;
            }
            else if (sheetExplainer.sheetNames.Key == FindingMethod.CONTAIN)
            {
                SelectedSheetNamesTypeIndex = 1;
            }
            else if (sheetExplainer.sheetNames.Key == FindingMethod.REGEX)
            {
                SelectedSheetNamesTypeIndex = 2;
            }
            else if (sheetExplainer.sheetNames.Key == FindingMethod.ALL)
            {
                SelectedSheetNamesTypeIndex = 3;
            }
            foreach (String str in sheetExplainer.sheetNames.Value)
            {
                TbSheetNamesText += $"{str}\n";
            }
        }

        private void BtnDeleteClick()
        {
            String path = $"{System.Environment.CurrentDirectory}\\SheetExplainers\\{SelectedSheetExplainersItem}.json";
            MessageBoxResult result = CustomizableMessageBox.MessageBox.Show($"{Application.Current.FindResource("Delete").ToString()}\n{path}", Application.Current.FindResource("Warning").ToString(), MessageBoxButton.OKCancel, MessageBoxImage.Warning);
            if (result == MessageBoxResult.OK)
            {
                File.Delete(path);
                SelectedSheetExplainersIndex = 0;
            }
        }

        private void BtnClearTempClick()
        {
            String path = $"{System.Environment.CurrentDirectory}\\SheetExplainers";
            string[] files = Directory.GetFiles(path, "*.json");
            List<string> deleteList = new List<string>();
            string deleteStr = "";
            foreach (string file in files)
            {
                if (new FileInfo(file).Name.StartsWith("Temp_Sheet_Explainer_"))
                {
                    deleteList.Add(file);
                    deleteStr = $"{deleteStr}\n{file}";
                }
            }

            if (deleteList.Count == 0)
            {
                return;
            }

            MessageBoxResult result = CustomizableMessageBox.MessageBox.Show(Application.Current.FindResource("ToBeDeletedSoon").ToString().Replace("{0}", deleteStr), Application.Current.FindResource("Warning").ToString(), MessageBoxButton.OKCancel, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Cancel)
            {
                return;
            }

            foreach (string deletePath in deleteList)
            {
                if (Path.GetFileNameWithoutExtension(deletePath) == SelectedSheetExplainersItem)
                {
                    SelectedSheetExplainersIndex = 0;
                }
            }

            foreach (string deletePath in deleteList)
            {
                File.Delete(deletePath);
            }
        }

        private void BtnSaveClick()
        {
            Save(true);
        }

        public void DragEnter(IDropInfo dropInfo)
        {
            // throw new NotImplementedException();
        }

        public void DragOver(IDropInfo dropInfo)
        {
            dropInfo.DropTargetAdorner = DropTargetAdorners.Highlight;
            dropInfo.Effects = DragDropEffects.Copy;
        }

        public void DragLeave(IDropInfo dropInfo)
        {
            // throw new NotImplementedException();
        }

        public void Drop(IDropInfo dropInfo)
        {
            if (dropInfo.DragInfo?.VisualSource is null
                && dropInfo.Data is DataObject dataObject
                && dataObject.GetDataPresent(DataFormats.FileDrop)
                && dataObject.ContainsFileDropList())
            {
                string parentName = ((TextBox)((Grid)dropInfo.TargetScrollViewer.Parent).TemplatedParent).Name;
                if (parentName == "tb_paths" || parentName == "tb_filenames")
                {
                    string msg = dataObject.GetFileDropList()[0];
                    Clipboard.SetText(msg);
                    CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetterWithTimmer(), new RefreshList { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, msg, Application.Current.FindResource("Copied").ToString(), MessageBoxImage.Information);
                }
            }
            else
            {
                GongSolutions.Wpf.DragDrop.DragDrop.DefaultDropHandler.Drop(dropInfo);
            }
        }

        // ---------------------------------------------------- Common Logic
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

        private void Save(bool isRename)
        {
            TextBox tbName = new TextBox();
            tbName.Margin = new Thickness(5);
            tbName.Height = 30;
            tbName.VerticalContentAlignment = VerticalAlignment.Center;

            string newName = "";
            if (SelectedSheetExplainersIndex >= 1)
            {
                newName = SelectedSheetExplainersItem;
            }
            if (isRename)
            {
                if (SelectedSheetExplainersIndex >= 1)
                {
                    tbName.Text = $"Copy Of {SelectedSheetExplainersItem}";
                }
                int result = CustomizableMessageBox.MessageBox.Show(new RefreshList() { tbName, Application.Current.FindResource("Ok").ToString(), Application.Current.FindResource("Cancel").ToString() }, Application.Current.FindResource("Name").ToString(), Application.Current.FindResource("Saving").ToString(), MessageBoxImage.Information);

                if (result != 1)
                {
                    return;
                }

                if (tbName.Text == "")
                {
                    CustomizableMessageBox.MessageBox.Show(new RefreshList { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, Application.Current.FindResource("FileNameEmptyError").ToString(), Application.Current.FindResource("Error").ToString(), MessageBoxImage.Error);
                    return;
                }

                newName = tbName.Text;
            }

            SheetExplainer sheetExplainer = new SheetExplainer();
            sheetExplainer.pathes = StringListDeteleBlank(TbPathsText.Split('\n').ToList());

            FindingMethod fileNamesFindingMethod = new FindingMethod();
            if (SelectedFileNamesTypeIndex == 0)
            {
                fileNamesFindingMethod = FindingMethod.SAME;
            }
            else if (SelectedFileNamesTypeIndex == 1)
            {
                fileNamesFindingMethod = FindingMethod.CONTAIN;
            }
            else if (SelectedFileNamesTypeIndex == 2)
            {
                fileNamesFindingMethod = FindingMethod.REGEX;
            }
            else if (SelectedFileNamesTypeIndex == 3)
            {
                fileNamesFindingMethod = FindingMethod.ALL;
            }
            KeyValuePair<FindingMethod, List<string>> fileNames = new KeyValuePair<FindingMethod, List<string>>(fileNamesFindingMethod, StringListDeteleBlank(TbFileNamesText.Split('\n').ToList()));
            sheetExplainer.fileNames = fileNames;

            FindingMethod sheetNamesFindingMethod = new FindingMethod();
            if (SelectedSheetNamesTypeIndex == 0)
            {
                sheetNamesFindingMethod = FindingMethod.SAME;
            }
            else if (SelectedSheetNamesTypeIndex == 1)
            {
                sheetNamesFindingMethod = FindingMethod.CONTAIN;
            }
            else if (SelectedSheetNamesTypeIndex == 2)
            {
                sheetNamesFindingMethod = FindingMethod.REGEX;
            }
            else if (SelectedSheetNamesTypeIndex == 3)
            {
                sheetNamesFindingMethod = FindingMethod.ALL;
            }
            KeyValuePair<FindingMethod, List<string>> sheetNames = new KeyValuePair<FindingMethod, List<string>>(sheetNamesFindingMethod, StringListDeteleBlank(TbSheetNamesText.Split('\n').ToList()));
            sheetExplainer.sheetNames = sheetNames;

            FileHelper.SavaSheetExplainerJson(newName, sheetExplainer, true);

            CustomizableMessageBox.MessageBox.Show(GlobalObjects.GlobalObjects.GetPropertiesSetterWithTimmer(), new RefreshList { new ButtonSpacer(), Application.Current.FindResource("Ok").ToString() }, Application.Current.FindResource("SuccessfullySaved").ToString(), Application.Current.FindResource("Save").ToString(), MessageBoxImage.Information);
        }
    }
}
