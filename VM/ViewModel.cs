using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using Model;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using Excel = Microsoft.Office.Interop.Excel;

namespace VM
{
    public class ViewModel : DependencyObject, INotifyPropertyChanged
    {
        // property changed event
        public event PropertyChangedEventHandler PropertyChanged;

        public delegate void SelectConflict(int index);
        public event SelectConflict OnSelectConflict;

        public ScrollViewer SW_Props;
        public StackPanel spStatus;
        ObservableCollection<StoneInfo> ocSI;

        #region Inputparams
        public static readonly DependencyProperty ExcelFilePathProperty = DependencyProperty.Register(
            "ExcelFilePath", typeof(string), typeof(ViewModel));
        public string ExcelFilePath
        {
            get { return (string)GetValue(ExcelFilePathProperty); }
            set
            {
                SetValue(ExcelFilePathProperty, value);
                OnPropertyChanged("ExcelFilePath");
            }
        }

        /*private string xLX_FilePath;
        public string XLX_FilePath
        {
            get { return xLX_FilePath; }
            set
            {
                xLX_FilePath = value;
                OnPropertyChanged("XLX_FilePath");
            }
        }*/

        public static readonly DependencyProperty StonesPathProperty = DependencyProperty.Register(
            "StonesPath", typeof(string), typeof(ViewModel));
        public string StonesPath
        {
            get { return (string)GetValue(StonesPathProperty); }
            set
            {
                SetValue(StonesPathProperty, value);
                OnPropertyChanged("StonesPath");
            }
        }

        /*private string stonesPath;
        public string StonesPath
        {
            get { return stonesPath; }
            set
            {
                stonesPath = value;
                OnPropertyChanged("StonesPath");
            }
        }*/

        public static readonly DependencyProperty PackageProperty = DependencyProperty.Register(
            "Package", typeof(string), typeof(ViewModel));
        public string Package
        {
            get { return (string)GetValue(PackageProperty); }
            set
            {
                SetValue(PackageProperty, value);
                OnPropertyChanged("Package");
            }
        }

        /*
         private string package = "Alrosa-USO";
        public string Package
        {
            get { return package; }
            set
            {
                package = value;
                OnPropertyChanged("Package");
            }
        }
         */
        #endregion

        private string stonesFoundStatus;
        public string StonesFoundStatus
        {
            get { return stonesFoundStatus; }
            set
            {
                stonesFoundStatus = value;
                OnPropertyChanged("StonesFoundStatus");
                OnPropertyChanged("Settings");
            }
        }

        private ObservableCollection<ControlLibrary.StoneControl> stoneControls;
        public ObservableCollection<ControlLibrary.StoneControl> StoneControls
        {
            get { return stoneControls; }
            set
            {
                stoneControls = value;
                OnPropertyChanged("StoneControls");
                OnPropertyChanged("Settings");
            }
        }

        private static string settingPath = $"{AppDomain.CurrentDomain.BaseDirectory}\\Settings.json";//.txt";//.json";
        private FileOperations.Settings settings;
        public FileOperations.Settings Settings
        {
            get { return settings; }
            set
            {
                settings = value;
                OnPropertyChanged("Settings");
            }
        }

        private FileOperations.ExcelOperations xObj;

        private bool loadEnabled;
        public bool LoadEnabled
        {
            get { return loadEnabled; }
            set
            {
                loadEnabled = value;
                OnPropertyChanged("LoadEnabled");
            }
        }

        private bool saveEnabled;
        public bool SaveEnabled
        {
            get { return saveEnabled; }
            set
            {
                saveEnabled = value;
                OnPropertyChanged("SaveEnabled");
            }
        }
        private ObservableCollection<ComboBoxItem> conflictModes;
        public ObservableCollection<ComboBoxItem> ConflictModes
        {
            get { return conflictModes; }
            set
            {
                conflictModes = value;
                OnPropertyChanged("ConflictModes");
            }
        }

        /*public ObservableCollection<ComboBoxItem> ConflictMode
        {
            get { return conflictModes; }
            set
            {
                conflictModes = value;
                OnPropertyChanged("ConflictModes");
            }
        }*/

        public FileOperations.ExcelOperations.DateConflictMode ConflictMode;

        public ViewModel()
        {
            stoneControls = new ObservableCollection<ControlLibrary.StoneControl>();
            Settings = new FileOperations.Settings();
            SaveEnabled = false;
            LoadEnabled = true;
            //BindingSettingsValues();
            if (System.IO.File.Exists(settingPath))
            {
                Settings = FileOperations.Settings.LoadSettings(settingPath);
                //GetValuesFromSettings();
            }
            BindingSettingsValues();
            xObj = new FileOperations.ExcelOperations();
            FillConflictModes();
            //SW_Props = new ScrollViewer();
        }

        private void FillConflictModes()
        {
            if (ConflictModes == null)
            {
                ConflictModes = new ObservableCollection<ComboBoxItem>();
            }

            AddNewComboBoxItem("From Excel");
            AddNewComboBoxItem("From File");
            AddNewComboBoxItem("Earliest");
        }

        public void SelectDefaultConflict()
        {
            try
            {
                //int cIndex = ConflictModes.IndexOf(ConflictModes.Single(x => x.Content.Equals("Earliest")));
                //cIndex = 2;
                OnSelectConflict?.Invoke(2);
            }
            catch (Exception e)
            {

            }
        }

        private void AddNewComboBoxItem(string conent)
        {
            ComboBoxItem cbI = new ComboBoxItem();
            cbI.Content = conent;
            cbI.VerticalContentAlignment = VerticalAlignment.Center;
            cbI.HorizontalContentAlignment = HorizontalAlignment.Left;
            cbI.Selected += ConflictMode_Changed;
            ConflictModes.Add(cbI);
        }

        private void ConflictMode_Changed(object sender, RoutedEventArgs e)
        {
            string senderName = null;

            PropertyInfo pi = sender.GetType().GetProperty("Content");
            if (pi != null)
            {
                try
                {
                    senderName = (string)pi.GetValue(sender);
                    if (!string.IsNullOrWhiteSpace(senderName))
                    {
                        switch (senderName)
                        {
                            case "From Excel":
                                this.ConflictMode = FileOperations.ExcelOperations.DateConflictMode.FromExcel;
                                break;
                            case "From File":
                                this.ConflictMode = FileOperations.ExcelOperations.DateConflictMode.FromFile;
                                break;
                            case "Earliest":
                                this.ConflictMode = FileOperations.ExcelOperations.DateConflictMode.Earliest;
                                break;
                        }
                    }
                }
                catch (Exception ex)
                {

                }
                
            }
        }

        private void BindingSettingsValues()
        {
            if (Settings != null)
            {
                SetTwoWayBinding(Settings, nameof(Settings.StartExcelPath), this, ExcelFilePathProperty);
                SetTwoWayBinding(Settings, nameof(Settings.StartStonesDir), this, StonesPathProperty);
                SetTwoWayBinding(Settings, nameof(Settings.Package), this, PackageProperty);
            }
        }

        public void OpenExcelFile()
        {
            CommonOpenFileDialog COFD = new CommonOpenFileDialog();
            COFD.IsFolderPicker = false;
            COFD.Title = "Select excel file";
            if (Settings != null)
            {
                if (!string.IsNullOrWhiteSpace(Settings.StartExcelPath))
                {
                    try
                    {
                        COFD.DefaultDirectory = System.IO.Path.GetDirectoryName(Settings.StartExcelPath);
                        COFD.DefaultFileName = Settings.StartExcelPath;
                    }
                    catch
                    {

                    }
                }
            }
            var DR = COFD.ShowDialog();

            if (DR == CommonFileDialogResult.Ok)
            {
                ExcelFilePath = COFD.FileName;
                //settings.StartExcelPath = XLX_FilePath;
                Settings.SaveSettings(settingPath);
                //XLX_FilePath = "C:\\Users\\Ko12A\\Downloads\\109E_AlrosaId_G.xlsx";
            }
        }

        public void OpenStonesFolderFile()
        {
            CommonOpenFileDialog COFD = new CommonOpenFileDialog();
            COFD.IsFolderPicker = true;
            COFD.Title = "Select stones folder";
            if (Settings != null)
            {
                if (!string.IsNullOrWhiteSpace(Settings.StartStonesDir))
                {
                    COFD.DefaultDirectory = Settings.StartStonesDir;
                }
            }
            var DR = COFD.ShowDialog();
            
            //bool DR = true;
            if (DR == CommonFileDialogResult.Ok)
            {
                string fileName = COFD.FileName;
                //fileName = "C:\\Users\\Ko12A\\Downloads\\";

                StonesPath = fileName;
                Settings.StartStonesDir= StonesPath;
                Settings.SaveSettings(settingPath);
            }
        }

        public void GetStonesFromExcel()
        {
            //Validation needed
            try
            {
                ValidateInputData();
            }
            catch (Exception e)
            {
                return;
            }
            Task.Run(() => a_GetStonesFromExcel());
        }

        public async void a_GetStonesFromExcel()
        {
            string stonesPath = string.Empty;
            string excelPath = string.Empty;
            string pkg = string.Empty;
            FileOperations.ExcelOperations exObj = null;
            
            Application.Current.Dispatcher.Invoke(new Action(() =>
            {
                stonesPath = StonesPath;
                excelPath = ExcelFilePath;
                exObj = xObj;
                pkg = Package;
                SaveEnabled = false;
                LoadEnabled = false;
                StonesFoundStatus = "Loading...";
                spStatus.Children.RemoveRange(1, spStatus.Children.Count - 1);
            }
            ));
            

            /*Task<ObservableCollection<StoneInfo>> task = new Task<ObservableCollection<StoneInfo>>(
                () => BusinessLogic.ModelLogic.GetStonesFromExcel(xObj, XLX_FilePath, Prefix, boxIdRegexString)
                );

            task.Start();
            await task;
            
            ocSI = task.Result;*/
            ObservableCollection<StoneInfo> oc = StoneInfo.GetStonesFromExcel(exObj, excelPath, pkg, FileOperations.StoneOperations.BoxIdRegexString);

            int curFile = 0;
            BusinessLogic.StoneSearcher searcher = new BusinessLogic.StoneSearcher();

            Viewbox vb1 = null;
            Viewbox vb2 = null;
            Application.Current.Dispatcher.Invoke(new Action(() =>
            {
                ocSI = oc;
                CreateStoneControls();
                StonesFoundStatus = $"Found {ocSI.Count} stones";

                ProgressBar progressBar = new ProgressBar();
                progressBar.MinWidth = 30;
                TextBlock tb = new TextBlock();
                tb.Text = "Searching files...";
                SolidColorBrush brush = new SolidColorBrush(Color.FromRgb(128, 128, 128));
                tb.Foreground = brush;
                //tb.FontSize = 10;

                vb1 = new Viewbox();
                vb1.Child = tb;
                vb1.Margin = new Thickness(10, 0, 0, 0);
                vb2 = new Viewbox();
                vb2.Child = progressBar;
                vb2.Margin = new Thickness(5, 0, 0, 0);

                progressBar.Maximum = ocSI.Count;
                progressBar.Value = curFile;

                SetOneWayBinding(searcher, nameof(searcher.CurrentFile), progressBar, ProgressBar.ValueProperty);
            }
            ));
            

           

            /*spStatus.Children.Add(vb1);
            spStatus.Children.Add(vb2);
            */
            

            


            /*await*/
            Task t = new Task(new Action(() =>
            {
                
                Application.Current.Dispatcher.Invoke(new Action(() =>
                {
                    spStatus.Children.Add(vb1);
                    spStatus.Children.Add(vb2);
                }
                ));

                searcher.FindFiles(ocSI, stonesPath);

                Application.Current.Dispatcher.Invoke(new Action(() =>
                {
                    spStatus.Children.Remove(vb1);
                    spStatus.Children.Remove(vb2);
                }
                ));
            }));

            t.Start();
            await t;

            Application.Current.Dispatcher.Invoke(new Action(() =>
            {
                Settings.SaveSettings(settingPath);
                SaveEnabled = true;
                LoadEnabled = true;
            }
            ));
            
        }

        private void ValidateInputData()
        {
            //Check Excel path
            if (string.IsNullOrWhiteSpace(ExcelFilePath))
            {
                throw new Exception("Excel file path is empty");
            }
            else
            {
                if (!FileOperations.FileOperations.CheckFileExists(ExcelFilePath))
                {
                    throw new Exception("Excel file path does not exists");
                }
            }

            //Check stones path
            if (string.IsNullOrWhiteSpace(StonesPath))
            {
                throw new Exception("Stones path is empty");
            }
            else
            {
                if (!FileOperations.FileOperations.IsDirectory(StonesPath))
                {
                    throw new Exception("Stones path is not directory path");
                }
                if (!FileOperations.FileOperations.CheckDirectoryExists(StonesPath))
                {
                    throw new Exception("Stones path does not exists");
                }
            }

            if (string.IsNullOrWhiteSpace(Package))
            {
                throw new Exception("Package is empty");
            }
        }

        public void SaveAllStonesScanDateToExcel()
        {
            try
            {
                ValidateInputData();
                ValidateInputData();
                if (ocSI == null)
                {
                    throw new Exception("Stones collection is empty");
                }
                else
                {
                    if (ocSI.Count < 1)
                    {
                        throw new Exception("Stones collection is empty");
                    }
                }
            }
            catch (Exception e)
            {
                return;
            }

            SaveEnabled = false;
            DataTable dt = CreateScanDateDataTable(ocSI);

            xObj.SaveScanDataDataTable(ExcelFilePath, dt, ConflictMode);
            MessageBox.Show("Successfully saved");
            SaveEnabled = true;
        }

        private DataTable CreateScanDateDataTable(ObservableCollection<StoneInfo> collection)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Scan Date"));
            int columnNum = dt.Columns.IndexOf("Scan Date");
            foreach (StoneInfo si in collection)
            {
                DataRow dr = dt.NewRow();
                dr[columnNum] = si.ScanDate;
                dt.Rows.Add(dr);
            }

            return dt;
        }

        private void CreateStoneControls()
        {
            if (StoneControls == null)
            {
                StoneControls = new ObservableCollection<ControlLibrary.StoneControl>();
            }
            StoneControls.Clear();

            foreach (StoneInfo si in ocSI)
            {
                if (si != null)
                {
                    if (stoneControls == null)
                    {
                        stoneControls = new ObservableCollection<ControlLibrary.StoneControl>();
                    }

                    ControlLibrary.StoneControl sc = CreateStoneControl(si);
                    StoneControls.Add(sc);
                }
            }
        }

        private ControlLibrary.StoneControl CreateStoneControl(StoneInfo si)
        {
            ControlLibrary.StoneControl sc = new ControlLibrary.StoneControl();
            sc.DataContext = si;
            SetOneWayBinding(si, nameof(si.FileFound), sc, ControlLibrary.StoneControl.FileFoundProperty);
            sc.RefreshFound();
            sc.Padding = new Thickness(1);
            sc.Selected = false;
            sc.OnSetSelected += SelectStone;

            return sc;
        }

        #region Binding
        private void SetOneWayBinding(object source, string sourcePropertyName, DependencyObject targetObject, DependencyProperty targetProperty)
        {
            SetBinding(source, sourcePropertyName, targetObject, targetProperty, BindingMode.OneWay);
        }

        private void SetTwoWayBinding(object source, string sourcePropertyName, DependencyObject targetObject, DependencyProperty targetProperty)
        {
            SetBinding(source, sourcePropertyName, targetObject, targetProperty, BindingMode.TwoWay);
        }

        private void SetBinding(object source, string sourcePropertyName, DependencyObject targetObject, DependencyProperty targetProperty, BindingMode mode)
        {
            Binding b = new Binding();
            b.Source = source;
            b.Path = new PropertyPath(sourcePropertyName);
            b.Mode = mode;
            b.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            BindingOperations.SetBinding(targetObject, targetProperty, b);
        }
        #endregion

        private void SelectStone(ControlLibrary.StoneControl sc)
        {
            //int lastSelIndex = StoneControls.IndexOf(XmlDataProvider => XmlDataProvider.)
            if (StoneControls != null)
            {
                if (StoneControls.Count > 0)
                {
                    ControlLibrary.StoneControl prevSelected = StoneControls.Where(x => x.Selected == true).FirstOrDefault();

                    if (prevSelected != null)
                    {
                        prevSelected.Selected = false;
                        
                    }

                    sc.Selected = true;

                    ControlLibrary.StonePropertyControl spC = CreateStonePropertyControl((StoneInfo)sc.DataContext);
                    if (spC != null)
                    {
                        RemovePropControl();
                        AddPropControl(spC);
                    }
                }
            }
        }

        private void AddPropControl(ControlLibrary.StonePropertyControl spC)
        {
            SW_Props.Content = spC;
        }

        private void RemovePropControl()
        {
            SW_Props.Content = null;
        }

        private ControlLibrary.StonePropertyControl CreateStonePropertyControl(StoneInfo si)
        {
            if (si == null)
            {
                return null;
            }
            ControlLibrary.StonePropertyControl spC = new ControlLibrary.StonePropertyControl();
            spC.DataContext = si;
            SetOneWayBinding(si, nameof(si.FileFound), spC, ControlLibrary.StonePropertyControl.FileFoundProperty);
            spC.RefreshFound();
            return spC;
        }

        internal void OnPropertyChanged(String property)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(property));
            }
        }
    }

    public class RelayCommand : ICommand
    {
        private Action<object> execute;
        private Func<object, bool> canExecute;

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public RelayCommand(Action<object> execute, Func<object, bool> canExecute = null)
        {
            this.execute = execute;
            this.canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            return this.canExecute == null || this.canExecute(parameter);
        }

        public void Execute(object parameter)
        {
            this.execute(parameter);
        }
    }

}
