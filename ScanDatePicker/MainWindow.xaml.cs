using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using VM;

namespace ScanDatePicker
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        VM.ViewModel VM;
        public MainWindow()
        {
            InitializeComponent();
            VM = new VM.ViewModel();
            this.DataContext = VM;
            VM.SW_Props = this.swProps;
            VM.spStatus = this.sb;
            if (VM.ConflictModes == null)
            {
                VM.ConflictModes = new System.Collections.ObjectModel.ObservableCollection<ComboBoxItem>();
            } 
            VM.OnSelectConflict += SelectConflictFromName;
            VM.SelectDefaultConflict();
        }

        private void SelectConflictFromName(int index)
        {
            cbConflict.SelectedIndex = index;
        }

        private void AddPropControl(ControlLibrary.StonePropertyControl spC)
        {
            swProps.Content = spC;
        }

        private void RemovePropControl()
        {
            swProps.Content = null;
        }

        // команда открытия excel из файла
        private RelayCommand openExcelCommand;
        public RelayCommand OpenExcelCommand
        {
            get
            {
                return openExcelCommand ??
                  (openExcelCommand = new RelayCommand(obj =>
                  {
                      if (VM == null)
                      {
                          VM = new ViewModel();
                      }
                      VM.OpenExcelFile();
                  }
                  ));
            }
        }

        // команда открытия папки с камнями
        private RelayCommand openStonesFolderCommand;
        public RelayCommand OpenStonesFolderCommand
        {
            get
            {
                return openStonesFolderCommand ??
                  (openStonesFolderCommand = new RelayCommand(obj =>
                  {
                      if (VM == null)
                      {
                          VM = new ViewModel();
                      }
                      VM.OpenStonesFolderFile();
                  }
                  ));
            }
        }

        // команда открытия excel из файла
        private RelayCommand getStonesFromExcelCommand;
        public RelayCommand GetStonesFromExcelCommand
        {
            get
            {
                return getStonesFromExcelCommand ??
                  (getStonesFromExcelCommand = new RelayCommand(obj =>
                  {
                      if (VM == null)
                      {
                          VM = new ViewModel();
                      }
                      VM.GetStonesFromExcel();
                  }
                  ));
            }
        }

        // команда сохранения даты сканирования в excel из файл
        private RelayCommand saveScanDateToExcelCommand;
        public RelayCommand SaveScanDateToExcelCommand
        {
            get
            {
                return saveScanDateToExcelCommand ??
                  (saveScanDateToExcelCommand = new RelayCommand(obj =>
                  {
                      if (VM == null)
                      {
                          VM = new ViewModel();
                      }
                      VM.SaveAllStonesScanDateToExcel();
                  }
                  ));
            }
        }
    }
}
