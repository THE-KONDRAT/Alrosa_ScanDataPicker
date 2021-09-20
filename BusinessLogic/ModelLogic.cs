using Model;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace BusinessLogic
{
    public class StoneSearcher : INotifyPropertyChanged
    {
        // property changed event
        public event PropertyChangedEventHandler PropertyChanged;

        private int currentFile;
        public int CurrentFile
        {
            get { return currentFile; }
            set
            {
                currentFile = value;
                OnPropertyChanged("CurrentFile");
            }
        }
        public StoneSearcher()
        {
            currentFile = 0;
        }
        public void FindFiles(ObservableCollection<StoneInfo> collection, string scanfolder)
        {
            if (!string.IsNullOrWhiteSpace(scanfolder))
            {


                System.IO.DirectoryInfo stonesDI = new System.IO.DirectoryInfo(scanfolder);
                System.IO.FileInfo[] glxFIs = stonesDI.GetFiles("*.glx", System.IO.SearchOption.AllDirectories);

                for (int i = 0; i < collection.Count; i++)
                {
                    StoneInfo si = collection[i];
                    SetScanDateFromFile(si, glxFIs);
                    //si.FileFound = FindFile(si, stonesDI);
                    //Thread.Sleep(200);

                    CurrentFile += 1;

                    /*Application.Current.Dispatcher.Invoke(new Action(() =>
                    {
                        curFile += 1;
                    }
                    ));*/
                }
            }

        }

        public static void SetScanDateFromFile(StoneInfo si, System.IO.FileInfo[] glxFIs)
        {
            try
            {
                bool success = false;

                DateTime fileScanDate = FileOperations.StoneOperations.GetFileCreationTime(si.GLX_DirectoryName, si.GLX_FileName, glxFIs, ref success);
                if (!success)
                {
                    return;
                }
                //string dt = scanDate.ToString();

                //if StoneInfo already has ScanDate
                if (!string.IsNullOrWhiteSpace(si.ScanDate))
                {
                    try
                    {
                        DateTime existedDate = DateTime.ParseExact(si.ScanDate, Settings.DateFormats.Internal, System.Globalization.CultureInfo.InvariantCulture);

                        string strExDate = FileOperations.DateOperations.ConvertDateTimeToString(existedDate, Settings.DateFormats.Internal);
                        string strScanDate = FileOperations.DateOperations.ConvertDateTimeToString(fileScanDate, Settings.DateFormats.Internal);
                        si.ScanDate = $"{strExDate} ({strScanDate})";

                    }
                    catch (Exception e)
                    {
                        //si.ScanDate = FileOperations.DateFormats.ConvertDateTimeToString(existedDate, FileOperations.DateFormats.Internal);
                    }

                }
                else
                {
                    try
                    {
                        string strScanDate = FileOperations.DateOperations.ConvertDateTimeToString(fileScanDate, Settings.DateFormats.Internal);
                        si.ScanDate = strScanDate;

                    }
                    catch (Exception e)
                    {
                        //si.ScanDate = FileOperations.DateFormats.ConvertDateTimeToString(existedDate, FileOperations.DateFormats.Internal);
                    }
                }

                /*if (si.ScanDate != dt)
                {
                    si.ScanDate += $" ({dt})";
                    var v = string.Format($"{{0:{FileOperations.DateFormats.Internal}}}", inputDate);
                }
                else
                {
                    si.ScanDate = dt;
                }*/
                si.FileFound = true;
            }
            catch (Exception e)
            {

            }
        }

        internal void OnPropertyChanged(String property)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(property));
            }
        }
    }
}
