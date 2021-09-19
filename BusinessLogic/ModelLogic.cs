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

                for (int i = 0; i < collection.Count; i++)
                {
                    StoneInfo si = collection[i];
                    StoneInfo.SetScanDateFromFile(si, stonesDI);
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

        internal void OnPropertyChanged(String property)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(property));
            }
        }
    }
}
