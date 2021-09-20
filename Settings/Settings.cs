using Newtonsoft.Json;
using System;
using System.ComponentModel;

namespace Settings
{
    public class Settings : INotifyPropertyChanged
    {
        // property changed event
        public event PropertyChangedEventHandler PropertyChanged;

        private string startStonesDir;
        public string StartStonesDir
        {
            get { return startStonesDir; }
            set
            {
                startStonesDir = value;
                OnPropertyChanged("StartStonesDir");
            }
        }

        private string startExcelPath;
        public string StartExcelPath
        {
            get { return startExcelPath; }
            set
            {
                startExcelPath = value;
                OnPropertyChanged("StartExcelPath");
            }
        }

        private string package;
        public string Package
        {
            get { return package; }
            set
            {
                package = value;
                OnPropertyChanged("Package");
            }
        }

        #region Excel Settings
        private ExcelDateColumnSettings dateColumnSettings;
        public ExcelDateColumnSettings DateColumnSettings
        {
            get { return dateColumnSettings; }
            set
            {
                dateColumnSettings = value;
                OnPropertyChanged("DateColumnSettings");
            }
        }
        private DateConflictMode dateConflictMode;
        public DateConflictMode DateConflictMode
        {
            get { return dateConflictMode; }
            set
            {
                dateConflictMode = value;
                OnPropertyChanged("DateConflictMode");
            }
        }
        #endregion

        public Settings()
        {
            dateColumnSettings = new ExcelDateColumnSettings();
        }

        public static Settings LoadSettings(string filepath)
        {
            Settings s = null;
            string jsonString = System.IO.File.ReadAllText(filepath);
            s = JsonConvert.DeserializeObject<Settings>(jsonString,
                new JsonSerializerSettings()
                {
                    TypeNameHandling = TypeNameHandling.Auto
                });
            return s;
        }

        public void SaveSettings(string filepath)
        {
            string jstring = JsonConvert.SerializeObject(this, Formatting.Indented,
                new JsonSerializerSettings()
                {
                    TypeNameHandling = TypeNameHandling.Auto
                });

            System.IO.File.WriteAllText(filepath, jstring);
        }

        internal void OnPropertyChanged(String property)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(property));
            }
        }
    }

    public class ExcelDateColumnSettings : INotifyPropertyChanged
    {
        // property changed event
        public event PropertyChangedEventHandler PropertyChanged;

        private string columnName;
        public string ColumnName
        {
            get { return columnName; }
            set
            {
                columnName = value;
                OnPropertyChanged("ColumnName");
            }
        }

        private string nameNumberFormat;
        public string NameNumberFormat
        {
            get { return nameNumberFormat; }
            set
            {
                nameNumberFormat = value;
                OnPropertyChanged("NameNumberFormat");
            }
        }

        private System.Drawing.Color nameBackgroundColor;
        public System.Drawing.Color NameBackgroundColor
        {
            get { return nameBackgroundColor; }
            set
            {
                nameBackgroundColor = value;
                OnPropertyChanged("NameBackgroundColor");
            }
        }

        private System.Drawing.Color nameForegroundColor;
        public System.Drawing.Color NameForegroundColor
        {
            get { return nameForegroundColor; }
            set
            {
                nameForegroundColor = value;
                OnPropertyChanged("NameForegroundColor");
            }
        }

        private string dateNumberFormat;
        public string DateNumberFormat
        {
            get { return dateNumberFormat; }
            set
            {
                dateNumberFormat = value;
                OnPropertyChanged("DateNumberFormat");
            }
        }

        public ExcelDateColumnSettings()
        {
            ColumnName = "Scan Date";
            NameNumberFormat = "General";
            NameBackgroundColor = System.Drawing.Color.Black;
            NameForegroundColor = System.Drawing.Color.White;
            DateNumberFormat = nameof(DateFormats.Short);
        }
        
        internal void OnPropertyChanged(String property)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(property));
            }
        }
    }

    public static class DateFormats
    {
        public static string Short = "m/d/yyyy";
        public static string Long = "[$-F800]dddd, mmmm dd, yyyy";
        public static string Internal = "dd.MM.yyyy HH:mm:ss";
    }

    public enum DateConflictMode
    {
        FromExcel,
        FromFile,
        Earliest
    }
}
