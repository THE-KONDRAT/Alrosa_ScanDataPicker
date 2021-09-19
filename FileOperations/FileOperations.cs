using Newtonsoft.Json;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace FileOperations
{
    public class ExcelOperations
    {
        //private Excel.Application xlApp;

        public ExcelOperations()
        {
            //xlApp = new Excel.Application(); //создаём приложение Excel
        }

        public object[,] LoadDataFromExcelSheet(string filePath)
        {
            object[,] dataArr = null;

            try
            {
                Excel.Application xlApp = new Excel.Application(); //создаём приложение Excel
                Excel.Range Rng;
                Excel.Workbook xlWB;
                Excel.Worksheet xlSht;
                
                xlWB = xlApp.Workbooks.Open(filePath); //открываем наш файл           
                                                       //xlSht = xlWB.Worksheets["Лист1"]; //или так xlSht = xlWB.ActiveSheet //активный лист
                xlSht = xlWB.ActiveSheet;

                //var aa = xlSht.Cells;
                //var lr = xlSht.Cells.[xlSht.Rows.Count, "A"];
                //iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row; //последняя заполненная строка в столбце А
                //iLastCol = xlSht.Cells[1, xlSht.Columns.Count].End[Excel.XlDirection.xlToLeft].Column; //последний заполненный столбец в 1-й строке
                //int countRows = xlSht.UsedRange.Rows.Count;
                //int countColumns = xlSht.UsedRange.Columns.Count;
                //object[,] data = xlSht.Range[xlSht.Cells[1, 1], xlSht.Cells[countRows-1, countColumns-1]].Cells.Value2;
                //object[,] data = xlSht.UsedRange.Cells.Value2;
                Rng = xlSht.UsedRange; //пример записи диапазона ячеек в переменную Rng
                //Rng = xlSht.UsedRange;
                //Rng = (Excel.Range)xlSht.Range["A1", xlSht.Cells[iLastRow, iLastCol]]; //пример записи диапазона ячеек в переменную Rng
                //Rng = xlSht.get_Range("A1", "B10"); //пример записи диапазона ячеек в переменную Rng
                //Rng = xlSht.get_Range("A1:B10"); //пример записи диапазона ячеек в переменную Rng
                //Rng = xlSht.UsedRange; //пример записи диапазона ячеек в переменную Rng

                dataArr = (object[,])Rng.Value; //чтение данных из ячеек в массив            
                                                //xlSht.get_Range("K1").get_Resize(dataArr.GetUpperBound(0), dataArr.GetUpperBound(1)).Value = dataArr; //выгрузка массива на лист

                //закрытие Excel
                xlWB.Close(true); //сохраняем и закрываем файл
                xlApp.Quit();
                releaseObject(xlSht);
                releaseObject(xlWB);
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message);
            }
            return dataArr;
        }

        public void SaveScanDataDataTable(string filePath, DataTable scanDateDT, DateConflictMode dateConflict)
        {
            string numberFormat  = DateFormats.Long;
            object[,] dataArr = null;
            Excel.Application xlApp = new Excel.Application(); //создаём приложение Excel
            try
            {
                Excel.Range Rng;
                Excel.Workbook xlWB;
                Excel.Worksheet xlSht;
                int iLastRow, iLastCol;


                xlWB = xlApp.Workbooks.Open(filePath); //открываем наш файл           
                                                       //xlSht = xlWB.Worksheets["Лист1"]; //или так xlSht = xlWB.ActiveSheet //активный лист
                xlSht = xlWB.ActiveSheet;

                Rng = xlSht.UsedRange; //пример записи диапазона ячеек в переменную Rng
                int columnsCount = 0;
                dataArr = (object[,])Rng.Value;
                //добавляем столбцы в DataTable, excel считает с 1
                for (int i = 1; i <= dataArr.GetUpperBound(1); i++)
                {
                    if (!string.IsNullOrWhiteSpace((string)dataArr[1, i]))
                    {
                        columnsCount++;
                    }
                }

                //rows and cols
                //Rng = (Excel.Range)xlSht.Range[(Excel.Range)xlSht.Cells[1, 1], (Excel.Range)xlSht.Cells[scanDateDT.Rows.Count, columnsCount + 1]];
                //Rng = (Excel.Range)xlSht.Cells[scanDateDT.Rows.Count, columnsCount + 1]; //пример записи диапазона ячеек в переменную Rng

                DataTable dt = new DataTable();
                dataArr = (object[,])Rng.Value; //чтение данных из ячеек в массив  
                //добавляем столбцы в DataTable
                for (int i = 1; i <= dataArr.GetUpperBound(1); i++)
                    dt.Columns.Add((string)dataArr[1, i]);
                //bool sdCol = dt.Columns.Contains("Scan Date");
                int scanDateColumn = dt.Columns.IndexOf("Scan Date");
                if (scanDateColumn < 0)
                {
                    //scanDateColumn = dt.Columns.Count + 1;
                    //xlApp.Cells[1, scanDateColumn] = "Scan Date";
                    scanDateColumn = columnsCount + 1;
                    xlApp.Cells[1, scanDateColumn] = "Scan Date";
                    SetCellsStyle(xlSht, scanDateColumn, numberFormat);
                }
                else
                {
                    scanDateColumn += 1; //Excel считает с 1
                }

                //var a = ((Excel.Range)xlApp.Cells[1, scanDateColumn]).NumberFormat;

                //throw new Exception("Debug");

                //Insert data
                for (int i = 0; i < scanDateDT.Rows.Count; i++)
                {
                    DataRow row = scanDateDT.Rows[i];
                    try
                    {
                        DateTime date = GetDateFromString((string)row["Scan Date"], dateConflict);
                        xlApp.Cells[i + 2, scanDateColumn] = date;
                        Excel.Range cell = (Excel.Range)xlSht.Cells[i + 2, scanDateColumn];
                        cell.NumberFormat = numberFormat;
                    }
                    catch (Exception e)
                    {
                        Excel.Range cell = (Excel.Range)xlSht.Cells[i + 2, scanDateColumn];
                        cell.NumberFormat = numberFormat;
                    }
                    
                }


                /*try
                {
                    //получение диапазона ячеек размером с число камней
                    Excel.Range ColumnRng = (Excel.Range)xlSht.Range[(Excel.Range)xlSht.Cells[2, scanDateColumn], (Excel.Range)xlSht.Cells[scanDateDT.Rows.Count + 1, scanDateColumn]]; 
                    
                }
                catch (Exception e)
                {

                }*/
                /*var c = xlSht.Cells[dt.Columns.Count, 1];
                xlSht.Cells[dt.Columns.Count, 1] = "Scan Date";
                var d = xlSht.Cells[dt.Columns.Count, 1];

                dataArr = (object[,])xlSht.UsedRange.Value; //чтение данных из ячеек в массив  
                DataTable dt2 = new DataTable();                                //добавляем столбцы в DataTable
                for (int i = 1; i <= dataArr.GetUpperBound(1); i++)
                    dt2.Columns.Add((string)dataArr[1, i]);
                bool sdCol2 = dt2.Columns.Contains("Scan Date");

                dataArr = (object[,])Rng.Value; //чтение данных из ячеек в массив            
                                                //xlSht.get_Range("K1").get_Resize(dataArr.GetUpperBound(0), dataArr.GetUpperBound(1)).Value = dataArr; //выгрузка массива на лист
                */
                //закрытие Excel
                xlWB.Close(true); //сохраняем и закрываем файл
                xlApp.Quit();
                releaseObject(xlSht);
                releaseObject(xlWB);
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message);
                xlApp.Quit();
            }
        }

        private static void SetCellsStyle(Excel.Worksheet xlSht, int scanDateColumn, string numberFormat)
        {
            //Get ScanDate column
            Excel.Range ColumnRng = ((Excel.Range)xlSht.Cells[1, scanDateColumn]).EntireColumn;
            /*Set cells format "data/time"
             * 
             * Long Date: "[$-F800]dddd, mmmm dd, yyyy"
             * Short Date: "m/d/yyyy"
             */
            //ColumnRng.NumberFormat = "DD.MM.YYYY";
            ColumnRng.NumberFormat = numberFormat;
            //ColumnRng.NumberFormat = "m/d/yyyy";
            //Set first cell format "General"
            ((Excel.Range)xlSht.Cells[1, scanDateColumn]).NumberFormat = "General";
            //Set first cell background
            ((Excel.Range)xlSht.Cells[1, scanDateColumn]).Interior.Color = Excel.XlRgbColor.rgbBlack;
            //Set first cell foreground
            ((Excel.Range)xlSht.Cells[1, scanDateColumn]).Font.Color = Excel.XlRgbColor.rgbWhite;
        }

        private static DateTime GetDateFromString(string strDate, DateConflictMode dateConflict)
        {
            Regex stRegex = new Regex(@"\((.+)\)");
            string strExDate = strDate;
            string strNewDate = null;
            Match match = stRegex.Match(strDate);
            while (match.Success)
            {
                string sMatch = match.Groups[0].Value;
                strExDate = strExDate.Replace(sMatch, "").Trim();
                strNewDate = match.Groups[1].Value;
                match = match.NextMatch();
            }

            switch (dateConflict)
            {
                case DateConflictMode.FromExcel:
                    return DateTime.ParseExact(strExDate/*si.ScanDate*/, DateFormats.Internal, System.Globalization.CultureInfo.InvariantCulture);

                case DateConflictMode.FromFile:
                    return DateTime.ParseExact(strNewDate, DateFormats.Internal, System.Globalization.CultureInfo.InvariantCulture);

                case DateConflictMode.Earliest:
                    DateTime existedDate = DateTime.ParseExact(strExDate/*si.ScanDate*/, DateFormats.Internal, System.Globalization.CultureInfo.InvariantCulture);

                    if (string.IsNullOrWhiteSpace(strNewDate))
                    {
                        return existedDate;
                    }

                    DateTime newDate = DateTime.ParseExact(strNewDate, DateFormats.Internal, System.Globalization.CultureInfo.InvariantCulture);
                    if (newDate < existedDate)
                    {
                        return newDate;
                    }
                    else
                    {
                        return existedDate;
                    }
                default:
                    throw new Exception("Wrong date conflict");
                    //return new DateTime();
            }  
        }

        public enum DateConflictMode
        {
            FromExcel,
            FromFile,
            Earliest
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                //MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }

    public static class StoneOperations
    {
        public static string BoxIdRegexString = @"\W{1}\(.+?\)";

        public static DateTime GetFileCreationTime(string GLX_DirectoryName, string GLX_FileName, System.IO.DirectoryInfo stonesDI, ref bool success)
        {
            //bool success = false;
            success = false;
            DateTime dt = new DateTime();
            if (stonesDI != null)
            {
                //System.IO.FileInfo fi = new System.IO.FileInfo(scanfolder);
                //System.IO.DirectoryInfo stonesDI = new System.IO.DirectoryInfo(scanfolder);
                if (FileOperations.IsDirectory(stonesDI.Attributes))
                {
                    /*var a = System.IO.Directory.GetFiles(scanfolder, "*.glx", System.IO.SearchOption.AllDirectories);
                    System.IO.DirectoryInfo stonesDI = new System.IO.DirectoryInfo(scanfolder);
                    System.IO.FileInfo[] fileInfos = stonesDI.GetFiles("*.glx", System.IO.SearchOption.AllDirectories);*/

                    /*
                    //2 var
                    System.IO.DirectoryInfo di = stonesDI.GetDirectories().Single(x => x.Name.Equals(si.GLX_DirectoryName));
                    System.IO.FileInfo ffi = di.GetFiles().Single(x => x.Name.Equals(si.GLX_FileName));
                    */

                    //3 var
                    /*string searchPattern = GLX_FileName;
                    System.IO.FileInfo[] fileInfos = stonesDI.GetFiles("*.glx"/*searchPattern*//*, System.IO.SearchOption.AllDirectories);
                    var result = fileInfos.Where(x => x.Name.Equals(GLX_FileName));*/

                    /*searchPattern*/
                    /*var fileInfos =
                        from file in stonesDI.GetFiles("*.glx", System.IO.SearchOption.AllDirectories)
                        where file.Name.Equals(GLX_FileName)
                        select file.CreationTime;

                    if (fileInfos != null)
                    {
                        //System.IO.FileInfo fi = null;
                        if (fileInfos.Count() == 1)
                        {
                            dt = fileInfos.First();
                            success = true;
                        }
                        else if (fileInfos.Count() > 1)
                        {
                            //bool c = fileInfos.Contains(x => x.Name.Equals(GLX_FileName));
                            dt = fileInfos.First();
                            success = true;
                            //fi = fileInfos.Single(x => x.Name.Equals(GLX_FileName));
                        }
                    }*/

                    
                    var dirinfos =
                        from directory in stonesDI.GetDirectories("", System.IO.SearchOption.AllDirectories)
                        where directory.Name.Equals(GLX_DirectoryName)
                        select directory;

                    var dinfo = dirinfos.FirstOrDefault();
                    if (dinfo != null)
                    {
                        var fileInfos2 =
                        from file in dinfo.GetFiles("*.glx", System.IO.SearchOption.AllDirectories)
                        where file.Name.Equals(GLX_FileName)
                        select file.CreationTime;

                        if (fileInfos2 != null)
                        {
                            //System.IO.FileInfo fi = null;
                            if (fileInfos2.Count() == 1)
                            {
                                dt = fileInfos2.First();
                                success = true;
                            }
                            else if (fileInfos2.Count() > 1)
                            {
                                //bool c = fileInfos.Contains(x => x.Name.Equals(GLX_FileName));
                                dt = fileInfos2.First();
                                success = true;
                                //fi = fileInfos.Single(x => x.Name.Equals(GLX_FileName));
                            }
                        }
                    }
                    
                    

                    

                }

                /*try
                {
                    
                }
                catch (Exception e)
                {
                    //MessageBox.Show(e.Message);
                }*/

            }

            return dt;
        }
    }

    public static class DateFormats
    {
        public static string Short = "m/d/yyyy";
        public static string Long = "[$-F800]dddd, mmmm dd, yyyy";
        public static string Internal = "dd.MM.yyyy HH:mm:ss";

        public static string ConvertDateTimeToString(DateTime dt, string format)
        {
            return string.Format($"{{0:{format}}}", dt);
        }
    }

    public static class FileOperations
    {
        public static bool CheckFileExists(string path)
        {
            return File.Exists(path);
        }

        public static bool CheckDirectoryExists(string path)
        {
            return Directory.Exists(path);
        }

        public static bool IsDirectory(string path)
        {
            DirectoryInfo di = new DirectoryInfo(path);
            if (di.Attributes.HasFlag(System.IO.FileAttributes.Directory))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool IsDirectory(FileAttributes attr)
        {
            if (attr.HasFlag(System.IO.FileAttributes.Directory))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }

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

        public Settings()
        {

        }

        /*public void LoadSettings(string filepath)
        {
            Settings s = null;
            string jsonString = System.IO.File.ReadAllText(filepath);
            s = JsonConvert.DeserializeObject<Settings>(jsonString,
                new JsonSerializerSettings()
                {
                    TypeNameHandling = TypeNameHandling.Auto
                });
            this.StartExcelPath = s.StartExcelPath;
            this.StartStonesDir = s.StartStonesDir;
        }*/

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
}
