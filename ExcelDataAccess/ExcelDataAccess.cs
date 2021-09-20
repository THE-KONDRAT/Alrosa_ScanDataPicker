using System;
using System.Data;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace DataAccess
{
    public static class ExcelDataAccess
    {
        public static object[,] LoadDataFromExcelSheet(string filePath)
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

        public static void SaveScanDataDataTable(string filePath, DataTable scanDateDT, Settings.ExcelDateColumnSettings columnSettings, Settings.DateConflictMode dateConflictMode)
        {
            string numberFormat = GetDateFormatFromName(columnSettings.DateNumberFormat);
            object[,] dataArr = null;
            Excel.Application xlApp = new Excel.Application(); //создаём приложение Excel
            try
            {
                Excel.Range Rng;
                Excel.Workbook xlWB;
                Excel.Worksheet xlSht;


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
                    SetCellsStyle(xlSht, scanDateColumn, columnSettings);
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
                        DateTime date = FileOperations.DateOperations.GetDateFromString((string)row["Scan Date"], dateConflictMode);
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

        private static void SetCellsStyle(Excel.Worksheet xlSht, int scanDateColumn, Settings.ExcelDateColumnSettings columnSettings)
        {
            //Get ScanDate column
            Excel.Range ColumnRng = ((Excel.Range)xlSht.Cells[1, scanDateColumn]).EntireColumn;

            ColumnRng.NumberFormat = GetDateFormatFromName(columnSettings.DateNumberFormat);
            //Set first cell format "General"
            ((Excel.Range)xlSht.Cells[1, scanDateColumn]).NumberFormat = columnSettings.NameNumberFormat;//"General";
            //Set first cell background
            ((Excel.Range)xlSht.Cells[1, scanDateColumn]).Interior.Color = columnSettings.NameBackgroundColor;
            //Set first cell foreground
            ((Excel.Range)xlSht.Cells[1, scanDateColumn]).Font.Color = columnSettings.NameForegroundColor;
        }

        private static string GetDateFormatFromName(string name)
        {
            string format = null;
            try
            {
                var props = typeof(Settings.DateFormats).GetProperties();
                var formatProperty = props.Where(x => x.Name.Equals("name")).FirstOrDefault();
                if (formatProperty != null)
                {
                    format = (string)formatProperty.GetValue(null);
                }
            }
            catch(Exception e)
            {

            }
            return format;
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
}
