using DataAccess;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Text.RegularExpressions;

namespace Model
{

    public class StoneInfo : INotifyPropertyChanged
    {
        // property changed event
        public event PropertyChangedEventHandler PropertyChanged;

        #region Stone properties
        #region Excel params
        private string fullBoxId;
        public string FullBoxId
        {
            get { return fullBoxId; }
            set
            {
                fullBoxId = value;
                OnPropertyChanged("FullBoxId");
            }
        }

        private string alrosaID;
        public string AlrosaID
        {
            get { return alrosaID; }
            set
            {
                alrosaID = value;
                OnPropertyChanged("AlrosaID");
            }
        }

        private string briefID;
        public string BriefID
        {
            get { return briefID; }
            set
            {
                briefID = value;
                OnPropertyChanged("BriefID");
            }
        }

        private string stoneID;
        public string StoneID
        {
            get { return stoneID; }
            set
            {
                stoneID = value;
                OnPropertyChanged("StoneID");
            }
        }

        private string stoneWeight;
        public string StoneWeight
        {
            get { return stoneWeight; }
            set
            {
                stoneWeight = value;
                OnPropertyChanged("StoneWeight");
            }
        }

        private string guid;
        public string GUID
        {
            get { return guid; }
            set
            {
                guid = value;
                OnPropertyChanged("GUID");
            }
        }

        private string measuredWeight;
        public string MeasuredWeight
        {
            get { return measuredWeight; }
            set
            {
                measuredWeight = value;
                OnPropertyChanged("MeasuredWeight");
            }
        }

        private string zvi_Result;
        public string ZVI_Result
        {
            get { return zvi_Result; }
            set
            {
                zvi_Result = value;
                OnPropertyChanged("ZVI_Result");
            }
        }

        private string color;
        public string Color
        {
            get { return color; }
            set
            {
                color = value;
                OnPropertyChanged("Color");
            }
        }

        private string blueUV;
        public string BlueUV
        {
            get { return blueUV; }
            set
            {
                blueUV = value;
                OnPropertyChanged("BlueUV");
            }
        }

        private string uv_Num;
        public string UV_Num
        {
            get { return uv_Num; }
            set
            {
                uv_Num = value;
                OnPropertyChanged("UV_Num");
            }
        }

        private string yellowUV;
        public string YellowUV
        {
            get { return yellowUV; }
            set
            {
                yellowUV = value;
                OnPropertyChanged("YellowUV");
            }
        }

        private string fancyYellow;
        public string FancyYellow
        {
            get { return fancyYellow; }
            set
            {
                fancyYellow = value;
                OnPropertyChanged("FancyYellow");
            }
        }

        private string type;
        public string Type
        {
            get { return type; }
            set
            {
                type = value;
                OnPropertyChanged("Type");
            }
        }

        private string brown;
        public string Brown
        {
            get { return brown; }
            set
            {
                brown = value;
                OnPropertyChanged("Brown");
            }
        }

        private string expertColor;
        public string ExpertColor
        {
            get { return expertColor; }
            set
            {
                expertColor = value;
                OnPropertyChanged("ExpertColor");
            }
        }

        private string expertUV;
        public string ExpertUV
        {
            get { return expertUV; }
            set
            {
                expertUV = value;
                OnPropertyChanged("ExpertUV");
            }
        }

        private string expertUV_Color;
        public string ExpertUV_Color
        {
            get { return expertUV_Color; }
            set
            {
                expertUV_Color = value;
                OnPropertyChanged("ExpertUV_Color");
            }
        }

        private string expertFancyYellow;
        public string ExpertFancyYellow
        {
            get { return expertFancyYellow; }
            set
            {
                expertFancyYellow = value;
                OnPropertyChanged("ExpertFancyYellow");
            }
        }

        private string comment;
        public string Comment
        {
            get { return comment; }
            set
            {
                comment = value;
                OnPropertyChanged("Comment");
            }
        }

        private string scanDate;
        public string ScanDate
        {
            get { return scanDate; }
            set
            {
                scanDate = value;
                OnPropertyChanged("ScanDate");
            }
        }
        #endregion

        private bool fileFound;
        public bool FileFound
        {
            get { return fileFound; }
            set
            {
                fileFound = value;
                OnPropertyChanged("FileFound");
            }
        }

        private string glx_DirectoryName;
        public string GLX_DirectoryName
        {
            get { return glx_DirectoryName; }
            set
            {
                glx_DirectoryName = value;
                OnPropertyChanged("GLX_DirectoryName");
            }
        }

        private string glx_FileName;
        public string GLX_FileName
        {
            get { return glx_FileName; }
            set
            {
                glx_FileName = value;
                OnPropertyChanged("GLX_FileName");
            }
        }

        private string packageName;
        public string PackageName
        {
            get { return packageName; }
            set
            {
                packageName = value;
                OnPropertyChanged("PackageName");
            }
        }

        #region search
        private string boxIdRegexString;
        public string BoxIdRegexString
        {
            get { return boxIdRegexString; }
            set
            {
                boxIdRegexString = value;
                OnPropertyChanged("BoxIdRegexString");
            }
        }

        #endregion



        private bool selected;
        public bool Selected
        {
            get { return selected; }
            set
            {
                selected = value;
                OnPropertyChanged("Selected");
            }
        }
        #endregion

        public StoneInfo()
        {

        }

        #region Methods

        public static ObservableCollection<StoneInfo> GetStonesFromExcel(string filePath, string packageName, string boxIdRegexString)
        {
            ObservableCollection<StoneInfo> ocSI = null;
            //var dataArr = xObj.LoadDataFromExcelSheet(filePath); //чтение данных из ячеек в массив            
            var dataArr = ExcelDataAccess.LoadDataFromExcelSheet(filePath);

            if (dataArr != null)
            {
                //заполняем DataTable для последующего заполнения dataGridView
                DataTable dt = new DataTable();
                DataRow dtRow;

                //добавляем столбцы в DataTable
                for (int i = 1; i <= dataArr.GetUpperBound(1); i++)
                    dt.Columns.Add((string)dataArr[1, i]);
                int scanDateIndex = dt.Columns.IndexOf("Scan Date");
                bool sdCol = dt.Columns.Contains("Scan Date");
                if (!sdCol)
                {
                    dt.Columns.Add("Scan Date");
                }

                //цикл по строкам массива
                for (int i = 2; i <= dataArr.GetUpperBound(0); i++)
                {
                    dtRow = dt.NewRow();
                    //цикл по столбцам массива
                    for (int n = 1; n <= dataArr.GetUpperBound(1); n++)
                    {
                        if (n == scanDateIndex + 1)
                        {
                            object obj = dataArr[i, n];
                            if (obj != null)
                            {
                                Type t = obj.GetType();
                                DateTime inputDate;
                                try
                                {
                                    if (t == typeof(DateTime))
                                    {
                                        inputDate = (DateTime)obj;
                                    }
                                    else
                                    {
                                        string s = (string)obj;
                                        double d = double.Parse((string)obj);

                                        inputDate = DateTime.FromOADate(d);
                                    }
                                    //var v = string.Format($"{{0:{FileOperations.DateFormats.Internal}}}", inputDate);
                                    dtRow[n - 1] = string.Format($"{{0:{Settings.DateFormats.Internal}}}", inputDate);
                                }
                                catch
                                {

                                }
                            }
                            
                        }
                        else
                        {
                            dtRow[n - 1] = dataArr[i, n];
                        }
                    }
                    dt.Rows.Add(dtRow);



                    StoneInfo si = CreateStoneInfo(dtRow, packageName, boxIdRegexString);

                    if (si != null)
                    {
                        if (ocSI == null)
                        {
                            ocSI = new ObservableCollection<StoneInfo>();
                        }
                        ocSI.Add(si);
                    }

                }
            }

            return ocSI;
            //find stones. invoke?
        }

        public static StoneInfo CreateStoneInfo(DataRow dtRow, string packageName, string boxIdRegexString)
        {
            StoneInfo si = new StoneInfo()
            {
                PackageName = packageName,
                FullBoxId = dtRow["Box ID"].ToString(),
                AlrosaID = dtRow["Alrosa ID"].ToString(),
                BriefID = dtRow["Brief ID"].ToString(),
                StoneID = dtRow["Stone ID"].ToString(),
                StoneWeight = dtRow["Stone weight"].ToString(),
                GUID = dtRow["GUID"].ToString(),
                MeasuredWeight = dtRow["Measured weight"].ToString(),
                ZVI_Result = dtRow["ZVI Result"].ToString(),
                Color = dtRow["Color"].ToString(),
                BlueUV = dtRow["Blue UV"].ToString(),
                UV_Num = dtRow["UV Num"].ToString(),
                YellowUV = dtRow["Yellow UV"].ToString(),
                FancyYellow = dtRow["Fancy Yellow"].ToString(),
                Type = dtRow["Type"].ToString(),
                Brown = dtRow["Brown"].ToString(),
                ExpertColor = dtRow["Expert Color"].ToString(),
                ExpertUV = dtRow["Expert UV"].ToString(),
                ExpertUV_Color = dtRow["Expert UV Color"].ToString(),
                ExpertFancyYellow = dtRow["Expert Fancy Yellow "].ToString(),
                Comment = dtRow["Comment"].ToString(),
                ScanDate = dtRow["Scan Date"].ToString(),

                BoxIdRegexString = boxIdRegexString
            };
            var a = dtRow["Scan Date"].ToString();
            string baseInfo = si.FullBoxId + si.AlrosaID + si.BriefID;
            baseInfo += si.StoneWeight + si.GUID;

            /*if (string.IsNullOrWhiteSpace(si.FullBoxId) && string.IsNullOrWhiteSpace(si.AlrosaID) && string.IsNullOrWhiteSpace(si.BriefID) 
                && string.IsNullOrWhiteSpace(si.StoneWeight) && string.IsNullOrWhiteSpace(si.GUID))*/
            if (string.IsNullOrWhiteSpace(baseInfo))
            {

                return null;
            }

            si.GLX_DirectoryName = CreateGLX_DirectoryName(si);
            si.glx_FileName = CreateGLX_FileName(si);

            //FileInfo fi = new FileInfo()
            si.FileFound = false;
            //si.FileFound = false; //FindFile(si, scanfolder)

            return si;
        }


        public void CreateGLX_FileName()
        {
            this.GLX_FileName = CreateGLX_FileName(this);
        }
        private static string CreateGLX_FileName(StoneInfo si)
        {
            if (si == null)
            {
                return null;
            }
            if (string.IsNullOrWhiteSpace(si.boxIdRegexString))
            {
                return null;
            }

            string fullBoxId = si.FullBoxId;
            Regex regexBoxId = new Regex(si.boxIdRegexString);
            string boxId = fullBoxId;
            Match match = regexBoxId.Match(fullBoxId);
            while (match.Success)
            {
                string sMatch = match.Groups[0].Value;
                boxId = boxId.Replace(sMatch, "");
                match = match.NextMatch();
            }
            if (string.IsNullOrWhiteSpace(si.PackageName))
            {
                return null;
            }

            return $"{si.PackageName}_{boxId}_{si.StoneID}.glx";
        }

        public void CreateGLX_DirectoryName()
        {
            this.GLX_DirectoryName = CreateGLX_DirectoryName(this);
        }
        private static string CreateGLX_DirectoryName(StoneInfo si)
        {
            if (si == null)
            {
                return null;
            }
            if (string.IsNullOrWhiteSpace(si.boxIdRegexString))
            {
                return null;
            }

            string fullBoxId = si.FullBoxId;
            Regex regexBoxId = new Regex(si.BoxIdRegexString);
            string boxId = fullBoxId;
            Match match = regexBoxId.Match(fullBoxId);
            while (match.Success)
            {
                string sMatch = match.Groups[0].Value;
                boxId = boxId.Replace(sMatch, "");
                match = match.NextMatch();
            }

            //Directory {Box ID:-- ()}_{Stone ID}
            //File {Prefix}_{Box ID:-- ()}_{Stone ID}.glx
            return $"{boxId}_{si.StoneID}";
        }
        #endregion

        internal void OnPropertyChanged(String property)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(property));
            }
        }
    }
}
