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
    public static class DateOperations
    {
        public static DateTime GetDateFromString(string strDate, Settings.DateConflictMode dateConflict)
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
                case Settings.DateConflictMode.FromExcel:
                    return DateTime.ParseExact(strExDate/*si.ScanDate*/, Settings.DateFormats.Internal, System.Globalization.CultureInfo.InvariantCulture);

                case Settings.DateConflictMode.FromFile:
                    return DateTime.ParseExact(strNewDate, Settings.DateFormats.Internal, System.Globalization.CultureInfo.InvariantCulture);

                case Settings.DateConflictMode.Earliest:
                    DateTime existedDate = DateTime.ParseExact(strExDate/*si.ScanDate*/, Settings.DateFormats.Internal, System.Globalization.CultureInfo.InvariantCulture);

                    if (string.IsNullOrWhiteSpace(strNewDate))
                    {
                        return existedDate;
                    }

                    DateTime newDate = DateTime.ParseExact(strNewDate, Settings.DateFormats.Internal, System.Globalization.CultureInfo.InvariantCulture);
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

        public static string ConvertDateTimeToString(DateTime dt, string format)
        {
            return string.Format($"{{0:{format}}}", dt);
        }
    }

    public static class StoneOperations
    {
        public static string BoxIdRegexString = @"\W{1}\(.+?\)";

        public static DateTime GetFileCreationTime(string GLX_DirectoryName, string GLX_FileName, System.IO.FileInfo[] glxFIs, ref bool success)
        {
            //bool success = false;
            success = false;
            DateTime dt = new DateTime();

            var fileInfos =
                        from file in glxFIs
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
            }
            #region old variants
            /*
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
                    */
                    
                    /*var fileInfos =
                        from file in glxFIs.GetFiles("", System.IO.SearchOption.AllDirectories)
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
                    }
                    */

                    /*
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
                    */
                    

                    
            /*
                }

                /*try
                {
                    
                }
                catch (Exception e)
                {
                    //MessageBox.Show(e.Message);
                }*/
                /*
            }
            */
            #endregion

            return dt;
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

}
