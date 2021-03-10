using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Pri.LongPath;

namespace DeliveryApplication.Class
{
    class clsDelivery : clsVariable
    {
        public static string RemoteSFtpPart = string.Empty;

        public static void Delivery(ProjectSet projectSet)
        {
            Console.Clear();

            #region Check new files
            switch (projectSet)
            {
                case ProjectSet.Caselaw:
                    projectPath = CaseLaw;
                    clsConsole.WriteLine("\r\n Checking Caselaw ...", ConsoleColor.Red);
                    listFileInfosCaselaw = new DirectoryInfo(Path.Combine(CaseLaw, "Next"))
                        .GetFiles("*.zip").Where(p => Regex.IsMatch(p.Name, @"([0-9A-Z\-]{20}00000-00|[0-9A-Z\-]{20}00000-00_Reports|PR[0-9]+|PR[0-9]+_Reports).zip", RegexOptions.IgnoreCase)).ToList();
                    break;
                case ProjectSet.NonVirgo:
                    projectPath = NonVirgo;
                    clsConsole.WriteLine("\r\n Checking Non-Virgo ...", ConsoleColor.Red);
                    listFileInfosNonVirgo = new DirectoryInfo(Path.Combine(NonVirgo, "Next"))
                        .GetFiles("*.zip").Where(p => Regex.IsMatch(p.Name, @"([0-9]+_TAT_PR[0-9A-Z]+|[0-9]+_TAT_PR[0-9A-Z]+_Reports).zip", RegexOptions.IgnoreCase)).ToList();
                    listDirectoryInfosNonVirgo = new DirectoryInfo(Path.Combine(NonVirgo, "Next"))
                        .GetDirectories().Where(p => Regex.IsMatch(p.Name, "D[A-Z0-9]{4}B[0-9]{6}$")).ToList();
                    break;
                case ProjectSet.Virgo:
                    projectPath = Virgo;
                    clsConsole.WriteLine("\r\n Checking Virgo ...", ConsoleColor.Red);
                    listFileInfosVirgo = new DirectoryInfo(Path.Combine(Virgo, "Next"))
                        .GetFiles("*.zip").Where(p => Regex.IsMatch(p.Name, @"([0-9A-Z\-]{20}00000-00|[0-9A-Z\-]{20}00000-00_Reports|PR[0-9]+|PR[0-9]+_Reports).zip", RegexOptions.IgnoreCase)).ToList();
                    break;
                case ProjectSet.CaseRelated:
                    projectPath = CaseRelated;
                    clsConsole.WriteLine("\r\n Checking Case Related ...", ConsoleColor.Red);
                    listDirectoryInfosCaseRelated = new DirectoryInfo(Path.Combine(CaseRelated, "Next"))
                        .GetDirectories().Where(p => Regex.IsMatch(p.Name, "D[A-Z0-9]{4}B[0-9]{6}$")).ToList();
                    break;
                case ProjectSet.StateNet:
                    projectPath = StateNet;
                    clsConsole.WriteLine("\r\n Checking State Net ...", ConsoleColor.Red);
                    listFileInfosStateNet = new DirectoryInfo(Path.Combine(StateNet, "Next"))
                        .GetFiles("*.zip").Where(p => Regex.IsMatch(p.Name, "(regulationtext|rdoctext)", RegexOptions.IgnoreCase)).ToList();
                    listDirectoryInfosStateNet = new DirectoryInfo(Path.Combine(StateNet, "Next"))
                        .GetDirectories().Where(p => Regex.IsMatch(p.Name, "(regulationtext|rdoctext)", RegexOptions.IgnoreCase)).ToList();
                    break;
                case ProjectSet.SecuritiesMosaic:
                    projectPath = SecuritiesMosaic;
                    clsConsole.WriteLine("\r\n Checking Securities Mosaic ...", ConsoleColor.Red);
                    listDirectoryInfosSecuritiesMosaic = new DirectoryInfo(Path.Combine(SecuritiesMosaic, "Next"))
                        .GetDirectories().Where(p => Regex.IsMatch(p.Name, "[0-9]+_TAT_PR(.*)BOUND$", RegexOptions.IgnoreCase)).ToList();
                    break;
            }

            listPartFolders = new DirectoryInfo(projectPath).GetDirectories().Where(p => Regex.IsMatch(p.Name, "Part [A-Z0-9]+", RegexOptions.RightToLeft)).ToList();

            #endregion

            #region Get next part
            if (listPartFolders.Count > 0)
            {
                var lastCreatedFolder = listPartFolders.OrderByDescending(p => p.CreationTime).Select(p => p.Name).First();
                NextPart = SetNextFolder(lastCreatedFolder);
            }
            else
            {
                string folderDone = Path.Combine(projectPath, "Done");
                if (Directory.Exists(folderDone))
                {
                    List<DirectoryInfo> listFoldersInsideDone = new DirectoryInfo(folderDone).GetDirectories().ToList();
                    var lastCreatedFolderInsideDone = listFoldersInsideDone.OrderByDescending(p => p.CreationTime).Select(p => p.Name).First();
                    NextPart = SetNextFolder(lastCreatedFolderInsideDone);

                }
                else
                {
                    NextPart = "Part A";
                }
            }
            #endregion

            Console.WriteLine();
            nextPartPath = Path.Combine(projectPath, NextPart);
            checkNextPartExist = Directory.Exists(nextPartPath);

            #region Delivery files
            switch (projectSet)
            {
                case ProjectSet.Caselaw:

                    totalCorrectFiles = listFileInfosCaselaw.Count;

                    if (listFileInfosCaselaw.Count > 0)
                    {
                        if (!checkNextPartExist)
                        {
                            clsConsole.WriteLine(string.Format(" Creating folder {0}", nextPartPath), ConsoleColor.Yellow);
                            Directory.CreateDirectory(nextPartPath);
                        }

                        clsConsole.WriteLine(string.Format(" - CASE LAW files for release {0}", totalCorrectFiles), ConsoleColor.Gray);
                        foreach (FileInfo fileInfo in listFileInfosCaselaw)
                        {
                            try
                            {
                                clsConsole.Write(string.Format("  * Moving file {0}", fileInfo.Name), ConsoleColor.DarkGray);
                                fileInfo.MoveTo(Path.Combine(nextPartPath, fileInfo.Name));
                            }
                            finally
                            {
                                clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                            }

                        }
                    }

                    clsReports.FilesCountReports.Caselaw = totalCorrectFiles;

                    break;
                case ProjectSet.NonVirgo:

                    totalCorrectFiles = listFileInfosNonVirgo.Count + listDirectoryInfosNonVirgo.Count;

                    if (listFileInfosNonVirgo.Count > 0)
                    {
                        if (!checkNextPartExist)
                        {
                            clsConsole.WriteLine(string.Format(" Creating folder {0}", nextPartPath), ConsoleColor.Yellow);
                            Directory.CreateDirectory(nextPartPath);
                        }

                        clsConsole.WriteLine(string.Format(" Non Virgo Agency files for release {0}", listFileInfosNonVirgo.Count), ConsoleColor.Gray);
                        foreach (FileInfo fileInfo in listFileInfosNonVirgo)
                        {
                            try
                            {
                                clsConsole.Write(string.Format("  * Moving file {0}", fileInfo.Name), ConsoleColor.DarkGray);
                                fileInfo.MoveTo(Path.Combine(nextPartPath, fileInfo.Name));
                            }
                            finally
                            {
                                clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                            }
                        }
                    }

                    if (listDirectoryInfosNonVirgo.Count > 0)
                    {
                        if (!checkNextPartExist)
                        {
                            clsConsole.WriteLine(string.Format(" Creating folder {0}", nextPartPath), ConsoleColor.Yellow);
                            Directory.CreateDirectory(nextPartPath);
                        }

                        clsConsole.WriteLine(string.Format(" Non Virgo folders for release {0}", listDirectoryInfosNonVirgo.Count), ConsoleColor.Gray);
                        foreach (DirectoryInfo directoryInfo in listDirectoryInfosNonVirgo)
                        {
                            try
                            {
                                clsConsole.Write(string.Format("  * Moving folder {0}", directoryInfo.Name), ConsoleColor.DarkGray);
                                directoryInfo.MoveTo(Path.Combine(nextPartPath, directoryInfo.Name));
                            }
                            finally
                            {
                                clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                            }
                        }
                    }

                    clsReports.FilesCountReports.NonVirgo = totalCorrectFiles;

                    break;
                case ProjectSet.Virgo:

                    totalCorrectFiles = listFileInfosVirgo.Count;

                    if (listFileInfosVirgo.Count > 0)
                    {
                        if (!checkNextPartExist)
                        {
                            clsConsole.WriteLine(string.Format(" Creating folder {0}", nextPartPath), ConsoleColor.Yellow);
                            Directory.CreateDirectory(nextPartPath);
                        }

                        clsConsole.WriteLine(string.Format(" - VIRGO files for release {0}", totalCorrectFiles), ConsoleColor.Gray);
                        foreach (FileInfo fileInfo in listFileInfosVirgo)
                        {
                            try
                            {
                                clsConsole.Write(string.Format("  * Moving file {0}", fileInfo.Name), ConsoleColor.DarkGray);
                                fileInfo.MoveTo(Path.Combine(nextPartPath, fileInfo.Name));
                            }
                            finally
                            {
                                clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                            }

                        }
                    }

                    clsReports.FilesCountReports.Virgo = totalCorrectFiles;

                    break;
                case ProjectSet.CaseRelated:

                    totalCorrectFiles = listDirectoryInfosCaseRelated.Count;

                    if (listDirectoryInfosCaseRelated.Count > 0)
                    {
                        if (!checkNextPartExist)
                        {
                            clsConsole.WriteLine(string.Format(" Creating folder {0}", nextPartPath), ConsoleColor.Yellow);
                            Directory.CreateDirectory(nextPartPath);
                        }

                        clsConsole.WriteLine(string.Format(" Case Related folders for release {0}", listDirectoryInfosCaseRelated.Count), ConsoleColor.Gray);
                        foreach (DirectoryInfo directoryInfo in listDirectoryInfosCaseRelated)
                        {
                            try
                            {
                                clsConsole.Write(string.Format("  * Moving folder {0}", directoryInfo.Name), ConsoleColor.DarkGray);
                                directoryInfo.MoveTo(Path.Combine(nextPartPath, directoryInfo.Name));
                            }
                            finally
                            {
                                clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                            }
                        }
                    }

                    clsReports.FilesCountReports.CaseRelated = totalCorrectFiles;

                    break;
                case ProjectSet.StateNet:

                    totalCorrectFiles = listFileInfosStateNet.Count + listDirectoryInfosStateNet.Count;

                    if (listFileInfosStateNet.Count > 0)
                    {
                        if (!checkNextPartExist)
                        {
                            clsConsole.WriteLine(string.Format(" Creating folder {0}", nextPartPath), ConsoleColor.Yellow);
                            Directory.CreateDirectory(nextPartPath);
                        }

                        clsConsole.WriteLine(string.Format("State Net files for release {0}", listFileInfosStateNet.Count), ConsoleColor.Gray);
                        foreach (FileInfo fileInfo in listFileInfosStateNet)
                        {
                            try
                            {
                                clsConsole.Write(string.Format("  * Moving file {0}", fileInfo.Name), ConsoleColor.DarkGray);
                                fileInfo.MoveTo(Path.Combine(nextPartPath, fileInfo.Name));
                            }
                            finally
                            {
                                clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                            }
                        }
                    }

                    if (listDirectoryInfosStateNet.Count > 0)
                    {
                        if (!checkNextPartExist)
                        {
                            clsConsole.WriteLine(string.Format(" Creating folder {0}", nextPartPath), ConsoleColor.Yellow);
                            Directory.CreateDirectory(nextPartPath);
                        }

                        clsConsole.WriteLine(string.Format(" State Net folders for release {0}", listDirectoryInfosStateNet.Count), ConsoleColor.Gray);
                        foreach (DirectoryInfo directoryInfo in listDirectoryInfosStateNet)
                        {
                            try
                            {
                                clsConsole.Write(string.Format("  * Moving folder {0}", directoryInfo.Name), ConsoleColor.DarkGray);
                                directoryInfo.MoveTo(Path.Combine(nextPartPath, directoryInfo.Name));
                            }
                            finally
                            {
                                clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                            }
                        }
                    }

                    clsReports.FilesCountReports.StateNet = totalCorrectFiles;

                    break;
                case ProjectSet.SecuritiesMosaic:

                    totalCorrectFiles = listDirectoryInfosSecuritiesMosaic.Count;

                    if (listDirectoryInfosSecuritiesMosaic.Count > 0)
                    {
                        if (!checkNextPartExist)
                        {
                            clsConsole.WriteLine(string.Format(" Creating folder {0}", nextPartPath), ConsoleColor.Yellow);
                            Directory.CreateDirectory(nextPartPath);
                        }

                        clsConsole.WriteLine(string.Format(" Securities Mosaic folders for release {0}", listDirectoryInfosSecuritiesMosaic.Count), ConsoleColor.Gray);
                        foreach (DirectoryInfo directoryInfo in listDirectoryInfosSecuritiesMosaic)
                        {
                            try
                            {
                                clsConsole.Write(string.Format("  * Moving folder {0}", directoryInfo.Name), ConsoleColor.DarkGray);
                                directoryInfo.MoveTo(Path.Combine(nextPartPath, directoryInfo.Name));
                            }
                            finally
                            {
                                clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                            }
                        }
                    }

                    clsReports.FilesCountReports.SecuritiesMosaic = totalCorrectFiles;
                    break;
            }
            #endregion

            if (Directory.Exists(nextPartPath))
            {
                Debug.WriteLine("===>>>" + nextPartPath);
                PreDelivery(projectSet);
                SheetName = clsExcel.GetWorkSheetName(projectSet);
                TransmittalProcess(projectSet);
                switch (projectSet)
                {
                    case ProjectSet.Caselaw:
                        RemoteSFtpPart = nextPartPath.Replace(projectPath, clsSFTP.rmCaselawPath).Replace(@"\", "/");
                        break;
                    case ProjectSet.NonVirgo:
                        RemoteSFtpPart = nextPartPath.Replace(projectPath, clsSFTP.rmNonVirgoPath).Replace(@"\", "/");
                        break;
                    case ProjectSet.Virgo:
                        RemoteSFtpPart = nextPartPath.Replace(projectPath, clsSFTP.rmVirgoPath).Replace(@"\", "/");
                        break;
                    case ProjectSet.CaseRelated:
                        RemoteSFtpPart = nextPartPath.Replace(projectPath, clsSFTP.rmCaseRelatedPath).Replace(@"\", "/");
                        break;
                    case ProjectSet.StateNet:
                        RemoteSFtpPart = nextPartPath.Replace(projectPath, clsSFTP.rmStateNetPath).Replace(@"\", "/");
                        break;
                    case ProjectSet.SecuritiesMosaic:
                        RemoteSFtpPart = nextPartPath.Replace(projectPath, clsSFTP.rmSecuritiesMosaicPath).Replace(@"\", "/");
                        break;
                }
            }
        }

        private static string SetNextFolder(string lastestPart)
        {
            Match match = Regex.Match(lastestPart, "[0-9A-Z]+", RegexOptions.RightToLeft);
            Match match1;
            int nextNum;
            if (match.Success)
            {
                if (match.Value.Contains("Z"))
                {
                    match1 = Regex.Match(match.Value, "[0-9]+", RegexOptions.RightToLeft);
                    if (match1.Success)
                    {
                        nextNum = Convert.ToInt32(match1.Value) + 1;

                        return "Part Z" + nextNum.ToString("D2");

                    }
                    else
                    {
                        return "Part Z01";
                    }
                }
                else
                {
                    match1 = Regex.Match(match.Value, "[A-Z]", RegexOptions.RightToLeft);
                    if (match1.Success)
                    {
                        nextNum = (int)char.ToUpper(Convert.ToChar(match1.Value)) + 1;

                        return "Part " + Convert.ToChar(nextNum);
                    }
                    else
                    {
                        throw new Exception("Can not get any letter");
                    }
                }
            }
            else
            {
                throw new Exception("Are you fucking kidding me ???");
            }
        }

        public static void PreDelivery(ProjectSet projectSet)
        {
            if (Directory.Exists(nextPartPath))
            {
                switch (projectSet)
                {
                    case ProjectSet.Caselaw:
                    case ProjectSet.Virgo:
                        foreach (FileInfo fileInfo in new DirectoryInfo(nextPartPath).GetFiles("*.zip").ToList())
                        {
                            clsFunction.ExtractStringFile(fileInfo.FullName, InputTransCaselaw);
                        }
                        break;

                    case ProjectSet.NonVirgo:
                        foreach (FileInfo fileInfo in new DirectoryInfo(nextPartPath).GetFiles("*.*", System.IO.SearchOption.AllDirectories)
                            .Where(p => Regex.IsMatch(p.Extension, "\\.(zip|visf|xml)", RegexOptions.IgnoreCase) && !Regex.IsMatch(p.Name, "Reports.zip", RegexOptions.IgnoreCase)
                            && !clsFunction.GetDirectoryName(p.DirectoryName).Equals("STRING", StringComparison.OrdinalIgnoreCase)))
                        {
                            switch (fileInfo.Extension.ToLower())
                            {
                                case ".zip":
                                    clsFunction.ExtractStringFile(fileInfo.FullName, InputTransNonVirgo);
                                    break;
                                default:
                                    fileInfo.CopyTo(Path.Combine(InputTransNonVirgo, fileInfo.Name), true);
                                    if (Regex.IsMatch(fileInfo.Name, "(D0KDV|D0VJE)", RegexOptions.IgnoreCase))
                                    {
                                        try
                                        {
                                            string zipOutput = Path.Combine(nextPartPath, clsFunction.GetDirectoryName(fileInfo.DirectoryName) + ".zip");

                                            if (File.Exists(zipOutput))
                                                File.Delete(zipOutput);

                                            clsConsole.Write(string.Format("  *** Starting compress D0KDV/D0VJE: {0}", Path.GetFileNameWithoutExtension(zipOutput)), ConsoleColor.DarkGray);
                                            ZipFile.CreateFromDirectory(fileInfo.DirectoryName, zipOutput, CompressionLevel.Optimal, true);
                                            clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                                        }
                                        finally
                                        {
                                            foreach (FileInfo f in new DirectoryInfo(fileInfo.DirectoryName).GetFiles("*.*", System.IO.SearchOption.AllDirectories))
                                            {
                                                f.Attributes = System.IO.FileAttributes.Normal;
                                            }

                                            fileInfo.Delete();
                                        }
                                    }
                                    break;
                            }
                        }
                        break;

                    case ProjectSet.CaseRelated:
                        clsConsole.WriteLine(string.Format("\r\n *** Starting merge and compress file(s) for Case Related\r\n      {0}\r\n", nextPartPath), ConsoleColor.Cyan);
                        foreach (DirectoryInfo directoryInfo in new DirectoryInfo(nextPartPath).GetDirectories())
                        {
                            string zipOutput = Path.Combine(nextPartPath, directoryInfo.Name + ".zip");

                            if (File.Exists(zipOutput))
                                File.Delete(zipOutput);

                            clsConsole.Write(string.Format("  *** Merging string {0}", directoryInfo.Name), ConsoleColor.DarkGray);
                            clsFunction.MergeString(directoryInfo, InputTransCaseRelated);
                            clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                            clsConsole.Write(string.Format("\r   *** Starting compress folder {0}", Path.GetFileNameWithoutExtension(zipOutput)), ConsoleColor.DarkGray);
                            ZipFile.CreateFromDirectory(directoryInfo.FullName, zipOutput, CompressionLevel.Optimal, true);
                            clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                        }
                        break;

                    case ProjectSet.StateNet:
                        foreach (FileInfo fileInfo in new DirectoryInfo(nextPartPath).GetFiles("*.zip", System.IO.SearchOption.AllDirectories)
                            .Where(p => Regex.IsMatch(p.Name, "(regulationtext|rdoctext)", RegexOptions.IgnoreCase)))
                        {
                            clsFunction.ExtractStringFile(fileInfo.FullName, InputTransStateNet);
                        }
                        break;
                    case ProjectSet.SecuritiesMosaic:
                        clsConsole.WriteLine(string.Format("\r\n *** Starting compress file(s) for Securities Mosaic\r\n      {0}\r\n", nextPartPath), ConsoleColor.Cyan);
                        foreach (DirectoryInfo directoryInfo in new DirectoryInfo(nextPartPath).GetDirectories())
                        {
                            string zipOutput = Path.Combine(nextPartPath, directoryInfo.Name + ".zip");
                            try
                            {
                                if (File.Exists(zipOutput))
                                    File.Delete(zipOutput);

                                clsFunction.GetStringMosaic(directoryInfo);
                                clsConsole.Write(string.Format("\r  *** Starting compress folder {0}", Path.GetFileNameWithoutExtension(zipOutput)), ConsoleColor.DarkGray);
                                ZipFile.CreateFromDirectory(directoryInfo.FullName, zipOutput, CompressionLevel.Optimal, false);
                                clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                            }
                            catch (Exception)
                            {
                                File.Delete(zipOutput);
                                clsConsole.WriteLine(" ... deleted", ConsoleColor.Red);
                            }

                        }
                        break;
                }
            }
        }

        public static bool TransmittalInitialize()
        {
            clsConsole.WriteLine("\r\n  Verify LN Schedule file ...", ConsoleColor.Red);
            if (!File.Exists(LNSchedule))
            {
                LNSchedule = Path.Combine(setTransmittalPath, @"Source\Lexis Nexis Delivery Schedule (" + DateTime.Now.ToString("MMMM yyyy") + ").xlsx");
                if (!File.Exists(LNSchedule))
                {
                    clsConsole.WriteLine(string.Format("\r\n   Schedule not found:\r\n{0}", LNSchedule), ConsoleColor.Red);
                    Environment.Exit(0);
                }
                else
                {
                    CopyofLNSchedule = Path.Combine(setTransmittalPath, @"Source\Copy of Lexis Nexis Delivery Schedule (" + DateTime.Now.ToString("MMMM yyyy") + ").xlsx");
                    SaveAsExcelFormat = Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook;
                }
            }

            Console.Clear();

            if (!clsFunction.FileInUsed(LNSchedule))
            {
                clsConsole.WriteLine("\r\n  Please wait while opening LN Schedule ...", ConsoleColor.Red);

                ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                ExcelApp.Visible = true;
                ExcelApp.DisplayAlerts = false;
                ExcelApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal;
                WorkbookExcel = ExcelApp.Workbooks.Open(LNSchedule, missing, false, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                return true;
            }
            else
            {
                return false;
            }


        }

        public static void TransmittalProcess(ProjectSet projectSet) 
        {
            Console.Clear();

            switch (projectSet)
            {
                case ProjectSet.Caselaw:
                    MessageBox.Show(projectSet.ToString() + "======" + clsReports.FilesCountReports.Caselaw);

                    clsConsole.WriteLine(string.Format("\r\n [{0}]CASELAW", clsReports.FilesCountReports.Caselaw), ConsoleColor.Cyan);
                    TransmittalFileInfos = new DirectoryInfo(InputTransCaselaw).GetFiles("*.*").Where(p => Regex.IsMatch(p.Extension, "\\.(xml|visf|txt|out)", RegexOptions.IgnoreCase)).ToList();
                    break;
                case ProjectSet.NonVirgo:
                    MessageBox.Show(projectSet.ToString() + "======" + clsReports.FilesCountReports.NonVirgo);

                    clsConsole.WriteLine(string.Format("\r\n [{0}]NON-VIRGO", clsReports.FilesCountReports.NonVirgo), ConsoleColor.Cyan);
                    TransmittalFileInfos = new DirectoryInfo(InputTransNonVirgo).GetFiles("*.*").Where(p => Regex.IsMatch(p.Extension, "\\.(xml|visf|txt|out)", RegexOptions.IgnoreCase)).ToList();
                    break;
                case ProjectSet.Virgo:
                    MessageBox.Show(projectSet.ToString() + "======" + clsReports.FilesCountReports.Virgo);

                    clsConsole.WriteLine(string.Format("\r\n [{0}]VIRGO", clsReports.FilesCountReports.NonVirgo), ConsoleColor.Cyan);
                    TransmittalFileInfos = new DirectoryInfo(InputTransCaselaw).GetFiles("*.*").Where(p => Regex.IsMatch(p.Extension, "\\.(xml|visf|txt|out)", RegexOptions.IgnoreCase)).ToList();
                    break;
                case ProjectSet.CaseRelated:
                    MessageBox.Show(projectSet.ToString() + "======" + clsReports.FilesCountReports.CaseRelated);

                    clsConsole.WriteLine(string.Format("\r\n [{0}]CASE RELATED", clsReports.FilesCountReports.CaseRelated), ConsoleColor.Cyan);
                    TransmittalFileInfos = new DirectoryInfo(InputTransCaseRelated).GetFiles("*.*").Where(p => Regex.IsMatch(p.Extension, "\\.(xml|visf|txt|out)", RegexOptions.IgnoreCase)).ToList();
                    break;
                case ProjectSet.StateNet:
                    MessageBox.Show(projectSet.ToString() + "======" + clsReports.FilesCountReports.StateNet);

                    clsConsole.WriteLine(string.Format("\r\n [{0}]STATE NET", clsReports.FilesCountReports.StateNet), ConsoleColor.Cyan);
                    TransmittalFileInfos = new DirectoryInfo(InputTransStateNet).GetFiles("*.*").Where(p => Regex.IsMatch(p.Extension, "\\.(xml|visf|txt|out)", RegexOptions.IgnoreCase)).ToList();
                    break;
                case ProjectSet.SecuritiesMosaic:
                    MessageBox.Show(projectSet.ToString() + "======" + clsReports.FilesCountReports.SecuritiesMosaic);

                    clsConsole.WriteLine(string.Format("\r\n [{0}]SECURITIES MOSAIC", clsReports.FilesCountReports.StateNet), ConsoleColor.Cyan);
                    TransmittalFileInfos = new DirectoryInfo(InputTransSecuritiesMosaic).GetFiles("*.*").Where(p => Regex.IsMatch(p.Extension, "\\.(xml|visf|txt|out)", RegexOptions.IgnoreCase)).ToList();
                    break;
            }

            MessageBox.Show("File count: " + TransmittalFileInfos.Count);

            SheetName = clsExcel.GetWorkSheetName(projectSet);

            if (!string.IsNullOrEmpty(SheetName))
            {
                if (TransmittalFileInfos.Count > 0)
                {
                    Debug.WriteLine(SheetName + "======" + TransmittalFileInfos.Count);
                    listJobnames = new List<clsReports.Jobnames>();
                    string fileName = string.Empty;
                    WorksheetExcel = (Microsoft.Office.Interop.Excel.Worksheet)WorkbookExcel.Worksheets[SheetName];
                    WorksheetExcel.Select();
                    TransmittalFileInfos.ForEach(f =>
                    {
                        fileName = Path.GetFileNameWithoutExtension(f.FullName);
                        Debug.WriteLine("Chay ky tu " + f.Name);
                        if (!fileName.Equals("manifest", StringComparison.OrdinalIgnoreCase))
                        {
                            clsConsole.Write(string.Format(" *** Searching {0} on delivery Schedule", fileName), ConsoleColor.DarkGray, false);
                            bool flag = clsExcel.SearchValue(fileName, WorksheetExcel);
                            if (flag)
                            {
                                range = (Microsoft.Office.Interop.Excel.Range)WorksheetExcel.Application.ActiveCell;
                                if (range.Value == fileName)
                                {
                                    clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                                    try
                                    {
                                        ImportValueToCell(SheetName, f);
                                        listJobnames.Add(new clsReports.Jobnames
                                        {
                                            FileName = fileName,
                                            ActualChars = Charcount
                                        });
                                    }
                                    catch (Exception ex)
                                    {
                                        errorCount++;
                                    }
                                }
                                else
                                {
                                    errorCount++;
                                    clsConsole.WriteLine(string.Format(" ... not found", fileName), ConsoleColor.Red);
                                }
                            }
                            else
                            {
                                clsConsole.WriteLine(string.Format(" ... try again", fileName), ConsoleColor.Red);
                                if (SheetName.Equals("CASE LAW") || SheetName.Equals("VIRGO"))
                                {
                                    tempFilesNotFound.Add(f);
                                }
                            }
                        }
                        else
                        {
                            f.Delete();
                        }
                    });

                    Action<string> runTrialCourt = strWorkSheet =>
                    {
                        Console.Clear();
                        clsConsole.WriteLine("\r\n ************ Trying to run with Trial Court ...", ConsoleColor.Red);
                        int indexPass = 0;
                        SheetName = strWorkSheet;
                        WorksheetExcel = (Microsoft.Office.Interop.Excel.Worksheet)WorkbookExcel.Worksheets[SheetName];
                        WorksheetExcel.Select();
                        tempFilesNotFound.ForEach(fnf =>
                        {
                            fileName = Path.GetFileNameWithoutExtension(fnf.FullName);
                            if (!fileName.Equals("manifest", StringComparison.OrdinalIgnoreCase))
                            {
                                clsConsole.Write(string.Format(" *** Searching {0} on delivery Schedule", fileName), ConsoleColor.DarkGray, false);
                                bool flag = clsExcel.SearchValue(fileName, WorksheetExcel);
                                if (flag)
                                {
                                    range = (Microsoft.Office.Interop.Excel.Range)WorksheetExcel.Application.ActiveCell;
                                    if (range.Value == fileName)
                                    {
                                        clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                                        try
                                        {
                                            ImportValueToCell(SheetName, fnf);
                                            listJobnames.Add(new clsReports.Jobnames
                                            {
                                                FileName = fileName,
                                                ActualChars = Charcount
                                            });
                                            indexPass++;
                                        }
                                        catch (Exception ex)
                                        {
                                            errorCount++;

                                        }
                                    }
                                    else
                                    {
                                        errorCount++;
                                        clsConsole.WriteLine(string.Format(" ... not found", fileName), ConsoleColor.Red);
                                    }
                                }
                                else
                                {
                                    clsConsole.WriteLine("Not found", ConsoleColor.Red);
                                }
                            }
                            else
                            {
                                fnf.Delete();
                            }
                        });
                        if (indexPass == tempFilesNotFound.Count)
                        {
                            tempFilesNotFound.Clear();
                        }
                        SheetName = clsExcel.GetWorkSheetName(projectSet);
                    };

                    if (tempFilesNotFound.Count > 0)
                    {
                        try
                        {
                            runTrialCourt("TRIAL COURT CASE LAW");
                        }
                        finally
                        {
                            if (tempFilesNotFound.Count > 0)
                            {
                                runTrialCourt("TRIAL COURT VIRGO");
                            }
                        }
                    }

                    Console.Clear();
                    clsConsole.WriteLine("\r\n Saving LN Schedule ...", ConsoleColor.Red);
                    WorkbookExcel.Save();

                    JobnameReports(projectSet);
                }
                TransmittalFileInfos.Clear();
            }
        }

        static void ImportValueToCell(string Project, FileInfo fileInfo)
        {
            Charcount = clsExcel.RunCharcount(Project, fileInfo.FullName);
            totalCharcount += Charcount;
            WorksheetExcel.Cells[range.Row, 11].Value = Charcount;
            WorksheetExcel.Cells[range.Row, 12].Value = DateTime.Now.ToString("MM/dd/yyyy");
            WorksheetExcel.Cells[range.Row, 13].Value = DateTime.Now.ToString("HH:mm:ss");
            WorksheetExcel.Cells[range.Row, 14].Value = "Released";
            WorksheetExcel.Cells[range.Row, 15].Value = clsExcel.DeliveryStatus(WorksheetExcel.Cells[range.Row, 2].Value.ToString(), WorksheetExcel.Cells[range.Row, 4].Value.ToString(), WorksheetExcel.Cells[range.Row, 12].Value.ToString(), WorksheetExcel.Cells[range.Row, 13].Value.ToString());
            fileInfo.Delete();
        }

        public static void JobnameReports(ProjectSet projectSet)
        {
            Console.Clear();

            switch (projectSet)
            {
                case ProjectSet.Caselaw:
                    JobnamesPath = Path.Combine(documentPath, "Jobnames_Caselaw_" + NextPart + ".xls");
                    break;
                case ProjectSet.NonVirgo:
                    JobnamesPath = Path.Combine(documentPath, "Jobnames_NonVirgo_" + NextPart + ".xls");
                    break;
                case ProjectSet.Virgo:
                    JobnamesPath = Path.Combine(documentPath, "Jobnames_Virgo_" + NextPart + ".xls");
                    break;
                case ProjectSet.CaseRelated:
                    JobnamesPath = Path.Combine(documentPath, "Jobnames_CaseRelated_" + NextPart + ".xls");
                    break;
                case ProjectSet.StateNet:
                    JobnamesPath = Path.Combine(documentPath, "Jobnames_StateNet_" + NextPart + ".xls");
                    break;
                case ProjectSet.SecuritiesMosaic:
                    JobnamesPath = Path.Combine(documentPath, "Jobnames_SecuritiesMosaic_" + NextPart + ".xls");
                    break;
            }

            if (listJobnames.Count > 0)
            {
                if (!string.IsNullOrEmpty(JobnamesPath))
                {
                    clsConsole.WriteLine("\r\n Creating Jobnames file", ConsoleColor.Cyan);
                    clsConsole.WriteLine("      " + JobnamesPath, ConsoleColor.Cyan);
                    int indexJN = 2;
                    Microsoft.Office.Interop.Excel.Application ExcelJN = new Microsoft.Office.Interop.Excel.Application();
                    ExcelJN.Visible = false;
                    Microsoft.Office.Interop.Excel.Workbook WorkBookJN = ExcelJN.Workbooks.Add();
                    Microsoft.Office.Interop.Excel.Worksheet WorkSheetJN = (Microsoft.Office.Interop.Excel.Worksheet)WorkBookJN.Worksheets[1];
                    WorkSheetJN.Cells[1, 1] = "Jobnames";
                    WorkSheetJN.Cells[1, 2] = "Actual Chars";
                    foreach (clsReports.Jobnames jn in listJobnames)
                    {
                        WorkSheetJN.Cells[indexJN, 1] = jn.FileName;
                        WorkSheetJN.Cells[indexJN, 2] = jn.ActualChars;
                        indexJN++;
                    }
                    WorkSheetJN.Name = "Search";

                    if (File.Exists(JobnamesPath))
                    {
                        if (!clsFunction.FileInUsed(JobnamesPath))
                        {
                            File.Delete(JobnamesPath);
                            WorkBookJN.SaveAs(JobnamesPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel8, missing, missing, missing, missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
                        }
                        else
                        {
                            JobnamesPath = Path.Combine(Path.GetDirectoryName(JobnamesPath), "Copy of " + Path.GetFileName(JobnamesPath));
                            WorkBookJN.SaveAs(JobnamesPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel8, missing, missing, missing, missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
                        }
                    }
                    else
                    {
                        WorkBookJN.SaveAs(JobnamesPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel8, missing, missing, missing, missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
                    }

                    ExcelJN.Quit();

                    listJobnames.Clear();
                }
            }

        }

        public static void SaveAsLNSchedule()
        {
            Console.Clear();
            try
            {
                clsConsole.WriteLine("\r\n Complete saving as Schedule ...", ConsoleColor.Red);
                WorkbookExcel.SaveAs(LNSchedule, SaveAsExcelFormat, missing, missing, missing, missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
            }
            catch (Exception)
            {
                clsConsole.WriteLine("\r\n Complete saving Copy of Schedule ...", ConsoleColor.Red);
                WorkbookExcel.SaveAs(CopyofLNSchedule, SaveAsExcelFormat, missing, missing, missing, missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
                ScheduleCopyExisted = true;
            }

            WorkbookExcel.Close();

            ExcelApp.Quit();
        }

        public static void LNScheduleBackup()
        {
            Console.Clear();
            clsConsole.WriteLine("\r\n **** Backup LN Schedule. ****", ConsoleColor.Cyan);

            if (!Directory.Exists(ScheduleBackup))
            {
                Directory.CreateDirectory(ScheduleBackup);
            }

            System.IO.FileStream fsCopy;
            if (ScheduleCopyExisted)
            {
                fsCopy = new System.IO.FileStream(CopyofLNSchedule, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            }
            else
            {
                fsCopy = new System.IO.FileStream(LNSchedule, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            }

            byte[] buffer = new byte[1024 * 1024];
            int bytesCount = 0;
            int bytesRead;
            string excelFormat = string.Empty;

            switch (Path.GetExtension(LNSchedule).ToLower())
            {
                case ".xlsx":
                    excelFormat = ".xlsx";
                    break;
                default:
                    excelFormat = ".xls";
                    break;
            }

            using (System.IO.FileStream fsPaste = new System.IO.FileStream(Path.Combine(ScheduleBackup, "Lexis Nexis Delivery Schedule (" + DateTime.Now.ToString("MMMM yyyy") + ")" + excelFormat), System.IO.FileMode.Create))
            {
                try
                {
                    while ((bytesRead = fsCopy.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        fsPaste.Write(buffer, 0, bytesRead);
                        bytesCount += bytesRead;
                        double per = (double)bytesCount * 100.0 / fsCopy.Length;

                        clsConsole.Write(string.Format("     Copying LN Schedule to Local drive D [{0}/{1}]", bytesCount, fsCopy.Length), ConsoleColor.DarkGray, false);
                    }
                    fsPaste.Close();
                }
                finally
                {
                    clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                    fsCopy.Close();
                }

            }

        }
    }
}
