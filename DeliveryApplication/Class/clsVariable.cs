using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Pri.LongPath;

namespace DeliveryApplication.Class
{
    class clsVariable
    {
        public static string setReleasePath = Properties.Settings.Default.ReleasePath;
        public static string setTransmittalPath = Properties.Settings.Default.TrasmittalPath;
        public static int setInterval = Properties.Settings.Default.Interval;
        public static string CaseLaw = Path.Combine(setReleasePath, "01-CASE LAW", DateTime.Now.ToString("MM-dd-yyyy"));
        public static string Virgo = Path.Combine(setReleasePath, "03-VIRGO", DateTime.Now.ToString("MM-dd-yyyy"));
        public static string NonVirgo = Path.Combine(setReleasePath, "04-NON VIRGO", DateTime.Now.ToString("MM-dd-yyyy"));
        public static string CaseRelated = Path.Combine(setReleasePath, "06-CASE RELATED", DateTime.Now.ToString("MM-dd-yyyy"));
        public static string SecuritiesMosaic = Path.Combine(setReleasePath, "10-SECURITIES MOSAIC", DateTime.Now.ToString("MM-dd-yyyy"));
        public static string StateNet = Path.Combine(setReleasePath, "11-STATE NET", DateTime.Now.ToString("MM-dd-yyyy"));
        public static string NextPart = string.Empty;
        public static bool checkNextPartExist;
        public static int totalCorrectFiles;
        public static string projectPath = string.Empty;
        public static string nextPartPath = string.Empty;

        public static List<DirectoryInfo> listPartFolders = null;
        public static List<FileInfo> listFileInfosCaselaw = null;
        public static List<FileInfo> listFileInfosNonVirgo = null;
        public static List<FileInfo> listFileInfosVirgo = null;
        public static List<FileInfo> listFileInfosStateNet = null;
        public static List<DirectoryInfo> listDirectoryInfosNonVirgo = null;
        public static List<DirectoryInfo> listDirectoryInfosCaseRelated = null;
        public static List<DirectoryInfo> listDirectoryInfosStateNet = null;
        public static List<DirectoryInfo> listDirectoryInfosSecuritiesMosaic = null;
        public static List<FileInfo> TransmittalFileInfos = null;
        public static List<clsReports.Jobnames> listJobnames = null;
        public static List<FileInfo> tempFilesNotFound = new List<FileInfo>();

        public static string documentPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        public static string InputTransmittal = Path.Combine(setTransmittalPath, "INPUT");
        public static string InputTransCaselaw = Path.Combine(setTransmittalPath, @"INPUT\1. CASE LAW & VIRGO");
        public static string InputTransNonVirgo = Path.Combine(setTransmittalPath, @"INPUT\2. NON VIRGO");
        public static string InputTransCaseRelated = Path.Combine(setTransmittalPath, @"INPUT\4. CASE RELATED");
        public static string InputTransSecuritiesMosaic = Path.Combine(setTransmittalPath, @"INPUT\5. SECURITY MOSAIC");
        public static string InputTransStateNet = Path.Combine(setTransmittalPath, @"INPUT\6. STATE NET");
        public static string LNSchedule = Path.Combine(setTransmittalPath, @"Source\Lexis Nexis Delivery Schedule (" + DateTime.Now.ToString("MMMM yyyy") + ").xls");
        public static string CopyofLNSchedule = Path.Combine(setTransmittalPath, @"Source\Copy of Lexis Nexis Delivery Schedule (" + DateTime.Now.ToString("MMMM yyyy") + ").xls");
        public static string ScheduleBackup = @"D:\LN Schedule\" + DateTime.Now.ToString("yyyy") + @"\" + DateTime.Now.ToString("MM-MMMM");
        public static bool ScheduleCopyExisted = false;
        public static string JobnamesPath = string.Empty;

        public static Microsoft.Office.Interop.Excel.Application ExcelApp;
        public static Microsoft.Office.Interop.Excel.Workbook WorkbookExcel;
        public static Microsoft.Office.Interop.Excel.Worksheet WorksheetExcel;
        public static Microsoft.Office.Interop.Excel.Range range;
        public static object missing = System.Reflection.Missing.Value;
        public static object SaveAsExcelFormat = Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel8;
        public static string SheetName = string.Empty;

        public static int totalCharcount = 0;
        public static int errorCount = 0;
        public static int Charcount = 0;

        public static bool isAutoRun = false;

        public enum ProjectSet
        {
            Caselaw,
            NonVirgo,
            Virgo,
            CaseRelated,
            StateNet,
            SecuritiesMosaic
        }
    }
}
