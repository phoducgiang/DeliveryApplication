using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeliveryApplication.Class
{
    class clsExcel : clsVariable
    {
        public static string GetWorkSheetName(ProjectSet projectSet)
        {
            string sName = string.Empty;
            switch (projectSet)
            {
                case ProjectSet.Caselaw:
                    sName = "CASE LAW";
                    break;
                case ProjectSet.NonVirgo:
                    sName = "NON-VIRGO";
                    break;
                case ProjectSet.Virgo:
                    sName = "VIRGO";
                    break;
                case ProjectSet.CaseRelated:
                    sName = "CASE RELATED";
                    break;
                case ProjectSet.StateNet:
                    sName = "STATE NET";
                    break;
                case ProjectSet.SecuritiesMosaic:
                    sName = "SECURITY MOSAIC";
                    break;
            }

            return sName;
        }

        public static int RunCharcount(string Project, string strFileName)
        {
            int Chars = 0;
            switch (Project)
            {
                case "CASE LAW":
                case "VIRGO":
                case "SECURITY MOSAIC":
                case "STATE NET":
                    Chars = clsLNCharcount.XMLCharcount(strFileName);
                    break;
                default:
                    Chars = clsLNCharcount.VISFCharcount(strFileName);
                    break;
            }
            return Chars;
        }

        public static bool SearchValue(object searchValue, Microsoft.Office.Interop.Excel.Worksheet objWs)
        {
            bool flag;
            try
            {
                // Find the last real row
                int rowindex = objWs.Cells.Find(searchValue, System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                               Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                // Find the last real column
                int colindex = objWs.Cells.Find(searchValue, System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                                Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

                objWs.Cells[rowindex, colindex].Select();
                flag = true;
            }
            catch (Exception)
            {
                flag = false;
            }
            return flag;
        }

        public static string DeliveryStatus(object strRSD, object strTimeDue, object strRSDActual, object strTimeActual)
        {
            Func<string, DateTime> ConvertFromOADate = ValueOADate =>
            {
                double d = double.Parse(ValueOADate);
                return DateTime.FromOADate(d);
            };

            DateTime RSD = Convert.ToDateTime(strRSD);
            DateTime TimeDue = ConvertFromOADate(Convert.ToString(strTimeDue));
            DateTime RSDActual = Convert.ToDateTime(strRSDActual);
            DateTime TimeActual = ConvertFromOADate(Convert.ToString(strTimeActual));
            DateTime CompareRSD = new DateTime(RSD.Year, RSD.Month, RSD.Day, TimeDue.Hour, TimeDue.Minute, 0);
            DateTime CompareActual = new DateTime(RSDActual.Year, RSDActual.Month, RSDActual.Day, TimeActual.Hour, TimeActual.Minute, 0);

            if (CompareActual < CompareRSD)
            {
                return "Ahead";
            }
            else if (CompareActual == CompareRSD)
            {
                return "OnTime";
            }
            else
            {
                return "Late";
            }
        }

        public bool VerifyExcelApp()
        {
            ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            if (ExcelApp != null)
            {
                ExcelApp.Quit();
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
