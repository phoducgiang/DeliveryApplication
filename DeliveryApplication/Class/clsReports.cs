using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeliveryApplication.Class
{
    class clsReports
    {
        public class FilesCountReports
        {
            public static int Caselaw;
            public static int NonVirgo;
            public static int Virgo;
            public static int CaseRelated;
            public static int StateNet;
            public static int SecuritiesMosaic;
        }

        public class sFtpPathReports
        {
            public static string Caselaw;
            public static string NonVirgo;
            public static string Virgo;
            public static string CaseRelated;
            public static string StateNet;
            public static string SecuritiesMosaic;
        }

        public class Jobnames
        {
            public string FileName;
            public int ActualChars;
        }
    }
}
