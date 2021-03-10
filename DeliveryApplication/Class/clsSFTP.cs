using Pri.LongPath;
using Renci.SshNet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DeliveryApplication.Class
{
    class clsSFTP
    {
        static string host = @"sftp2-sp.spi-global.com";
        static string sFtpUser = "spisp_l115";
        static string sFtpPassword = "waWRat22";
        static SftpClient sFtpClient;
        static ConnectionInfo connectionInfo;
        static string RemoteSFtpDir = string.Empty;
        static string RemoteSFtpFilePath = string.Empty;
        static string RemoteSFtpPart = string.Empty;
        public static string rmCaselawPath = "/TO SPI/Lexis-Nexis/CL_MOD FILES/CASELAW_XML/" + DateTime.Now.ToString("MM-dd-yyyy");
        public static string rmNonVirgoPath = "/TO SPI/Lexis-Nexis/NON VIRGO/" + DateTime.Now.ToString("MM-dd-yyyy");
        public static string rmVirgoPath = "/TO SPI/Lexis-Nexis/CL_MOD FILES/VIRGO_XML/" + DateTime.Now.ToString("MM-dd-yyyy");
        public static string rmCaseRelatedPath = "/TO SPI/Lexis-Nexis/CASE RELATED/" + DateTime.Now.ToString("MM-dd-yyyy");
        public static string rmStateNetPath = "/TO SPI/Lexis-Nexis/STATE_NET/" + DateTime.Now.ToString("MM-dd-yyyy");
        public static string rmSecuritiesMosaicPath = "/TO SPI/Lexis-Nexis/SECURITIES MOSAIC/" + DateTime.Now.ToString("MM-dd-yyyy");
        static string remoteDirectory = string.Empty;

        private static void sFtpInitialize()
        {
            connectionInfo = new ConnectionInfo(host, sFtpUser, new PasswordAuthenticationMethod(sFtpUser, sFtpPassword));
        }

        private static bool sFtpCheckConnection()
        {
            sFtpInitialize();

            using (sFtpClient = new SftpClient(connectionInfo))
            {
                try
                {
                    sFtpClient.Connect();
                    if (sFtpClient.IsConnected)
                    {
                        sFtpClient.Disconnect();
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    return false;
                }
            }
        }

        private static bool sFtpCheckFolderExists(string strRemoteSFtpPath)
        {
            return sFtpClient.Exists(strRemoteSFtpPath) ? true : false;
        }

        public static void sFtpCreateFolderByCurrentDate()
        {
            if (sFtpCheckConnection())
            {
                Console.Clear();

                using (sFtpClient = new SftpClient(connectionInfo))
                {
                    sFtpClient.Connect();

                    if (sFtpClient.IsConnected)
                    {
                        clsConsole.WriteLine("\r\n ======= sFtp check folder existed =======", ConsoleColor.Red);

                        if (!sFtpCheckFolderExists(rmCaselawPath))
                        {
                            clsConsole.WriteLine(string.Format(" ***** Folder on sFtp {0} has been created. ***** ", rmCaselawPath), ConsoleColor.Cyan);
                            sFtpClient.CreateDirectory(rmCaselawPath);
                        }

                        if (!sFtpCheckFolderExists(rmNonVirgoPath))
                        {
                            clsConsole.WriteLine(string.Format(" ***** Folder on sFtp {0} has been created. ***** ", rmNonVirgoPath), ConsoleColor.Cyan);
                            sFtpClient.CreateDirectory(rmNonVirgoPath);
                        }

                        if (!sFtpCheckFolderExists(rmVirgoPath))
                        {
                            clsConsole.WriteLine(string.Format(" ***** Folder on sFtp {0} has been created. ***** ", rmVirgoPath), ConsoleColor.Cyan);
                            sFtpClient.CreateDirectory(rmVirgoPath);
                        }

                        if (!sFtpCheckFolderExists(rmCaseRelatedPath))
                        {
                            clsConsole.WriteLine(string.Format(" ***** Folder on sFtp {0} has been created. ***** ", rmCaseRelatedPath), ConsoleColor.Cyan);
                            sFtpClient.CreateDirectory(rmCaseRelatedPath);
                        }

                        if (!sFtpCheckFolderExists(rmStateNetPath))
                        {
                            clsConsole.WriteLine(string.Format(" ***** Folder on sFtp {0} has been created. ***** ", rmStateNetPath), ConsoleColor.Cyan);
                            sFtpClient.CreateDirectory(rmStateNetPath);
                        }

                        if (!sFtpCheckFolderExists(rmSecuritiesMosaicPath))
                        {
                            clsConsole.WriteLine(string.Format(" ***** Folder on sFtp {0} has been created. ***** ", rmSecuritiesMosaicPath), ConsoleColor.Cyan);
                            sFtpClient.CreateDirectory(rmSecuritiesMosaicPath);
                        }

                        sFtpClient.Disconnect();
                    }
                }
            }
        }

        public static void sFtpUploadFiles(clsVariable.ProjectSet projectSet, string strProjectPath, string strNextPart)
        {
            if (Directory.Exists(strNextPart))
            {
                Console.Clear();
                clsConsole.WriteLine(string.Format("\r\n  *****************  Starting upload: {0}", strNextPart), ConsoleColor.Cyan);
                if (sFtpCheckConnection())
                {
                    using (sFtpClient = new SftpClient(connectionInfo))
                    {
                        sFtpClient.Connect();

                        if (sFtpClient.IsConnected)
                        {
                            switch (projectSet)
                            {
                                case clsVariable.ProjectSet.Caselaw:
                                    RemoteSFtpPart = strNextPart.Replace(strProjectPath, rmCaselawPath).Replace(@"\", "/");
                                    break;
                                case clsVariable.ProjectSet.NonVirgo:
                                    RemoteSFtpPart = strNextPart.Replace(strProjectPath, rmNonVirgoPath).Replace(@"\", "/");
                                    break;
                                case clsVariable.ProjectSet.Virgo:
                                    RemoteSFtpPart = strNextPart.Replace(strProjectPath, rmVirgoPath).Replace(@"\", "/");
                                    break;
                                case clsVariable.ProjectSet.CaseRelated:
                                    RemoteSFtpPart = strNextPart.Replace(strProjectPath, rmCaseRelatedPath).Replace(@"\", "/");
                                    break;
                                case clsVariable.ProjectSet.StateNet:
                                    RemoteSFtpPart = strNextPart.Replace(strProjectPath, rmStateNetPath).Replace(@"\", "/");
                                    break;
                                case clsVariable.ProjectSet.SecuritiesMosaic:
                                    RemoteSFtpPart = strNextPart.Replace(strProjectPath, rmSecuritiesMosaicPath).Replace(@"\", "/");
                                    break;
                            }

                            if (!sFtpCheckFolderExists(RemoteSFtpPart))
                            {
                                clsConsole.WriteLine(string.Format("\r\n  ** Create remote sFtp Path: {0}", RemoteSFtpPart), ConsoleColor.Green);

                                sFtpClient.CreateDirectory(RemoteSFtpPart);

                                Match m = Regex.Match(RemoteSFtpPart, "(CASELAW_XML|NON VIRGO|VIRGO_XML|CASE RELATED|STATE_NET|SECURITIES MOSAIC)", RegexOptions.RightToLeft);
                                switch (m.Value)
                                {
                                    case "CASELAW_XML":
                                        clsReports.sFtpPathReports.Caselaw = RemoteSFtpPart;
                                        break;
                                    case "NON VIRGO":
                                        clsReports.sFtpPathReports.NonVirgo = RemoteSFtpPart;
                                        break;
                                    case "VIRGO_XML":
                                        clsReports.sFtpPathReports.Virgo = RemoteSFtpPart;
                                        break;
                                    case "CASE RELATED":
                                        clsReports.sFtpPathReports.CaseRelated = RemoteSFtpPart;
                                        break;
                                    case "STATE_NET":
                                        clsReports.sFtpPathReports.StateNet = RemoteSFtpPart;
                                        break;
                                    case "SECURITIES MOSAIC":
                                        clsReports.sFtpPathReports.SecuritiesMosaic = RemoteSFtpPart;
                                        break;
                                }
                            }

                            string filePath = string.Empty;
                            DirectoryInfo directoryInfo = new DirectoryInfo(strNextPart);

                            switch (projectSet)
                            {
                                case clsVariable.ProjectSet.NonVirgo:
                                case clsVariable.ProjectSet.StateNet:
                                    foreach (DirectoryInfo dir in directoryInfo.GetDirectories("*", System.IO.SearchOption.AllDirectories))
                                    {
                                        RemoteSFtpDir = dir.FullName.Replace(strNextPart, RemoteSFtpPart).Replace(@"\", "/") + "/";

                                        if (!sFtpCheckFolderExists(RemoteSFtpDir))
                                        {
                                            try
                                            {
                                                clsConsole.Write(string.Format("  *** Create remote sFtp Folders & Sub Folders: {0}", RemoteSFtpDir), ConsoleColor.DarkGray);
                                                sFtpClient.CreateDirectory(RemoteSFtpDir);
                                            }
                                            finally
                                            {
                                                clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                                            }
                                        }
                                    }

                                    clsConsole.WriteLine("\r\n  ** Starting upload file ...", ConsoleColor.Green);
                                    foreach (FileInfo file in directoryInfo.GetFiles("*.*", System.IO.SearchOption.AllDirectories))
                                    {
                                        if (!Regex.IsMatch(file.Name, "(D0KDV|D0VJE)", RegexOptions.IgnoreCase))
                                        {
                                            using (var FileStream = File.OpenRead(file.FullName))
                                            {
                                                try
                                                {
                                                    RemoteSFtpDir = file.DirectoryName.Replace(strNextPart, RemoteSFtpPart).Replace(@"\", "/") + "/";
                                                    if (!sFtpCheckFolderExists(RemoteSFtpDir))
                                                    {
                                                        sFtpClient.CreateDirectory(RemoteSFtpDir);
                                                    }
                                                    RemoteSFtpFilePath = RemoteSFtpDir + file.Name;
                                                    clsConsole.Write(string.Format("  *** Uploading file: {0}", RemoteSFtpFilePath), ConsoleColor.DarkGray);
                                                    sFtpClient.UploadFile(FileStream, RemoteSFtpFilePath, true);

                                                }
                                                finally
                                                {
                                                    clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                                                }
                                            }
                                        }
                                    }
                                    break;
                                default:
                                    foreach (FileInfo file in directoryInfo.GetFiles("*.zip"))
                                    {
                                        using (var FileStream = File.OpenRead(file.FullName))
                                        {
                                            try
                                            {
                                                RemoteSFtpDir = file.DirectoryName.Replace(strNextPart, RemoteSFtpPart).Replace(@"\", "/") + "/";
                                                if (!sFtpCheckFolderExists(RemoteSFtpDir))
                                                {
                                                    sFtpClient.CreateDirectory(RemoteSFtpDir);
                                                }
                                                RemoteSFtpFilePath = RemoteSFtpDir + file.Name;
                                                clsConsole.Write(string.Format("  *** Uploading file: {0}", RemoteSFtpFilePath), ConsoleColor.DarkGray);
                                                sFtpClient.UploadFile(FileStream, RemoteSFtpFilePath, true);
                                            }
                                            finally
                                            {
                                                clsConsole.WriteLine(" ... ok", ConsoleColor.Green);
                                            }
                                        }
                                    }
                                    break;
                            }

                            sFtpClient.Disconnect();
                        }
                    }
                }
            }
        }
    }
}
