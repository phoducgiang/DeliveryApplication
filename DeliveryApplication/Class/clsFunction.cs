using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Pri.LongPath;

namespace DeliveryApplication.Class
{
    class clsFunction
    {
        [DllImport("KERNEL32.DLL", CallingConvention = CallingConvention.StdCall, SetLastError = true)]
        private static extern bool SetProcessWorkingSetSize(IntPtr pProcess, long dwMinimumWorkingSetSize, long dwMaximumWorkingSetSize);

        public static void ExtractStringFile(string strFileInput, string strDestination)
        {
            using (ZipArchive zip = ZipFile.Open(strFileInput, ZipArchiveMode.Read))
            {
                var selection = zip.Entries.Where(p => Regex.IsMatch(p.Name, @"([0-9A-Z\-]{20}00000-00|(regulationtext|rdoctext)(.*)).(xml|visf)", RegexOptions.RightToLeft));
                if (selection.Count() > 0)
                {
                    foreach (var item in selection)
                    {
                        item.ExtractToFile(Path.Combine(strDestination, item.Name), true);
                    }
                }
            }
        }

        public static void MergeString(DirectoryInfo dirInfoFolder, string strOutputFolder)
        {
            List<DirectoryInfo> lstStringFolder = dirInfoFolder.GetDirectories().Where(p => Regex.IsMatch(p.Name, "up(.*)isf|for(.*)ing", RegexOptions.IgnoreCase)).ToList();
            if (lstStringFolder.Count > 0)
            {
                foreach (DirectoryInfo i in lstStringFolder)
                {
                    Match m = Regex.Match(i.FullName, "D[A-Z0-9]{4}B[0-9]{6}", RegexOptions.IgnoreCase);
                    if (m.Success)
                    {
                        string outputFile = Path.Combine(strOutputFolder, m.Value + ".txt");

                        if (File.Exists(outputFile))
                        {
                            File.Delete(outputFile);
                        }

                        List<string> lstFiles = Directory.GetFiles(i.FullName, "*.*").Where(p => Regex.IsMatch(Path.GetExtension(p), "\\.(visf|xml)")).ToList();
                        if (lstFiles.Count > 0)
                        {
                            System.IO.StreamWriter sw = new System.IO.StreamWriter(outputFile, false, Encoding.Default);
                            int index = 1;
                            foreach (string f in lstFiles)
                            {
                                System.IO.StreamReader sr = new System.IO.StreamReader(f, Encoding.Default);
                                string Reader = sr.ReadToEnd().Trim();
                                sr.Close();
                                if (index == lstFiles.Count)
                                {
                                    sw.Write(Reader.TrimEnd());
                                }
                                else
                                {
                                    sw.WriteLine(Reader.TrimEnd());
                                }
                                index++;
                            }
                            sw.Close();
                        }
                    }
                }
            }
        }

        public static void GetStringMosaic(DirectoryInfo dirInfoFolder)
        {
            List<FileInfo> fileInfos = dirInfoFolder.GetFiles("*.xml", System.IO.SearchOption.AllDirectories).Where(p => Regex.IsMatch(p.DirectoryName, "02_spi", RegexOptions.IgnoreCase)).ToList();
            if (fileInfos.Count > 0)
            {
                foreach (FileInfo f in fileInfos)
                {
                    f.CopyTo(Path.Combine(clsVariable.InputTransSecuritiesMosaic, f.Name), true);
                }
            }
        }

        public static string GetDirectoryName(string inputDir)
        {
            return new FileInfo(inputDir).Name;
        }

        public static bool FileInUsed(string inputFile)
        {
            FileInfo fi = new FileInfo(inputFile);
            System.IO.FileStream fs = null;
            try
            {
                fs = fi.Open(System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.None);
            }
            catch (Exception)
            {
                return true;
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }
            return false;
        }

        public static void ReleaseMemory()
        {
            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                if (Environment.OSVersion.Platform == PlatformID.Win32NT)
                {
                    if (!Debugger.IsAttached)
                    {
                        SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1);
                    }
                    Thread.Sleep(1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static void OpenFormSTA(ThreadStart thrStart)
        {
            Thread thread = new Thread(thrStart);
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        public static void RestartApp(int Interval)
        {
            Console.Clear();
            DateTime dateTime = DateTime.Now.AddMinutes(Convert.ToDouble(Interval));
            int num = Convert.ToInt32(Interval) * 60;
            clsConsole.WriteLine("\r\n Next checking will be on: " + dateTime.ToLongTimeString(), ConsoleColor.Red);
            for (int i = 0; i <= num; i++)
            {
                Thread.Sleep(1000);
                ReleaseMemory();
            }
            Console.Clear();
            Program.Main(new string[]
            {
                "-auto"
            });
        }

        public static clsVariable.ProjectSet StringToEnum(string strConvert)
        {
            clsVariable.ProjectSet projectSet;
            if (Enum.TryParse(strConvert, true, out projectSet))
            {
                return projectSet;
            }
            else
            {
                throw new Exception("Can not parse string to enum.");
                return 0;
            }
        }

        public static void ClipboardSetText(string strPath)
        {
            bool checkStrEmpty = string.IsNullOrEmpty(strPath) ? true : false;
            if (!checkStrEmpty)
            {
                Clipboard.SetText(strPath, TextDataFormat.Text);
            }
        }

        public static void ClearJobnames()
        {
            foreach (FileInfo fiJN in new DirectoryInfo(clsVariable.documentPath).GetFiles("*.xls").Where(p => Regex.IsMatch(p.Name, "(Jobnames|Copy of Jobnames)_(Caselaw|CaseRelated|NonVirgo|Virgo|StateNet|SecuritiesMosaic)_Part [A-Z0-9]+.xls", RegexOptions.IgnoreCase)).ToList())
            {
                if (!clsFunction.FileInUsed(fiJN.FullName))
                {
                    fiJN.Delete();
                }
            }
        }

        public static void CleanupInput()
        {
            using (System.Diagnostics.Process cmd = new System.Diagnostics.Process())
            {
                System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo()
                {
                    WorkingDirectory = clsVariable.InputTransmittal,
                    WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden,
                    FileName = "cmd.exe",
                    Arguments = "/C DEL /F /S *.txt;*.visf;*.xml;*.out"
                };
                cmd.StartInfo = startInfo;
                cmd.Start();
                cmd.WaitForExit();
            }
        }
    }
}
