using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DeliveryApplication.Class
{
    class clsLNCharcount
    {
        private static int intEntChar, intTagChar;

        private static int GetXmlTagCount(string sLine)
        {
            foreach (object obj in Regex.Matches(sLine, "</?[^>]+>"))
            {
                Match matItem = (Match)obj;
                string strItem = matItem.Value;
            }
            return Regex.Matches(sLine, "</?[^>]+>").Count;
        }

        private static int GetXmlCharCount(string sLine)
        {
            string sChar = Regex.Replace(sLine, "</?[^>]+>", "");
            sChar = Regex.Replace(sChar, "&[^\\s]+;", "");
            return sChar.Length;
        }

        private static int GetEntityCount(string sline)
        {
            int intCnt = 0;
            checked
            {
                foreach (object obj in Regex.Matches(sline, "\\&[^\\s]+\\;"))
                {
                    Match matItem = (Match)obj;
                    string strItem = matItem.Value;
                    bool flag = CountOccurrences(strItem, ";") == 1;
                    if (flag)
                    {
                        intEntChar += strItem.Length;
                        intCnt++;
                    }
                    else
                    {
                        for (int intSpos = Strings.InStr(strItem, "&", CompareMethod.Binary); intSpos != 0; intSpos = Strings.InStr(intSpos + 2, strItem, "&", CompareMethod.Binary))
                        {
                            int intEpos = Strings.InStr(intSpos, strItem, ";", CompareMethod.Binary);
                            flag = (intEpos != 0);
                            if (flag)
                            {
                                string strEnt = Strings.Mid(strItem, intSpos, intEpos - intSpos + 1);
                                intEntChar += strEnt.Length;
                                intCnt++;
                            }
                        }
                    }
                }
                return intCnt;
            }
        }

        private static int CountOccurrences(string data, string find)
        {
            return Regex.Matches(data, Regex.Escape(find), RegexOptions.IgnoreCase).Count;
        }

        private static string CheckMultiSpaceBlankLine(string inputFile)
        {
            try
            {
                StreamReader sr = new StreamReader(inputFile, Encoding.Default);
                string Reader = sr.ReadToEnd();
                sr.Close();

                if (Regex.IsMatch(Reader, "  ", RegexOptions.IgnoreCase))
                {
                    throw new Exception("Multiple spaces found.");
                }
                else if (Regex.IsMatch(Reader, @"(\r?\n\s*){2,}", RegexOptions.IgnoreCase))
                {
                    throw new Exception("Blank line found.");
                }
                else
                {
                    return Reader;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public static int XMLCharcount(string inputFile)
        {
            int intChar = 0;
            int intElement = 0;
            int intEntity = 0;
            int intTotal = 0;
            intTagChar = 0;
            intEntChar = 0;
            checked
            {
                try
                {
                    FileSystem.FileOpen(1, inputFile, OpenMode.Input, OpenAccess.Read, OpenShare.Default, -1);
                    while (!FileSystem.EOF(1))
                    {
                        string strLine = FileSystem.LineInput(1);
                        intChar += GetXmlCharCount(strLine);
                        intElement += GetXmlTagCount(strLine);
                        intEntity += GetEntityCount(strLine);
                    }
                    intTagChar = intElement * 2;
                    intTotal = intChar + intTagChar + intEntChar;
                }
                catch (Exception)
                {
                    return 0;
                }
                finally
                {
                    FileSystem.FileClose(1);
                }

                return intTotal;
            }
        }

        public static int VISFCharcount(string inputFile)
        {
            StreamReader sr = new StreamReader(inputFile, Encoding.Default);
            string Reader = sr.ReadToEnd();
            sr.Close();

            Reader = Reader.Replace("\r", "");
            Reader = Reader.Replace("\n", "");
            Reader = Reader.Replace("\r\n", "");
            Reader = Reader.Replace("  ", " ");
            Reader = Regex.Replace(Reader, @"(\r?\n\s*){2,}", "", RegexOptions.IgnoreCase);

            return Reader.Trim().Length;
        }

        public static int CANADACharcount(string inputFile)
        {
            return File.ReadAllText(inputFile).Length;
        }
    }
}
