using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeliveryApplication.Class
{
    class clsConsole
    {
        public static void WriteLine(string text, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(text);
            Console.ForegroundColor = ConsoleColor.Gray;
        }

        public static void Write(string text, ConsoleColor color, bool noNewLine = true)
        {
            Console.ForegroundColor = color;
            if (noNewLine)
            {
                Console.Write(text);
            }
            else
            {
                Console.Write("\r" + text);
            }
            Console.ForegroundColor = ConsoleColor.Gray;
        }
    }
}
