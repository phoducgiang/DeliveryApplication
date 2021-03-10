using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DeliveryApplication.Class;
using DeliveryApplication.Forms;

namespace DeliveryApplication
{
    class Program
    {
        [DllImport("kernel32.dll")] static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        public static IntPtr handle;

        public static void Main(string[] args)
        {
            Console.SetWindowSize(150, 8);
            handle = GetConsoleWindow();
            clsConsole.WriteLine("\r\n Checking connection to sFtp ...", ConsoleColor.Red);
            clsSFTP.sFtpCreateFolderByCurrentDate();
            clsFunction.OpenFormSTA(FormMain);
            ShowWindow(handle, SW_HIDE);
        }

        static void FormMain()
        {
            frmMain form = new frmMain();
            form.ShowDialog();
        }
    }
}
