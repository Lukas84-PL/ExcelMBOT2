using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Windows;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Diagnostics;

namespace ExcelDataManipulation
{
    class ExcelInstance
    {
        [DllImport("Oleacc.dll")]
        public static extern int AccessibleObjectFromWindow(
       int hwnd, uint dwObjectID, byte[] riid,
       ref Microsoft.Office.Interop.Excel.Window ptr);

        public delegate bool EnumChildCallback(int hwnd, ref int lParam);

        [DllImport("User32.dll")]
        public static extern bool EnumChildWindows(
        int hWndParent, EnumChildCallback lpEnumFunc,
        ref int lParam);


        [DllImport("User32.dll")]
        public static extern int GetClassName(
        int hWnd, StringBuilder lpClassName, int nMaxCount);

        public static bool EnumChildProc(int hwndChild, ref int lParam)
        {
            StringBuilder buf = new StringBuilder(128);
            GetClassName(hwndChild, buf, 128);
            if (buf.ToString() == "EXCEL7")
            {
                lParam = hwndChild;
                return false;
            }
            return true;
        }
        public string Instance(string workbookname, string visible, out Workbook workbook, out Application application, out Worksheet sheet, string sheetname = "")
        {
            Excel.Application app = new Excel.Application();
            EnumChildCallback cb;
            List<Process> procs = new List<Process>();
            procs.AddRange(Process.GetProcessesByName("excel"));

            foreach (Process p in procs)
            {
                if ((int)p.MainWindowHandle > 0)
                {
                    int childWindow = 0;
                    cb = new EnumChildCallback(EnumChildProc);
                    EnumChildWindows((int)p.MainWindowHandle, cb, ref childWindow);

                    if (childWindow > 0)
                    {
                        const uint OBJID_NATIVEOM = 0xFFFFFFF0;
                        Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");
                        Excel.Window window = null;
                        int res = AccessibleObjectFromWindow(childWindow, OBJID_NATIVEOM, IID_IDispatch.ToByteArray(), ref window);
                        if (res >= 0)
                        {
                            app = window.Application;
                            Console.WriteLine(app.Name);
                            try
                            {
                                workbook = app.Workbooks.get_Item(workbookname);
                                app.DisplayAlerts = false;
                                app.EnableEvents = false;
                                application = app;
                                if (sheetname == "")
                                {
                                    sheet = (Excel.Worksheet)workbook.ActiveSheet;
                                }
                                else
                                {
                                    sheet = (Excel.Worksheet)workbook.Worksheets[sheetname];
                                }

                                if (visible == "yes" || visible == "Yes" || visible == "YES")
                                {
                                    app.Visible = true;
                                }
                                else
                                {
                                    app.Visible = false;
                                }
                                return "Workbook found";
                            }
                            catch (Exception)
                            {

                            }



                        }
                    }
                }
            }
            workbook = null;
            sheet = null;
            application = null;
            return "Excel not found";
        }
        #region SKU COSTING SPECIFIC
        public string ActivateSheet(string workbookname, string visible, string sheetname)
        {
            try
            {

                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet, sheetname);
                sheet.Select();
                sheet.Activate();
                return "Workbook found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

    }
    
    
    class Program
    {
        static void Main(string[] args)
        {
            ExcelInstance obj = new ExcelInstance();
            List<string> output = new List<string>();

             obj.ActivateSheet("CEDC analysis_FR05.xlsb","Yes", " Intercompany");
        }
    }

}
