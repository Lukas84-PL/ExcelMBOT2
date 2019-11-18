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
using System.IO;
using System.Data.OleDb;
using System.Data;
//using System.Windows.Forms;

namespace ExcelDataManipulation
{
    //class ExcelInstance
    //{
    //    [DllImport("Oleacc.dll")]
    //    public static extern int AccessibleObjectFromWindow(
    //   int hwnd, uint dwObjectID, byte[] riid,
    //   ref Microsoft.Office.Interop.Excel.Window ptr);

    //    public delegate bool EnumChildCallback(int hwnd, ref int lParam);

    //    [DllImport("User32.dll")]
    //    public static extern bool EnumChildWindows(
    //    int hWndParent, EnumChildCallback lpEnumFunc,
    //    ref int lParam);


    //    [DllImport("User32.dll")]
    //    public static extern int GetClassName(
    //    int hWnd, StringBuilder lpClassName, int nMaxCount);

    //    public static bool EnumChildProc(int hwndChild, ref int lParam)
    //    {
    //        StringBuilder buf = new StringBuilder(128);
    //        GetClassName(hwndChild, buf, 128);
    //        if (buf.ToString() == "EXCEL7")
    //        {
    //            lParam = hwndChild;
    //            return false;
    //        }
    //        return true;
    //    }
    //    public string Instance(string workbookname, string visible, out Workbook workbook, out Application application, out Worksheet sheet, string sheetname = "")
    //    {
    //        Excel.Application app = new Excel.Application();
    //        EnumChildCallback cb;
    //        List<Process> procs = new List<Process>();
    //        procs.AddRange(Process.GetProcessesByName("excel"));

    //        foreach (Process p in procs)
    //        {
    //            if ((int)p.MainWindowHandle > 0)
    //            {
    //                int childWindow = 0;
    //                cb = new EnumChildCallback(EnumChildProc);
    //                EnumChildWindows((int)p.MainWindowHandle, cb, ref childWindow);

    //                if (childWindow > 0)
    //                {
    //                    const uint OBJID_NATIVEOM = 0xFFFFFFF0;
    //                    Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");
    //                    Excel.Window window = null;
    //                    int res = AccessibleObjectFromWindow(childWindow, OBJID_NATIVEOM, IID_IDispatch.ToByteArray(), ref window);
    //                    if (res >= 0)
    //                    {
    //                        app = window.Application;
    //                        Console.WriteLine(app.Name);
    //                        try
    //                        {
    //                            workbook = app.Workbooks.get_Item(workbookname);
    //                            app.DisplayAlerts = false;
    //                            app.EnableEvents = false;
    //                            application = app;

    //                            if (sheetname == "")
    //                            {
    //                                sheet = (Excel.Worksheet)workbook.ActiveSheet;
    //                            }
    //                            else
    //                            {
    //                                sheet = (Excel.Worksheet)workbook.Worksheets[sheetname];
    //                            }

    //                            if (visible == "yes" || visible == "Yes" || visible == "YES")
    //                            {
    //                                app.Visible = true;
    //                            }
    //                            else
    //                            {
    //                                app.Visible = false;
    //                            }
    //                            return "Workbook found";
    //                        }
    //                        catch (Exception)
    //                        {

    //                        }



    //                    }
    //                }
    //            }
    //        }
    //        workbook = null;
    //        sheet = null;
    //        application = null;
    //        return "Excel not found";
    //    }


    //}

    class Program
        {
        private Tuple<int?, int?, string, int?> ObtainExcelData(string workbookname, string newsheetname = "")
        {
            try
            {
                string visible = "Yes";
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                try
                {
                    xlapp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    xlapp.Visible = true;
                    workbook = xlapp.Workbooks.get_Item(workbookname);
                    sheet = (Excel.Worksheet)workbook.ActiveSheet;
                }
                catch (Exception)
                {

                    //workbook = null;
                    //xlapp = null;
                    //sheet = null;
                    //ExcelInstance instance = new ExcelInstance();
                    //instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                }
                xlapp.Visible = true;
                xlapp.DisplayAlerts = false;
                xlapp.EnableEvents = false;


                string worksheetname = "Not found";
                int? rowcount = 0;
                int? columncount = 0;
                int? worksheetcount = 0;


                //Get worksheet name
                worksheetname = sheet.Name;
                //Get last used column count
                columncount = sheet.UsedRange.Columns.Count;
                //Get last used row count
                rowcount = sheet.UsedRange.Rows.Count;
                //Count number of sheets in workbook
                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    worksheetcount++;
                }
                //If new name provided then change sheet name
                if (newsheetname != "")
                {
                    sheet.Name = newsheetname;
                }


                return Tuple.Create(columncount, rowcount, worksheetname, worksheetcount);

            }
            catch (Exception e)
            {
                int? value = null;
                int? value2 = null;
                int? value3 = null;
                string exc = "<#EXCEL INTEGRATION MBOT FAILED#> " + e.ToString();
                return Tuple.Create(value, value2, exc, value3);
            }
        }

        #region FILTER OUT VALUES
        public string FilterOutValues(string workbookname, string visible, int columnfrom, int rowfrom, int columnto, int rowto, int filtercolumn, string[] filterlist)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                try
                {
                    xlapp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    xlapp.Visible = true;
                    workbook = xlapp.Workbooks.get_Item(workbookname);
                    sheet = (Excel.Worksheet)workbook.ActiveSheet;
                }
                catch (Exception)
                {

                    //workbook = null;
                    //xlapp = null;
                    //sheet = null;
                    //ExcelInstance instance = new ExcelInstance();
                    //instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                }
                xlapp.Visible = true;
                //xlapp.DisplayAlerts = false;
                //xlapp.EnableEvents = false;
                workbook.Activate();
                sheet.Activate();
                //sheet.Select();

                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];


                Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                range.AutoFilter(filtercolumn, filterlist,
                            Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                return "Workbook found";


            }
            catch (Exception e)
            {
                return "<#EXCEL INTEGRATION MBOT FAILED#> " + e.ToString();
            }
        }
        #endregion

        static void Main(string[] args)
            {

            Program obj = new Program();
            //string[] rowarray = new string[] { "Purchase Org.", "Vendor", "Material" };
            //string[] nosubs = new string[] { "Purchase Org.", "Vendor", "Material" };
            // rowarray[0] = "a";
            // rowarray[1] = "c";
            //obj.OpenSpreadsheet("from.xlsx", "Yes", @"C:\Users\LXB0906\Desktop\sheetcopy\");
            //obj.OpenSpreadsheet("to.xlsx", "Yes", @"C:\Users\LXB0906\Desktop\sheetcopy\");
            //Console.WriteLine(obj.LoopThroughRowsAndInsertData("BE11VATReport1.xlsx","Yes",14,1,14999, "NL821496098B01","VAT correct",36));
            string[] rowarray = new string[] { "Howard, George", "Zhang, Jennifer" };
            obj.FilterOutValues("PSF Report_10212019.xlsx", "yes",1,1,25,4400,7, rowarray);

            }
        }

    
}
