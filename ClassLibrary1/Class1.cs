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
        public string Instance(string workbookname,string visible,out Workbook workbook, out Application application, out Worksheet sheet, string sheetname = "")
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


    }
    public class ExcelMBOT 
    {
       
        #region ACTIONS
        #region SAVEAS WORKBOOK
        public string SaveAs(string workbookname, string visible, string newfilenamefullpath = "")
        {
            try
            {
                string status = "Failed";
                Application xlapp = (Application)Marshal.GetActiveObject("Excel.Application");
                xlapp.DisplayAlerts = false;
                xlapp.EnableEvents = false;
                Workbook workbook = xlapp.Workbooks.get_Item(workbookname);
                if (visible == "yes" || visible == "Yes" || visible == "YES")
                {
                    xlapp.Visible = true;
                }
                else
                {
                    xlapp.Visible = false;
                }
                workbook.SaveAs(newfilenamefullpath);

                foreach (Excel.Workbook wb in xlapp.Workbooks)
                {
                    if ((workbook.Path + @"\" + workbook.Name) == newfilenamefullpath)
                    {
                        status = "Completed";
                    }
                }
            return status;
            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region SAVE AS CSV
        public string SaveAsCSV(string workbookname, string visible, string newfilenamefullpath = "")
        {
            try
            {
                string status = "Failed";
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);

                workbook.SaveAs(newfilenamefullpath, XlFileFormat.xlCSV);

                foreach (Excel.Workbook wb in xlapp.Workbooks)
                {
                    if ((workbook.Path + @"\" + workbook.Name) == newfilenamefullpath)
                    {
                        status = "Completed";
                    }
                }
                return status;
            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region SELECT RANGE
        public string SelectRange(string workbookname, string visible, int columnfrom, int rowfrom, int columnto, int rowto)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);

                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];

                //Select cells in a range
                    
                Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                range.Select();
                range.Activate();
                return "Workbook found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region INSERT STRING IN CELL
        public string InsertDataInCell(string workbookname, string visible, int column, int row, string insertdata)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);

                //Select cells in a range
                sheet.Cells[row, column] = insertdata;

                return "Workbook found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region GET VALUE OF CELL
        public string GetValueOfCell(string workbookname, string visible, int column, int row)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);

                return sheet.Cells[row, column].text;

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region COPY RANGE
        public string CopyRange(string workbookname, string visible, int columnfrom, int rowfrom, int columnto, int rowto)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];

                //Select cells in a range

                Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                range.Copy();
                return "Workbook found";


            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region COPY SELECTION
        public string CopySelection(string workbookname, string visible)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cells in a range
                Excel.Range range = xlapp.ActiveWindow.Selection;
                range.Copy();
                return "Workbook found";


            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region COPY COLUMN
        public string CopyColumn(string workbookname, string visible, string column)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cells in a range
                Excel.Range range = (Excel.Range)sheet.Columns[column + ":" + column];
                range.Copy();
                return "Workbook found";


            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region COPY ROW
        public string CopyRow(string workbookname, string visible, int row)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cells in a range
                Excel.Range range = (Excel.Range)sheet.Rows[row + ":" + row];
                range.Copy();
                return "Workbook found";


            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region COPY AND PASTE VALUES OF COLUMN
        public string CopyAndPasteValuesOfColumn(string workbookname, string visible, string columnfrom, string columnto)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cells in a range
                Excel.Range rangefrom = (Excel.Range)sheet.Columns[columnfrom + ":" + columnfrom];
                rangefrom.Copy();
                Excel.Range rangeto = (Excel.Range)sheet.Columns[columnto + ":" + columnto];
                rangeto.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                return "Workbook found";


            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region COPY AND PASTE VALUES OF ROW
        public string CopyAndPasteValuesOfRow(string workbookname, string visible, int rowfrom, int rowto)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cells in a range
                Excel.Range rangefrom = (Excel.Range)sheet.Rows[rowfrom + ":" + rowfrom];
                rangefrom.Copy();
                Excel.Range rangeto = (Excel.Range)sheet.Rows[rowto + ":" + rowto];
                rangeto.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                return "Workbook found";
 

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region COPY AND PASTE COLUMN
        public string CopyAndPasteColumn(string workbookname, string visible, string columnfrom, string columnto)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cells in a range
                Excel.Range rangefrom = (Excel.Range)sheet.Columns[columnfrom + ":" + columnfrom];
                rangefrom.Copy();
                Excel.Range rangeto = (Excel.Range)sheet.Columns[columnto + ":" + columnto];
                rangeto.PasteSpecial();
                return "Workbook found";


            }
            catch (Exception e)
            {
                return e.ToString();
            }
}
        #endregion
        #region COPY AND PASTE ROW
        public string CopyAndPasteRow(string workbookname, string visible, int rowfrom, int rowto)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cells in a range
                Excel.Range rangefrom = (Excel.Range)sheet.Rows[rowfrom + ":" + rowfrom];
                rangefrom.Copy();
                Excel.Range rangeto = (Excel.Range)sheet.Rows[rowto + ":" + rowto];
                rangeto.PasteSpecial();
                return "Workbook found";
                
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion

        #region INSERT FORMULA
        public string InsertFormula(string workbookname, string visible, int column, int row, string inputformula)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cell
                Excel.Range range = sheet.Cells[row, column];
                //Enter forumla
                range.Formula = inputformula;
                return "Workbook found";
               

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region REPLACE DATA IN RANGE
        public string ReplaceDataInRange(string workbookname, string visible, int columnfrom, int rowfrom, int columnto, int rowto, string oldstring, string newstring)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];
                //Select cells in a range
                Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                range.Replace(oldstring, newstring);
                return "Workbook found";
            
            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region REPLACE DATA IN SELECTION
        public string ReplaceDataInSelection(string workbookname, string visible, string oldstring, string newstring)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cells in a range
                Excel.Range range = xlapp.ActiveWindow.Selection;
                range.Replace(oldstring, newstring);
                return "Workbook found";
              

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region CHANGE FONT IN RANGE TO BOLD
        public string ChangeFontInRangeToBold(string workbookname, string visible, int columnfrom, int rowfrom, int columnto, int rowto)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];

                //Select cells in a range
                Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                range.Font.Bold = true;
                return "Workbook found";
             

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region CHANGE FONT IN SELECTION TO BOLD
        public string ChangeFontInSelectionToBold(string workbookname, string visible)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cells in a range
                Excel.Range range = xlapp.ActiveWindow.Selection;
                range.Font.Bold = true;
                return "Workbook found";
              

            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion

        #region PASTE VALUES IN SELECTION
        public string PasteValuesInSelection(string workbookname, string visible)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cells in a range
                Excel.Range range = xlapp.ActiveWindow.Selection;
                range.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                return "Workbook found";
              

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region PASTE VALUES IN CELL
        public string PasteValuesInCell(string workbookname, string visible, int column, int row)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cell
                Excel.Range range = sheet.Cells[row, column];
                //Enter forumla
                range.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                return "Workbook found";
              
            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region PASTE IN SELECTION
        public string PasteInSelection(string workbookname, string visible)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cells in a range
                Excel.Range range = xlapp.ActiveWindow.Selection;
                range.PasteSpecial();
                return "Workbook found";
             

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region PASTE IN CELL
        public string PasteInCell(string workbookname, string visible, int column, int row)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cell
                Excel.Range range = sheet.Cells[row, column];
                //Enter forumla
                range.PasteSpecial();
                return "Workbook found";
                

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region SELECT BLANK CELLS OF SELECTION
        public string SelectBlankCellsOfSelection(string workbookname, string visible)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Create object from selection
                Excel.Range range = xlapp.ActiveWindow.Selection;
                //Select blank cells in a range
                Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                //Delete blank cells rows
                blankcells.Select();
                return "Workbook found";

              

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region SELECT BLANK CELLS OF RANGE
        public string SelectBlankCellsOfRange(string workbookname, string visible, int column, int rowfrom, int rowto)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, column];
                Excel.Range rngTo = xlapp.Cells[rowto, column];

                //Select cells in a range
                Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                //Select blank cells in a range
                Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                //Delete blank cells rows
                blankcells.Select();
                return "Workbook found";
               
            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region SELECT BLANK CELLS OF COLUMN
        public string SelectBlankCellsOfColumn(string workbookname, string visible, string column)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select column
                Excel.Range range = (Excel.Range)sheet.Columns[column + ":" + column];
                //Select blank cells in a range
                Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                //Delete blank cells rows
                blankcells.Select();
                return "Workbook found";

               

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region SELECT BLANK CELLS OF ROW
        public string SelectBlankCellsOfRow(string workbookname, string visible, int row)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select column
                Excel.Range range = (Excel.Range)sheet.Rows[row + ":" + row];
                //Select blank cells in a range
                Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                //Delete blank cells rows
                blankcells.Select();
                return "Workbook found";
               

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region DELETE BLANK ROWS OF SELECTION
        public string DeleteBlankRowsInSelection(string workbookname, string visible)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Create object from selection
                Excel.Range range = xlapp.ActiveWindow.Selection;
                //Select blank cells in a range
                Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                //Delete blank cells rows
                blankcells.EntireRow.Delete();
                return "Workbook found";
              

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region DELETE BLANK ROWS OF RANGE
        public string DeleteBlankRowsOfRange(string workbookname, string visible, int column, int rowfrom, int rowto)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, column];
                Excel.Range rngTo = xlapp.Cells[rowto, column];

                //Select cells in a range
                Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                //Select blank cells in a range
                Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                //Delete blank cells rows
                blankcells.EntireRow.Delete();
                return "Workbook found";
                

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region DELETE BLANK ROWS OF COLUMN
        public string DeleteBlankRowsOfColumn(string workbookname, string visible, string column)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select column
                Excel.Range range = (Excel.Range)sheet.Columns[column + ":" + column];
                //Select blank cells of column
                Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                //Delete blank cells rows
                blankcells.EntireRow.Delete();
                return "Workbook found";
               

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region DELETE BLANK COLUMNS IN A SELECTION
        public string DeleteBlankColumnsOfSelection(string workbookname, string visible)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cells in a range
                Excel.Range range = xlapp.ActiveWindow.Selection;
                //Select blank cells in a range
                Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                //Delete blank cells columns
                blankcells.EntireColumn.Delete();
                return "Workbook found";
                

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region DELETE BLANK COLUMNS IN A RANGE
        public string DeleteBlankColumnsOfRange(string workbookname, string visible, int columnfrom, int columnto, int row)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[row, columnfrom];
                Excel.Range rngTo = xlapp.Cells[row, columnto];

                //Select cells in a range
                Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                //Select blank cells in a range
                Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                //Delete blank cells columns
                blankcells.EntireColumn.Delete();
                return "Workbook found";
               

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region DELETE BLANK COLUMNS OF ROW
        public string DeleteBlankColumnsOfRow(string workbookname, string visible, int row)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cells in a range
                Excel.Range range = (Excel.Range)sheet.Rows[row + ":" + row];
                //Select blank cells in a range
                Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                //Delete blank cells columns
                blankcells.EntireColumn.Delete();
                return "Workbook found";
               

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion


        #region DELETE ROWS OF SELECTED CELLS
        public string DeleteRowsOfSelectedCells(string workbookname, string visible)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Create object from selection
                Excel.Range range = xlapp.ActiveWindow.Selection;
                //Delete selected rows
                range.EntireRow.Delete();
                return "Workbook found";
               

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region CLEAR SELECTED CELLS
        public string ClearSelectedCells(string workbookname, string visible)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Create object from selection
                Excel.Range range = xlapp.ActiveWindow.Selection;
                //Delete selected rows
                range.ClearContents();
                return "Workbook found";
             

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region DELETE COLUMNS OF SELECTED CELLS
        public string DeleteColumnsOfSelectedCells(string workbookname, string visible)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Create object from selection
                Excel.Range range = xlapp.ActiveWindow.Selection;
                //Delete selected rows
                range.EntireColumn.Delete();
                return "Workbook found";
               

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region CLEAR CELLS IN RANGE
        public string ClearCellsInRange(string workbookname, string visible, int columnfrom, int rowfrom, int columnto, int rowto)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];

                //Create object from selection
                Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                //Delete selected rows
                range.ClearContents();
                return "Workbook found";
               

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region SELECT ROW
        public string SelectEntireRow(string workbookname, string visible, int row)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cells in a range
                Excel.Range range = (Excel.Range)sheet.Rows[row + ":" + row];
                //Select row
                range.Select();
                return "Workbook found";
               

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region SELECT COLUMN
        public string SelectEntireColumn(string workbookname, string visible, string column)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select column
                Excel.Range range = (Excel.Range)sheet.Columns[column + ":" + column];
                //Delete blank cells rows
                range.Select();
                return "Workbook found";
            

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region INSERT ROW IN SELECTION
        public string InsertRowInSelection(string workbookname, string visible)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cells in a range
                Excel.Range range = xlapp.ActiveWindow.Selection;
                //Delete blank cells rows
                range.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                return "Workbook found";
               

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region INSERT ROW IN RANGE
        public string InsertRowInRange(string workbookname, string visible, int row)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cells in a range
                Excel.Range range = (Excel.Range)sheet.Rows[row + ":" + row];
                //Delete blank cells rows
                range.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                return "Workbook found";
               

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region INSERT COLUMN IN SELECTION
        public string InsertColumnInSelection(string workbookname, string visible)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select column
                Excel.Range range = xlapp.ActiveWindow.Selection;
                //Delete blank cells rows
                range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
                return "Workbook found";
               

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region INSERT COLUMN IN RANGE
        public string InsertColumnInRange(string workbookname, string visible, string column)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select column
                Excel.Range range = (Excel.Range)sheet.Columns[column + ":" + column];
                //Delete blank cells rows
                range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
                return "Workbook found";
               

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region CHANGE SHEET NAME
        public string ChangeSheetName(string workbookname, string newsheetname)
        {
            try
            {

            ExcelMBOT obj = new ExcelMBOT();


            var result = obj.ObtainExcelData(workbookname, newsheetname);
            return "Finished";

            }
            catch (Exception e)
            {

                return e.ToString();
            }

        }
        #endregion

        #region GET SHEET NAME
        public string GetSheetName(string workbookname)
        {
            try
            {

            ExcelMBOT obj = new ExcelMBOT();

            var result = obj.ObtainExcelData(workbookname);

            return result.Item3;

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region ACTIVATE SHEET
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
        #region GET SHEET COUNT
        public string GetSheetCount(string workbookname)
        {
            try
            {

            ExcelMBOT obj = new ExcelMBOT();

            var result = obj.ObtainExcelData(workbookname);

            return result.Item4.ToString();

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region GET ROW COUNT
        public string GetRowsCount(string workbookname)
        {
            try
            {

            ExcelMBOT obj = new ExcelMBOT();

            var result = obj.ObtainExcelData(workbookname);

            return result.Item2.ToString();

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region GET COLUMN COUNT
        public string GetColumnsCount(string workbookname)
        {
            try
            {

            ExcelMBOT obj = new ExcelMBOT();

            var result = obj.ObtainExcelData(workbookname);

            return result.Item1.ToString();

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region GET LAST ROW OF SPECIFIC COLUMN
        public string GetLastRowOfSpecificColumn(string workbookname, string visible, int column, int rowstart)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);


                while (sheet.Cells[rowstart +1, column].value != null || sheet.Cells[rowstart +2, column].value != null)
                    {
                        ++rowstart;
                    }

                
            
            return rowstart.ToString();

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region GET LAST COLUMN OF SPECIFIC ROW
        public string GetLastColumnOfSpecificRow(string workbookname, string visible, int row, int columnstart)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select column

                while (sheet.Cells[row, columnstart +1].value != null || sheet.Cells[row, columnstart +2].value != null)
                    {
                        ++columnstart;
                    }

     
            return columnstart.ToString();

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region AUTOFIT COLUMN
        public string AutofitColumn(string workbookname, string visible, string column)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select column
                Excel.Range range = (Excel.Range)sheet.Columns[column + ":" + column];
                //Autofit column
                range.AutoFit();
                return "Workbook found";
               


            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region AUTOFIT ROW
        public string AutofitRow(string workbookname, string visible, int row)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Select cells in a range
                Excel.Range range = (Excel.Range)sheet.Rows[row + ":" + row];
                //Autofit row
                range.AutoFit();
                return "Workbook found";
               

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region DELETE SHEET
        public string DeleteSheet(string workbookname, string visible)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Delete sheet
                sheet.Delete();
                return "Workbook found";
               

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region COPY SHEET TO NEW WORKBOOK
        public string CopySheetToNewWorkbook(string workbookname, string visible)
        {
            try
            {


                List<string> workbooklist = new List<string>();
                Excel.Workbook wkb = null;
                string wkbname;
                string newworkbook = "Unable to find new workbook name";
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Find all open workbooks and add their names to the list
                foreach (Excel.Workbook activewkbs in xlapp.Workbooks)
                {
                    wkbname = Convert.ToString(activewkbs.Name);
                    workbooklist.Add(wkbname);
                    if (activewkbs.Name == workbookname)
                    {
                        wkb = activewkbs;
                    }
                }

                if (wkb != null)
                {

                //Copy sheet
                    sheet.Copy();
                    //Obtain name of newly created worksheet
                    foreach (Excel.Workbook wkbnew in xlapp.Workbooks)
                    {
                        wkbname = Convert.ToString(wkbnew.Name);
                        if (workbooklist.Contains(wkbname) == false)
                        {
                            newworkbook = wkbname;
                        }
                }

                return newworkbook;
            }
            else
            {
                return "Wokbook not available";
            }

            }
            catch (Exception e)
            {
                return e.ToString();
            }

        }
        #endregion
        #region COPY SHEET TO SPECIFIC WORKBOOK
        public string CopySheetToSpecificWorkbook(string workbooknameFROM, string workbooknameTO, string visible)
        {
            try
            {
                Excel.Workbook wkbFrom = null;
                Excel.Workbook wkbTo = null;
                Application xlappFrom = null;
                Application xlappTo = null;
                Worksheet sheetFrom = null;
                Worksheet sheetTo = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbooknameFROM, visible, out wkbFrom, out xlappFrom, out sheetFrom);
                instance.Instance(workbooknameTO, visible, out wkbTo, out xlappTo, out sheetTo);
              

                if (wkbFrom != null && wkbTo != null)
                {

                    sheetFrom.Copy(wkbTo.Worksheets[1]);

                    return "Object copied";
                }
                else
                {
                    return "Workbooks not available";
                }


            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion

        #region CREATE PIVOT IN NEW TAB
        public string CreatePivotTableInNewTab(string workbookname, string visible, string columnfrom, int rowfrom, string columnto, int rowto, string newsheetname, string pivotname, string pivotcelldesticnation, string[] rowfieldlist = null, string[] columnfieldlist = null, string[] valuefieldlist = null, string[] filterfieldlist = null)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);

                Excel.PivotCaches pCaches = workbook.PivotCaches();
                string worksheetname = sheet.Name;
                string pivotdata = worksheetname + "!" + columnfrom + rowfrom + ":" + columnto + rowto;
                Excel.Worksheet sheet2 = workbook.Worksheets.Add();
                sheet2.Name = newsheetname;

                Excel.Range rngDes = sheet2.get_Range(pivotcelldesticnation);
                //Send range to cache and use it to create pivot
                Excel.PivotCache cache = pCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, pivotdata, Excel.XlPivotTableVersionList.xlPivotTableVersion14);
                //Excel.PivotTable pTable = cache.CreatePivotTable(TableDestination: rngDes, TableName: pivotname, DefaultVersion: Excel.XlPivotTableVersionList.xlPivotTableVersion14);
                Excel.PivotTable pTable = cache.CreatePivotTable(TableDestination: rngDes, TableName: "PivotTable1", DefaultVersion: Excel.XlPivotTableVersionList.xlPivotTableVersion14);


                if (rowfieldlist != null)
                {
                    foreach (var rowfield in rowfieldlist)
                    {
                        pTable.PivotFields(rowfield).Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                    }
                }

                if (columnfieldlist != null)
                {
                    foreach (var columnfield in columnfieldlist)
                    {
                        pTable.PivotFields(columnfield).Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
                    }
                }

                if (valuefieldlist != null)
                {
                    foreach (var valuefield in valuefieldlist)
                    {
                        pTable.PivotFields(valuefield).Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                    }
                }

                if (filterfieldlist != null)
                {
                    foreach (var filterfield in filterfieldlist)
                    {
                        pTable.PivotFields(filterfield).Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                    }
                }


                return pTable.Name + " created";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region FILTER ON VALUE PIVOT IN SELECTED SHEET
        public string FilterOnValuePivotInSelectedSheet(string workbookname, string visible, string pivotname, string[] filtervalues, string filterfield)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);

                Excel.PivotCaches pCaches = workbook.PivotCaches();

                Excel.PivotTable pivot = (Excel.PivotTable)sheet.PivotTables(pivotname);
                Excel.PivotField pivotfield = pivot.PivotFields(filterfield);
                pivotfield.ClearAllFilters();
                int count = pivot.PivotFields(1).PivotItems.Count;
                for (int i = 1; i <= count; i++)
                {
                    if (Array.IndexOf(filtervalues, pivotfield.PivotItems(i).Name) > -1)
                    {
                        pivotfield.PivotItems(i).visible = true;
                    }
                    else
                    {
                        pivotfield.PivotItems(i).visible = false;
                    }

                    return "Completed";
                }
            return "Failed";

            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region FILTER OUT VALUE PIVOT IN SELECTED SHEET
        public string FilterOutValuePivotInSelectedSheet(string workbookname, string visible, string pivotname, string[] filtervalues, string filterfield)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.PivotCaches pCaches = workbook.PivotCaches();

                Excel.PivotTable pivot = (Excel.PivotTable)sheet.PivotTables(pivotname);
                Excel.PivotField pivotfield = pivot.PivotFields(filterfield);
                pivotfield.ClearAllFilters();
                int count = pivot.PivotFields(1).PivotItems.Count;
                for (int i = 1; i <= count; i++)
                {
                    if (Array.IndexOf(filtervalues, pivotfield.PivotItems(i).Name) > -1)
                    {
                        pivotfield.PivotItems(i).visible = false;
                    }
                    else
                    {
                        pivotfield.PivotItems(i).visible = true;
                    }
                    return "Completed";
                }

            return "Failed";

            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion

        #region FILTER ON VALUES
        public string FilterOutValues(string workbookname, string visible, int columnfrom, int rowfrom, int columnto, int rowto, int filtercolumn, string[] filterlist)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];


                Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                range.AutoFilter(filtercolumn, filterlist,
                            Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                return "Workbook found";


            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region CHANGE RANGE FONT COLOR USING HTML
        public string ChangeRangeFontColor(string workbookname, string visible, int columnfrom, int rowfrom, int columnto, int rowto, string color)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];

                Excel.Range range = sheet.get_Range(rngFrom, rngTo);

                range.Font.Color = System.Drawing.ColorTranslator.FromHtml(color);
                return "Workbook found";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region CHANGE RANGE FILL COLOR USING HTML
        public string ChangeRangeFillColorHTML(string workbookname, string visible, int columnfrom, int rowfrom, int columnto, int rowto, string color)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];

                Excel.Range range = sheet.get_Range(rngFrom, rngTo);

                range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(color);
                return "Workbook found";

            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region SET TEXT TO COLUMNS OF RANGE
        public string TextToColumns(string workbookname, string visible, int columnfrom, int rowfrom, int columnto, int rowto, string delimiter)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];

                Excel.Range MyRange = sheet.get_Range(rngFrom, rngTo);

                MyRange.TextToColumns(MyRange,
                XlTextParsingType.xlDelimited,
                XlTextQualifier.xlTextQualifierDoubleQuote,
                true,        // Consecutive Delimiter
                Type.Missing,// Tab
                Type.Missing,// Semicolon
                false,        // Comma
                false,       // Space
                true,// Other
                delimiter,         // Other Char
                Type.Missing,// Field Info
                Type.Missing,// Decimal Separator
                Type.Missing,// Thousands Separator
                Type.Missing);// Trailing Minus Numbers
                return "Workbook found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region CHANGE FORMAT OF RANGE
        public string ChangeFormatOfRange(string workbookname, string visible, int columnfrom, int rowfrom, int columnto, int rowto, string format)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];

                if (format == "@")
                {
                    Array fieldInfoArray = new int[,] { { 1, 2 } };
                    //Select cells in a range

                    Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                    range.NumberFormat = format;
                    range.TextToColumns(range,
                    XlTextParsingType.xlDelimited,
                    XlTextQualifier.xlTextQualifierDoubleQuote,
                    true,        // Consecutive Delimiter
                    Type.Missing,// Tab
                    Type.Missing,// Semicolon
                    false,        // Comma
                    false,       // Space
                    false,// Other
                    "",         // Other Char
                    fieldInfoArray,// Field Info
                    Type.Missing,// Decimal Separator
                    Type.Missing,// Thousands Separator
                    Type.Missing);// Trailing Minus Numbers
                }
                else
                {


                    Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                    range.TextToColumns();
                    range.NumberFormat = format;
                    range.TextToColumns();
                }
                return "Workbook found";


            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region SELECT VISIBLE CELLS IN RANGE
        public string SelectVisibleCellsInRange(string workbookname, string visible, int columnfrom, int rowfrom, int columnto, int rowto)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];

                Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                range.SpecialCells(XlCellType.xlCellTypeVisible).Select();
                return "Workbook found";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region SORT RANGE
        public string SortRange(string workbookname, string visible, int columnfrom, int rowfrom, int columnto, int rowto, int sortcolumnnumber)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];
                Excel.Range rngsortFrom = xlapp.Cells[rowfrom, sortcolumnnumber];
                Excel.Range rngsortTo = xlapp.Cells[rowto, sortcolumnnumber];

                //Select cells in a range

                Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                Excel.Range range2 = sheet.get_Range(rngsortFrom, rngsortTo);
                range.Sort(range2,
                            XlSortOrder.xlAscending,
                            Type.Missing, Type.Missing,
                            XlSortOrder.xlAscending,
                            Type.Missing,
                            XlSortOrder.xlAscending,
                            XlYesNoGuess.xlYes,
                            Type.Missing,
                            Type.Missing,
                            XlSortOrientation.xlSortColumns,
                            XlSortMethod.xlPinYin,
                            XlSortDataOption.xlSortNormal,
                            XlSortDataOption.xlSortNormal,
                            XlSortDataOption.xlSortNormal);

                return "Workbook found";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region OPEN SPREADSHEET
        public string OpenSpreadsheet(string workbookname, string visible, string path)
        {
            try
            {

                Application xlapp = new Application();
                xlapp.DisplayAlerts = false;
                xlapp.EnableEvents = false;
                if (visible == "yes" || visible == "Yes" || visible == "YES")
                {
                    xlapp.Visible = true;
                }
                else
                {
                    xlapp.Visible = false;
                }
                xlapp.Workbooks.Open( path + workbookname ,false,false);
                return "Workbook found";
            }
            catch (Exception e)
            {
                return e.ToString();
            }

        }
        #endregion
        #region INSERT OBJECT
        public string InsertObject(string workbookname, string visible, int column, int row, string filepath, string iconname, int iconindex, int iconwidth = 10)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range range = sheet.Cells[row, column];

                var obj = xlapp.ActiveSheet.OLEObjects.Add(Filename: filepath,
                    Link: false,
                    DisplayAsIcon: true,
                    IconFileName: iconname,
                    IconIndex: iconindex,
                    IconLabel: iconname,
                    Left: range.Left,
                    Top: range.Top,
                    Width: iconwidth,
                    Height: 10);


                obj.Width = iconwidth;
                return "Workbook found";


            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region DRAG CELL VALUE TO RANGE
        public string DragCellValueToRange(string workbookname, string visible, int column, int rowfrom, int rowto)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, column];
                Excel.Range rngTo = xlapp.Cells[rowto, column];


                Excel.Range rng = xlapp.get_Range(rngFrom, rngFrom);

                rng.AutoFill(xlapp.get_Range(rngFrom, rngTo),
                Excel.XlAutoFillType.xlFillWeekdays);
                return "Workbook found";

            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region FILTER OUT VALUES
        public string FilterOnValues(string workbookname, string visible, int columnfrom, int rowfrom, int columnto, int rowto, int filtercolumn, string[] filterlist)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];

                Excel.Range range = sheet.get_Range(rngFrom, rngTo);

                foreach (var item in filterlist)
                {
                    range.AutoFilter(filtercolumn, "<>" + item,
                    Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                }
                return "Workbook found";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region REMOVE DUPLICATES FROM RANGE
        public string RemoveDuplicatesFromRange(string workbookname, string visible, int columnfrom, int rowfrom, int columnto, int rowto)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];

                Range range = sheet.get_Range(rngFrom, rngTo);

                range.RemoveDuplicates(
                Excel.XlYesNoGuess.xlNo);
                return "Workbook found";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region REMOVE DUPLICATES FROM COLUMNS
        public string RemoveDuplicatesFromColumns(string workbookname, string visible, string columnfrom, string columnto)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);

                Excel.Range range = (Excel.Range)sheet.Columns[columnfrom + ":" + columnto];

                range.RemoveDuplicates(
                Excel.XlYesNoGuess.xlNo);
                return "Workbook found";

            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region SELECT VALUE IN RANGE
        public string SelectValueInRange(string workbookname, string visible, int columnfrom, int rowfrom, int columnto, int rowto, string searchvalue)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];


                Range range = sheet.get_Range(rngFrom, rngTo);

                range.Find(searchvalue).Select();
                return "Workbook found";


            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region FIND VALUE IN RANGE AND GET ADRESS
        public string GetAdressOfValue(string workbookname, string visible, int columnfrom, int rowfrom, int columnto, int rowto, string searchvalue)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];

                Range range = sheet.get_Range(rngFrom, rngTo);

                Range adressrange = range.Find(searchvalue);

                return adressrange.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region REPLACE VALUES IN RANGE
        public string ReplaceValuesInRange(string workbookname, string visible, int columnfrom , int rowfrom, int columnto , int rowto, string replace, string replaceto)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
                Excel.Range rngTo = xlapp.Cells[rowto, columnto];

                Range range = sheet.get_Range(rngFrom, rngTo);

                sheet.Cells.Replace(replace, replaceto);
                return "Workbook found";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region GO TO LAST ROW OF SPECIFIC COLUMN
        public string GoToLastRowOfSpecificColumn(string workbookname, string visible, int column, int rowstart)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);

                while (sheet.Cells[rowstart +1, column].value != null || sheet.Cells[rowstart +2, column].value != null)
                {
                    ++rowstart;
                }

                Excel.Range lastcell = xlapp.Cells[rowstart, column];

                lastcell.Activate();
                lastcell.Select();

                return "Workbook found";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region GO TO LAST COLUMN OF SPECIFIC ROW
        public string GoToLastColumnOfSpecificRow(string workbookname, string visible, int columnstart, int row)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);

                while (sheet.Cells[row, columnstart + 1].value != null || sheet.Cells[row, columnstart + 2].value != null)
                {
                    ++columnstart;
                }

                Excel.Range lastcell = xlapp.Cells[row, columnstart];
                lastcell.Activate();
                lastcell.Select();
                return "Workbook found";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region GO TO LAST COLUMN OF USED RANGE
        public string GoToLastColumnOfUsedRange(string workbookname, string visible, int row)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                int columncount = 0;

                columncount = sheet.UsedRange.Columns.Count;
                Range lastcolumn = xlapp.Cells[row, columncount];
                lastcolumn.Select();
                lastcolumn.Activate();
                return "Workbook found";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region GO TO LAST ROW OF USED RANGE
        public string GoToLastRowOfUsedRange(string workbookname, string visible, int column)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                int rowcount = 0;
                rowcount = sheet.UsedRange.Rows.Count;
                Range lastrow = xlapp.Cells[rowcount, column];
                lastrow.Select();
                lastrow.Activate();
                return "Workbook found";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region CLOSE SPREADSHEET WITH SAVING
        public string CloseSpreadsheetWithSaving(string workbookname, string visible)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                workbook.Close(true);
                return "Workbook found";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region CLOSE SPREADSHEET WITHOUT SAVING
        public string CloseSpreadsheetWithoutSaving(string workbookname, string visible)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                workbook.Close(false);
                return "Workbook found";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region QUIT EXCEL APP
        public string QuitExcelApp()
        {
            try
            {
                Application xlapp = (Application)Marshal.GetActiveObject("Excel.Application");
                xlapp.DisplayAlerts = false;

                xlapp.Quit();
                return "Workbook found";
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region CREATE EXCEL WORKBOOK
        public string CreateExcelWorkbook(string workbookname, string visible)
        {
            try
            {
                Excel.Application xlapp = new Application();
                xlapp.DisplayAlerts = false;

                Workbook newWorkbook = xlapp.Application.Workbooks.Add();
                if (visible == "yes" || visible == "Yes" || visible == "YES")
                {
                    xlapp.Visible = true;
                }
                else
                {
                    xlapp.Visible = false;
                }
                newWorkbook.SaveAs(workbookname);
                newWorkbook.Close();
                return "Workbook created";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region LOOP THROUGH ALL ROWS IN 1 COLUMN AND SET 3 CELLS
        public string Search1ValuesIn1ColumnsSet3Cells(string workbookname, string visible, int loopcolumn, int startrow, string searchvalue1, int searchcolumn1, int setcolumn1, string setvalue1, int setcolumn2, string setvalue2, int setcolumn3, string setvalue3)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                //Workbook workbook = xlapp.Workbooks.get_Item(workbookname);

                List<string> returnlist = new List<string> { };
                //Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;

                while (sheet.Cells[startrow, loopcolumn].value != null || sheet.Cells[startrow + 1, loopcolumn].value != null || sheet.Cells[startrow + 2, loopcolumn].value != null)
                {
                    string s1 = sheet.Cells[startrow, searchcolumn1].text;

                    if (s1.Contains(searchvalue1))
                    {
                        if (setvalue1 != "")
                        {
                            sheet.Cells[startrow, setcolumn1].value = setvalue1;
                        }
                        if (setvalue2 != "")
                        {
                            sheet.Cells[startrow, setcolumn2].value = setvalue2;
                        }
                        if (setvalue3 != "")
                        {
                            sheet.Cells[startrow, setcolumn3].value = setvalue3;
                        }

                    }
                    ++startrow;
                }


                return "Excel found";

            }
            catch (Exception e)
            {
                string exc = e.ToString();
                return exc;
            }
        }
        #endregion
        #region LOOP THROUGH ALL ROWS IN 2 COLUMNS AND SET 3 CELLS
        public string Search2ValuesIn2ColumnsSet3Cells(string workbookname, string visible, int loopcolumn, int startrow, string searchvalue1, int searchcolumn1, string searchvalue2, int searchcolumn2, int setcolumn1, string setvalue1, int setcolumn2, string setvalue2, int setcolumn3, string setvalue3)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);

                List<string> returnlist = new List<string> { };

                while (sheet.Cells[startrow, loopcolumn].value != null || sheet.Cells[startrow + 1, loopcolumn].value != null || sheet.Cells[startrow + 2, loopcolumn].value != null)
                {
                    string s1 = sheet.Cells[startrow, searchcolumn1].text;
                    string s2 = sheet.Cells[startrow, searchcolumn2].text;

                    if (s1.Contains(searchvalue1) && s2.Contains(searchvalue2))
                    {
                        if (setvalue1 != "")
                        {
                            sheet.Cells[startrow, setcolumn1].value = setvalue1;
                        }
                        if (setvalue2 != "")
                        {
                            sheet.Cells[startrow, setcolumn2].value = setvalue2;
                        }
                        if (setvalue3 != "")
                        {
                            sheet.Cells[startrow, setcolumn3].value = setvalue3;
                        }

                    }
                    ++startrow;
                }


                return "Excel found";

            }
            catch (Exception e)
            {
                string exc = e.ToString();
                return exc;
            }
        }
        #endregion
        #region LOOP THROUGH ROWS IN COLUMN AND CHECK IF CELL CONTAINS STRING
        public List<string> Search1ValueIn1Column(string workbookname, string visible, int loopcolumn, int startrow, string searchvalue, int searchcolumn)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                string cellvalue = null;
                string celladress = null;
                List<string> returnlist = new List<string> { };

                while (sheet.Cells[startrow, loopcolumn].value != null || sheet.Cells[startrow + 1, loopcolumn].value != null || sheet.Cells[startrow + 2, loopcolumn].value != null)
                {
                    string s = sheet.Cells[startrow, searchcolumn].text;

                    if (s.Contains(searchvalue))
                    {
                        cellvalue = sheet.Cells[startrow, searchcolumn].text;
                        Excel.Range adressrange = xlapp.Cells[startrow, searchcolumn];
                        celladress = adressrange.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress);
                        returnlist.Add(cellvalue);
                        break;
                    }
                    ++startrow;
                }


                return returnlist;

            }
            catch (Exception e)
            {
                string exc = e.ToString();
                List<string> returnlist = new List<string> { exc };
                return returnlist;
            }
        }
        #endregion
        #region LOOP THROUGH ROWS IN 2 COLUMNS AND CHECK IF CELL CONTAINS STRING
        public List<string> Search2ValuesIn2Columns(string workbookname, string visible, int loopcolumn, int startrow, string searchvalue1, int searchcolumn1, string searchvalue2, int searchcolumn2)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                string cellvalue1 = null;
                string celladress1 = null;
                string cellvalue2 = null;
                string celladress2 = null;
                List<string> returnlist = new List<string> { };

                while (sheet.Cells[startrow, loopcolumn].value != null || sheet.Cells[startrow + 1, loopcolumn].value != null || sheet.Cells[startrow + 2, loopcolumn].value != null)
                {
                    string s1 = sheet.Cells[startrow, searchcolumn1].text;
                    string s2 = sheet.Cells[startrow, searchcolumn2].text;

                    if (s1.Contains(searchvalue1) && s2.Contains(searchvalue2))
                    {
                        cellvalue1 = sheet.Cells[startrow, searchcolumn1].text;
                        Excel.Range adressrange1 = xlapp.Cells[startrow, searchcolumn1];
                        celladress1 = adressrange1.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress1);
                        returnlist.Add(cellvalue1);
                        cellvalue2 = sheet.Cells[startrow, searchcolumn2].text;
                        Excel.Range adressrange2 = xlapp.Cells[startrow, searchcolumn2];
                        celladress2 = adressrange2.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress2);
                        returnlist.Add(cellvalue2);
                        break;
                    }
                    ++startrow;
                }


                return returnlist;

            }
            catch (Exception e)
            {
                string exc = e.ToString();
                List<string> returnlist = new List<string> { exc };
                return returnlist;
            }
        }
        #endregion
        #region LOOP THROUGH ALL ROWS IN 1 COLUMN AND GET 3 CELLS
        public List<string> Search1ValuesIn1ColumnsAll(string workbookname, string visible, int loopcolumn, int startrow, string searchvalue1, int searchcolumn1, int getcolumn1, int getcolumn2, int getcolumn3)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                string cellvalue = null;

                List<string> returnlist = new List<string> { };

                while (sheet.Cells[startrow, loopcolumn].value != null || sheet.Cells[startrow + 1, loopcolumn].value != null || sheet.Cells[startrow + 2, loopcolumn].value != null)
                {
                    string s1 = sheet.Cells[startrow, searchcolumn1].text;

                    if (s1.Contains(searchvalue1))
                    {
                        cellvalue = sheet.Cells[startrow, getcolumn1].text;
                        returnlist.Add(cellvalue);
                        cellvalue = sheet.Cells[startrow, getcolumn2].text;
                        returnlist.Add(cellvalue);
                        cellvalue = sheet.Cells[startrow, getcolumn3].text;
                        returnlist.Add(cellvalue);

                    }
                    ++startrow;
                }


                return returnlist;

            }
            catch (Exception e)
            {
                string exc = e.ToString();
                List<string> returnlist = new List<string> { exc };
                return returnlist;
            }
        }
        #endregion
        #region LOOP THROUGH ALL ROWS IN 1 ROW AND GET ALL VALUES
        public List<string> LoopThroughRows(string workbookname, string visible, int loopcolumn, int startrow)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                string cellvalue = null;

                List<string> returnlist = new List<string> { };

                while (sheet.Cells[startrow, loopcolumn].value != null || sheet.Cells[startrow + 1, loopcolumn].value != null || sheet.Cells[startrow + 2, loopcolumn].value != null)
                {
                    cellvalue = sheet.Cells[startrow, loopcolumn].text;
                    returnlist.Add(cellvalue);
                   
                    ++startrow;
                }


                return returnlist;

            }
            catch (Exception e)
            {
                string exc = e.ToString();
                List<string> returnlist = new List<string> { exc };
                return returnlist;
            }
        }
        #endregion
        #region LOOP THROUGH ALL ROWS IN 2 COLUMN AND GET 3 CELLS
        public List<string> Search2ValuesIn2ColumnsAll(string workbookname, string visible, int loopcolumn, int startrow, string searchvalue1, int searchcolumn1, string searchvalue2, int searchcolumn2, int getcolumn1, int getcolumn2, int getcolumn3)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                string cellvalue = null;

                List<string> returnlist = new List<string> { };

                while (sheet.Cells[startrow, loopcolumn].value != null || sheet.Cells[startrow + 1, loopcolumn].value != null || sheet.Cells[startrow + 2, loopcolumn].value != null)
                {
                    string s1 = sheet.Cells[startrow, searchcolumn1].text;
                    string s2 = sheet.Cells[startrow, searchcolumn2].text;

                    if (s1.Contains(searchvalue1) && s2.Contains(searchvalue2))
                    {
                        cellvalue = sheet.Cells[startrow, getcolumn1].text;
                        returnlist.Add(cellvalue);
                        cellvalue = sheet.Cells[startrow, getcolumn2].text;
                        returnlist.Add(cellvalue);
                        cellvalue = sheet.Cells[startrow, getcolumn3].text;
                        returnlist.Add(cellvalue);

                    }
                    ++startrow;
                }


                return returnlist;

            }
            catch (Exception e)
            {
                string exc = e.ToString();
                List<string> returnlist = new List<string> { exc };
                return returnlist;
            }
        }
        #endregion
        #region LOOP THROUGH ROWS IN 3 COLUMNS AND CHECK IF CELL CONTAINS STRING
        public List<string> Search3ValuesIn3Columns(string workbookname, string visible, int loopcolumn, int startrow, string searchvalue1, int searchcolumn1, string searchvalue2, int searchcolumn2, string searchvalue3, int searchcolumn3)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                string cellvalue1 = null;
                string celladress1 = null;
                string cellvalue2 = null;
                string celladress2 = null;
                string cellvalue3 = null;
                string celladress3 = null;
                List<string> returnlist = new List<string> { };

                while (sheet.Cells[startrow, loopcolumn].value != null || sheet.Cells[startrow + 1, loopcolumn].value != null || sheet.Cells[startrow + 2, loopcolumn].value != null)
                {
                    string s1 = sheet.Cells[startrow, searchcolumn1].text;
                    string s2 = sheet.Cells[startrow, searchcolumn2].text;
                    string s3 = sheet.Cells[startrow, searchcolumn3].text;

                    if (s1.Contains(searchvalue1) && s2.Contains(searchvalue2) && s3.Contains(searchvalue3))
                    {
                        cellvalue1 = sheet.Cells[startrow, searchcolumn1].text;
                        Excel.Range adressrange1 = xlapp.Cells[startrow, searchcolumn1];
                        celladress1 = adressrange1.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress1);
                        returnlist.Add(cellvalue1);
                        cellvalue2 = sheet.Cells[startrow, searchcolumn2].text;
                        Excel.Range adressrange2 = xlapp.Cells[startrow, searchcolumn2];
                        celladress2 = adressrange2.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress2);
                        returnlist.Add(cellvalue2);
                        cellvalue3 = sheet.Cells[startrow, searchcolumn3].text;
                        Excel.Range adressrange3 = xlapp.Cells[startrow, searchcolumn3];
                        celladress3 = adressrange3.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress3);
                        returnlist.Add(cellvalue3);
                        break;
                    }
                    ++startrow;
                }


                return returnlist;

            }
            catch (Exception e)
            {
                string exc = e.ToString();
                List<string> returnlist = new List<string> { exc };
                return returnlist;
            }
        }
        #endregion
        #region LOOP THROUGH ROWS IN 4 COLUMNS AND CHECK IF CELL CONTAINS STRING
        public List<string> Search4ValuesIn4Columns(string workbookname, string visible, int loopcolumn, int startrow, string searchvalue1, int searchcolumn1, string searchvalue2, int searchcolumn2, string searchvalue3, int searchcolumn3, string searchvalue4, int searchcolumn4)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                string cellvalue1 = null;
                string celladress1 = null;
                string cellvalue2 = null;
                string celladress2 = null;
                string cellvalue3 = null;
                string celladress3 = null;
                string cellvalue4 = null;
                string celladress4 = null;
                List<string> returnlist = new List<string> { };

                while (sheet.Cells[startrow, loopcolumn].value != null || sheet.Cells[startrow + 1, loopcolumn].value != null || sheet.Cells[startrow + 2, loopcolumn].value != null)
                {
                    string s1 = sheet.Cells[startrow, searchcolumn1].text;
                    string s2 = sheet.Cells[startrow, searchcolumn2].text;
                    string s3 = sheet.Cells[startrow, searchcolumn3].text;
                    string s4 = sheet.Cells[startrow, searchcolumn4].text;

                    if (s1.Contains(searchvalue1) && s2.Contains(searchvalue2) && s3.Contains(searchvalue3) && s4.Contains(searchvalue4))
                    {
                        cellvalue1 = sheet.Cells[startrow, searchcolumn1].text;
                        Excel.Range adressrange1 = xlapp.Cells[startrow, searchcolumn1];
                        celladress1 = adressrange1.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress1);
                        returnlist.Add(cellvalue1);
                        cellvalue2 = sheet.Cells[startrow, searchcolumn2].text;
                        Excel.Range adressrange2 = xlapp.Cells[startrow, searchcolumn2];
                        celladress2 = adressrange2.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress2);
                        returnlist.Add(cellvalue2);
                        cellvalue3 = sheet.Cells[startrow, searchcolumn3].text;
                        Excel.Range adressrange3 = xlapp.Cells[startrow, searchcolumn3];
                        celladress3 = adressrange3.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress3);
                        returnlist.Add(cellvalue3);
                        cellvalue4 = sheet.Cells[startrow, searchcolumn4].text;
                        Excel.Range adressrange4 = xlapp.Cells[startrow, searchcolumn4];
                        celladress4 = adressrange4.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress4);
                        returnlist.Add(cellvalue4);
                        break;
                    }
                    ++startrow;
                }


                return returnlist;

            }
            catch (Exception e)
            {
                string exc = e.ToString();
                List<string> returnlist = new List<string> { exc };
                return returnlist;
            }
        }
        #endregion
        #region LOOP THROUGH COLUMNS IN ROW AND CHECK IF CELL CONTAINS STRING
        public List<string> Search1ValueIn1Row(string workbookname, string visible, int looprow, int startcolumn, string searchvalue, int searchrow)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                string cellvalue = null;
                string celladress = null;
                List<string> returnlist = new List<string> { };

                while (sheet.Cells[looprow, startcolumn].value != null || sheet.Cells[looprow, startcolumn + 1].value != null || sheet.Cells[looprow, startcolumn + 2].value != null)
                {
                    string s = sheet.Cells[searchrow, startcolumn].text;

                    if (s.Contains(searchvalue))
                    {
                        cellvalue = sheet.Cells[searchrow, startcolumn].text;
                        Excel.Range adressrange = xlapp.Cells[searchrow, startcolumn];
                        celladress = adressrange.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress);
                        returnlist.Add(cellvalue);
                        break;
                    }
                    ++startcolumn;
                }


                return returnlist;

            }
            catch (Exception e)
            {
                string exc = e.ToString();
                List<string> returnlist = new List<string> { exc };
                return returnlist;
            }
        }
        #endregion
        #region LOOP THROUGH COLUMNS IN 2 ROWS AND CHECK IF CELL CONTAINS STRING
        public List<string> Search2ValuesIn2Rows(string workbookname, string visible, int looprow, int startcolumn, string searchvalue1, int searchrow1, string searchvalue2, int searchrow2)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                string cellvalue1 = null;
                string celladress1 = null;
                string cellvalue2 = null;
                string celladress2 = null;
                List<string> returnlist = new List<string> { };

                while (sheet.Cells[looprow, startcolumn ].value != null || sheet.Cells[looprow, startcolumn + 1].value != null || sheet.Cells[looprow, startcolumn + 2].value != null)
                {
                    string s1 = sheet.Cells[searchrow1, startcolumn].text;
                    string s2 = sheet.Cells[searchrow2, startcolumn].text;

                    if (s1.Contains(searchvalue1) && s2.Contains(searchvalue2))
                    {
                        cellvalue1 = sheet.Cells[searchrow1, startcolumn].text;
                        Excel.Range adressrange1 = xlapp.Cells[searchrow1, startcolumn];
                        celladress1 = adressrange1.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress1);
                        returnlist.Add(cellvalue1);
                        cellvalue2 = sheet.Cells[searchrow2, startcolumn].text;
                        Excel.Range adressrange2 = xlapp.Cells[searchrow2, startcolumn];
                        celladress2 = adressrange2.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress2);
                        returnlist.Add(cellvalue2);
                        break;
                    }
                    ++startcolumn;
                }


                return returnlist;

            }
            catch (Exception e)
            {
                string exc = e.ToString();
                List<string> returnlist = new List<string> { exc };
                return returnlist;
            }
        }
        #endregion
        #region LOOP THROUGH ALL COLUMNS IN 1 ROW AND GET 3 CELLS
        public List<string> Search2ValuesIn2RowsAll(string workbookname, string visible, int looprow, int startcolumn, string searchvalue1, int searchrow1, string searchvalue2, int searchrow2, int getcolumn1, int getcolumn2, int getcolumn3)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                string cellvalue1 = null;
                string cellvalue2 = null;
                string cellvalue3 = null;
                List<string> returnlist = new List<string> { };

                while (sheet.Cells[looprow, startcolumn ].value != null || sheet.Cells[looprow, startcolumn + 1].value != null || sheet.Cells[looprow, startcolumn + 2].value != null)
                {
                    string s1 = sheet.Cells[searchrow1, startcolumn].text;
                    string s2 = sheet.Cells[searchrow2, startcolumn].text;

                    if (s1.Contains(searchvalue1) && s2.Contains(searchvalue2))
                    {
                        cellvalue1 = sheet.Cells[searchrow1, getcolumn1].text;
                        cellvalue2 = sheet.Cells[searchrow1, getcolumn2].text;
                        cellvalue3 = sheet.Cells[searchrow1, getcolumn3].text;
                        returnlist.Add(cellvalue1);
                        returnlist.Add(cellvalue2);
                        returnlist.Add(cellvalue3);
                        
                    }
                    ++startcolumn;
                }


                return returnlist;

            }
            catch (Exception e)
            {
                string exc = e.ToString();
                List<string> returnlist = new List<string> { exc };
                return returnlist;
            }
        }
        #endregion
        #region LOOP THROUGH ALL COLUMNS IN 1 ROW AND GET ALL VALUES
        public List<string> LoopThroughColumns(string workbookname, string visible, int looprow, int startcolumn)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                string cellvalue1 = null;
                List<string> returnlist = new List<string> { };

                while (sheet.Cells[looprow, startcolumn].value != null || sheet.Cells[looprow, startcolumn + 1].value != null || sheet.Cells[looprow, startcolumn + 2].value != null)
                {
                    cellvalue1 = sheet.Cells[looprow, startcolumn].text;
                    returnlist.Add(cellvalue1);

                    ++startcolumn;
                }


                return returnlist;

            }
            catch (Exception e)
            {
                string exc = e.ToString();
                List<string> returnlist = new List<string> { exc };
                return returnlist;
            }
        }
        #endregion
        #region LOOP THROUGH COLUMNS IN 3 ROWS AND CHECK IF CELL CONTAINS STRING
        public List<string> Search3ValuesIn3Rows(string workbookname, string visible, int looprow, int startcolumn, string searchvalue1, int searchrow1, string searchvalue2, int searchrow2, string searchvalue3, int searchrow3)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                string cellvalue1 = null;
                string celladress1 = null;
                string cellvalue2 = null;
                string celladress2 = null;
                string cellvalue3 = null;
                string celladress3 = null;
                List<string> returnlist = new List<string> { };

                while (sheet.Cells[looprow, startcolumn ].value != null || sheet.Cells[looprow, startcolumn + 1].value != null || sheet.Cells[looprow, startcolumn + 2].value != null)
                {
                    string s1 = sheet.Cells[searchrow1, startcolumn].text;
                    string s2 = sheet.Cells[searchrow2, startcolumn].text;
                    string s3 = sheet.Cells[searchrow3, startcolumn].text;

                    if (s1.Contains(searchvalue1) && s2.Contains(searchvalue2))
                    {
                        cellvalue1 = sheet.Cells[searchrow1, startcolumn].text;
                        Excel.Range adressrange1 = xlapp.Cells[searchrow1, startcolumn];
                        celladress1 = adressrange1.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress1);
                        returnlist.Add(cellvalue1);
                        cellvalue2 = sheet.Cells[searchrow2, startcolumn].text;
                        Excel.Range adressrange2 = xlapp.Cells[searchrow2, startcolumn];
                        celladress2 = adressrange2.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress2);
                        returnlist.Add(cellvalue2);
                        cellvalue3 = sheet.Cells[searchrow3, startcolumn].text;
                        Excel.Range adressrange3 = xlapp.Cells[searchrow3, startcolumn];
                        celladress3 = adressrange3.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress3);
                        returnlist.Add(cellvalue3);
                        break;
                    }
                    ++startcolumn;
                }


                return returnlist;

            }
            catch (Exception e)
            {
                string exc = e.ToString();
                List<string> returnlist = new List<string> { exc };
                return returnlist;
            }
        }
        #endregion
        #region LOOP THROUGH COLUMNS IN 4 ROWS AND CHECK IF CELL CONTAINS STRING
        public List<string> Search4ValuesIn4Rows(string workbookname, string visible, int looprow, int startcolumn, string searchvalue1, int searchrow1, string searchvalue2, int searchrow2, string searchvalue3, int searchrow3, string searchvalue4, int searchrow4)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                string cellvalue1 = null;
                string celladress1 = null;
                string cellvalue2 = null;
                string celladress2 = null;
                string cellvalue3 = null;
                string celladress3 = null;
                string cellvalue4 = null;
                string celladress4 = null;
                List<string> returnlist = new List<string> { };

                while (sheet.Cells[looprow, startcolumn ].value != null || sheet.Cells[looprow, startcolumn + 1].value != null || sheet.Cells[looprow, startcolumn + 2].value != null)
                {
                    string s1 = sheet.Cells[searchrow1, startcolumn].text;
                    string s2 = sheet.Cells[searchrow2, startcolumn].text;
                    string s3 = sheet.Cells[searchrow3, startcolumn].text;
                    string s4 = sheet.Cells[searchrow4, startcolumn].text;

                    if (s1.Contains(searchvalue1) && s2.Contains(searchvalue2))
                    {
                        cellvalue1 = sheet.Cells[searchrow1, startcolumn].text;
                        Excel.Range adressrange1 = xlapp.Cells[searchrow1, startcolumn];
                        celladress1 = adressrange1.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress1);
                        returnlist.Add(cellvalue1);
                        cellvalue2 = sheet.Cells[searchrow2, startcolumn].text;
                        Excel.Range adressrange2 = xlapp.Cells[searchrow2, startcolumn];
                        celladress2 = adressrange2.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress2);
                        returnlist.Add(cellvalue2);
                        cellvalue3 = sheet.Cells[searchrow3, startcolumn].text;
                        Excel.Range adressrange3 = xlapp.Cells[searchrow3, startcolumn];
                        celladress3 = adressrange3.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress3);
                        returnlist.Add(cellvalue3);
                        cellvalue4 = sheet.Cells[searchrow3, startcolumn].text;
                        Excel.Range adressrange4 = xlapp.Cells[searchrow3, startcolumn];
                        celladress4 = adressrange4.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();
                        returnlist.Add(celladress4);
                        returnlist.Add(cellvalue4);
                        break;
                    }
                    ++startcolumn;
                }


                return returnlist;

            }
            catch (Exception e)
            {
                string exc = e.ToString();
                List<string> returnlist = new List<string> { exc };
                return returnlist;
            }
        }
        #endregion
        #region LOOP 1 COLUMN AND GET HEADER NAME OF A COLUMN THAT HAS SPECIFIC VALUE IN IT
        public List<string> Loop1ColumnAndGetHeaderNameOfColumnWithSpecificValue(string workbookname, string visible, int loopcolumn, int startrow, int looprow, int startcolumn, string searchvalue, int headerrow)
        {
            try
            {
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
                string cellvalue = null;
                long lastrow;
                long lastcolumn;
                List<string> returnlist = new List<string> { };

                lastrow = startrow;
                while (sheet.Cells[lastrow + 1, loopcolumn].value != null || sheet.Cells[lastrow + 2, loopcolumn].value != null)
                {
                    ++lastrow;
                }

                lastcolumn = startcolumn;
                while (sheet.Cells[looprow, lastcolumn + 1].value != null || sheet.Cells[looprow, lastcolumn + 2].value != null)
                {
                    ++lastcolumn;
                }


                while (sheet.Cells[startrow, loopcolumn].value != null || sheet.Cells[startrow + 1, loopcolumn].value != null || sheet.Cells[startrow + 2, loopcolumn].value != null)
                {

                    cellvalue = sheet.Cells[startrow, loopcolumn].text;
                    returnlist.Add(cellvalue);
                    for (int i = 1; i <= lastcolumn; i++)
                    {
                        cellvalue = sheet.Cells[startrow, i].text;
                        if (cellvalue == searchvalue)
                        {
                            returnlist.Add(sheet.Cells[headerrow, i].text);
                        }
                    }
                    returnlist.Add("next record");

                    ++startrow;
                }


                return returnlist;

            }
            catch (Exception e)
            {
                string exc = e.ToString();
                List<string> returnlist = new List<string> { exc };
                return returnlist;
            }
        }
        #endregion

        #region OBTAIN DATA FROM EXCEL

        #region GET: SHEETNAME, COLUMN COUNT, ROW COUNT, CHANGE SHEET NAME, SHEET COUNT
        private Tuple<int?, int?, string, int?> ObtainExcelData(string workbookname, string newsheetname = "")
        {
            try
            {
                string visible = "Yes";
                Workbook workbook = null;
                Application xlapp = null;
                Worksheet sheet = null;
                ExcelInstance instance = new ExcelInstance();
                instance.Instance(workbookname, visible, out workbook, out xlapp, out sheet);
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
                string exc = e.ToString();
                return Tuple.Create(value, value2,exc, value3);
            }
        }
        #endregion
        #endregion

    }
}
#endregion