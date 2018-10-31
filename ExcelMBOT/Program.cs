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
using System.Data.Linq;
using System.Xml;
using System.Xml.Linq;

using System.Data;
using System.Text.RegularExpressions;
using System.Globalization;

namespace ConsoleApp1
{
    public class Program
    {
        Application xlapp = (Application)Marshal.GetActiveObject("Excel.Application");
        #region GET: SHEETNAME, COLUMN COUNT, ROW COUNT, CHANGE SHEET NAME, SHEET COUNT
        public Tuple<int, int, string, int> GetSheetName(Excel.Application xlapp, string workbookname, string newsheetname = "")
        {
            string worksheetname = "Not found";
            int rowcount = 0;
            int columncount = 0;
            int worksheetcount = 0;

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
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
                }
            }

            return Tuple.Create(columncount, rowcount, worksheetname, worksheetcount);
        }
        #endregion
        #region SAVEAS WORKBOOK
        public string SaveAs(Excel.Application xlapp, string workbookname, string newfilename = "")
        {
            string status = "Failed";
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    //SaveAs
                    workbook.SaveAs(newfilename);

                    foreach (Excel.Workbook wb in xlapp.Workbooks)
                    {
                        if ((workbook.Path + @"\" + workbook.Name) == newfilename)
                        {
                            status = "Completed";
                        }
                    }

                }
            }

            return status;
        }
        #endregion
        #region SELECT RANGE
        public void SelectRange(Excel.Application xlapp, string workbookname, string columnfrom, int rowfrom, string columnto, int rowto)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = sheet.get_Range(columnfrom + rowfrom, columnto + rowto);
                    range.Select();
                }
            }
        }
        #endregion

        #region SELECT BLANK CELLS OF SELECTION
        public void SelectBlankCellsOfSelection(Excel.Application xlapp, string workbookname)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Create object from selection
                    Excel.Range range = xlapp.ActiveWindow.Selection;
                    //Select blank cells in a range
                    Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                    //Delete blank cells rows
                    blankcells.Select();

                }
            }
        }
        #endregion
        #region SELECT BLANK CELLS OF RANGE
        public void SelectBlankCellsOfRange(Excel.Application xlapp, string workbookname, string column, int rowfrom, int rowto)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = sheet.get_Range(column + rowfrom, column + rowto);
                    //Select blank cells in a range
                    Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                    //Delete blank cells rows
                    blankcells.Select();

                }
            }
        }
        #endregion
        #region SELECT BLANK CELLS OF COLUMN
        public void SelectBlankCellsOfColumn(Excel.Application xlapp, string workbookname, string column)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select column
                    Excel.Range range = (Excel.Range)sheet.Columns[column + ":" + column];
                    //Select blank cells in a range
                    Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                    //Delete blank cells rows
                    blankcells.Select();

                }
            }
        }
        #endregion
        #region SELECT BLANK CELLS OF ROW
        public void SelectBlankCellsOfRow(Excel.Application xlapp, string workbookname, int row)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select column
                    Excel.Range range = (Excel.Range)sheet.Rows[row + ":" + row];
                    //Select blank cells in a range
                    Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                    //Delete blank cells rows
                    blankcells.Select();

                }
            }
        }
        #endregion

        #region DELETE BLANK ROWS OF SELECTION
        public void DeleteBlankRowsInSelection(Excel.Application xlapp, string workbookname)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Create object from selection
                    Excel.Range range = xlapp.ActiveWindow.Selection;
                    //Select blank cells in a range
                    Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                    //Delete blank cells rows
                    blankcells.EntireRow.Delete();

                }
            }
        }
        #endregion
        #region DELETE BLANK ROWS OF RANGE
        public void DeleteBlankRowsOfRange(Excel.Application xlapp, string workbookname, string column, int rowfrom, int rowto)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = sheet.get_Range(column + rowfrom, column + rowto);
                    //Select blank cells in a range
                    Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                    //Delete blank cells rows
                    blankcells.EntireRow.Delete();

                }
            }
        }
        #endregion
        #region DELETE BLANK ROWS OF COLUMN
        public void DeleteBlankRowsOfColumn(Excel.Application xlapp, string workbookname, string column)
        {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select column
                    Excel.Range range = (Excel.Range)sheet.Columns[column + ":" + column];
                    //Select blank cells of column
                    Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                    //Delete blank cells rows
                    blankcells.EntireRow.Delete();

                }
            }
        }
        #endregion

        #region DELETE BLANK COLUMNS IN A SELECTION
        public void DeleteBlankColumnsOfSelection(Excel.Application xlapp, string workbookname)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = xlapp.ActiveWindow.Selection;
                    //Select blank cells in a range
                    Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                    //Delete blank cells columns
                    blankcells.EntireColumn.Delete();

                }
            }
        }
        #endregion
        #region DELETE BLANK COLUMNS IN A RANGE
        public void DeleteBlankColumnsOfRange(Excel.Application xlapp, string workbookname, string columnfrom, string columnto, int row)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = sheet.get_Range(columnfrom + row, columnto + row);
                    //Select blank cells in a range
                    Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                    //Delete blank cells columns
                    blankcells.EntireColumn.Delete();

                }
            }
        }
        #endregion
        #region DELETE BLANK COLUMNS OF ROW
        public void DeleteBlankColumnsOfRow(Excel.Application xlapp, string workbookname, int row)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = (Excel.Range)sheet.Rows[row + ":" + row];
                    //Select blank cells in a range
                    Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                    //Delete blank cells columns
                    blankcells.EntireColumn.Delete();

                }
            }
        }
        #endregion


        #region DELETE ROWS OF SELECTED CELLS
        public void DeleteRowsOfSelectedCells(Excel.Application xlapp, string workbookname)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Create object from selection
                    Excel.Range range = xlapp.ActiveWindow.Selection;
                    //Delete selected rows
                    range.EntireRow.Delete();

                }
            }
        }
        #endregion
        #region DELETE COLUMNS OF SELECTED CELLS
        public void DeleteColumnsOfSelectedCells(Excel.Application xlapp, string workbookname)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Create object from selection
                    Excel.Range range = xlapp.ActiveWindow.Selection;
                    //Delete selected rows
                    range.EntireColumn.Delete();

                }
            }
        }
        #endregion

        #region SELECT ROW
        public void SelectEntireRow(Excel.Application xlapp, string workbookname, int row)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = (Excel.Range)sheet.Rows[row + ":" + row];
                    //Select row
                    range.Select();
                }
            }
        }
        #endregion
        #region SELECT COLUMN
        public void SelectEntireColumn(Excel.Application xlapp, string workbookname, string column)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select column
                    Excel.Range range = (Excel.Range)sheet.Columns[column + ":" + column];
                    //Delete blank cells rows
                    range.Select();
                }
            }
        }
        #endregion
        #region AUTOFIT ROW
        public void AutofitRow(string workbookname, int row)
        {

            Excel.Application xlapp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = (Excel.Range)sheet.Rows[row + ":" + row];
                    //Select row
                    range.AutoFit();
                }
            }
        }
        #endregion

        #region DRAG FORMULA
        public void DragFromula(string workbookname)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Create object from selection
                    string formula = sheet.Cells[7][2].Formula;
                    sheet.Cells[7][3].Formula = xlapp.ConvertFormula(formula, Excel.XlReferenceStyle.xlA1, Excel.XlReferenceStyle.xlR1C1, Excel.XlReferenceType.xlRelative, sheet.Cells[7][2]);

                    Excel.Range oRng = sheet.get_Range("H2").get_Resize(100, 1);
                    oRng = xlapp.ConvertFormula(formula, Excel.XlReferenceStyle.xlA1, Excel.XlReferenceStyle.xlR1C1, Excel.XlReferenceType.xlRelative, sheet.Cells[7][2]);
                    //range.Formula = "IF(AND(A" + 1 + "<> 0,B" + 1 + "<>2),\"YES\",\"NO\")";

                    //Delete selected rows
                    //range.ClearContents();

                }
            }
        }
        #endregion

        #region GO TO LAST ROW OF SPECIFIC COLUMN
        public void GoToLastRowOfSpecificasColumn(string workbookname, int column, int rowstart)
        {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select column

                    while (sheet.Cells[rowstart, column].value != null)
                    {
                        ++rowstart;
                    }
                    rowstart -= 1; 
                    Excel.Range lastcell = xlapp.Cells[rowstart, column];

                    lastcell.Activate();
                    lastcell.Select();
                }
            }
        }
        #endregion
        #region GO TO LAST COLUMN OF USED RANGE
        public void GoToLastColumnOfUsedRange(string workbookname, int row)
        {
            int columncount = 0;
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    columncount = sheet.UsedRange.Columns.Count;
                    Range lastcolumn = xlapp.Cells[row, columncount];
                    lastcolumn.Select();
                    lastcolumn.Activate();
                }
            }

        }
        #endregion
        #region CLOSE SPREADSHEET WITH SAVING
        public void QuitExcelApp(string workbookname)
        {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    xlapp.Quit();
                }
            }
        }
        #endregion
        #region CLOSE SPREADSHEET WITHOUT SAVING
        public void CloseSpreadsheetWithoutSaving(string workbookname)
        {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    workbook.Close(false);
                }
            }
        }
        #endregion

        #region GET EXCEL RANGE TO ARRAY
        public void OpenSpreadsheet(string workbookname, string path)
        {
            xlapp.Workbooks.Open(path + workbookname, false, false);

        }
        #endregion
        #region SORT RANGE
        #region DRAG CELL VALUE TO RANGE
        public void DragCellValueToRange(string workbookname,int column, int rowfrom, int rowto)
        {
            Excel.Range rngFrom = xlapp.Cells[rowfrom, column];
            Excel.Range rngTo = xlapp.Cells[rowto, column];

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;

                    Excel.Range rng = xlapp.get_Range(rngFrom, rngFrom);

                    rng.AutoFill(xlapp.get_Range(rngFrom, rngTo),
                        Excel.XlAutoFillType.xlFillWeekdays);
                }
            }
        }
        #endregion
        #endregion
        //#region CREATE EXCEL WORKBOOK
        //public string CreateExcelWrokbook(string workbookname)
        //{
        //    try
        //    {
        //        Application xlApp = new Application();

        //        xlApp.Workbooks.Add("test.xlsx");
        //        return "Workbook created";

        //    }
        //    catch (Exception e)
        //    {

        //        return e.ToString();
        //    }
        //}
        //#endregion
        #region LOOP THROUGH 2 ROWS IN COLUMN TO FIND 2 VALUES
        public List<string> LoopThrough3RowsInColumnsFor3Values(string workbookname, int loopcolumn, int startrow, string searchvalue1, int searchcolumn1, string searchvalue2, int searchcolumn2, string searchvalue3, int searchcolumn3)
        {
            try
            {
                Application xlapp = (Application)Marshal.GetActiveObject("Excel.Application");
                Workbook workbook = xlapp.Workbooks.get_Item(workbookname);
                string cellvalue1 = null;
                string cellvalue2 = null;
                string cellvalue3 = null;
                string celladress1 = null;
                string celladress2 = null;
                string celladress3 = null;
                List<string> returnlist = new List<string> {};
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;

                while (sheet.Cells[startrow + 1, loopcolumn].value != null || sheet.Cells[startrow + 2, loopcolumn].value != null)
                {

                    if (sheet.Cells[startrow, searchcolumn1].text == searchvalue1 && sheet.Cells[startrow, searchcolumn2].text == searchvalue2 && sheet.Cells[startrow, searchcolumn3].text == searchvalue3)
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
        #region LOOP THROUGH ROWS IN COLUMN AND CHECK IF CELL CONTAINS STRING
        public string ActivateSheet(string workbookname, string visible, string sheetname)
        {
            try
            {

                Application xlapp = (Application)Marshal.GetActiveObject("Excel.Application");
                xlapp.DisplayAlerts = false;
                if (visible == "yes" || visible == "Yes" || visible == "YES")
                {
                    xlapp.Visible = true;
                }
                else
                {
                    xlapp.Visible = false;
                }
                Workbook workbook = xlapp.Workbooks.get_Item(workbookname);


                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[sheetname];
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


        static void Main(string[] args)
        {

            //.Application xlapp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");

            Program sheetname = new Program();

            //List<string> list = new List<string> { "e", "f" };
            //sheetname.DeleteBlankColumnsOfSelection(xlapp, "Ctest.xlsx");
            // sheetname.DeleteBlankColumnsOfSelection(xlapp, "Ctest.xlsx", "A", "C", 7);
            //string[] rowlist = { "dupa"};
            //string[] columnlist = { "ColumnTest1"};
            //string[] valuefieldlist = { "ColumnTest2"};
            //string[] filterfieldlist = {"e", "f"  };
            //object cols = new object[] { 1,2 };
            sheetname.ActivateSheet("Ctest.xlsx","yes", "What is SFS 2.0");
            //string wynik = sheetname.GetAdressOfValue("Ctest.xlsx", 1, 1, 8, 3, "dupa", "kupa");
            //Console.WriteLine(result);
            //Console.WriteLine(result.Item1);
            //Console.WriteLine(result.Item2);
            //Console.WriteLine(result.Item3);
            //Console.WriteLine(result.Item4);
            //Console.WriteLine(workbookname);


        }
    }
}
