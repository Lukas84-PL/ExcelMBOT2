using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Drawing;





namespace ConsoleApp1
{
    public class Program
    {
        Excel.Application xlapp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
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

        #region FILTER ON VALUE PIVOT IN SELECTED SHEET
        public string FilterOnValuePivotInSelectedSheet(string workbookname, string pivotname, string[] filtervalues, string filterfield)
        {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.PivotCaches pCaches = workbook.PivotCaches();
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;

                    Excel.PivotTable pivot = (Excel.PivotTable)sheet.PivotTables(pivotname);
                    Excel.PivotField pivotfield = pivot.PivotFields(filterfield);
                    pivotfield.ClearAllFilters();
                    int count = pivot.PivotFields(1).PivotItems.Count;
                    for (int i = 1; i <= count; i++)
                    // string nm = pf.PivotItems(i).Name;
                    {
                        if (Array.IndexOf(filtervalues, pivotfield.PivotItems(i).Name) > -1)
                        {
                            pivotfield.PivotItems(i).visible = true;
                        }
                        else
                        {
                            pivotfield.PivotItems(i).visible = false;
                        }

                        //Select cells in a range
                        //Send range to cache and use it to create pivot

                    }
                }
            }
            return "Pivot creating failed";
        }
        #endregion

        static void Main(string[] args)
        {

            Excel.Application xlapp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");

            Program sheetname = new Program();

            //sheetname.DeleteBlankColumnsOfSelection(xlapp, "Ctest.xlsx");
            // sheetname.DeleteBlankColumnsOfSelection(xlapp, "Ctest.xlsx", "A", "C", 7);
            string[] rowlist = { "Question", "Answer", "Test" };
            string[] columnlist = { "ColumnTest1"};
            string[] valuefieldlist = { "ColumnTest2"};
            string[] filterfieldlist = {"1","3" };
            string result = sheetname.FilterPivotInSelectedSheet("Ctest.xlsx", "pivot1", filterfieldlist, "ColumnTest3");
            Console.WriteLine(result);
            //Console.WriteLine(result.Item1);
            //Console.WriteLine(result.Item2);
            //Console.WriteLine(result.Item3);
            //Console.WriteLine(result.Item4);
            //Console.WriteLine(workbookname);


            }
    }
}
