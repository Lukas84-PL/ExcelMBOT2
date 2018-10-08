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

        #region APPLY FILTER
        public void ApplyFilter(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto, int filtercolumn,string[] filterlist)
        {
            Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
            Excel.Range rngTo = xlapp.Cells[rowto, columnto];

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                    //range.AutoFilter(3, "e");

                    foreach (var item in filterlist)
                    {
                        range.AutoFilter(filtercolumn, "<>" + item,
                        Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                        
                    }


                }
            }
        }
        #endregion

        #region GET EXCEL RANGE TO ARRAY
        public Array GetExcelRangeToArray(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto)
        {
            Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
            Excel.Range rngTo = xlapp.Cells[rowto, columnto];
            string[,] result = null;

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {

                //ArrayList arrForValues = new ArrayList();
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range

                    Excel.Range range = sheet.get_Range(rngFrom, rngTo);

                    object[,] value = range.Value; //the value 


                    //ArrayList[,] arr = new ArrayList[,](value);

                    int rank = ((Array)value).Rank;
                    if (rank == 2)
                    {
                        int columnCount = columnto - (columnfrom - 1);
                        int rowCount = rowto - (rowfrom - 1);
                        result = new string[rowCount, columnCount];

                        for (int i = 1; (i - 1) < rowCount; i++)
                        {
                            for (int j = 1; (j - 1) < columnCount; j++)
                            {
                                if (value.GetValue(i, j) != null)
                                {
                                    result[i - 1, j - 1] = ((Array)value).GetValue(i, j).ToString();
                                }
                                else
                                {
                                    result[i - 1, j - 1] = "";
                                }
                                
                            }
                        }
                    }

                    //result.Cast<int>()
                    //    .Select((x, i) => new { x, index = i / result.GetLength(1) })
                    //    .GroupBy(x => x.index)
                    //    .Select(x => x.Select(s => s.x).ToList()).Dump();


                    //myArrayList.AddRange(result);
                    return result;
                }
            }
            return result;
        }
        #endregion
        #region SORT RANGE
        #region PASTE IN CELL
        public void InsertObject(string workbookname, int column, int row, string filepath, string iconname ,int iconindex, int iconwidth = 1, int iconheight = 3)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cell
                    Excel.Range range = sheet.Cells[row, column];
                    //Enter forumla

                    range.Activate();

                    //var obj = xlapp.ActiveSheet.OLEObjects.Add(Filename: filepath,
                    //    Link: false,
                    //    DisplayAsIcon: true,
                    //    IconFileName: iconname,
                    //    IconIndex: iconindex,
                    //    IconLabel: iconname,
                    //    Left: 3,
                    //    Top: 1,
                    //    Width: 3,
                    //    Height: 2);
                    

                    Excel.OLEObject testshape = xlapp.ActiveSheet.OLEObjects.Add(Filename: filepath,
                        Link: false,
                        DisplayAsIcon: true,
                        IconFileName: iconname,
                        IconIndex: iconindex,
                        IconLabel: iconname,
                        Left: false,
                        Top: false,
                        Width: 180,
                        Height: 200);

                    //object theObject = testshape.Object;
                    testshape.Width = 200;
                    testshape.Height = 200;


                    //theObject.GetType().InvokeMember("Caption", System.Reflection.BindingFlags.SetProperty, null, theObject, new object[] { "caption" })

                    //Excel.OLEObjects oleObjects = (sheet as Excel._Worksheet).OLEObjects() as Excel.OLEObjects;
                    //Excel.OLEObject testshape = oleObjects.Item("Object 6") as OLEObject;
                    //object theObject = obj.Object;

                    //object activeSheet = xlapp.ActiveSheet;
                    //Excel.OLEObjects oleObjects = (activeSheet as Excel._Worksheet).OLEObjects() as Excel.OLEObjects;
                    //Excel.OLEObject testshape = oleObjects.Item(iconname) as OLEObject;
                    ////object theObject = testshape.Object;
                    //testshape.Width = 150;
                    //double left = theObject.
                    //Console.WriteLine(left);

                }
            }
        }
        #endregion
        #endregion


        static void Main(string[] args)
        {

            Excel.Application xlapp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");

            Program sheetname = new Program();

            //List<string> list = new List<string> { "e", "f" };
            //sheetname.DeleteBlankColumnsOfSelection(xlapp, "Ctest.xlsx");
            // sheetname.DeleteBlankColumnsOfSelection(xlapp, "Ctest.xlsx", "A", "C", 7);
            //string[] rowlist = { "Question", "Answer", "Test" };
            //string[] columnlist = { "ColumnTest1"};
            //string[] valuefieldlist = { "ColumnTest2"};
            string[] filterfieldlist = {"e", "f"  };
            Array output = sheetname.GetExcelRangeToArray("Ctest.xlsx",1,1, 3, 5);
            //Console.WriteLine(result);
            //Console.WriteLine(result.Item1);
            //Console.WriteLine(result.Item2);
            //Console.WriteLine(result.Item3);
            //Console.WriteLine(result.Item4);
            //Console.WriteLine(workbookname);


            }
    }
}
