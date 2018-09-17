using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelMBOT
{
    public class ExcelMBOT
    {
        Excel.Application xlapp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");

        #region ACTIONS
        #region SAVEAS WORKBOOK
        public string SaveAs(string workbookname, string newfilename = "")
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
        public void SelectRange(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto)
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
                    range.Select();
                }
            }
        }
        #endregion
        #region COPY RANGE
        public void CopyRange(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto)
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
                    range.Copy();
                }
            }
        }
        #endregion
        #region COPY SELECTION
        public void CopySelection(string workbookname)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = xlapp.ActiveWindow.Selection;
                    range.Copy();
                }
            }
        }
        #endregion
        #region INSERT FORMULA
        public void InsertFormula(string workbookname, int column, int row, string inputformula)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cell
                    Excel.Range range = sheet.Cells[row, column];
                    //Enter forumla
                    range.Formula = inputformula;
                }
            }
        }
        #endregion

        #region REPLACE DATA IN RANGE
        public void ReplaceDataInRange(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto, string oldstring, string newstring)
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
                    range.Replace(oldstring, newstring);
                }
            }
        }
        #endregion
        #region REPLACE DATA IN SELECTION
        public void ReplaceDataInSelection(string workbookname, string oldstring, string newstring)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = xlapp.ActiveWindow.Selection;
                    range.Replace(oldstring, newstring);
                }
            }
        }
        #endregion

        #region CHANGE FONT IN RANGE TO BOLD
        public void ChangeFontInRangeToBold(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto)
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
                    range.Font.Bold = true;

                }
            }
        }
        #endregion
        #region CHANGE FONT IN SELECTION TO BOLD
        public void ChangeFontInSelectionToBold(string workbookname)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = xlapp.ActiveWindow.Selection;
                    range.Font.Bold = true;
                }
            }
        }
        #endregion

        #region PASTE VALUES IN SELECTION
        public void PasteValuesInSelection(string workbookname, string oldstring, string newstring)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = xlapp.ActiveWindow.Selection;
                    range.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                }
            }
        }
        #endregion
        #region PASTE VALUES IN CELL
        public void PasteValuesInCell(string workbookname, int column, int row, string inputformula)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cell
                    Excel.Range range = sheet.Cells[row, column];
                    //Enter forumla
                    range.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                }
            }
        }
        #endregion

        #region SELECT BLANK CELLS OF SELECTION
        public void SelectBlankCellsOfSelection(string workbookname)
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
        public void SelectBlankCellsOfRange(string workbookname, int column, int rowfrom, int rowto)
        {
            Excel.Range rngFrom = xlapp.Cells[rowfrom, column];
            Excel.Range rngTo = xlapp.Cells[rowto, column];

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                    //Select blank cells in a range
                    Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                    //Delete blank cells rows
                    blankcells.Select();

                }
            }
        }
        #endregion
        #region SELECT BLANK CELLS OF COLUMN
        public void SelectBlankCellsOfColumn(string workbookname, string column)
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
        public void SelectBlankCellsOfRow(string workbookname, int row)
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
        public void DeleteBlankRowsInSelection(string workbookname)
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
        public void DeleteBlankRowsOfRange(string workbookname, int column, int rowfrom, int rowto)
        {
            Excel.Range rngFrom = xlapp.Cells[rowfrom, column];
            Excel.Range rngTo = xlapp.Cells[rowto, column];

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                    //Select blank cells in a range
                    Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                    //Delete blank cells rows
                    blankcells.EntireRow.Delete();

                }
            }
        }
        #endregion
        #region DELETE BLANK ROWS OF COLUMN
        public void DeleteBlankRowsOfColumn(string workbookname, string column)
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
        public void DeleteBlankColumnsOfSelection(string workbookname)
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
        public void DeleteBlankColumnsOfRange(string workbookname, int columnfrom, int columnto, int row)
        {
            Excel.Range rngFrom = xlapp.Cells[row, columnfrom];
            Excel.Range rngTo = xlapp.Cells[row, columnto];

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                    //Select blank cells in a range
                    Excel.Range blankcells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                    //Delete blank cells columns
                    blankcells.EntireColumn.Delete();

                }
            }
        }
        #endregion
        #region DELETE BLANK COLUMNS OF ROW
        public void DeleteBlankColumnsOfRow(string workbookname, int row)
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
        public void DeleteRowsOfSelectedCells(string workbookname)
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
        public void DeleteColumnsOfSelectedCells(string workbookname)
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
        public void SelectEntireRow(string workbookname, int row)
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
        public void SelectEntireColumn(string workbookname, string column)
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

        #region INSERT ROW IN SELECTION
        public void InsertRowInSelection(string workbookname)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = xlapp.ActiveWindow.Selection();
                    //Delete blank cells rows
                    range.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                }
            }
        }
        #endregion
        #region INSERT ROW IN RANGE
        public void InsertRowInRange(string workbookname, int row)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = (Excel.Range)sheet.Rows[row + ":" + row];
                    //Delete blank cells rows
                    range.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                }
            }
        }
        #endregion
        #region INSERT COLUMN IN SELECTION
        public void InsertColumnInSelection(string workbookname)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select column
                    Excel.Range range = xlapp.ActiveWindow.Selection();
                    //Delete blank cells rows
                    range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
                }
            }
        }
        #endregion
        #region INSERT COLUMN IN RANGE
        public void InsertColumnInRange(string workbookname, string column)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select column
                    Excel.Range range = (Excel.Range)sheet.Columns[column + ":" + column];
                    //Delete blank cells rows
                    range.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
                }
            }
        }
        #endregion

        #region CHANGE SHEET NAME
        public void ChangeSheetName(string workbookname, string newsheetname)
        {
            ExcelMBOT obj = new ExcelMBOT();

            var result = obj.ObtainExcelData(workbookname, newsheetname);
        }
        #endregion

        #region GET SHEET NAME
        public string GetSheetName(string workbookname)
        {
            ExcelMBOT obj = new ExcelMBOT();

            var result = obj.ObtainExcelData(workbookname);

            return result.Item3;
        }
        #endregion
        #region GET SHEET COUNT
        public int GetSheetCount(string workbookname)
        {
            ExcelMBOT obj = new ExcelMBOT();

            var result = obj.ObtainExcelData(workbookname);

            return result.Item4;
        }
        #endregion
        #region GET ROW COUNT
        public int GetRowsCount(string workbookname)
        {
            ExcelMBOT obj = new ExcelMBOT();

            var result = obj.ObtainExcelData(workbookname);

            return result.Item2;
        }
        #endregion
        #region GET COLUMN COUNT
        public int GetColumnsCount(string workbookname)
        {
            ExcelMBOT obj = new ExcelMBOT();

            var result = obj.ObtainExcelData(workbookname);

            return result.Item1;
        }
        #endregion

        #region AUTOFIT COLUMN
        public void AutofitColumn(string workbookname, string column)
        {

            Excel.Application xlapp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select column
                    Excel.Range range = (Excel.Range)sheet.Columns[column + ":" + column];
                    //Autofit column
                    range.AutoFit();
                }
            }
        }
        #endregion
        #region AUTOFIT ROW
        public void AutofitRow(string workbookname, int row)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = (Excel.Range)sheet.Rows[row + ":" + row];
                    //Autofit row
                    range.AutoFit();
                }
            }
        }
        #endregion
        #region DELETE SHEET
        public void DeleteSheet(string workbookname)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Delete sheet
                    sheet.Delete();
                }
            }
        }
        #endregion
        #region COPY SHEET TO NEW WORKBOOK
        public string CopySheetToNewWorkbook(string workbookname)
        {
            Excel.Application xlapp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            List<string> workbooklist = new List<string>();
            string wkbname;
            string newworkbook = "Unable to find new workbook name";

            //Find all open workbooks and add their names to the list
            foreach (Excel.Workbook activewkbs in xlapp.Workbooks)
            {
                wkbname = Convert.ToString(activewkbs.Name);
                workbooklist.Add(wkbname);
            }

            //Find workbook from input
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Copy sheet
                    sheet.Copy();
                    //Obtain name of newly created worksheet
                    foreach (Excel.Workbook wkb in xlapp.Workbooks)
                    {
                        wkbname = Convert.ToString(wkb.Name);
                        if (workbooklist.Contains(wkbname) == false)
                        {
                            newworkbook = wkbname;
                        }
                    }
                }
            }
            return newworkbook;
        }
        #endregion
        #endregion



        #region OBTAIN DATA FROM EXCEL

        #region GET: SHEETNAME, COLUMN COUNT, ROW COUNT, CHANGE SHEET NAME, SHEET COUNT
        private Tuple<int, int, string, int> ObtainExcelData(string workbookname, string newsheetname = "")
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
        #endregion

    }
}
