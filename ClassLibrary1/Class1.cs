using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelDataManipulation
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
        #region COPY COLUMN
        public void CopyColumn(string workbookname, string column)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = (Excel.Range)sheet.Columns[column + ":" + column];
                    range.Copy();
                }
            }
        }
        #endregion
        #region COPY ROW
        public void CopyRow(string workbookname, int row)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = (Excel.Range)sheet.Rows[row + ":" + row];
                    range.Copy();
                }
            }
        }
        #endregion
        #region COPY AND PASTE VALUES OFCOLUMN
        public void CopyAndPasteValuesOfColumn(string workbookname, string columnfrom, string columnto)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range rangefrom = (Excel.Range)sheet.Columns[columnfrom + ":" + columnfrom];
                    rangefrom.Copy();
                    Excel.Range rangeto = (Excel.Range)sheet.Columns[columnto + ":" + columnto];
                    rangeto.PasteSpecial(Excel.XlPasteType.xlPasteValues);

                }
            }
        }
        #endregion
        #region COPY AND PASTE VALUES OF ROW
        public void CopyAndPasteValuesOfRow(string workbookname, int rowfrom, int rowto)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range rangefrom = (Excel.Range)sheet.Rows[rowfrom + ":" + rowfrom];
                    rangefrom.Copy();
                    Excel.Range rangeto = (Excel.Range)sheet.Rows[rowto + ":" + rowto];
                    rangeto.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                }
            }
        }
        #endregion
        #region COPY AND PASTE COLUMN
        public void CopyAndPasteColumn(string workbookname, string columnfrom, string columnto)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range rangefrom = (Excel.Range)sheet.Columns[columnfrom + ":" + columnfrom];
                    rangefrom.Copy();
                    Excel.Range rangeto = (Excel.Range)sheet.Columns[columnto + ":" + columnto];
                    rangeto.PasteSpecial();

                }
            }
        }
        #endregion
        #region COPY AND PASTE ROW
        public void CopyAndPasteRow(string workbookname, int rowfrom, int rowto)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range rangefrom = (Excel.Range)sheet.Rows[rowfrom + ":" + rowfrom];
                    rangefrom.Copy();
                    Excel.Range rangeto = (Excel.Range)sheet.Rows[rowto + ":" + rowto];
                    rangeto.PasteSpecial();
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
        #region PASTE IN SELECTION
        public void PasteInSelection(string workbookname, string oldstring, string newstring)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = xlapp.ActiveWindow.Selection;
                    range.PasteSpecial();
                }
            }
        }
        #endregion
        #region PASTE IN CELL
        public void PasteInCell(string workbookname, int column, int row, string inputformula)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cell
                    Excel.Range range = sheet.Cells[row, column];
                    //Enter forumla
                    range.PasteSpecial();
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
        #region CLEAR SELECTED CELLS
        public void ClearSelectedCells(string workbookname)
        {
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Create object from selection
                    Excel.Range range = xlapp.ActiveWindow.Selection;
                    //Delete selected rows
                    range.ClearContents();

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
        #region CLEAR CELLS IN RANGE
        public void ClearCellsInRange(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto)
        {
            Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
            Excel.Range rngTo = xlapp.Cells[rowto, columnto];

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Create object from selection
                    Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                    //Delete selected rows
                    range.ClearContents();

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
        #region GET LAST ROW OF SPECIFIC COLUMN
        public int GetLastRowOfSpecificColumn(string workbookname, int column, int rowstart)
        {
            int rowcount = 0;
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select column

                    while (sheet.Cells[rowstart, column].value != null)
                    {
                        ++rowstart;
                        ++rowcount;
                    }
                }
            }
            return rowcount;
        }
        #endregion
        #region GET LAST COLUMN OF SPECIFIC ROW
        public int GetLastColumnOfSpecificRow(string workbookname, int row, int columnstart)
        {
            int columncount = 0;
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select column

                    while (sheet.Cells[row, columnstart].value != null)
                    {
                        ++columnstart;
                        ++columncount;
                    }
                }
            }
            return columncount;
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
            Excel.Workbook wkb = null;
            string wkbname;
            string newworkbook = "Unable to find new workbook name";

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
                Excel.Worksheet sheet = (Excel.Worksheet)wkb.ActiveSheet;
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
        #endregion
        #region COPY SHEET TO SPECIFIC WORKBOOK
        public string CopySheetToSpecificWorkbook(string workbooknameFROM, string workbooknameTO)
        {
            Excel.Application xlapp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            List<string> workbooklist = new List<string>();
            Excel.Workbook wkbFrom = null;
            Excel.Workbook wkbTo = null;
            string wkbname;

            foreach (Excel.Workbook activewkbs in xlapp.Workbooks)
            {
                wkbname = Convert.ToString(activewkbs.Name);
                workbooklist.Add(wkbname);
                if (activewkbs.Name == workbooknameFROM)
                {
                    wkbFrom = activewkbs;
                }
                else if (activewkbs.Name == workbooknameTO)
                {
                    wkbTo = activewkbs;
                }
            }

            if (wkbFrom != null && wkbTo != null)
            {
                Excel.Worksheet sheet = (Excel.Worksheet)wkbFrom.ActiveSheet;

                sheet.Copy(wkbTo.Worksheets[1]);

                return "Object copied";
            }
            else
            {
                return "Workbooks not available";
            }


        }
        #endregion

        #region CREATE PIVOT IN NEW TAB
        public string CreatePivotTableInNewTab(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto, string newsheetname, string pivotname, string pivotcelldesticnation, string[] rowfieldlist, string[] columnfieldlist, string[] valuefieldlist, string[] filterfieldlist)
        {
            Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
            Excel.Range rngTo = xlapp.Cells[rowto, columnto];

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.PivotCaches pCaches = workbook.PivotCaches();
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    Excel.Worksheet sheet2 = workbook.Worksheets.Add();
                    sheet2.Name = newsheetname;
                    //Select cells in a range
                    Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                    Excel.Range rngDes = sheet2.get_Range(pivotcelldesticnation);
                    //Send range to cache and use it to create pivot
                    Excel.PivotCache cache = pCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, range, Excel.XlPivotTableVersionList.xlPivotTableVersion14);
                    Excel.PivotTable pTable = cache.CreatePivotTable(TableDestination: rngDes, TableName: pivotname, DefaultVersion: Excel.XlPivotTableVersionList.xlPivotTableVersion14);


                    foreach (var rowfield in rowfieldlist)
                    {
                        pTable.PivotFields(rowfield).Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                    }

                    foreach (var columnfield in columnfieldlist)
                    {
                        pTable.PivotFields(columnfield).Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
                    }
                    foreach (var valuefield in valuefieldlist)
                    {
                        pTable.PivotFields(valuefield).Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                    }
                    foreach (var filterfield in filterfieldlist)
                    {
                        pTable.PivotFields(filterfield).Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                    }

                    return pTable.Name + " created";
                }
            }
            return "Pivot creatin failed";
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

                        return "Completed";
                    }
                }
            }
            return "Failed";
        }
        #endregion
        #region FILTER OUT VALUE PIVOT IN SELECTED SHEET
        public string FilterOutValuePivotInSelectedSheet(string workbookname, string pivotname, string[] filtervalues, string filterfield)
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
                            pivotfield.PivotItems(i).visible = false;
                        }
                        else
                        {
                            pivotfield.PivotItems(i).visible = true;
                        }

                        //Select cells in a range
                        //Send range to cache and use it to create pivot

                        return "Completed";
                    }
                }
            }
            return "Failed";
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
