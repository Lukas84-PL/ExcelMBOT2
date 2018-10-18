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

namespace ExcelDataManipulation
{
    public class ExcelMBOT
    {

        Excel.Application xlapp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");

        #region ACTIONS
        #region SAVEAS WORKBOOK
        public string SaveAs(string workbookname, string newfilename = "")
        {
            try
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
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region SELECT RANGE
        public string SelectRange(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto)
        {
            try
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
                    range.Activate();
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region COPY RANGE
        public string CopyRange(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region COPY SELECTION
        public string CopySelection(string workbookname)
        {
            try
            {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = xlapp.ActiveWindow.Selection;
                    range.Copy();
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region COPY COLUMN
        public string CopyColumn(string workbookname, string column)
        {
            try
            {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = (Excel.Range)sheet.Columns[column + ":" + column];
                    range.Copy();
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion
        #region COPY ROW
        public string CopyRow(string workbookname, int row)
        {
            try
            {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = (Excel.Range)sheet.Rows[row + ":" + row];
                    range.Copy();
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region COPY AND PASTE VALUES OFCOLUMN
        public string CopyAndPasteValuesOfColumn(string workbookname, string columnfrom, string columnto)
        {
            try
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
                    return "Workbook found";

                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region COPY AND PASTE VALUES OF ROW
        public string CopyAndPasteValuesOfRow(string workbookname, int rowfrom, int rowto)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region COPY AND PASTE COLUMN
        public string CopyAndPasteColumn(string workbookname, string columnfrom, string columnto)
        {
            try
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
                    return "Workbook found";

                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {
                return e.ToString();
            }
}
        #endregion
        #region COPY AND PASTE ROW
        public string CopyAndPasteRow(string workbookname, int rowfrom, int rowto)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion

        #region INSERT FORMULA
        public string InsertFormula(string workbookname, int column, int row, string inputformula)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region REPLACE DATA IN RANGE
        public string ReplaceDataInRange(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto, string oldstring, string newstring)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region REPLACE DATA IN SELECTION
        public string ReplaceDataInSelection(string workbookname, string oldstring, string newstring)
        {
            try
            {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = xlapp.ActiveWindow.Selection;
                    range.Replace(oldstring, newstring);
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region CHANGE FONT IN RANGE TO BOLD
        public string ChangeFontInRangeToBold(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region CHANGE FONT IN SELECTION TO BOLD
        public string ChangeFontInSelectionToBold(string workbookname)
        {
            try
            {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = xlapp.ActiveWindow.Selection;
                    range.Font.Bold = true;
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }
        #endregion

        #region PASTE VALUES IN SELECTION
        public string PasteValuesInSelection(string workbookname, string oldstring, string newstring)
        {
            try
            {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = xlapp.ActiveWindow.Selection;
                    range.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region PASTE VALUES IN CELL
        public string PasteValuesInCell(string workbookname, int column, int row, string inputformula)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region PASTE IN SELECTION
        public string PasteInSelection(string workbookname, string oldstring, string newstring)
        {
            try
            {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range
                    Excel.Range range = xlapp.ActiveWindow.Selection;
                    range.PasteSpecial();
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region PASTE IN CELL
        public string PasteInCell(string workbookname, int column, int row, string inputformula)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region SELECT BLANK CELLS OF SELECTION
        public string SelectBlankCellsOfSelection(string workbookname)
        {
            try
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
                    return "Workbook found";

                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region SELECT BLANK CELLS OF RANGE
        public string SelectBlankCellsOfRange(string workbookname, int column, int rowfrom, int rowto)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region SELECT BLANK CELLS OF COLUMN
        public string SelectBlankCellsOfColumn(string workbookname, string column)
        {
            try
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
                    return "Workbook found";

                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region SELECT BLANK CELLS OF ROW
        public string SelectBlankCellsOfRow(string workbookname, int row)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region DELETE BLANK ROWS OF SELECTION
        public string DeleteBlankRowsInSelection(string workbookname)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region DELETE BLANK ROWS OF RANGE
        public string DeleteBlankRowsOfRange(string workbookname, int column, int rowfrom, int rowto)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region DELETE BLANK ROWS OF COLUMN
        public string DeleteBlankRowsOfColumn(string workbookname, string column)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region DELETE BLANK COLUMNS IN A SELECTION
        public string DeleteBlankColumnsOfSelection(string workbookname)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region DELETE BLANK COLUMNS IN A RANGE
        public string DeleteBlankColumnsOfRange(string workbookname, int columnfrom, int columnto, int row)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region DELETE BLANK COLUMNS OF ROW
        public string DeleteBlankColumnsOfRow(string workbookname, int row)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion


        #region DELETE ROWS OF SELECTED CELLS
        public string DeleteRowsOfSelectedCells(string workbookname)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region CLEAR SELECTED CELLS
        public string ClearSelectedCells(string workbookname)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region DELETE COLUMNS OF SELECTED CELLS
        public string DeleteColumnsOfSelectedCells(string workbookname)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region CLEAR CELLS IN RANGE
        public string ClearCellsInRange(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region SELECT ROW
        public string SelectEntireRow(string workbookname, int row)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region SELECT COLUMN
        public string SelectEntireColumn(string workbookname, string column)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region INSERT ROW IN SELECTION
        public string InsertRowInSelection(string workbookname)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region INSERT ROW IN RANGE
        public string InsertRowInRange(string workbookname, int row)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region INSERT COLUMN IN SELECTION
        public string InsertColumnInSelection(string workbookname)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region INSERT COLUMN IN RANGE
        public string InsertColumnInRange(string workbookname, string column)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

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
        public string GetLastRowOfSpecificColumn(string workbookname, int column, int rowstart)
        {
            try
            {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select column

                    while (sheet.Cells[rowstart +1, column].value != null || sheet.Cells[rowstart +2, column].value != null)
                    {
                        ++rowstart;
                    }

                }
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
        public string GetLastColumnOfSpecificRow(string workbookname, int row, int columnstart)
        {
            try
            {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select column

                    while (sheet.Cells[row, columnstart +1].value != null || sheet.Cells[row, columnstart +2].value != null)
                    {
                        ++columnstart;
                    }

                }
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
        public string AutofitColumn(string workbookname, string column)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region AUTOFIT ROW
        public string AutofitRow(string workbookname, int row)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region DELETE SHEET
        public string DeleteSheet(string workbookname)
        {
            try
            {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Delete sheet
                    sheet.Delete();
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region COPY SHEET TO NEW WORKBOOK
        public string CopySheetToNewWorkbook(string workbookname)
        {
            try
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
            catch (Exception e)
            {

                return e.ToString();
            }

        }
        #endregion
        #region COPY SHEET TO SPECIFIC WORKBOOK
        public string CopySheetToSpecificWorkbook(string workbooknameFROM, string workbooknameTO)
        {
            try
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
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region CREATE PIVOT IN NEW TAB
        public string CreatePivotTableInNewTab(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto, string newsheetname, string pivotname, string pivotcelldesticnation, string[] rowfieldlist, string[] columnfieldlist, string[] valuefieldlist, string[] filterfieldlist)
        {
            try
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
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region FILTER ON VALUE PIVOT IN SELECTED SHEET
        public string FilterOnValuePivotInSelectedSheet(string workbookname, string pivotname, string[] filtervalues, string filterfield)
        {
            try
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
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region FILTER OUT VALUE PIVOT IN SELECTED SHEET
        public string FilterOutValuePivotInSelectedSheet(string workbookname, string pivotname, string[] filtervalues, string filterfield)
        {
            try
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
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region FILTER ON VALUES
        public string FilterOutValues(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto, int filtercolumn, string[] filterlist)
        {
            try
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
                    range.AutoFilter(filtercolumn, filterlist,
                                     Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                        return "Workbook found";
                }
            }
                return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region CHANGE RANGE FONT COLOR USING HTML
        public string ChangeRangeFontColor(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto, string color)
        {
            try
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

                    range.Font.Color = System.Drawing.ColorTranslator.FromHtml(color);
                    return "Workbook found";
                    }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region CHANGE RANGE FILL COLOR USING HTML
        public string ChangeRangeFillColorHTML(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto, string color)
        {
            try
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

                    range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(color);
                    return "Workbook found";
                }
            }
            return "Workbook not found";
            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region SET TEXT TO COLUMNS OF RANGE
        public string TextToColumns(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto, string delimiter)
        {
            try
            {

            Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
            Excel.Range rngTo = xlapp.Cells[rowto, columnto];

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {

                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select cells in a range

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
            }
            return "Workbook not found";
            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region CHANGE FORMAT OF RANGE
        public string ChangeFormatOfRange(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto, string format)
        {
            try
            {

            Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
            Excel.Range rngTo = xlapp.Cells[rowto, columnto];

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                
                if (workbook.Name == workbookname)
                {
                    if (format == "@")
                    {
                        Array fieldInfoArray = new int[,] { { 1, 2 } };
                        Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
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
                        Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                        //Select cells in a range

                        Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                        range.TextToColumns();
                        range.NumberFormat = format;
                        range.TextToColumns();
                    }
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region SELECT VISIBLE CELLS IN RANGE
        public string SelectVisibleCellsInRange(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto)
        {
            try
            {

            Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
            Excel.Range rngTo = xlapp.Cells[rowto, columnto];

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;

                    Excel.Range range = sheet.get_Range(rngFrom, rngTo);
                    range.SpecialCells(XlCellType.xlCellTypeVisible).Select();
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region SORT RANGE
        public string SortRange(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto, int sortcolumnnumber)
        {
            try
            {

            Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
            Excel.Range rngTo = xlapp.Cells[rowto, columnto];
            Excel.Range rngsortFrom = xlapp.Cells[rowfrom, sortcolumnnumber];
            Excel.Range rngsortTo = xlapp.Cells[rowto, sortcolumnnumber];

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
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
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region OPEN SPREADSHEET
        public string OpenSpreadsheet(string workbookname, string path)
        {
            try
            {


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
        public string InsertObject(string workbookname, int column, int row, string filepath, string iconname, int iconindex, int iconwidth = 10)
        {
            try
            {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
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
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region DRAG CELL VALUE TO RANGE
        public string DragCellValueToRange(string workbookname, int column, int rowfrom, int rowto)
        {
            try
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
                    return "Workbook found";
                 }
            }
            return "Workbook not found";
            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region FILTER OUT VALUES
        public string FilterOnValues(string workbookname, int columnfrom, int rowfrom, int columnto, int rowto, int filtercolumn, string[] filterlist)
        {
            try
            {

            Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
            Excel.Range rngTo = xlapp.Cells[rowto, columnto];

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    Excel.Range range = sheet.get_Range(rngFrom, rngTo);

                    foreach (var item in filterlist)
                    {
                        range.AutoFilter(filtercolumn, "<>" + item,
                        Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                    }
                    return "Workbook found";

                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region REMOVE DUPLICATES FROM RANGE
        public string RemoveDuplicatesFromRange(string workbookname, int rowfrom, int columnfrom, int rowto, int columnto)
        {
            try
            {

            Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
            Excel.Range rngTo = xlapp.Cells[rowto, columnto];

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;

                    Range range = sheet.get_Range(rngFrom, rngTo);

                    range.RemoveDuplicates(
                        Excel.XlYesNoGuess.xlNo);
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region REMOVE DUPLICATES FROM COLUMNS
        public string RemoveDuplicatesFromColumns(string workbookname, string columnfrom, string columnto)
        {
            try
            {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;

                    Excel.Range range = (Excel.Range)sheet.Columns[columnfrom + ":" + columnto];

                    range.RemoveDuplicates(
                        Excel.XlYesNoGuess.xlNo);
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region SELECT VALUE IN RANGE
        public string SelectValueInRange(string workbookname, int rowfrom, int columnfrom, int rowto, int columnto, string searchvalue)
        {
            try
            {

            Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
            Excel.Range rngTo = xlapp.Cells[rowto, columnto];

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;

                    Range range = sheet.get_Range(rngFrom, rngTo);

                    range.Find(searchvalue).Select();
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region FIND VALUE IN RANGE AND GET ADRESS
        public string GetAdressOfValue(string workbookname, int rowfrom, int columnfrom, int rowto, int columnto, string searchvalue)
        {
            try
            {

            Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
            Excel.Range rngTo = xlapp.Cells[rowto, columnto];

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;

                    Range range = sheet.get_Range(rngFrom, rngTo);

                    Range adressrange = range.Find(searchvalue);

                    return adressrange.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1).ToString();


                }
            }
            return "not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region REPLACE VALUES IN RANGE
        public string ReplaceValuesInRange(string workbookname, int rowfrom, int columnfrom, int rowto, int columnto, string replace, string replaceto)
        {
            try
            {

            Excel.Range rngFrom = xlapp.Cells[rowfrom, columnfrom];
            Excel.Range rngTo = xlapp.Cells[rowto, columnto];

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;

                    Range range = sheet.get_Range(rngFrom, rngTo);

                    sheet.Cells.Replace(replace, replaceto);
                    return "Workbook found";

                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region GO TO LAST ROW OF SPECIFIC COLUMN
        public string GoToLastRowOfSpecificColumn(string workbookname, int column, int rowstart)
        {
            try
            {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    while (sheet.Cells[rowstart +1, column].value != null || sheet.Cells[rowstart +2, column].value != null)
                    {
                        ++rowstart;
                    }

                    Excel.Range lastcell = xlapp.Cells[rowstart, column];

                    lastcell.Activate();
                    lastcell.Select();

                    return "Workbook found";
                }
            }
            return "Workbook not found";


            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region GO TO LAST COLUMN OF SPECIFIC ROW
        public string GoToLastColumnOfSpecificRow(string workbookname, int row, int columnstart)
        {
            try
            {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    //Select column

                    while (sheet.Cells[row, columnstart + 1].value != null || sheet.Cells[row, columnstart + 2].value != null)
                    {
                        ++columnstart;
                    }

                    Excel.Range lastcell = xlapp.Cells[row, columnstart];

                    lastcell.Activate();
                    lastcell.Select();
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region GO TO LAST COLUMN OF USED RANGE
        public string GoToLastColumnOfUsedRange(string workbookname, int row)
        {
            try
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
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region GO TO LAST ROW OF USED RANGE
        public string GoToLastRowOfUsedRange(string workbookname, int column)
        {
            try
            {

            int rowcount = 0;
            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.ActiveSheet;
                    rowcount = sheet.UsedRange.Rows.Count;
                    Range lastrow = xlapp.Cells[rowcount, column];
                    lastrow.Select();
                    lastrow.Activate();
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region CLOSE SPREADSHEET WITH SAVING
        public string CloseSpreadsheetWithSaving(string workbookname)
        {
            try
            {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    workbook.Close(true);
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region CLOSE SPREADSHEET WITHOUT SAVING
        public string CloseSpreadsheetWithoutSaving(string workbookname)
        {
            try
            {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    workbook.Close(false);
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion
        #region QUIT EXCEL APP
        public string QuitExcelApp(string workbookname)
        {
            try
            {

            foreach (Excel.Workbook workbook in xlapp.Workbooks)
            {
                if (workbook.Name == workbookname)
                {
                    xlapp.Quit();
                    return "Workbook found";
                }
            }
            return "Workbook not found";

            }
            catch (Exception e)
            {

                return e.ToString();
            }
        }
        #endregion

        #region OBTAIN DATA FROM EXCEL

        #region GET: SHEETNAME, COLUMN COUNT, ROW COUNT, CHANGE SHEET NAME, SHEET COUNT
        private Tuple<int?, int?, string, int?> ObtainExcelData(string workbookname, string newsheetname = "")
        {
            try
            {

            string worksheetname = "Not found";
            int? rowcount = 0;
            int? columncount = 0;
            int? worksheetcount = 0;

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