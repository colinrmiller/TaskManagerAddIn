using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using TaskManager;
using Excel = Microsoft.Office.Interop.Excel;

namespace TaskManagemer
{
    public partial class TaskRibbon : RibbonBase
    {
        private Excel.Application excelApp;

        //public TaskRibbon()
        //{
        //    InitializeComponent();
        //}

        private void TaskRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
        }

        private void btnTaskItem_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selection = excelApp.Selection;
            selection.End[Excel.XlDirection.xlToLeft].Select();
            Excel.Range activeCell = excelApp.ActiveCell;
            Excel.Range newRange = activeCell.Range["A1:D1"];
            newRange.Select();
            newRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

            newRange.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            newRange.Interior.PatternColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
            newRange.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorLight2;
            newRange.Interior.TintAndShade = 0.749992370372631;

            activeCell.Offset[0, 3].Select();
            int lastColumnIndex = excelApp.ActiveCell.Column;

            Excel.Range checkboxCell = excelApp.ActiveCell;
            checkboxCell.ClearContents();
            checkboxCell.Value = "FALSE";  // This simulates a checkbox

            checkboxCell.Offset[0, -1].Value = DateTime.Now.ToShortDateString();

            excelApp.Selection.End[Excel.XlDirection.xlToLeft].Select();

            AddAllBordersToRow(lastColumnIndex);
        }

        private void btnNewTask_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selection = excelApp.Selection;
            selection.End[Excel.XlDirection.xlDown].Select();
            selection.End[Excel.XlDirection.xlUp].Select();
            Excel.Range activeCell = excelApp.ActiveCell;
            Excel.Range newRange = activeCell.Offset[2, 0].Range["A1:D1"];
            newRange.Select();

            newRange.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            newRange.Interior.PatternColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
            newRange.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorLight2;
            newRange.Interior.TintAndShade = 0.499984740745262;

            activeCell.Offset[3, 0].Range["A1:D1"].Select();
            Excel.Range secondRow = excelApp.Selection;

            secondRow.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            secondRow.Interior.PatternColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
            secondRow.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorLight2;
            secondRow.Interior.TintAndShade = 0.749992370372631;

            activeCell.Offset[2, 3].Select();
            Excel.Range checkboxCell1 = excelApp.ActiveCell;
            checkboxCell1.ClearContents();
            checkboxCell1.Value = "FALSE";  // This simulates a checkbox

            activeCell.Offset[3, 3].Select();
            Excel.Range checkboxCell2 = excelApp.ActiveCell;
            checkboxCell2.ClearContents();
            checkboxCell2.Value = "FALSE";  // This simulates a checkbox

            excelApp.Selection.End[Excel.XlDirection.xlToLeft].Select();
            activeCell.Offset[2, 0].Select();
        }

        private void btnDeleteRow_Click(object sender, RibbonControlEventArgs e)
        {
            excelApp.ActiveCell.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
        }

        private void AddAllBordersToRow(int endingColumn)
        {
            Excel.Range selection = excelApp.Selection;
            selection.End[Excel.XlDirection.xlToLeft].Select();
            Excel.Range rangeToFormat = excelApp.Range[excelApp.ActiveCell, excelApp.ActiveCell.Offset[0, endingColumn - 1]];
            rangeToFormat.Select();

            Excel.Borders borders = rangeToFormat.Borders;
            borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;

            Excel.Border[] borderArray = new Excel.Border[]
            {
                borders[Excel.XlBordersIndex.xlEdgeLeft],
                borders[Excel.XlBordersIndex.xlEdgeTop],
                borders[Excel.XlBordersIndex.xlEdgeBottom],
                borders[Excel.XlBordersIndex.xlEdgeRight],
                borders[Excel.XlBordersIndex.xlInsideVertical],
                borders[Excel.XlBordersIndex.xlInsideHorizontal]
            };

            foreach (Excel.Border border in borderArray)
            {
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.ThemeColor = 3;
                border.TintAndShade = -0.499984740745262;
                border.Weight = Excel.XlBorderWeight.xlThin;
            }
        }
    }
}