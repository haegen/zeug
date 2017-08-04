using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using Microsoft.Office.Tools.Excel;

namespace InsoBaseAddin
{
    public static class MyFormat
    {

        public static void SetCustomPageHeader(Excel.Worksheet ws, string name)
        {
            ws.PageSetup.CenterHeader = "&\"Arial\" &B &11 " + name;
        }

        public static void SetBorder(Excel.Worksheet ws)
        {
            Excel.Range range = ws.UsedRange;
            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.Color = Color.Black;
        }

        public static void SetBorders(Excel.Workbook wb)
        {
            for (int sheetIndex = 1; sheetIndex <= wb.Worksheets.Count; sheetIndex++)
            {
                if (wb.Worksheets[sheetIndex].Name != "Tabelle1")
                {
                    Excel.Worksheet ws = wb.Worksheets[sheetIndex];
                    SetBorder(ws);
                }
            }
        }

        public static void FormatHeader(Excel.Worksheet ws, bool autofilter)
        {
            int columnCount = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
            var lastColum = ws.Cells[1, columnCount];

            Excel.Range header = (Excel.Range)ws.Range["A1", lastColum];
            header.Font.Name = "Arial";
            header.Font.Size = 10;
            header.Font.Bold = true;
            header.Interior.Color = Color.FromArgb(200, 200, 200);
            header.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            header.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
            header.RowHeight = header.RowHeight * 3 - 5.5;

            if (!ws.AutoFilterMode && autofilter)
                ws.Rows[1].AutoFilter();

            header.WrapText = true;
        }

        public static void FormatHeaders(Excel.Workbook wb, bool autofilter)
        {
            for (int sheetIndex = 1; sheetIndex <= wb.Worksheets.Count; sheetIndex++)
            {
                if (wb.Worksheets[sheetIndex].Name != "Tabelle1")
                {
                    Excel.Worksheet ws = wb.Worksheets[sheetIndex];
                    FormatHeader(ws, autofilter);
                }
            }
        }

        public static void FormatTableData(Excel.Worksheet ws)
        {
            int rowCount = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            int columnCount = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

            var col1 = ws.Cells[2, 1];
            var col2 = ws.Cells[rowCount, columnCount];

            Excel.Range tableData = ws.Range[col1, col2];
            tableData.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            tableData.Font.Name = "Arial";
            tableData.Font.Size = 9;
            tableData.Interior.Color = Color.FromArgb(255, 255, 255);

            for (int i = 1; i <= columnCount; i++)
            {
                ws.Columns[i].AutoFit();
                ws.Columns[i].ColumnWidth += 1;
            }
        }

        public static void FormatTableDatas(Excel.Workbook wb)
        {
            for (int sheetIndex = 1; sheetIndex <= wb.Worksheets.Count; sheetIndex++)
            {
                if (wb.Worksheets[sheetIndex].Name != "Tabelle1")
                {
                    Excel.Worksheet ws = wb.Worksheets[sheetIndex];
                    FormatTableData(ws);
                }
            }
        }

        public static void PageSetup(Excel.Worksheet ws, string verfahrensname)
        {
            ws.Activate();
            Globals.ThisAddIn.Application.ActiveWindow.View = Excel.XlWindowView.xlPageLayoutView;
            ws.PageSetup.CenterHeader = "&\"Arial\" &B &11 " + verfahrensname;
            ws.PageSetup.LeftMargin = 23;
            ws.PageSetup.RightMargin = 23;
            ws.PageSetup.TopMargin = 71;
            ws.PageSetup.BottomMargin = 57;
            ws.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
            ws.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
            ws.PageSetup.FirstPageNumber = (int)Excel.Constants.xlAutomatic;
            ws.PageSetup.Order = Excel.XlOrder.xlDownThenOver;
            ws.PageSetup.Zoom = false;
            ws.PageSetup.PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed;
            ws.PageSetup.ScaleWithDocHeaderFooter = true;
            ws.PageSetup.AlignMarginsHeaderFooter = true;
            ws.PageSetup.CenterFooter = "&\"Arial\" &B &11 Seite &P von &N";
            ws.PageSetup.RightFooterPicture.Filename = @"\\SSTOR01\Inso\Vorlagen\Logo\KDLB Kirstein\KDLB_Kirstein_Logo_neu.jpg";
            ws.PageSetup.RightFooter = "&G";
            ws.PageSetup.PrintTitleRows = "$1:$1";
            ws.PageSetup.CenterHorizontally = true;
            ws.PageSetup.PrintComments = Excel.XlPrintLocation.xlPrintNoComments;
            ws.PageSetup.FitToPagesWide = 1;
            ws.PageSetup.FitToPagesTall = false;
        }

        public static void PageSetups(Excel.Workbook wb, string verfahrensname)
        {
            for (int sheetIndex = 1; sheetIndex <= wb.Worksheets.Count; sheetIndex++)
            {
                if (wb.Worksheets[sheetIndex].Name != "Tabelle1")
                {
                    Excel.Worksheet ws = wb.Worksheets[sheetIndex];
                    PageSetup(ws, verfahrensname);
                }
            }
        }
    }
}
