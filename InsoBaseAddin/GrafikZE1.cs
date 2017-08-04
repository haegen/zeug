using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace InsoBaseAddin
{
    class GrafikZE1
    {
        public Excel.Worksheet Quelle { get; private set; }
        public Excel.Worksheet shAuswertung1 { get; private set; }
        public Excel.Worksheet shAuswertung2 { get; private set; }
        public Excel.Chart shAuswertung1Chart { get; private set; }
        public Excel.Chart shAuswertung2Chart { get; private set; }

        public bool IsSourceValid { get; private set; }
        public int IsMarked { get; private set; }

        private string colHeader1 = "Datum";
        private string colHeader2 = "Fällige Verbindlichkeiten der Woche";
        private string colHeader3 = "Darauf geleistete Zahlungen und ver. Überzahlung der Woche";

        private bool isColHeader1 = false;
        private bool isColHeader2 = false;
        private bool isColHeader3 = false;

        private int lastRow;
        private int lastColumn;

        // die verwendeten farben die in der ZE_Tabelle gesucht werden
        private Color c1 = Color.FromArgb(255, 192, 0);     // organge
        private Color c2 = Color.FromArgb(0, 176, 80);      // grün

        public GrafikZE1(Excel.Worksheet ws)
        {
            IsSourceValid = false;
            this.Quelle = ws;
            SetSourceValid();

            lastRow = Quelle.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            lastColumn = Quelle.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

            if (IsSourceValid)
            {
                IsMarked = isMarked(c1, c2);
            }
        }

        public void DiagrammErstellen()
        {
            int start = 0;
            int ende = 0;

            if (IsMarked == 3)
            {
                start = getFirstColoredRow(c1);
                ende = getLastColoredRow(c2);
                if (start < ende)
                {
                    CopyData(start, ende);

                    start = getFirstColoredRow(c2);
                    ende = getLastColoredRow(c2);

                    CopyData(start, ende);
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Die orange markierung darf nur vor der grünen markierung kommen.");
                }
            }
            if (IsMarked == 1)
            {
                start = getFirstColoredRow(c1);
                ende = getLastColoredRow(c1);

                CopyData(start, ende);
            }
            if (IsMarked == 2)
            {
                start = getFirstColoredRow(c2);
                ende = getLastColoredRow(c2);

                CopyData(start, ende);
            }

            Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[4].Move(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[3]);
            Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[4].Move(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[3]);
            Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[4].Activate();
        }

        private void CopyData(int start, int ende)
        {
            int range = ende - start + 1;
            shAuswertung1 = AddWorksheet("Tab_ZE_" + range + "W");
            Quelle.Range["A1"].EntireRow.Copy();
            shAuswertung1.Cells[1, 1].PasteSpecial(Excel.XlPasteType.xlPasteAllUsingSourceTheme);
            Quelle.Range[Quelle.Cells[start, 1], Quelle.Cells[ende, lastColumn]].Copy();
            shAuswertung1.Cells[2, 1].PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats);
            int rowIndex = ende - start + 2;

            shAuswertung1Chart = AddChart();
            EditChart(shAuswertung1Chart, rowIndex);
            shAuswertung1Chart.Location(Excel.XlChartLocation.xlLocationAsNewSheet).Name = "Grafik_ZE_I_" + range + "W"; // jetzt wrid aus dem chart ein sheet!
        }

        private void EditChart(Excel.Chart chart, int index)
        {
            chart.ChartType = Excel.XlChartType.xlColumnClustered;

            var range1 = shAuswertung1.Range[shAuswertung1.Cells[1, 1], shAuswertung1.Cells[index, 1]];
            var range2 = shAuswertung1.Range[shAuswertung1.Cells[1, 5], shAuswertung1.Cells[index, 5]];
            var range3 = shAuswertung1.Range[shAuswertung1.Cells[1, 9], shAuswertung1.Cells[index, 9]];
            var range4 = shAuswertung1.Range[shAuswertung1.Cells[1, 10], shAuswertung1.Cells[index, 10]];

            chart.SetSourceData(Source: Globals.ThisAddIn.Application.Union(range1, range2, range3, range4));
            chart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;
            chart.Legend.Font.Size = 8;
            chart.Legend.Font.Name = "Arial";
            chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = Color.FromArgb(192, 0, 0);
            chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = Color.FromArgb(0, 112, 192);
            chart.SeriesCollection(3).Format.Fill.ForeColor.RGB = Color.FromArgb(255, 140, 0);
            chart.Axes(Excel.XlAxisType.xlCategory).CategoryType = Excel.XlCategoryType.xlCategoryScale;
            chart.Axes(Excel.XlAxisType.xlCategory).TickLabels.Font.Size = 8;
            chart.Axes(Excel.XlAxisType.xlCategory).TickLabels.Font.Name = "Arial";
            chart.Axes(Excel.XlAxisType.xlCategory).TickLabels.Orientation = 90;
            chart.Axes(Excel.XlAxisType.xlValue).TickLabels.Font.Size = 8;
            chart.Axes(Excel.XlAxisType.xlValue).TickLabels.Font.Name = "Arial";
            chart.Axes(Excel.XlAxisType.xlValue).MinimumScale = 0;
        }

        private Excel.Chart AddChart()
        {
            Excel.ChartObjects charts = (Excel.ChartObjects)Quelle.ChartObjects(Type.Missing);
            Excel.Chart newCh = charts.Add(50, 50, 500, 300).Chart;

            return newCh;
        }

        private Excel.Worksheet AddWorksheet(string name)
        {
            Excel.Worksheet newWs = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(After: Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count]);
            newWs.Name = name;

            return newWs;
        }

        private void SetSourceValid()
        {
            var cell1 = Quelle.Cells[1, 1];
            var cell2 = Quelle.Cells[1, 5];
            var cell3 = Quelle.Cells[1, 9];

            if (cell1.Value == colHeader1)
                isColHeader1 = true;
            if (cell2.Value == colHeader2)
                isColHeader2 = true;
            if (cell3.Value == colHeader3)
                isColHeader3 = true;

            if (isColHeader1 && isColHeader2 && isColHeader3)
                IsSourceValid = true;
        }

        /// <summary>
        /// Methode überprüft ob in dem Tabellenblatt die übergebenen Farben vorkommen.
        /// </summary>
        /// <param name="rgb1">Farbwert der gesucht wird.</param>
        /// <param name="rgb2">Farbwert der gesucht wird.</param>
        /// <returns>gibt 0 zurück wenn keine Farbe gefunden wurde
        ///     1 wenn NUR der erste Farbwert vorhanden ist
        ///     2 wenn NUR der zweite Farbwert vorhanden ist und
        ///     3 wenn beide Farbwerte vorhanden sind</returns>
        private int isMarked(Color rgb1, Color rgb2)
        {
            int result = 0;
            bool isRgb1 = false;
            bool isRgb2 = false;

            for (int counter = 2; counter <= lastRow; counter++)
            {
                Color interiorColor = ColorTranslator.FromOle((int)Quelle.Cells[counter, 1].Interior.Color);

                if (interiorColor == rgb1)
                {
                    isRgb1 = true;
                    continue;
                }

                if(interiorColor == rgb2)
                {
                    isRgb2 = true;
                    continue;
                }
            }

            if (isRgb1 && isRgb2)
                result = 3;
            else
            {
                if (isRgb2)
                    result = 2;
                if (isRgb1)
                    result = 1;
            }

            return result;
        }

        private int getFirstColoredRow(Color rgb)
        {
            int firstRow = 0;

            for (int counter = 2; counter <= lastRow; counter++)
            {
                Color interiorColor = ColorTranslator.FromOle((int)Quelle.Cells[counter, 1].Interior.Color);

                if (interiorColor == rgb)
                {
                    firstRow = counter;
                    break;
                }
            }

            return firstRow;
        }

        private int getLastColoredRow(Color rgb)
        {
            int lastRow = 0;

            for (int counter = this.lastRow; counter >= 2; counter--)
            {
                Color interiorColor = ColorTranslator.FromOle((int)Quelle.Cells[counter, 1].Interior.Color);

                if (interiorColor == rgb)
                {
                    lastRow = counter;
                    break;
                }
            }

            return lastRow;
        }
    }
}
