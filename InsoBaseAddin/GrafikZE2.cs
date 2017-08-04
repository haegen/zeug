using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;


namespace InsoBaseAddin
{
    class GrafikZE2
    {
        public Excel.Worksheet Quelle { get; private set; }
        public Excel.Worksheet shAuswertung { get; private set; }

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

        public GrafikZE2(Excel.Worksheet ws)
        {
            IsSourceValid = false;
            this.Quelle = ws;
            SetSourceValid();

            lastRow = Quelle.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Row;
            lastColumn = Quelle.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Column;

            if (IsSourceValid)
            {
                IsMarked = isMarked(c1, c2);
            }
        }

        public void DiagrammErstellen()
        {
            shAuswertung = AddWorksheet("Tab_ZE_2");

            // Überschrift einfügen
            Quelle.Range["A1"].EntireRow.Copy();
            shAuswertung.Cells[1, 1].PasteSpecial(Excel.XlPasteType.xlPasteAllUsingSourceTheme);

            // letzte farbige Zeile ermitteln
            int row1 = getLastColoredRow(c1);
            int row2 = getLastColoredRow(c2);
            int lastColoredRow = 0;

            if (row1 > row2)
                lastColoredRow = row1;
            else
                lastColoredRow = row2;

            // Daten kopieren
            var cell1 = Quelle.Cells[lastColoredRow + 1, 1];
            var cell2 = Quelle.Cells[lastRow, lastColumn];

            Quelle.Range[cell1, cell2].Copy();
            shAuswertung.Cells[2, 1].PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats);


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

        private Excel.Worksheet AddWorksheet(string name)
        {
            Excel.Worksheet newWs = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(After: Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count]);
            newWs.Name = name;

            return newWs;
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

                if (interiorColor == rgb2)
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
    }
}
