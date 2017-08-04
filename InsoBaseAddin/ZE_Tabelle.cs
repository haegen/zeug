using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace InsoBaseAddin
{
    class ZE_Tabelle
    {
        private Excel.Workbook workbook;

        public Excel.Worksheet Quelle { get; private set; }
        public Excel.Worksheet ZE { get; private set; }
        public Excel.Worksheet Quelle_Kopie { get; private set; }

        public bool IsSourceValid { get; private set; }

        private string colHeader1 = "Datum";
        private string colHeader2 = "Soll";
        private string colHeader3 = "Haben";
        private string colHeader4 = "Haben(fällig)";
        private string colHeader5 = "Buchsaldo";
        private string colHeader6 = "Fälligkeitssaldo";
        private string colHeader7 = "überfällige Verbindlichkeiten";
        private string colHeader8 = "Überzahlung";

        private bool isColHeader1 = false;
        private bool isColHeader2 = false;
        private bool isColHeader3 = false;
        private bool isColHeader4 = false;
        private bool isColHeader5 = false;
        private bool isColHeader6 = false;
        private bool isColHeader7 = false;
        private bool isColHeader8 = false;

        private int lastColumn;
        private int lastRow;

        private string verfahrensname;

        public ZE_Tabelle(Excel.Workbook wb)
        {
            IsSourceValid = false;
            this.workbook = wb;

            // setze zweites Tabellenblatt als Quelle
            this.Quelle = wb.ActiveSheet;

            SetSourceValid();

            verfahrensname = Quelle.PageSetup.CenterHeader;
        }

        public void Format()
        {
            MyFormat.FormatHeaders(workbook, true);
            MyFormat.FormatTableDatas(workbook);
            MyFormat.SetBorders(workbook);
            MyFormat.PageSetups(workbook, verfahrensname);

            ZE.Columns["C"].Hidden = true;
            ZE.Columns["F:H"].Hidden = true;
            ZE.Columns["L:N"].Hidden = true;
            ZE.Columns["O"].Delete();
        }

        public void EditZESheet()
        {
            ZE.Columns["E"].Delete();

            ZE.Cells[1, 1].Value = "Datum";
            ZE.Cells[1, 2].Value = "Fällige Verbindlichkeiten Beginn der Woche";
            ZE.Cells[1, 3].Value = "Fällig werdende Rechnungen der Woche(original)";
            ZE.Cells[1, 4].Value = "Fällige Verbindlichkeiten der Woche";
            ZE.Cells[1, 5].Value = "Darauf geleistete Zahlungen der Woche";
            ZE.Cells[1, 6].Value = "Überzahlung";
            ZE.Cells[1, 7].Value = "Verrechnete Überzahlung";
            ZE.Cells[1, 8].Value = "Darauf geleistete Zahlungen und ver. Überzahlung der Woche";
            ZE.Cells[1, 9].Value = "Nicht gedeckte fällige Verbindlichkeiten Ende der Woche";
            ZE.Cells[1, 10].Value = "Gedeckte fällige Verbindlichkeiten der Woche in Prozent";
            ZE.Cells[1, 11].Value = "Überfällige Verbindlichkeiten in Prozent";
            ZE.Cells[1, 12].Value = "ø Rechnungseingang im Jahr";
            ZE.Cells[1, 13].Value = "Fällige Verbindlichkeiten Ende der Woche";
            ZE.Cells[1, 14].Value = "Fällige Rechnungen der Woche";

            int counter = 2;

            while (lastRow >= counter)
            {
                ZE.Cells[counter, 4].FormulaR1C1 = "=RC[-2] + RC[-1]";

                if (counter == 2)
                {
                    ZE.Cells[counter, 7].FormulaR1C1 = "=RC[-1]";
                }
                else
                {
                    ZE.Cells[counter, 7].FormulaR1C1 = "=RC[-1] - R[-1]C[-1]";
                }

                ZE.Cells[counter, 8].FormulaR1C1 = "=RC[-3] - RC[-1]";
                ZE.Cells[counter, 9].FormulaR1C1 = "=RC[-5] - RC[-1]";
                ZE.Cells[counter, 10].FormulaR1C1 = "=IF(RC[-6] = 0, 0, IF(((RC[-5] / RC[-6])) < 0, 0, (RC[-5] / RC[-6])))";
                ZE.Cells[counter, 11].FormulaR1C1 = "=IF(RC[-2] > 0.001, 1, 0";
                ZE.Cells[counter, 14].FormulaR1C1 = "=IF(RC[-7] < 0, RC[-11] + RC[-7], RC[-11])";

                counter++;
            }

            ZE.Columns["N"].Cut();
            ZE.Columns["D"].Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

            ZE.Columns["K:L"].NumberFormat = "0,00%";
            ZE.Columns["B:J"].NumberFormat = "#.##0,00 €";
            ZE.Columns["M:N"].NumberFormat = "#.##0,00 €";
        }

        public void CopyZEData()
        {
            Quelle.Range[Quelle.Cells[2, 1], Quelle.Cells[lastRow, 1]].Copy();
            ZE.Cells[2, 1].PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            Quelle.Range[Quelle.Cells[2, 7], Quelle.Cells[lastRow, 7]].Copy();
            ZE.Cells[3, 2].PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            Quelle.Range[Quelle.Cells[2, 7], Quelle.Cells[lastRow, 7]].Copy();
            ZE.Cells[2, 14].PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            Quelle.Range[Quelle.Cells[2, 4], Quelle.Cells[lastRow, 4]].Copy();
            ZE.Cells[2, 3].PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            Quelle.Range[Quelle.Cells[2, 2], Quelle.Cells[lastRow, 2]].Copy();
            ZE.Cells[2, 6].PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            Quelle.Range[Quelle.Cells[2, 8], Quelle.Cells[lastRow, 8]].Copy();
            ZE.Cells[2, 7].PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);

            Quelle.Range[Quelle.Cells[2, 11], Quelle.Cells[lastRow, 11]].Copy();
            ZE.Cells[2, 13].PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
        }

        public void EditQuellSheet()
        {
            List<Jahr> arJahr;
            Jahr tmpJahr;
            DateTime date;
            int counter;

            Quelle.Cells[1, 3].Value = "Rechnungseingang";
            Quelle.Cells[1, ++lastColumn].Value = "kumulierte Rechnungen";
            Quelle.Cells[1, ++lastColumn].Value = "kumulierte Zahlungen";
            Quelle.Cells[1, ++lastColumn].Value = "ø Rechnungseingang";

            if (Quelle.Cells[2, 4].Value < 0)
                Quelle.Cells[2, 9].FormulaR1C1 = "=RC[-5] * -1";
            else
                Quelle.Cells[2, 9].FormulaR1C1 = "=RC[-5]";

            Quelle.Cells[2, 10].FormulaR1C1 = "=RC[-8]";

            arJahr = new List<Jahr>();
            tmpJahr = new Jahr();

            date = Convert.ToDateTime(Quelle.Cells[2, 1].Value);
            tmpJahr.SetJahr(date.Year);
            tmpJahr.Init(date.Month);

            tmpJahr.SetSummeHaben(Convert.ToDecimal(Quelle.Cells[2, 3].Value));

            arJahr.Add(tmpJahr);

            counter = 1;
            while (lastRow >= counter + 2)
            {
                Quelle.Cells[2 + counter, 9].FormulaR1C1 = "=(R[-1]C - RC[-5])";
                Quelle.Cells[2 + counter, 10].FormulaR1C1 = "=(R[-1]C + RC[-8])";

                date = Convert.ToDateTime(Quelle.Cells[2 + counter, 1].Value);
                if (IsInList(date.Year, arJahr) == false)
                {
                    tmpJahr = new Jahr();
                    tmpJahr.SetJahr(date.Year);
                    tmpJahr.Init(date.Month);
                    arJahr.Add(tmpJahr);
                }
                else
                {
                    arJahr.ElementAt(arJahr.Count - 1).SetSummeHaben(Convert.ToDecimal(Quelle.Cells[2 + counter, 3].Value));
                    arJahr.ElementAt(arJahr.Count - 1).AddMonat(date.Month);
                }
                counter++;
            }

            counter = 0;
            while (lastRow >= counter + 2)
            {
                for (int index = 0; index < arJahr.Count; index++)
                {
                    date = Convert.ToDateTime(Quelle.Cells[2 + counter, 1].Value);
                    if (arJahr.ElementAt(index).GetJahr() == date.Year)
                    {
                        Quelle.Cells[2 + counter, 11].Value = arJahr.ElementAt(index).GetUmsatz();
                    }
                }
                counter++;
            }

            Quelle.Columns["B:K"].NumberFormat = "#.##0,00 €";
            Quelle.Columns["C"].ColumnWidth = 10;
            Quelle.Columns["G"].ColumnWidth = 18;
            Quelle.Columns["K"].ColumnWidth = 10;
        }

        public void CopyQuelleTable(Excel.Worksheet dest)
        {
            Quelle.UsedRange.Copy();
            dest.Cells[1, 1].PasteSpecial(Excel.XlPasteType.xlPasteValuesAndNumberFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
        }

        public void AddWorksheet(string name)
        {
            Excel.Worksheet newWs = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(After: Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count]);
            newWs.Name = name;

            if (name == Quelle.Name + "_neu")
                Quelle_Kopie = newWs;
           
            if (name == "ZE_Tabelle")
                ZE = newWs;
        }

        private bool IsInList(int value, List<Jahr> liste)
        {
            bool inList = false;

            foreach (Jahr element in liste)
            {
                if (element.GetJahr() == value)
                {
                    inList = true;
                    break;
                }
            }

            return inList;
        }

        private void SetSourceValid()
        {
            lastColumn = Quelle.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
            lastRow = Quelle.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            var cell1 = Quelle.Cells[1, 1];
            var cell2 = Quelle.Cells[1, lastColumn];

            Excel.Range range = Quelle.Range[cell1, cell2];

            for (int i = 1; i <= range.Cells.Count; i++)
            {
                var value = range.Cells[1, i].Value;

                if (value == colHeader1)
                    isColHeader1 = true;
                if (value == colHeader2)
                    isColHeader2 = true;
                if (value == colHeader3)
                    isColHeader3 = true;
                if (value == colHeader4)
                    isColHeader4 = true;
                if (value == colHeader5)
                    isColHeader5 = true;
                if (value == colHeader6)
                    isColHeader6 = true;
                if (value == colHeader7)
                    isColHeader7 = true;
                if (value == colHeader8)
                    isColHeader8 = true;
            }

            if (isColHeader1 && isColHeader2 && isColHeader3 && isColHeader4 && 
                isColHeader5 && isColHeader6 && isColHeader7 && isColHeader8)
            {
                IsSourceValid = true;
            }
        }
    }
}
