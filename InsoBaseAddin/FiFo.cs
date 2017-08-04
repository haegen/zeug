using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace InsoBaseAddin
{
    class FiFo
    {
        Worksheet sheet;

        public bool IsValid { get; private set; }
        int colBuDat = 0;
        int colFaelDat = 0;
        int colSoll = 0;
        int colHaben = 0;
        int colBelegNr = 0;
        int lastCol;
        int lastRow;
        
        public FiFo(Worksheet pTargetSheet)
        {
            IsValid = false;
            sheet = pTargetSheet;
            SetValid();
        }

        private void SetValid()
        {
            int col = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
            var cell1  = sheet.Cells[1,1];
            var cell2  = sheet.Cells[1,col];

            Excel.Range range = sheet.Range[cell1, cell2];
            for (int i = 1; i <= range.Cells.Count; i++)
			{
                var value = range.Cells[1, i].Value;
                if (value == "korr. Belegdatum")
                    colBuDat = i;
                if (value == "Haben")
                    colHaben = i;
                if (value == "Soll")
                    colSoll = i;
                if (value == "Fälligkeitsdatum")
                    colFaelDat = i;
                if (value == "BelegNr")
                    colBelegNr = i;
			}
            if (colBuDat > 0 && colFaelDat > 0 && colHaben > 0 && colSoll > 0)
            {
                IsValid = true;
                lastCol = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column + 1;
                lastRow = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            }
        }

        public void SetColumnHeader()
        {
            //int lastCol = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column + 1;
            sheet.Cells[1, lastCol].Value2 = "Tage zum Ausgleich";
            sheet.Columns[lastCol].ColumnWidth = 17;
        }

        public void CalcAusgleichsdatum2()
        {
            //int lastRow = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            //int lastCol = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

            List<DatumValuePair> SollPairs = new List<DatumValuePair>();
            List<DatumValuePair> HabenPairs = new List<DatumValuePair>();
            

            for (int i = 2; i <= lastRow; i++)
            {
                decimal? sollValue = null;
                decimal? habenValue = null;

                sollValue = (decimal?)sheet.Cells[i, colSoll].Value2;
                habenValue = (decimal?)sheet.Cells[i, colHaben].Value2;

                if (sollValue != null && sollValue > 0)
                {
                    double dateSoll = ((Excel.Range)sheet.Cells[i, colBuDat]).Value2;
                    DateTime SollBelegDat;
                    SollBelegDat = DateTime.FromOADate(dateSoll);

                    DatumValuePair position = new DatumValuePair(sollValue.Value, SollBelegDat, i);
                    SollPairs.Add(position);
                }

                if (habenValue != null && habenValue != 0)
                {
                    double dateHaben = ((Excel.Range)sheet.Cells[i, colFaelDat]).Value2;
                    DateTime HabenFaelDat;
                    HabenFaelDat = DateTime.FromOADate(dateHaben);

                    DatumValuePair position = new DatumValuePair(sollValue.Value, HabenFaelDat, i);
                    HabenPairs.Add(position);
                }
            }

            foreach (DatumValuePair soll in SollPairs)
            {
                while (soll.Wert > 0)
                {

                }
            }

        }

        public void CalcAusgleichsdatumNachBelegNr()
        {
            DateTime SollBelegDat;
            DateTime HabenFaelDat;
            decimal? sollValue = null;
            decimal? habenValue = null;

            for (int i = 2; i <= lastRow; i++)
            {
                habenValue = Convert.ToDecimal(sheet.Cells[i, colHaben].Value);
                if (habenValue != null && habenValue > 0)
                {
                    double dateHaben = ((Excel.Range)sheet.Cells[i, colFaelDat]).Value2;
                    HabenFaelDat = DateTime.FromOADate(dateHaben);
                    var BelegNrHaben = ((Excel.Range)sheet.Cells[i, colBelegNr]).Value2;
                    TimeSpan diff;

                    for (int j = 2; j < lastRow; j++)
                    {
                        var BelegNrSoll = ((Excel.Range)sheet.Cells[j, colBelegNr]).Value2;
                        sollValue = Convert.ToDecimal(sheet.Cells[j, colSoll].Value);
                        if (BelegNrHaben == BelegNrSoll && sollValue != null && sollValue == habenValue)
                        {
                            double dateSoll = ((Excel.Range)sheet.Cells[j, colFaelDat]).Value2;
                            SollBelegDat = DateTime.FromOADate(dateSoll);
                            diff = SollBelegDat.Subtract(HabenFaelDat);
                            sheet.Cells[i, lastCol].Value2 = diff.TotalDays.ToString();
                        }
                        else
                            continue;
                    }
                }
            }
        }

        public void CalcAusgleichsdatumNachFaelligkeitFiLo()
        {
            int countHaben = lastRow;
            int countHabenOld = lastRow;
            DateTime SollBelegDat;
            DateTime HabenFaelDat;
            decimal? sollValue = null;
            decimal? habenValue = null;

            for (int i = lastRow; i >= 2; i--)
            {
                sollValue = Convert.ToDecimal(sheet.Cells[i, colSoll].Value);
                if (sollValue != null && sollValue > 0)
                {
                    double dateSoll = ((Excel.Range)sheet.Cells[i, colBuDat]).Value2;
                    SollBelegDat = DateTime.FromOADate(dateSoll);
                    TimeSpan diff;
                    while (sollValue > 0)
                    {
                        if (countHaben < countHabenOld)
                        {
                            var test = sheet.Cells[countHaben, colHaben].Value2;
                            habenValue = (decimal?)sheet.Cells[countHaben, colHaben].Value2;
                            countHabenOld = countHaben;
                        }
                        if (habenValue != null)
                        {
                            if (habenValue != 0)
                            {
                                if (habenValue < 0)
                                {
                                    sollValue -= habenValue;
                                    countHaben++;
                                    continue;
                                }
                                double dateHaben = ((Excel.Range)sheet.Cells[countHaben, colFaelDat]).Value2;
                                HabenFaelDat = DateTime.FromOADate(dateHaben);
                                decimal? tempDiff = sollValue - habenValue;
                                if (tempDiff >= 0)
                                {
                                    diff = HabenFaelDat.Subtract(SollBelegDat);
                                    sheet.Cells[countHaben, lastCol].Value2 = diff.TotalDays.ToString();
                                    sollValue -= habenValue;
                                }
                                else
                                {
                                    habenValue -= sollValue;
                                    break;
                                }
                            }
                        }
                        countHaben--;
                    }

                }

            }
        }

        public void CalcAusgleichsdatumNachFaelligkeitFiFo()
        {
            //int lastRow = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            //int newCol = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column + 1;
            int countHaben = 2;
            int countHabenOld = 0;
            DateTime SollBelegDat;
            DateTime HabenFaelDat;
            decimal? sollValue = null;
            decimal? habenValue = null;

            for (int i = 2; i <= lastRow; i++)
            {
                sollValue = Convert.ToDecimal(sheet.Cells[i, colSoll].Value);
                if (sollValue != null && sollValue > 0)
                {
                    double dateSoll = ((Excel.Range)sheet.Cells[i, colBuDat]).Value2;
                    SollBelegDat = DateTime.FromOADate(dateSoll);
                    TimeSpan diff;
                    while(sollValue > 0)
                    {
                        if (countHaben > countHabenOld)
                        {
                            var test = sheet.Cells[countHaben, colHaben].Value2;
                            habenValue = (decimal?)sheet.Cells[countHaben, colHaben].Value2;
                            countHabenOld = countHaben;
                        }
                        if (habenValue != null)
                        {
                            if (habenValue != 0)
                            {
                                if (habenValue < 0)
                                {
                                    sollValue -= habenValue;
                                    countHaben++;
                                    continue;
                                }
                                double dateHaben = ((Excel.Range)sheet.Cells[countHaben, colFaelDat]).Value2;
                                HabenFaelDat = DateTime.FromOADate(dateHaben);
                                decimal? tempDiff = sollValue - habenValue;
                                if (tempDiff >= 0)
                                {
                                    diff = SollBelegDat.Subtract(HabenFaelDat);
                                    sheet.Cells[countHaben, lastCol].Value2 = diff.TotalDays.ToString();
                                    sollValue -= habenValue;
                                }
                                else
                                {
                                    habenValue -= sollValue;
                                    break;
                                }
                            }
                            //else if(habenValue < 0)
                            //{
                            //    sollValue += (habenValue * -1);
                            //}
                        }
                        countHaben++;
                    }

                }

            }

            //SollWert ermitteln
            //Belegdatum festschreiben
            //Habenwert abziehen
            //soll wert größer 0
            //belegdatum vom faelligkeitsdatum abziehen
            //nächsten haben wert ermitteln
            //haben wert abziehen
            //soll wert gleich 0 und haben wert größer 0
            //nächsten soll wert ermitteln
            //neues belegdatum festschreiben
            //haben wert abziehen
            //soll wert gleich 0 und haben wert gleich 0
            //nächsten soll wert ermitteln
            //belegdatum festschreiben
            //nächsten haben wert ermitteln
            //soll wert gleich 0 oder kleiner 0 und haben wert größer 0
            //nächsten soll wert ermitteln
            //belegdatum schreiben
            //haben wert abziehen
        }
    }
}
