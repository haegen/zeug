using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System.Drawing;

namespace MyExcelAddin
{
    class Table
    {
        public List<String> headers { private get; set; }

        private Worksheet worksheet { get; set; }
        private List<Boolean> isHeaders { get; set; }

        private int lastRow { get; set; }
        private int lastColumn { get; set; }

        private Color mark1 { get; set; }
        private Color mark2 { get; set; }

        public Table(object worksheet, string name)
        {
            this.worksheet = (Worksheet)worksheet;
            this.worksheet.Name = name;

            lastRow = this.worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            lastColumn = this.worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

            mark1 = Color.FromArgb(255, 192, 0);     // organge
            mark2 = Color.FromArgb(0, 176, 80);      // grün
        }

        public bool IsSourceValid()
        {
            bool isValid = true;

            var cell1 = this.worksheet.Cells[1, 1];
            var cell2 = this.worksheet.Cells[1, lastColumn];

            Excel.Range range = this.worksheet.Range[cell1, cell2];

            isHeaders = new List<bool>();

            for (int columnIndex = 1; columnIndex <= lastColumn; columnIndex++)
            {
                var value = range.Cells[1, columnIndex].Value;

                foreach(String element in headers)
                {
                    if (value == element)
                        isHeaders.Add(true);
                    else
                        isHeaders.Add(false);
                }
            }

            foreach (Boolean element in isHeaders)
            {
                if (!element)
                    isValid = false;
            }

            return isValid;
        }
    }
}
