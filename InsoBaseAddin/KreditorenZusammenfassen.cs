using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Excel;
using System.Drawing;

namespace InsoBaseAddin
{
    class KreditorenZusammenfassen
    {
        Workbook workbook;

        public Worksheet ErrorSheet { get; private set; }
        public int ErrorRowIndex { get; private set; }

        public bool AreTablesValid { get; private set; }
        public bool IsSummaryPage { get; private set; }

        public string sheetName { get; private set; }

        public KreditorenZusammenfassen(Workbook pTargetbook)
        {
            sheetName = "Schriftverkehrliste";
            workbook = pTargetbook;

            AreTablesValid = true;
            IsSummaryPage = IsSummary();

            SetTablesValid();
        }

        private bool IsSummary()
        {
            bool summary = false;

            if (workbook.Worksheets[workbook.Worksheets.Count].Name == sheetName)
                summary = true;

            return summary;
        }

        public void Format()
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            Excel.Worksheet ws = workbook.Worksheets[workbook.Worksheets.Count];
            MyFormat.FormatHeader(ws, true);
            MyFormat.FormatTableData(ws);
            MyFormat.SetBorder(ws);
            MyFormat.PageSetup(ws, sheetName);

            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        public void CopyData()
        {
            for (int sheetIndex = 2; sheetIndex <= workbook.Worksheets.Count - 1; sheetIndex++)
            {
                int rowCount = workbook.Worksheets[sheetIndex].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                int summaryRowCount = workbook.Worksheets[workbook.Worksheets.Count].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;

                for (int rowIndex = 2; rowIndex <= rowCount; rowIndex++)
                {
                    if (IsRowEmpty(rowIndex, sheetIndex))
                    {
                        continue;
                    }

                    Excel.Range r1 = (Excel.Range)workbook.Worksheets[sheetIndex].Rows[rowIndex].EntireRow;
                    Excel.Range r2 = (Excel.Range)workbook.Worksheets[workbook.Worksheets.Count].Cells(summaryRowCount, 1);
                    r1.Copy(r2);
                    workbook.Worksheets[workbook.Worksheets.Count].Cells(summaryRowCount, 7).value = workbook.Worksheets[sheetIndex].name;

                    summaryRowCount++;
                }
            }
        }

        public void AddSummaryPage()
        {
            Excel.Worksheet newWs = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            newWs.Name = sheetName;

            AddSummaryHeader();
        }

        private void AddSummaryHeader()
        {
            Excel.Range r1 = (Excel.Range)workbook.Worksheets[workbook.Worksheets.Count - 1].Range("A1", "F1");
            Excel.Range r2 = (Excel.Range)workbook.Worksheets[workbook.Worksheets.Count].Cells(1, 1);
            r1.Copy(r2);

            workbook.Worksheets[workbook.Worksheets.Count].Cells(1, 7).Value = "Kreditoren";
        }

        private void SetTablesValid()
        {
            for (int sheetIndex = 2; sheetIndex <= workbook.Sheets.Count; sheetIndex++)
            {
                var rowCount = workbook.Worksheets[sheetIndex].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                var worksheet = workbook.Worksheets[sheetIndex];
                for (int rowIndex = 2; rowIndex <= rowCount; rowIndex++)
                {
                    if (IsRowEmpty(rowIndex, sheetIndex))
                    {
                        continue;
                    }
                    else
                    {
                        if (!IsDate(Convert.ToString(worksheet.Cells(rowIndex, 1).Value)) || worksheet.Cells(rowIndex, 3).Value == String.Empty)
                        {
                            ErrorSheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[sheetIndex]);
                            ErrorSheet.Activate();
                            ErrorSheet.Rows[rowIndex].Select();
                            ErrorRowIndex = rowIndex;
                            AreTablesValid = false;
                            break;
                        }
                    }
                }
                if (!AreTablesValid)
                {
                    break;
                }
            }
        }

        private bool IsRowEmpty(int rowIndex, int sheetIndex)
        {
            bool isEmpty = true;

            for (int column = 1; column <= 6; column++)
            {
                if (workbook.Worksheets[sheetIndex].Cells[rowIndex, column].Value != null)
                {
                    isEmpty = false;
                    break;
                }
            }

            return isEmpty;
        }

        private bool IsDate(string value)
        {
            try
            {
                DateTime dt = DateTime.Parse(value);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
