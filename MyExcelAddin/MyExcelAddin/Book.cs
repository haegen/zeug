using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;

namespace MyExcelAddin
{
    class Book
    {
        private Workbook workbook { get; set; }
        private List<Table> tables { get; set; }

        public Book(Workbook workbook)
        {
            this.workbook = workbook;

        }
    }
}
