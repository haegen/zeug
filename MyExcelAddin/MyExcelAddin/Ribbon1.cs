using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data.OleDb;

namespace MyExcelAddin
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnZE_Tabelle_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook workbook = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);

            List<String> headers = new List<string>();
            headers.Add("Datum");
            headers.Add("Soll");
            headers.Add("Haben");
            headers.Add("Haben(fällig)");
            headers.Add("Buchsaldo");
            headers.Add("Fälligkeitssaldo");
            headers.Add("überfällige Verbindlichkeiten");
            headers.Add("Überzahlung");

            string name = "test123";
            Table tbl = new Table(workbook.ActiveSheet, name);

            System.Windows.Forms.MessageBox.Show(tbl.IsSourceValid().ToString());
        }
    }
}
