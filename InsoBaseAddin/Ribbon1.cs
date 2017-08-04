using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace InsoBaseAddin
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnFiFo_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

            FiFo instance = new FiFo(worksheet);
            if (instance.IsValid)
            {
                instance.SetColumnHeader();
                instance.CalcAusgleichsdatumNachFaelligkeitFiFo();
            }
            else
                System.Windows.Forms.MessageBox.Show("Das ausgewählte Tabellenblatt entspricht nicht den Anforderungen der Funktion.", "Ungültige Auswahl", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
        }

        private void btnBelegNr_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

            FiFo instance = new FiFo(worksheet);
            if (instance.IsValid)
            {
                instance.SetColumnHeader();
                instance.CalcAusgleichsdatumNachBelegNr();
            }
            else
                System.Windows.Forms.MessageBox.Show("Das ausgewählte Tabellenblatt entspricht nicht den Anforderungen der Funktion.", "Ungültige Auswahl", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
        }

        private void btnFiLo_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

            FiFo instance = new FiFo(worksheet);
            if (instance.IsValid)
            {
                instance.SetColumnHeader();
                instance.CalcAusgleichsdatumNachFaelligkeitFiLo();
            }
            else
                System.Windows.Forms.MessageBox.Show("Das ausgewählte Tabellenblatt entspricht nicht den Anforderungen der Funktion.", "Ungültige Auswahl", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
        }

        private void btnKreditorenZusammenfassen_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook workbook = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);

            KreditorenZusammenfassen instance = new KreditorenZusammenfassen(workbook);

            //if (instance.IsValidFileName)
            //{
                if (instance.AreTablesValid)
                {
                    if (!instance.IsSummaryPage)
                    {
                        instance.AddSummaryPage();
                        instance.CopyData();
                        instance.Format();
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Das Tabellenblatt " + instance.sheetName + " exitiert bereits.");
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("In dem Tabellenblatt " + instance.ErrorSheet.Name + " Zeile " + instance.ErrorRowIndex + " ist ein Fehler aufgetreten.\nEs muss mindestens ein Datum und ein Ereignis vorhanden sein.", "Eingabe überprüfen.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
            //}
            //else
            //{
            //    System.Windows.Forms.MessageBox.Show("Das ausgewählte Datei entspricht nicht den Anforderungen der Funktion.", "Ungültige Auswahl", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            //}
        }

        private void btnZETabelleErstellen_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            ZE_Tabelle instance = new ZE_Tabelle(workbook);

            if (instance.IsSourceValid)
            {
                instance.AddWorksheet(instance.Quelle.Name + "_neu");
                instance.AddWorksheet("ZE_Tabelle");

                instance.CopyQuelleTable(workbook.Worksheets[instance.Quelle.Name + "_neu"]);
                instance.EditQuellSheet();

                instance.CopyZEData();
                instance.EditZESheet();

                instance.Format();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Das ausgewählte Datei entspricht nicht den Anforderungen der Funktion.", "Ungültige Auswahl", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void btnGrafikZE1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            GrafikZE1 instance = new GrafikZE1(worksheet);

            if (instance.IsSourceValid)
            {
                if (instance.IsMarked != 0)
                {
                    instance.DiagrammErstellen();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Es wurde nichts in der Tabelle markiert.");
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Das ausgewählte Tabellenblatt entspricht nicht den Anforderungen der Funktion.", "Ungültige Auswahl", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void btnGrafikZE2_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            GrafikZE2 instance = new GrafikZE2(worksheet);

            if (instance.IsSourceValid)
            {
                if (instance.IsMarked != 0)
                {
                    instance.DiagrammErstellen();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Es wurde nichts in der Tabelle markiert.");
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Das ausgewählte Tabellenblatt entspricht nicht den Anforderungen der Funktion.", "Ungültige Auswahl", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
    }
}
