namespace InsoBaseAddin
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">"true", wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls "false".</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für Designerunterstützung -
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.TabKDLB_KI = this.Factory.CreateRibbonTab();
            this.grFaelligkeit = this.Factory.CreateRibbonGroup();
            this.btnBelegNr = this.Factory.CreateRibbonButton();
            this.btnFiFo = this.Factory.CreateRibbonButton();
            this.btnFiLo = this.Factory.CreateRibbonButton();
            this.btnKreditorenZusammenfassen = this.Factory.CreateRibbonButton();
            this.btnZETabelleErstellen = this.Factory.CreateRibbonButton();
            this.btnGrafikZE1 = this.Factory.CreateRibbonButton();
            this.btnGrafikZE2 = this.Factory.CreateRibbonButton();
            this.TabKDLB_KI.SuspendLayout();
            this.grFaelligkeit.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabKDLB_KI
            // 
            this.TabKDLB_KI.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabKDLB_KI.Groups.Add(this.grFaelligkeit);
            this.TabKDLB_KI.Label = "KDLB KI";
            this.TabKDLB_KI.Name = "TabKDLB_KI";
            // 
            // grFaelligkeit
            // 
            this.grFaelligkeit.Items.Add(this.btnBelegNr);
            this.grFaelligkeit.Items.Add(this.btnFiFo);
            this.grFaelligkeit.Items.Add(this.btnFiLo);
            this.grFaelligkeit.Items.Add(this.btnKreditorenZusammenfassen);
            this.grFaelligkeit.Items.Add(this.btnZETabelleErstellen);
            this.grFaelligkeit.Items.Add(this.btnGrafikZE1);
            this.grFaelligkeit.Items.Add(this.btnGrafikZE2);
            this.grFaelligkeit.Label = "Ermittlung";
            this.grFaelligkeit.Name = "grFaelligkeit";
            // 
            // btnBelegNr
            // 
            this.btnBelegNr.Label = "Ausgleichstage nach BelegNr";
            this.btnBelegNr.Name = "btnBelegNr";
            this.btnBelegNr.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBelegNr_Click);
            // 
            // btnFiFo
            // 
            this.btnFiFo.Label = "Ausgleichstage nach FiFo berechnen";
            this.btnFiFo.Name = "btnFiFo";
            this.btnFiFo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFiFo_Click);
            // 
            // btnFiLo
            // 
            this.btnFiLo.Label = "Ausgleichstage nach FiLo berechnen";
            this.btnFiLo.Name = "btnFiLo";
            this.btnFiLo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFiLo_Click);
            // 
            // btnKreditorenZusammenfassen
            // 
            this.btnKreditorenZusammenfassen.Label = "Kreditoren zusammenfassen";
            this.btnKreditorenZusammenfassen.Name = "btnKreditorenZusammenfassen";
            this.btnKreditorenZusammenfassen.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnKreditorenZusammenfassen_Click);
            // 
            // btnZETabelleErstellen
            // 
            this.btnZETabelleErstellen.Label = "ZE Tabelle erstellen";
            this.btnZETabelleErstellen.Name = "btnZETabelleErstellen";
            this.btnZETabelleErstellen.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnZETabelleErstellen_Click);
            // 
            // btnGrafikZE1
            // 
            this.btnGrafikZE1.Label = "Grafik ZE 1";
            this.btnGrafikZE1.Name = "btnGrafikZE1";
            this.btnGrafikZE1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGrafikZE1_Click);
            // 
            // btnGrafikZE2
            // 
            this.btnGrafikZE2.Label = "Grafik ZE 2";
            this.btnGrafikZE2.Name = "btnGrafikZE2";
            this.btnGrafikZE2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGrafikZE2_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.TabKDLB_KI);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.TabKDLB_KI.ResumeLayout(false);
            this.TabKDLB_KI.PerformLayout();
            this.grFaelligkeit.ResumeLayout(false);
            this.grFaelligkeit.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab TabKDLB_KI;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grFaelligkeit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFiFo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBelegNr;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFiLo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnKreditorenZusammenfassen;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnZETabelleErstellen;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGrafikZE1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGrafikZE2;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
