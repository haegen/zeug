namespace MyExcelAddin
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
        /// <param name="disposing">"true", wenn verwaltete Ressourcen gelöscht werden sollen, andernfalls "false".</param>
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
            this.grErmittlung = this.Factory.CreateRibbonGroup();
            this.btnZE_Tabelle = this.Factory.CreateRibbonButton();
            this.TabKDLB_KI.SuspendLayout();
            this.grErmittlung.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabKDLB_KI
            // 
            this.TabKDLB_KI.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabKDLB_KI.Groups.Add(this.grErmittlung);
            this.TabKDLB_KI.Label = "Test";
            this.TabKDLB_KI.Name = "TabKDLB_KI";
            // 
            // grErmittlung
            // 
            this.grErmittlung.Items.Add(this.btnZE_Tabelle);
            this.grErmittlung.Label = "Ermittlung";
            this.grErmittlung.Name = "grErmittlung";
            // 
            // btnZE_Tabelle
            // 
            this.btnZE_Tabelle.Label = "Test ZE Tabelle anlgen";
            this.btnZE_Tabelle.Name = "btnZE_Tabelle";
            this.btnZE_Tabelle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnZE_Tabelle_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.TabKDLB_KI);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.TabKDLB_KI.ResumeLayout(false);
            this.TabKDLB_KI.PerformLayout();
            this.grErmittlung.ResumeLayout(false);
            this.grErmittlung.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab TabKDLB_KI;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grErmittlung;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnZE_Tabelle;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
