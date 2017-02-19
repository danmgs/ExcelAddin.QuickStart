namespace MyExcelAddIn
{
    partial class CustomRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public CustomRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.tabMyExcelAddin = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.checkBoxTaskPanel = this.Factory.CreateRibbonCheckBox();
            this.checkBoxShowTab2 = this.Factory.CreateRibbonCheckBox();
            this.toggleButtonPositionPanel = this.Factory.CreateRibbonToggleButton();
            this.tab1.SuspendLayout();
            this.tabMyExcelAddin.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Name = "group1";
            // 
            // tabMyExcelAddin
            // 
            this.tabMyExcelAddin.Groups.Add(this.group2);
            this.tabMyExcelAddin.Label = "My Excel Add-In";
            this.tabMyExcelAddin.Name = "tabMyExcelAddin";
            // 
            // group2
            // 
            this.group2.Items.Add(this.checkBoxTaskPanel);
            this.group2.Items.Add(this.checkBoxShowTab2);
            this.group2.Items.Add(this.toggleButtonPositionPanel);
            this.group2.Label = "Task Panel Options";
            this.group2.Name = "group2";
            // 
            // checkBoxTaskPanel
            // 
            this.checkBoxTaskPanel.Checked = true;
            this.checkBoxTaskPanel.Label = "Show Task Panel";
            this.checkBoxTaskPanel.Name = "checkBoxTaskPanel";
            this.checkBoxTaskPanel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBoxShowTaskPanel_Click);
            // 
            // checkBoxShowTab2
            // 
            this.checkBoxShowTab2.Label = "Switch Tab";
            this.checkBoxShowTab2.Name = "checkBoxShowTab2";
            this.checkBoxShowTab2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.switchPanelTab_Click);
            // 
            // toggleButtonPositionPanel
            // 
            this.toggleButtonPositionPanel.Label = "Change Panel Position";
            this.toggleButtonPositionPanel.Name = "toggleButtonPositionPanel";
            this.toggleButtonPositionPanel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonPositionPanel_Click);
            // 
            // CustomRibbon
            // 
            this.Name = "CustomRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tabMyExcelAddin);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.CustomRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tabMyExcelAddin.ResumeLayout(false);
            this.tabMyExcelAddin.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabMyExcelAddin;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxTaskPanel;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxShowTab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonPositionPanel;
    }

    partial class ThisRibbonCollection
    {
        internal CustomRibbon CustomRibbon
        {
            get { return this.GetRibbon<CustomRibbon>(); }
        }
    }
}
