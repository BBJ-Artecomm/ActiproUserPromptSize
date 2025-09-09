namespace ActiproTestVSTO
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            this.ActiproTests = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.openDialogBtn = this.Factory.CreateRibbonButton();
            this.ActiproTests.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // ActiproTests
            // 
            this.ActiproTests.Groups.Add(this.group1);
            this.ActiproTests.Label = "ActiproTests";
            this.ActiproTests.Name = "ActiproTests";
            // 
            // group1
            // 
            this.group1.Items.Add(this.openDialogBtn);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // openDialogBtn
            // 
            this.openDialogBtn.Label = "Open dialog";
            this.openDialogBtn.Name = "openDialogBtn";
            this.openDialogBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.openDialogBtn_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.ActiproTests);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.ActiproTests.ResumeLayout(false);
            this.ActiproTests.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab ActiproTests;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton openDialogBtn;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
