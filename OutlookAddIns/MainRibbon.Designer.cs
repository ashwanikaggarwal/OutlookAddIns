namespace OutlookAddIns
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabMain = this.Factory.CreateRibbonTab();
            this.grpTickets = this.Factory.CreateRibbonGroup();
            this.btnSetPath = this.Factory.CreateRibbonButton();
            this.tabMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMain
            // 
            this.tabMain.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabMain.Groups.Add(this.grpTickets);
            this.tabMain.Label = "LandscapeShow";
            this.tabMain.Name = "tabMain";
            // 
            // grpTickets
            // 
            this.grpTickets.Label = "Ticket Requests";
            this.grpTickets.Name = "grpTickets";
            // 
            // btnSetPath
            // 
            this.btnSetPath.Label = "Database Path";
            this.btnSetPath.Name = "btnSetPath";
            this.btnSetPath.OfficeImageId = "MicrosoftAccess";
            this.btnSetPath.ShowImage = true;
            this.btnSetPath.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetPath_Click);
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            // 
            // MainRibbon.OfficeMenu
            // 
            this.OfficeMenu.Items.Add(this.btnSetPath);
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tabMain);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tabMain.ResumeLayout(false);
            this.tabMain.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabMain;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpTickets;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetPath;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon Ribbon1
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
