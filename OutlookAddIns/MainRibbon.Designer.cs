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
            this.grpManualTickets = this.Factory.CreateRibbonGroup();
            this.grpSettings = this.Factory.CreateRibbonGroup();
            this.btnSetPath = this.Factory.CreateRibbonButton();
            this.btnTicketRequest = this.Factory.CreateRibbonButton();
            this.btnRegister = this.Factory.CreateRibbonButton();
            this.btnUpdate = this.Factory.CreateRibbonButton();
            this.btnSend = this.Factory.CreateRibbonButton();
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.tabMain.SuspendLayout();
            this.grpTickets.SuspendLayout();
            this.grpManualTickets.SuspendLayout();
            this.grpSettings.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMain
            // 
            this.tabMain.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabMain.Groups.Add(this.grpTickets);
            this.tabMain.Groups.Add(this.grpManualTickets);
            this.tabMain.Groups.Add(this.grpSettings);
            this.tabMain.Label = "LandscapeShow";
            this.tabMain.Name = "tabMain";
            // 
            // grpTickets
            // 
            this.grpTickets.Items.Add(this.btnTicketRequest);
            this.grpTickets.Label = "OneClick Tickets";
            this.grpTickets.Name = "grpTickets";
            // 
            // grpManualTickets
            // 
            this.grpManualTickets.Items.Add(this.btnRegister);
            this.grpManualTickets.Items.Add(this.btnUpdate);
            this.grpManualTickets.Items.Add(this.btnSend);
            this.grpManualTickets.Label = "Ticket Requests";
            this.grpManualTickets.Name = "grpManualTickets";
            // 
            // grpSettings
            // 
            this.grpSettings.Items.Add(this.menu1);
            this.grpSettings.Label = "Settings";
            this.grpSettings.Name = "grpSettings";
            // 
            // btnSetPath
            // 
            this.btnSetPath.Label = "Database Path";
            this.btnSetPath.Name = "btnSetPath";
            this.btnSetPath.OfficeImageId = "MicrosoftAccess";
            this.btnSetPath.ShowImage = true;
            this.btnSetPath.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetPath_Click);
            // 
            // btnTicketRequest
            // 
            this.btnTicketRequest.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnTicketRequest.Label = "Ticket Request";
            this.btnTicketRequest.Name = "btnTicketRequest";
            this.btnTicketRequest.OfficeImageId = "MailMergeRecipientsEditList";
            this.btnTicketRequest.ShowImage = true;
            // 
            // btnRegister
            // 
            this.btnRegister.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRegister.Label = "Register Request";
            this.btnRegister.Name = "btnRegister";
            this.btnRegister.OfficeImageId = "AddUserToPermissionGroup";
            this.btnRegister.ShowImage = true;
            this.btnRegister.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRegister_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdate.Label = "Update Details in Database";
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.OfficeImageId = "DatabaseModelingReverse";
            this.btnUpdate.ShowImage = true;
            // 
            // btnSend
            // 
            this.btnSend.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSend.Label = "Send Confirmation";
            this.btnSend.Name = "btnSend";
            this.btnSend.OfficeImageId = "ConversationsMenu";
            this.btnSend.ShowImage = true;
            // 
            // menu1
            // 
            this.menu1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu1.Label = "Modify";
            this.menu1.Name = "menu1";
            this.menu1.OfficeImageId = "AddInManager";
            this.menu1.ShowImage = true;
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
            this.grpTickets.ResumeLayout(false);
            this.grpTickets.PerformLayout();
            this.grpManualTickets.ResumeLayout(false);
            this.grpManualTickets.PerformLayout();
            this.grpSettings.ResumeLayout(false);
            this.grpSettings.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabMain;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpTickets;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetPath;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTicketRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpManualTickets;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRegister;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSend;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon Ribbon1
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
