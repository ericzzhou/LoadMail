namespace LoadMail
{
    partial class Rbb_PPManager : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Rbb_PPManager()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Rbb_PPManager));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnLoadMail = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "PPManager";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnLoadMail);
            this.group1.Label = "操作";
            this.group1.Name = "group1";
            // 
            // btnLoadMail
            // 
            this.btnLoadMail.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLoadMail.Image = ((System.Drawing.Image)(resources.GetObject("btnLoadMail.Image")));
            this.btnLoadMail.Label = "导入PPManager";
            this.btnLoadMail.Name = "btnLoadMail";
            this.btnLoadMail.ShowImage = true;
            this.btnLoadMail.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadMail_Click);
            // 
            // Rbb_PPManager
            // 
            this.Name = "Rbb_PPManager";
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Rbb_PPManager_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadMail;
    }

    partial class ThisRibbonCollection
    {
        internal Rbb_PPManager Ribbon1
        {
            get { return this.GetRibbon<Rbb_PPManager>(); }
        }
    }
}
