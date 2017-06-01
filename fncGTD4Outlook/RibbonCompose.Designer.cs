namespace fncGTD4Outlook
{
    partial class RibbonCompose : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonCompose()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnDelegarEnviar = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabNewMailMessage";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabNewMailMessage";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnDelegarEnviar);
            this.group1.Label = "fncGTD4Outlook";
            this.group1.Name = "group1";
            // 
            // btnDelegarEnviar
            // 
            this.btnDelegarEnviar.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDelegarEnviar.Label = "Delegar y Enviar";
            this.btnDelegarEnviar.Name = "btnDelegarEnviar";
            this.btnDelegarEnviar.OfficeImageId = "DirectRepliesTo";
            this.btnDelegarEnviar.ShowImage = true;
            this.btnDelegarEnviar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDelegarEnviar_Click);
            // 
            // RibbonCompose
            // 
            this.Name = "RibbonCompose";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose, Microsoft.Outlook.Post.Compose";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonCompose_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDelegarEnviar;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonCompose RibbonCompose
        {
            get { return this.GetRibbon<RibbonCompose>(); }
        }
    }
}
