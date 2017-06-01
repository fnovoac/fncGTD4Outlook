namespace fncGTD4Outlook
{
    partial class RibbonExplorer : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonExplorer()
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
            this.btnArchivar = this.Factory.CreateRibbonButton();
            this.btnCompletado = this.Factory.CreateRibbonButton();
            this.btnDelegar = this.Factory.CreateRibbonButton();
            this.btnDiferir = this.Factory.CreateRibbonButton();
            this.btnEliminar = this.Factory.CreateRibbonButton();
            this.btnConservar = this.Factory.CreateRibbonButton();
            this.btnReferencia = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabMail";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabMail";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnArchivar);
            this.group1.Items.Add(this.btnCompletado);
            this.group1.Items.Add(this.btnDelegar);
            this.group1.Items.Add(this.btnDiferir);
            this.group1.Items.Add(this.btnEliminar);
            this.group1.Items.Add(this.btnConservar);
            this.group1.Items.Add(this.btnReferencia);
            this.group1.Label = "fncGTD4Outlook";
            this.group1.Name = "group1";
            // 
            // btnArchivar
            // 
            this.btnArchivar.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnArchivar.Label = "Archivar";
            this.btnArchivar.Name = "btnArchivar";
            this.btnArchivar.OfficeImageId = "SaveSentItemsMenu";
            this.btnArchivar.ShowImage = true;
            this.btnArchivar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnArchivar_Click);
            // 
            // btnCompletado
            // 
            this.btnCompletado.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCompletado.Label = "Completado";
            this.btnCompletado.Name = "btnCompletado";
            this.btnCompletado.OfficeImageId = "WorkflowComplete";
            this.btnCompletado.ShowImage = true;
            this.btnCompletado.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCompletado_Click);
            // 
            // btnDelegar
            // 
            this.btnDelegar.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDelegar.Label = "Delegar";
            this.btnDelegar.Name = "btnDelegar";
            this.btnDelegar.OfficeImageId = "DirectRepliesTo";
            this.btnDelegar.ShowImage = true;
            this.btnDelegar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDelegar_Click);
            // 
            // btnDiferir
            // 
            this.btnDiferir.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDiferir.Label = "Diferir";
            this.btnDiferir.Name = "btnDiferir";
            this.btnDiferir.OfficeImageId = "DelayDeliveryOutlook";
            this.btnDiferir.ShowImage = true;
            this.btnDiferir.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDiferir_Click);
            // 
            // btnEliminar
            // 
            this.btnEliminar.Label = "Eliminar";
            this.btnEliminar.Name = "btnEliminar";
            this.btnEliminar.OfficeImageId = "AdpDiagramDeleteTable";
            this.btnEliminar.ShowImage = true;
            this.btnEliminar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEliminar_Click);
            // 
            // btnConservar
            // 
            this.btnConservar.Label = "Revisar luego";
            this.btnConservar.Name = "btnConservar";
            this.btnConservar.OfficeImageId = "FunctionsLogicalInsertGallery";
            this.btnConservar.ShowImage = true;
            this.btnConservar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConservar_Click);
            // 
            // btnReferencia
            // 
            this.btnReferencia.Label = "Referencia";
            this.btnReferencia.Name = "btnReferencia";
            this.btnReferencia.OfficeImageId = "Pushpin";
            this.btnReferencia.ShowImage = true;
            this.btnReferencia.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReferencia_Click);
            // 
            // RibbonExplorer
            // 
            this.Name = "RibbonExplorer";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonExplorer_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnArchivar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCompletado;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDelegar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDiferir;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEliminar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConservar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReferencia;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonExplorer RibbonExplorer
        {
            get { return this.GetRibbon<RibbonExplorer>(); }
        }
    }
}
