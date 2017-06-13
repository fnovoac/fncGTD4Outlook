namespace fncGTD4Outlook
{
    [System.ComponentModel.ToolboxItemAttribute(false)]
    partial class regionConversationItems : Microsoft.Office.Tools.Outlook.FormRegionBase
    {
        public regionConversationItems(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            : base(Globals.Factory, formRegion)
        {
            this.InitializeComponent();
        }

        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
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
            this.listViewConversation = new System.Windows.Forms.ListView();
            this.SuspendLayout();
            // 
            // listViewConversation
            // 
            this.listViewConversation.BackColor = System.Drawing.Color.DimGray;
            this.listViewConversation.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listViewConversation.ForeColor = System.Drawing.Color.WhiteSmoke;
            this.listViewConversation.Location = new System.Drawing.Point(0, 0);
            this.listViewConversation.Name = "listViewConversation";
            this.listViewConversation.Size = new System.Drawing.Size(875, 134);
            this.listViewConversation.TabIndex = 0;
            this.listViewConversation.UseCompatibleStateImageBehavior = false;
            this.listViewConversation.DoubleClick += new System.EventHandler(this.listViewConversation_DoubleClick);
            this.listViewConversation.Resize += new System.EventHandler(this.listViewConversation_Resize);
            // 
            // regionConversationItems
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(106)))), ((int)(((byte)(106)))), ((int)(((byte)(106)))));
            this.Controls.Add(this.listViewConversation);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "regionConversationItems";
            this.Size = new System.Drawing.Size(875, 134);
            this.FormRegionShowing += new System.EventHandler(this.regionConversationItems_FormRegionShowing);
            this.FormRegionClosed += new System.EventHandler(this.regionConversationItems_FormRegionClosed);
            this.Load += new System.EventHandler(this.regionConversationItems_LoadAsync);
            this.ResumeLayout(false);

        }

        #endregion

        #region Form Region Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private static void InitializeManifest(Microsoft.Office.Tools.Outlook.FormRegionManifest manifest, Microsoft.Office.Tools.Outlook.Factory factory)
        {
            manifest.FormRegionName = "Conversation Items";
            manifest.FormRegionType = Microsoft.Office.Tools.Outlook.FormRegionType.Adjoining;
            manifest.ShowInspectorCompose = false;

        }

        #endregion

        private System.Windows.Forms.ListView listViewConversation;

        public partial class regionConversationItemsFactory : Microsoft.Office.Tools.Outlook.IFormRegionFactory
        {
            public event Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler FormRegionInitializing;

            private Microsoft.Office.Tools.Outlook.FormRegionManifest _Manifest;

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public regionConversationItemsFactory()
            {
                this._Manifest = Globals.Factory.CreateFormRegionManifest();
                regionConversationItems.InitializeManifest(this._Manifest, Globals.Factory);
                this.FormRegionInitializing += new Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler(this.regionConversationItemsFactory_FormRegionInitializing);
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public Microsoft.Office.Tools.Outlook.FormRegionManifest Manifest
            {
                get
                {
                    return this._Manifest;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.IFormRegion Microsoft.Office.Tools.Outlook.IFormRegionFactory.CreateFormRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            {
                regionConversationItems form = new regionConversationItems(formRegion);
                form.Factory = this;
                return form;
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            byte[] Microsoft.Office.Tools.Outlook.IFormRegionFactory.GetFormRegionStorage(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                throw new System.NotSupportedException();
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            bool Microsoft.Office.Tools.Outlook.IFormRegionFactory.IsDisplayedForItem(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                if (this.FormRegionInitializing != null)
                {
                    Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs cancelArgs = Globals.Factory.CreateFormRegionInitializingEventArgs(outlookItem, formRegionMode, formRegionSize, false);
                    this.FormRegionInitializing(this, cancelArgs);
                    return !cancelArgs.Cancel;
                }
                else
                {
                    return true;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.FormRegionKindConstants Microsoft.Office.Tools.Outlook.IFormRegionFactory.Kind
            {
                get
                {
                    return Microsoft.Office.Tools.Outlook.FormRegionKindConstants.WindowsForms;
                }
            }
        }
    }

    partial class WindowFormRegionCollection
    {
        internal regionConversationItems regionConversationItems
        {
            get
            {
                foreach (var item in this)
                {
                    if (item.GetType() == typeof(regionConversationItems))
                        return (regionConversationItems)item;
                }
                return null;
            }
        }
    }
}
