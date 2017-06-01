namespace fncGTD4Outlook
{
    partial class frmDelegarEnviar
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtContacto = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.lblPlazo = new System.Windows.Forms.Label();
            this.cboPlazo = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.monthCalendar = new System.Windows.Forms.MonthCalendar();
            this.dateTimePickerRecordatorio = new System.Windows.Forms.DateTimePicker();
            this.chkRecordatorio = new System.Windows.Forms.CheckBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnDelegar = new System.Windows.Forms.Button();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtContacto
            // 
            this.txtContacto.Location = new System.Drawing.Point(94, 18);
            this.txtContacto.Name = "txtContacto";
            this.txtContacto.Size = new System.Drawing.Size(406, 23);
            this.txtContacto.TabIndex = 1;
            this.txtContacto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtContacto_KeyDown);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 15);
            this.label1.TabIndex = 2;
            this.label1.Text = "Delegar a:";
            // 
            // lblPlazo
            // 
            this.lblPlazo.AutoSize = true;
            this.lblPlazo.Location = new System.Drawing.Point(281, 63);
            this.lblPlazo.Name = "lblPlazo";
            this.lblPlazo.Size = new System.Drawing.Size(56, 15);
            this.lblPlazo.TabIndex = 16;
            this.lblPlazo.Text = "No fijado";
            // 
            // cboPlazo
            // 
            this.cboPlazo.FormattingEnabled = true;
            this.cboPlazo.Location = new System.Drawing.Point(94, 59);
            this.cboPlazo.Name = "cboPlazo";
            this.cboPlazo.Size = new System.Drawing.Size(181, 23);
            this.cboPlazo.TabIndex = 14;
            this.cboPlazo.SelectedIndexChanged += new System.EventHandler(this.cboPlazo_SelectedIndexChanged);
            this.cboPlazo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cboPlazo_KeyDown);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 62);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(76, 15);
            this.label2.TabIndex = 15;
            this.label2.Text = "Vencimiento:";
            // 
            // monthCalendar
            // 
            this.monthCalendar.BackColor = System.Drawing.Color.White;
            this.monthCalendar.CalendarDimensions = new System.Drawing.Size(2, 1);
            this.monthCalendar.Location = new System.Drawing.Point(6, 88);
            this.monthCalendar.MaxSelectionCount = 1;
            this.monthCalendar.Name = "monthCalendar";
            this.monthCalendar.TabIndex = 17;
            this.monthCalendar.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar_DateChanged);
            // 
            // dateTimePickerRecordatorio
            // 
            this.dateTimePickerRecordatorio.CustomFormat = "hh:mm tt";
            this.dateTimePickerRecordatorio.Enabled = false;
            this.dateTimePickerRecordatorio.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePickerRecordatorio.Location = new System.Drawing.Point(144, 256);
            this.dateTimePickerRecordatorio.Name = "dateTimePickerRecordatorio";
            this.dateTimePickerRecordatorio.ShowUpDown = true;
            this.dateTimePickerRecordatorio.Size = new System.Drawing.Size(85, 23);
            this.dateTimePickerRecordatorio.TabIndex = 19;
            // 
            // chkRecordatorio
            // 
            this.chkRecordatorio.AutoSize = true;
            this.chkRecordatorio.Location = new System.Drawing.Point(22, 258);
            this.chkRecordatorio.Name = "chkRecordatorio";
            this.chkRecordatorio.Size = new System.Drawing.Size(116, 19);
            this.chkRecordatorio.TabIndex = 18;
            this.chkRecordatorio.Text = "Fijar recordatorio";
            this.chkRecordatorio.UseVisualStyleBackColor = true;
            this.chkRecordatorio.CheckedChanged += new System.EventHandler(this.chkRecordatorio_CheckedChanged);
            this.chkRecordatorio.EnabledChanged += new System.EventHandler(this.chkRecordatorio_EnabledChanged);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.Control;
            this.panel1.Controls.Add(this.btnDelegar);
            this.panel1.Controls.Add(this.btnCancelar);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 311);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(514, 49);
            this.panel1.TabIndex = 20;
            // 
            // btnDelegar
            // 
            this.btnDelegar.Location = new System.Drawing.Point(289, 14);
            this.btnDelegar.Name = "btnDelegar";
            this.btnDelegar.Size = new System.Drawing.Size(106, 23);
            this.btnDelegar.TabIndex = 5;
            this.btnDelegar.Text = "Delegar y Enviar";
            this.btnDelegar.UseVisualStyleBackColor = true;
            this.btnDelegar.Click += new System.EventHandler(this.btnDelegar_Click);
            // 
            // btnCancelar
            // 
            this.btnCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancelar.Location = new System.Drawing.Point(401, 14);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(75, 23);
            this.btnCancelar.TabIndex = 6;
            this.btnCancelar.Text = "Cancelar";
            this.btnCancelar.UseVisualStyleBackColor = true;
            this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
            // 
            // frmDelegarEnviar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.CancelButton = this.btnCancelar;
            this.ClientSize = new System.Drawing.Size(514, 360);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.dateTimePickerRecordatorio);
            this.Controls.Add(this.chkRecordatorio);
            this.Controls.Add(this.monthCalendar);
            this.Controls.Add(this.lblPlazo);
            this.Controls.Add(this.cboPlazo);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtContacto);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmDelegarEnviar";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Delegar y Enviar";
            this.Load += new System.EventHandler(this.frmDelegarEnviar_Load);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtContacto;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblPlazo;
        private System.Windows.Forms.ComboBox cboPlazo;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.MonthCalendar monthCalendar;
        private System.Windows.Forms.DateTimePicker dateTimePickerRecordatorio;
        private System.Windows.Forms.CheckBox chkRecordatorio;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnDelegar;
        private System.Windows.Forms.Button btnCancelar;
    }
}