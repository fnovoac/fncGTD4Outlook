namespace fncGTD4Outlook
{
    partial class frmDiferir
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnDiferir = new System.Windows.Forms.Button();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.monthCalendar = new System.Windows.Forms.MonthCalendar();
            this.lblPlazo = new System.Windows.Forms.Label();
            this.cboPlazo = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.lblCantidadEmails = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.Control;
            this.panel1.Controls.Add(this.btnDiferir);
            this.panel1.Controls.Add(this.btnCancelar);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 247);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(514, 49);
            this.panel1.TabIndex = 21;
            // 
            // btnDiferir
            // 
            this.btnDiferir.Location = new System.Drawing.Point(289, 14);
            this.btnDiferir.Name = "btnDiferir";
            this.btnDiferir.Size = new System.Drawing.Size(106, 23);
            this.btnDiferir.TabIndex = 2;
            this.btnDiferir.Text = "Diferir";
            this.btnDiferir.UseVisualStyleBackColor = true;
            this.btnDiferir.Click += new System.EventHandler(this.btnDiferir_Click);
            // 
            // btnCancelar
            // 
            this.btnCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancelar.Location = new System.Drawing.Point(401, 14);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(75, 23);
            this.btnCancelar.TabIndex = 3;
            this.btnCancelar.Text = "Cancelar";
            this.btnCancelar.UseVisualStyleBackColor = true;
            this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
            // 
            // monthCalendar
            // 
            this.monthCalendar.BackColor = System.Drawing.Color.White;
            this.monthCalendar.CalendarDimensions = new System.Drawing.Size(2, 1);
            this.monthCalendar.Location = new System.Drawing.Point(7, 47);
            this.monthCalendar.MaxSelectionCount = 1;
            this.monthCalendar.Name = "monthCalendar";
            this.monthCalendar.TabIndex = 1;
            this.monthCalendar.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar_DateChanged);
            // 
            // lblPlazo
            // 
            this.lblPlazo.AutoSize = true;
            this.lblPlazo.Location = new System.Drawing.Point(281, 21);
            this.lblPlazo.Name = "lblPlazo";
            this.lblPlazo.Size = new System.Drawing.Size(56, 15);
            this.lblPlazo.TabIndex = 25;
            this.lblPlazo.Text = "No fijado";
            // 
            // cboPlazo
            // 
            this.cboPlazo.FormattingEnabled = true;
            this.cboPlazo.Location = new System.Drawing.Point(94, 18);
            this.cboPlazo.Name = "cboPlazo";
            this.cboPlazo.Size = new System.Drawing.Size(181, 23);
            this.cboPlazo.TabIndex = 0;
            this.cboPlazo.SelectedIndexChanged += new System.EventHandler(this.cboPlazo_SelectedIndexChanged);
            this.cboPlazo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cboPlazo_KeyDown);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 21);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(42, 15);
            this.label2.TabIndex = 24;
            this.label2.Text = "Diferir:";
            // 
            // lblCantidadEmails
            // 
            this.lblCantidadEmails.BackColor = System.Drawing.Color.Coral;
            this.lblCantidadEmails.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lblCantidadEmails.ForeColor = System.Drawing.Color.White;
            this.lblCantidadEmails.Location = new System.Drawing.Point(0, 226);
            this.lblCantidadEmails.Name = "lblCantidadEmails";
            this.lblCantidadEmails.Size = new System.Drawing.Size(514, 21);
            this.lblCantidadEmails.TabIndex = 26;
            this.lblCantidadEmails.Text = "1 Email seleccionado";
            this.lblCantidadEmails.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblCantidadEmails.Visible = false;
            // 
            // frmDiferir
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.CancelButton = this.btnCancelar;
            this.ClientSize = new System.Drawing.Size(514, 296);
            this.Controls.Add(this.lblCantidadEmails);
            this.Controls.Add(this.lblPlazo);
            this.Controls.Add(this.cboPlazo);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.monthCalendar);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmDiferir";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Diferir emails";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmDiferir_FormClosing);
            this.Load += new System.EventHandler(this.frmDiferir_Load);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnDiferir;
        private System.Windows.Forms.Button btnCancelar;
        private System.Windows.Forms.MonthCalendar monthCalendar;
        private System.Windows.Forms.Label lblPlazo;
        private System.Windows.Forms.ComboBox cboPlazo;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblCantidadEmails;
    }
}