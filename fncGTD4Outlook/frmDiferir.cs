using fncGTD4Outlook.Comun;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace fncGTD4Outlook
{
    public partial class frmDiferir : Form
    {
        List<Outlook.MailItem> emails = null;

        public frmDiferir()
        {
            InitializeComponent();
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmDiferir_Load(object sender, EventArgs e)
        {
            
            //configuramos la opción de autocompletado del combo de plazos
            AutoCompleteStringCollection stringPlazos = new AutoCompleteStringCollection();
            stringPlazos.Add("Ninguno");
            stringPlazos.Add("Mañana");
            stringPlazos.Add("En 2 días");
            stringPlazos.Add("Proxima semana");
            stringPlazos.Add("En 2 semanas");
            stringPlazos.Add("En 30 días");
            stringPlazos.Add("Personalizado");

            cboPlazo.AutoCompleteCustomSource = stringPlazos;
            cboPlazo.AutoCompleteSource = AutoCompleteSource.CustomSource;
            cboPlazo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;

            //Llenamos el combo de plazo
            cboPlazo.Items.Add("Ninguno");
            cboPlazo.Items.Add("Mañana");
            cboPlazo.Items.Add("En 2 días");
            cboPlazo.Items.Add("Proxima semana");
            cboPlazo.Items.Add("En 2 semanas");
            cboPlazo.Items.Add("En 30 días");
            cboPlazo.Items.Add("Personalizado");
            cboPlazo.SelectedIndex = 0;

            try
            {
                emails = Utils.GetMailItems();

                if (emails != null)
                {
                    if (emails.Count > 1)
                    {
                        lblCantidadEmails.Visible = true;
                        lblCantidadEmails.Text = String.Format("{0} Emails seleccionados", emails.Count);
                    }
                }

            }
            catch (Exception)
            {
                MessageBox.Show("No se pudieron obtener los emails seleccionados", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            

        }

        private void cboPlazo_SelectedIndexChanged(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Today;

            switch (cboPlazo.SelectedIndex)
            {
                case 0:
                    dt = DateTime.Today;
                    monthCalendar.SetDate(dt);
                    lblPlazo.Text = "No fijado";
                    monthCalendar.Enabled = false;
                    break;

                case 1:
                    monthCalendar.Enabled = true;
                    dt = DateTime.Today.AddDays(1);
                    monthCalendar.SetDate(dt);
                    lblPlazo.Text = String.Format("{0:D}", dt);
                    break;

                case 2:
                    monthCalendar.Enabled = true;
                    dt = DateTime.Today.AddDays(2);
                    monthCalendar.SetDate(dt);
                    lblPlazo.Text = String.Format("{0:D}", dt);
                    break;

                case 3:
                    monthCalendar.Enabled = true;
                    dt = DateTime.Today.AddDays(7);
                    monthCalendar.SetDate(dt);
                    lblPlazo.Text = String.Format("{0:D}", dt);
                    break;

                case 4:
                    monthCalendar.Enabled = true;
                    dt = DateTime.Today.AddDays(14);
                    monthCalendar.SetDate(dt);
                    lblPlazo.Text = String.Format("{0:D}", dt);
                    break;

                case 5:
                    monthCalendar.Enabled = true;
                    dt = DateTime.Today.AddDays(30);
                    monthCalendar.SetDate(dt);
                    lblPlazo.Text = String.Format("{0:D}", dt);
                    break;

                case 6:
                    monthCalendar.Enabled = true;
                    dt = DateTime.Today;
                    monthCalendar.SetDate(dt);
                    lblPlazo.Text = String.Format("{0:D}", dt);
                    break;
            }
        }

        private void cboPlazo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                btnDiferir.Focus();
            }
        }

        private void monthCalendar_DateChanged(object sender, DateRangeEventArgs e)
        {
            DateTime dt = monthCalendar.SelectionEnd;
            lblPlazo.Text = String.Format("{0:D}", dt);
        }

        private void frmDiferir_FormClosing(object sender, FormClosingEventArgs e)
        {
            for (int i = 0; i < emails.Count; i++)
            {
                if (emails[i] != null) Marshal.ReleaseComObject(emails[i]);
            }
        }

        private void btnDiferir_Click(object sender, EventArgs e)
        {
            Outlook.NameSpace ns = null;
            //Outlook.AppointmentItem apptItem = null;
            try
            {
                if (cboPlazo.SelectedIndex > 0)
                {
                    for (int i = 0; i < emails.Count; i++)
                    {
                        //obtenemos el folder donde moveremos el email (debe existir -> ver ThisAddIn.cs)
                        Outlook.MAPIFolder objfolder = Utils.GetFolderByName(Constants.folderDiferir);
                        if (objfolder == null)
                            objfolder = Globals.ThisAddIn.Application.Session.DefaultStore.GetRootFolder().Folders.Add(Constants.folderDiferir);

                        emails[i].TaskDueDate = DateTime.Parse(lblPlazo.Text);
                        emails[i].UnRead = false;
                        emails[i].Save();

                        try
                        {
                            emails[i].Move(objfolder);
                        }
                        catch (Exception)
                        {

                        }

                        //ns = Globals.ThisAddIn.Application.GetNamespace("MAPI");
                        //apptItem = ns.Session.Application.CreateItem(Outlook.OlItemType.olAppointmentItem);
                        //apptItem.Subject = emails[i].Subject;
                        //apptItem.RTFBody = emails[i].RTFBody;
                        //apptItem.AllDayEvent = true;
                        //apptItem.Start = DateTime.Parse(lblPlazo.Text);
                        //apptItem.End = DateTime.Parse(lblPlazo.Text);
                        //apptItem.Attachments.Add(emails[i]);
                        //apptItem.BillingInformation = emails[i].EntryID;
                        //apptItem.Save();
                        //if (apptItem != null) Marshal.ReleaseComObject(apptItem);
                    }

                    this.Close();
                }
                else
                {
                    MessageBox.Show("Debe seleccionar una fecha", "Diferir", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (ns != null) Marshal.ReleaseComObject(ns);
                for (int i = 0; i < emails.Count; i++)
                {
                    if (emails[i] != null) Marshal.ReleaseComObject(emails[i]);
                }
            }
            
        }
    }
}
