using fncGTD4Outlook.Comun;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace fncGTD4Outlook
{
    public partial class frmDelegarEnviar : Form
    {
        private Outlook.MailItem email = null;

        public frmDelegarEnviar()
        {
            InitializeComponent();
        }

        private void frmDelegarEnviar_Load(object sender, EventArgs e)
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


            email = Utils.GetMailItem();

            if (email != null)
            {
                string asignadoA = Utils.GetFirstReceiverFromTo(email.To);
                txtContacto.Text = Utils.SubstringBefore(asignadoA, "(").Trim();
            }


            //configuramos la opción de autocompletado del textbox
            List<Outlook.ContactItem> contactos = new List<Outlook.ContactItem>();
            contactos = Utils.GetListOfContacts(false);

            AutoCompleteStringCollection stringCol = new AutoCompleteStringCollection();
            for (int i = 0; i < contactos.Count; i++)
            {
                stringCol.Add(contactos[i].FullName);
                Marshal.ReleaseComObject(contactos[i]);
            }

            txtContacto.AutoCompleteCustomSource = stringCol;
            txtContacto.AutoCompleteSource = AutoCompleteSource.CustomSource;
            txtContacto.AutoCompleteMode = AutoCompleteMode.SuggestAppend;

            //Completamos los status
            cboStatus.Items.Add("0 No iniciado");
            cboStatus.Items.Add("1 En proceso");
            cboStatus.Items.Add("2 Completado");
            cboStatus.Items.Add("3 Esperando a");
            cboStatus.Items.Add("4 Diferido");
            cboStatus.SelectedIndex = 0;

        }

        private void txtContacto_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                cboPlazo.Focus();
            }
        }

        private void cboPlazo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                chkRecordatorio.Focus();
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
                    chkRecordatorio.Enabled = false;
                    break;

                case 1:
                    monthCalendar.Enabled = true;
                    chkRecordatorio.Enabled = true;
                    dt = DateTime.Today.AddDays(1);
                    monthCalendar.SetDate(dt);
                    lblPlazo.Text = String.Format("{0:D}", dt);
                    break;

                case 2:
                    monthCalendar.Enabled = true;
                    chkRecordatorio.Enabled = true;
                    dt = DateTime.Today.AddDays(2);
                    monthCalendar.SetDate(dt);
                    lblPlazo.Text = String.Format("{0:D}", dt);
                    break;

                case 3:
                    monthCalendar.Enabled = true;
                    chkRecordatorio.Enabled = true;
                    dt = DateTime.Today.AddDays(7);
                    monthCalendar.SetDate(dt);
                    lblPlazo.Text = String.Format("{0:D}", dt);
                    break;

                case 4:
                    monthCalendar.Enabled = true;
                    chkRecordatorio.Enabled = true;
                    dt = DateTime.Today.AddDays(14);
                    monthCalendar.SetDate(dt);
                    lblPlazo.Text = String.Format("{0:D}", dt);
                    break;

                case 5:
                    monthCalendar.Enabled = true;
                    chkRecordatorio.Enabled = true;
                    dt = DateTime.Today.AddDays(30);
                    monthCalendar.SetDate(dt);
                    lblPlazo.Text = String.Format("{0:D}", dt);
                    break;

                case 6:
                    monthCalendar.Enabled = true;
                    chkRecordatorio.Enabled = true;
                    dt = DateTime.Today;
                    monthCalendar.SetDate(dt);
                    lblPlazo.Text = String.Format("{0:D}", dt);
                    break;
            }
        }

        private void monthCalendar_DateChanged(object sender, DateRangeEventArgs e)
        {
            DateTime dt = monthCalendar.SelectionEnd;
            lblPlazo.Text = String.Format("{0:D}", dt);
        }

        private void chkRecordatorio_CheckedChanged(object sender, EventArgs e)
        {
            dateTimePickerRecordatorio.Enabled = chkRecordatorio.Checked;
        }

        private void chkRecordatorio_EnabledChanged(object sender, EventArgs e)
        {
            dateTimePickerRecordatorio.Enabled = chkRecordatorio.Checked;
        }

        private void btnDelegar_Click(object sender, EventArgs e)
        {
            bool retValue = false;
            Outlook.UserProperties itemProps = null;
            Outlook.UserProperty newProp = null;
            Outlook.Recipients recipients = null;
            Outlook.Recipient recipientBCC = null;

            try
            {
                recipients = email.Recipients;
                recipientBCC = recipients.Add(Constants.myEmail);
                recipientBCC.Type = (int)Outlook.OlMailRecipientType.olBCC;
                retValue = recipients.ResolveAll();

                email.ClearTaskFlag();

                itemProps = email.UserProperties;

                newProp = itemProps.Find("MarkAsTask");
                if (newProp != null)
                {
                    newProp.Value = Outlook.OlMarkInterval.olMarkNoDate;
                }
                else
                {
                    newProp = itemProps.Add("MarkAsTask", Outlook.OlUserPropertyType.olInteger, true);
                    newProp.Value = Outlook.OlMarkInterval.olMarkNoDate;
                }

                if (cboPlazo.SelectedIndex > 0)
                {
                    DateTime vencimiento = DateTime.Parse(lblPlazo.Text);

                    newProp = itemProps.Find("TaskDueDate");
                    if (newProp != null)
                    {
                        newProp.Value = vencimiento;
                    }
                    else
                    {
                        newProp = itemProps.Add("TaskDueDate", Outlook.OlUserPropertyType.olDateTime, true);
                        newProp.Value = vencimiento;
                    }

                    if (chkRecordatorio.Checked)
                    {
                        TimeSpan st = dateTimePickerRecordatorio.Value.TimeOfDay;
                        DateTime dt = vencimiento + st;

                        newProp = itemProps.Find("ReminderSet");
                        if (newProp != null)
                        {
                            newProp.Value = true;
                        }
                        else
                        {
                            newProp = itemProps.Add("ReminderSet", Outlook.OlUserPropertyType.olYesNo, true);
                            newProp.Value = true;
                        }

                        newProp = itemProps.Find("ReminderTime");
                        if (newProp != null)
                        {
                            newProp.Value = dt;
                        }
                        else
                        {
                            newProp = itemProps.Add("ReminderTime", Outlook.OlUserPropertyType.olDateTime, true);
                            newProp.Value = dt;
                        }
                    }
                }

                email.BillingInformation = txtContacto.Text;
                email.Send();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (recipientBCC != null) Marshal.ReleaseComObject(recipientBCC);
                if (recipients != null) Marshal.ReleaseComObject(recipients);
                if (newProp != null) Marshal.ReleaseComObject(newProp);
                if (itemProps != null) Marshal.ReleaseComObject(itemProps);

                this.Close();
            }

        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
