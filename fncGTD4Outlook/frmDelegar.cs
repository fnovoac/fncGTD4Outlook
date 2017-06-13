using fncGTD4Outlook.Comun;
using fncGTD4Outlook.Controles;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace fncGTD4Outlook
{
    public partial class frmDelegar : Form
    {
        private bool _EmailHasData = false;

        public frmDelegar()
        {
            InitializeComponent();
        }

        private void frmDelegar_Load(object sender, EventArgs e)
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


            //Verificamos la data que el email ya pueda tener (en caso estemos haciendo una actualización)
            List<Outlook.MailItem> mails = new List<Outlook.MailItem>();
            mails = Utils.GetMailItems();

            if (mails != null && mails.Count > 0)
            {
                if (mails.Count == 1)
                {
                    //Si está marcada es porque ya hemos llenado información antes
                    if (mails[0].IsMarkedAsTask)
                    {
                        _EmailHasData = true;

                        txtContacto.Text = mails[0].BillingInformation;
                        if (mails[0].TaskDueDate.Year == 4501)
                        {
                            //Si es este año es porque no tiene asignado una fecha en realidad
                            cboPlazo.SelectedIndex = 0;
                        }
                        else
                        {
                            lblPlazo.Text = mails[0].TaskDueDate.ToString();
                            monthCalendar.SetDate(DateTime.Parse(lblPlazo.Text));
                            cboPlazo.SelectedIndex = 6;
                        }

                        if (mails[0].ReminderSet)
                        {
                            chkRecordatorio.Checked = true;
                            dateTimePickerRecordatorio.Value = mails[0].ReminderTime;
                        }
                    }

                }
                else
                {
                    lblCantidadEmails.Visible = true;
                    lblCantidadEmails.Text=  String.Format("{0} Emails seleccionados", mails.Count);
                }

                for (int i = 0; i < mails.Count; i++)
                {
                    if (mails[i] != null) Marshal.ReleaseComObject(mails[i]);
                }

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

            #region ? Listar Categorias
            //Outlook.Categories categories = Globals.ThisAddIn.Application.Session.Categories;
            //foreach (Outlook.Category category in categories)
            //    MessageBox.Show(category.Name);
            #endregion



        }

        private void btnDelegar_Click(object sender, EventArgs e)
        {
            List<Outlook.MailItem> mails = new List<Outlook.MailItem>();

            try
            {
                mails = Utils.GetMailItems();

                if (mails != null && mails.Count > 0)
                {
                    for (int i = 0; i < mails.Count; i++)
                    {
                        mails[i].ClearTaskFlag();

                        #region ? Propiedades de usuario
                        //esto debe ir al inicio:
                        //Outlook.UserProperties itemProps = null;
                        //Outlook.UserProperty newProp = null;

                        //esto debe ir en el Finally
                        //if (newProp != null) Marshal.ReleaseComObject(newProp);
                        //if (itemProps != null) Marshal.ReleaseComObject(itemProps);

                        //itemProps = mails[0].UserProperties;

                        //newProp = itemProps.Find("DelegadoNombreFull");
                        //if (newProp != null)
                        //{
                        //    newProp.Value = txtContacto.Text;
                        //}
                        //else
                        //{
                        //    newProp = itemProps.Add("DelegadoNombreFull", Outlook.OlUserPropertyType.olText, true);
                        //    newProp.Value = txtContacto.Text;
                        //}
                        #endregion

                        mails[i].MarkAsTask(Outlook.OlMarkInterval.olMarkNoDate);

                        if (cboPlazo.SelectedIndex > 0)
                        {
                            DateTime vencimiento = DateTime.Parse(lblPlazo.Text);
                            mails[i].TaskDueDate = vencimiento;

                            if (chkRecordatorio.Checked)
                            {
                                TimeSpan st = dateTimePickerRecordatorio.Value.TimeOfDay;
                                DateTime dt = vencimiento + st;
                                mails[i].ReminderSet = true;
                                mails[i].ReminderTime = dt;
                            }
                        }

                        mails[i].BillingInformation = txtContacto.Text;
                        mails[i].Save();
                        try
                        {
                            Utils.MoverEmail(mails[i], Constants.folderDelegar);
                        }
                        catch (Exception)
                        {

                        }

                    }
                }
            }
            catch (Exception ex)
            {
                //Número de error al mover el email a un folder (salió cuando traté de moverlo al mismo folder donde ya estaba)
                if (ex.HResult != -2147352567)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                
            }
            finally
            {
                for (int i = 0; i < mails.Count; i++)
                {
                    if (mails[i] != null) Marshal.ReleaseComObject(mails[i]);
                }

                this.Close();
            }
            
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void chkRecordatorio_CheckedChanged(object sender, EventArgs e)
        {
            dateTimePickerRecordatorio.Enabled = chkRecordatorio.Checked;
        }

        private void monthCalendar_DateChanged(object sender, DateRangeEventArgs e)
        {
            DateTime dt = monthCalendar.SelectionEnd;
            lblPlazo.Text = String.Format("{0:D}", dt);
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
                    if (!_EmailHasData)
                    {
                        dt = DateTime.Today;
                        monthCalendar.SetDate(dt);
                        lblPlazo.Text = String.Format("{0:D}", dt);
                    }
                    break;
            }
                
        }

        private void cboPlazo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                chkRecordatorio.Focus();
            }
        }

        private void txtContacto_TextChanged(object sender, EventArgs e)
        {
            //txtContacto.Width = 441;
        }

        private void txtContacto_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                cboPlazo.Focus();
            }
        }

        private void txtContacto_Leave(object sender, EventArgs e)
        {
            //if (txtContacto.TextLength > 0)
            //{
            //    Size size = TextRenderer.MeasureText(txtContacto.Text, txtContacto.Font);
            //    txtContacto.Width = size.Width;
            //    //txtContacto.Height = size.Height;
            //}
            //else
            //{
            //    txtContacto.Width = 441;
            //}
        }

        private void chkRecordatorio_EnabledChanged(object sender, EventArgs e)
        {
            dateTimePickerRecordatorio.Enabled = chkRecordatorio.Checked;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
