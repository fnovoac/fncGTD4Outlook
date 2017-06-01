using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using fncGTD4Outlook.Comun;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace fncGTD4Outlook
{
    public partial class RibbonReadEmail
    {
        private void RibbonReadEmail_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnArchivar_Click(object sender, RibbonControlEventArgs e)
        {
            Utils.MoverEmail(Utils.getMailItems(), Constants.folderArchivar);
        }

        private void btnCompletado_Click(object sender, RibbonControlEventArgs e)
        {
            List<Outlook.MailItem> emails = new List<Outlook.MailItem>();

            try
            {
                emails = Utils.getMailItems();

                for (int i = 0; i < emails.Count; i++)
                {
                    if (emails[i] != null)
                    {
                        emails[i].TaskCompletedDate = DateTime.Today;
                        emails[i].Save();
                        Utils.MoverEmail(emails[i], Constants.folderArchivar);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                for (int i = 0; i < emails.Count; i++)
                    if (emails[i] != null) Marshal.ReleaseComObject(emails[i]);
            }
        }

        private void btnDelegar_Click(object sender, RibbonControlEventArgs e)
        {
            frmDelegar f = new frmDelegar();
            f.ShowDialog();
            f.Dispose();
            f = null;
        }

        private void btnDiferir_Click(object sender, RibbonControlEventArgs e)
        {
            frmDiferir f = new frmDiferir();
            f.ShowDialog();
            f.Dispose();
            f = null;
        }

        private void btnEliminar_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                List<Outlook.MailItem> mailItems = Utils.getMailItems();
                if (mailItems != null)
                {
                    for (int i = 0; i < mailItems.Count; i++)
                    {

                        mailItems[i].Delete();
                        if (mailItems[i] != null) Marshal.ReleaseComObject(mailItems[i]);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnConservar_Click(object sender, RibbonControlEventArgs e)
        {
            Utils.MoverEmail(Utils.getMailItems(), Constants.folderConservar);
        }

        private void btnReferencia_Click(object sender, RibbonControlEventArgs e)
        {
            Utils.MoverEmail(Utils.getMailItems(), Constants.folderReferencia);
        }
    }
}
