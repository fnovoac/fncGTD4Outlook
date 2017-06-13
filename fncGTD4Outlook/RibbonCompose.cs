using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using fncGTD4Outlook.Comun;
using System.Runtime.InteropServices;

namespace fncGTD4Outlook
{
    public partial class RibbonCompose
    {
        private void RibbonCompose_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnDelegarEnviar_Click(object sender, RibbonControlEventArgs e)
        {
            frmDelegarEnviar f = new frmDelegarEnviar();
            f.Show();
            f = null;
        }

        private void btnCompletarEnviar_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.MailItem email = null;
            Outlook.MailItem emailTask = null;
            object oItem = null;

            try
            {
                email = Utils.GetMailItem();
                if (email != null)
                {
                    if (email.Subject != string.Empty)
                    {
                        if (Utils.originalEmailsAsTask != null)
                        {
                            for (int i = 0; i < Utils.originalEmailsAsTask.Count; i++)
                            {
                                oItem = Globals.ThisAddIn.Application.Session.GetItemFromID(Utils.originalEmailsAsTask[i]);
                                if (oItem is Outlook.MailItem)
                                {
                                    emailTask = oItem as Outlook.MailItem;
                                    emailTask.TaskCompletedDate = DateTime.Today;
                                    emailTask.Save();
                                    Utils.MoverEmail(emailTask, Constants.folderArchivar);
                                }

                            }
                        }
                    }
                }
            }
            catch (Exception)
            {

            }

            if (email != null) Marshal.ReleaseComObject(email);
            if (oItem != null) Marshal.ReleaseComObject(email);
        }
    }
}
