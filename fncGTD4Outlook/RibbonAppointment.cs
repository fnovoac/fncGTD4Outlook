using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace fncGTD4Outlook
{
    public partial class RibbonAppointment
    {
        private void RibbonAppointment_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnAbrilEmail_Click(object sender, RibbonControlEventArgs e)
        {
            //Outlook.AppointmentItem appItem = null;
            //Outlook.MailItem email = null;
            //object obj = null;

            //try
            //{
            //    appItem = Utils.getAppointmentItem();
            //    if (appItem != null && appItem.BillingInformation != string.Empty)
            //    {
            //        obj = Globals.ThisAddIn.Application.Session.GetItemFromID(appItem.BillingInformation);
            //        if (obj is Outlook.MailItem)
            //        {
            //            email = obj as Outlook.MailItem;
            //            email.Display(true);
            //        }
            //    }

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
            //finally
            //{
            //    if (appItem != null) Marshal.ReleaseComObject(appItem);
            //    if (obj != null) Marshal.ReleaseComObject(obj);
            //    if (email != null) Marshal.ReleaseComObject(email);
            //}
        }
    }
}
