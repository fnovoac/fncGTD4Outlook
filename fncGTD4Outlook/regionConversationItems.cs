using fncGTD4Outlook.Comun;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace fncGTD4Outlook
{
    partial class regionConversationItems
    {
        private BackgroundWorker bgWorker;

        //List<string[]> listConversaciones = new List<string[]>();


        #region Form Region Factory 

        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Note)]
        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Post)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("fncGTD4Outlook.regionConversationItems")]
        public partial class regionConversationItemsFactory
        {
            // Occurs before the form region is initialized.
            // To prevent the form region from appearing, set e.Cancel to true.
            // Use e.OutlookItem to get a reference to the current Outlook item.
            private void regionConversationItemsFactory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {
            }
        }

        #endregion

        // Occurs before the form region is displayed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void regionConversationItems_FormRegionShowing(object sender, System.EventArgs e)
        {
            
        }

        // Occurs when the form region is closed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void regionConversationItems_FormRegionClosed(object sender, System.EventArgs e)
        {
        }

        private async void regionConversationItems_Load(object sender, EventArgs e)
        {
            listViewConversation.View = View.Details;
            listViewConversation.FullRowSelect = true;
            listViewConversation.MultiSelect = false;
            listViewConversation.HeaderStyle = ColumnHeaderStyle.Nonclickable;

            listViewConversation.Columns.Add("EntryID");
            listViewConversation.Columns.Add("Recibido");
            listViewConversation.Columns.Add("Ubicación");
            listViewConversation.Columns.Add("De");
            listViewConversation.Columns.Add("Mensaje");
            listViewConversation.Columns.Add("Tarea");
            listViewConversation.Columns.Add("Adjuntos");

            listViewConversation.Columns[0].Width = 0;

            Outlook.MailItem selectedMailm = null;
            Outlook.MailItem mailItem = null;
            Outlook.Conversation conv = null;
            Outlook.Table oTable = null;
            Outlook.Row oRow = null;
            object oItem = null;
            Outlook.Folder inFolder = null;

            try
            {
                ListViewItem[] result = await Task.Run(() =>
                {
                    string lastID = string.Empty;

                    selectedMailm = this.OutlookItem as Outlook.MailItem;

                    List<string[]> listData = new List<string[]>();

                    if (Utils.originalEmailsAsTask != null) Utils.originalEmailsAsTask.Clear();

                    if (selectedMailm != null)
                    {
                        if (selectedMailm.ConversationID != null)
                        {
                            // Obtain a Conversation object. 
                            conv = selectedMailm.GetConversation();
                            if (conv != null)
                            {
                                oTable = conv.GetTable();

                                while (!oTable.EndOfTable)
                                {
                                    oRow = oTable.GetNextRow();
                                    oItem = Globals.ThisAddIn.Application.Session.GetItemFromID(oRow["EntryID"]);
                                    if (oItem is Outlook.MailItem)
                                    {
                                        mailItem = oItem as Outlook.MailItem;
                                        inFolder = mailItem.Parent as Outlook.Folder;

                                        //por alguna razón algunas veces algunos elementos se agregan 2 veces
                                        if (lastID != oRow["EntryID"])
                                        {
                                            string[] row = {oRow["EntryID"],
                                                        mailItem.ReceivedTime.ToShortDateString() + " " + mailItem.ReceivedTime.ToShortTimeString(),
                                                        inFolder.Name,
                                                        mailItem.SenderName,
                                                        mailItem.Body.Substring(0,180) + "...",
                                                        mailItem.IsMarkedAsTask.ToString(),
                                                        (_cantidadAdjuntos(mailItem)>0)?_cantidadAdjuntos(mailItem).ToString():""
                                                        };

                                            listData.Add(row);

                                            if (mailItem.IsMarkedAsTask)
                                            {
                                                Utils.originalEmailsAsTask.Add(oRow["EntryID"]);
                                            }
                                        }

                                        lastID = oRow["EntryID"];

                                        //Marshal.ReleaseComObject(mailItem);
                                        //Marshal.ReleaseComObject(inFolder);
                                    }
                                    //Marshal.ReleaseComObject(oItem);
                                    //Marshal.ReleaseComObject(oRow);
                                }

                            }
                        }
                    }

                    ListViewItem[] listaRange = new ListViewItem[listData.Count];
                    for (int i = listData.Count - 1; i >= 0; i--)
                    {
                        listaRange[i] = new ListViewItem(listData[i]);
                    }

                    return listaRange;
                });



                listViewConversation.Items.AddRange(result);

                listViewConversation.SuspendLayout();
                listViewConversation.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
                listViewConversation.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
                listViewConversation.Columns[0].Width = 0;
                listViewConversation.ResumeLayout();

                //////listViewConversation.Items.Clear();
                //////listViewConversation.BeginUpdate();
                ////////for (int i = result.Count - 1; i >= 0; i--)
                ////////{
                ////////    listViewConversation.Items.Add(new ListViewItem(result[i]));
                ////////}
                //////listViewConversation.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
                //////listViewConversation.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
                //////listViewConversation.Columns[0].Width = 0;
                //////listViewConversation.EndUpdate();

            }
            catch (Exception)
            {

            }

            if (selectedMailm != null) Marshal.ReleaseComObject(selectedMailm);
            if (mailItem != null) Marshal.ReleaseComObject(mailItem);
            if (inFolder != null) Marshal.ReleaseComObject(inFolder);
            if (oRow != null) Marshal.ReleaseComObject(oRow);
            if (oTable != null) Marshal.ReleaseComObject(oTable);
            if (oItem != null) Marshal.ReleaseComObject(oItem);
            if (conv != null) Marshal.ReleaseComObject(conv);

            //////////////// Set up background worker object & hook up handlers
            //////////////bgWorker = new BackgroundWorker();
            //////////////bgWorker.DoWork += BgWorker_DoWork;
            //////////////bgWorker.RunWorkerCompleted += BgWorker_RunWorkerCompleted;

            //////////////// Launch background thread to do the work of reading the file.  This will
            //////////////// trigger BackgroundWorker.DoWork().  Note that we pass the filename to
            //////////////// process as a parameter.
            //////////////bgWorker.RunWorkerAsync();

        }

        private void LlenarListViewConversaciones()
        {
            //Task task = Task.Run(() =>
            //{
            //    ObtenerMailsDeConversacion();
            //});

            //Task UITask = task.ContinueWith(_ =>
            //{
            //    listViewConversation.Items.Clear();
            //    //MessageBox.Show("bien!");
            //    listViewConversation.BeginUpdate();
            //    for (int i = listConversaciones.Count - 1; i >= 0; i--)
            //    {
            //        listViewConversation.Items.Add(new ListViewItem(listConversaciones[i]));
            //    }
            //    listViewConversation.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            //    listViewConversation.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
            //    listViewConversation.Columns[0].Width = 0;
            //    listViewConversation.EndUpdate();
            //}, TaskScheduler.FromCurrentSynchronizationContext());



        }

        private void ObtenerMailsDeConversacion()
        {
            //Outlook.MailItem selectedMailm = null;
            //Outlook.MailItem mailItem = null;
            //Outlook.Conversation conv = null;
            //Outlook.Table oTable = null;
            //Outlook.Row oRow = null;
            //object oItem = null;
            //Outlook.Folder inFolder = null;

            //string lastID = string.Empty;

            ////List<string[]> listConversaciones = new List<string[]>();

            //selectedMailm = this.OutlookItem as Outlook.MailItem;

            //if (Utils.originalEmailsAsTask != null) Utils.originalEmailsAsTask.Clear();

            //try
            //{
            //    if (selectedMailm != null)
            //    {
            //        if (selectedMailm.ConversationID != null)
            //        {
            //            // Obtain a Conversation object. 
            //            conv = selectedMailm.GetConversation();
            //            if (conv != null)
            //            {
            //                oTable = conv.GetTable();

            //                while (!oTable.EndOfTable)
            //                {
            //                    oRow = oTable.GetNextRow();
            //                    oItem = Globals.ThisAddIn.Application.Session.GetItemFromID(oRow["EntryID"]);
            //                    if (oItem is Outlook.MailItem)
            //                    {
            //                        mailItem = oItem as Outlook.MailItem;
            //                        inFolder = mailItem.Parent as Outlook.Folder;

            //                        //por alguna razón algunas veces algunos elementos se agregan 2 veces
            //                        if (lastID != oRow["EntryID"])
            //                        {
            //                            string[] row = {oRow["EntryID"],
            //                                            mailItem.ReceivedTime.ToShortDateString() + " " + mailItem.ReceivedTime.ToShortTimeString(),
            //                                            inFolder.Name,
            //                                            mailItem.SenderName,
            //                                            mailItem.Body.Substring(0,180) + "...",
            //                                            mailItem.IsMarkedAsTask.ToString(),
            //                                            (_cantidadAdjuntos(mailItem)>0)?_cantidadAdjuntos(mailItem).ToString():""
            //                                            };

            //                            listConversaciones.Add(row);

            //                            if (mailItem.IsMarkedAsTask)
            //                            {
            //                                Utils.originalEmailsAsTask.Add(oRow["EntryID"]);
            //                            }
            //                        }

            //                        lastID = oRow["EntryID"];

            //                        //Marshal.ReleaseComObject(mailItem);
            //                        //Marshal.ReleaseComObject(inFolder);
            //                    }
            //                    //Marshal.ReleaseComObject(oItem);
            //                    //Marshal.ReleaseComObject(oRow);
            //                }
                            
            //            }
            //        }
            //    }

            //    if (selectedMailm != null) Marshal.ReleaseComObject(selectedMailm);
            //    if (mailItem != null) Marshal.ReleaseComObject(mailItem);
            //    if (inFolder != null) Marshal.ReleaseComObject(inFolder);
            //    if (oRow != null) Marshal.ReleaseComObject(oRow);
            //    if (oTable != null) Marshal.ReleaseComObject(oTable);
            //    if (oItem != null) Marshal.ReleaseComObject(oItem);
            //    if (conv != null) Marshal.ReleaseComObject(conv);

            //}
            //catch (Exception)
            //{
            //    throw;
            //}
            //finally
            //{
            //    if (selectedMailm != null) Marshal.ReleaseComObject(selectedMailm);
            //    if (mailItem != null) Marshal.ReleaseComObject(mailItem);
            //    if (inFolder != null) Marshal.ReleaseComObject(inFolder);
            //    if (oRow != null) Marshal.ReleaseComObject(oRow);
            //    if (oTable != null) Marshal.ReleaseComObject(oTable);
            //    if (oItem != null) Marshal.ReleaseComObject(oItem);
            //    if (conv != null) Marshal.ReleaseComObject(conv);
            //}
        }


        private void BgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            listViewConversation.Items.Clear();

            if (e.Error == null)
            {
                try
                {
                    List<string[]> listData = (List<string[]>)e.Result;
                    listViewConversation.BeginUpdate();
                    for (int i = listData.Count - 1; i >= 0; i--)
                    {
                        listViewConversation.Items.Add(new ListViewItem(listData[i]));
                    }
                    listViewConversation.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
                    listViewConversation.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
                    listViewConversation.Columns[0].Width = 0;
                    listViewConversation.EndUpdate();
                }
                catch (Exception)
                {

                }
                
            }
        }


        private void BgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            Outlook.MailItem selectedMailm = null;
            Outlook.MailItem mailItem = null;
            Outlook.Conversation conv = null;
            Outlook.Table oTable = null;
            Outlook.Row oRow = null;
            object oItem = null;
            Outlook.Folder inFolder = null;

            string lastID = string.Empty;

            List<string[]> listData = new List<string[]>();

            selectedMailm = this.OutlookItem as Outlook.MailItem;

            if(Utils.originalEmailsAsTask != null) Utils.originalEmailsAsTask.Clear();

            try
            {
                if (selectedMailm != null)
                {
                    if (selectedMailm.ConversationID != null)
                    {
                        // Obtain a Conversation object. 
                        conv = selectedMailm.GetConversation();
                        if (conv != null)
                        {
                            oTable = conv.GetTable();

                            while (!oTable.EndOfTable)
                            {
                                oRow = oTable.GetNextRow();
                                oItem = Globals.ThisAddIn.Application.Session.GetItemFromID(oRow["EntryID"]);
                                if (oItem is Outlook.MailItem)
                                {
                                    mailItem = oItem as Outlook.MailItem;
                                    inFolder = mailItem.Parent as Outlook.Folder;

                                    //por alguna razón algunas veces algunos elementos se agregan 2 veces
                                    if (lastID != oRow["EntryID"])
                                    {
                                        string[] row = {oRow["EntryID"],
                                                        mailItem.ReceivedTime.ToShortDateString() + " " + mailItem.ReceivedTime.ToShortTimeString(),
                                                        inFolder.Name,
                                                        mailItem.SenderName,
                                                        mailItem.Body.Substring(0,180) + "...",
                                                        mailItem.IsMarkedAsTask.ToString(),
                                                        (_cantidadAdjuntos(mailItem)>0)?_cantidadAdjuntos(mailItem).ToString():""
                                                        };

                                        listData.Add(row);

                                        if (mailItem.IsMarkedAsTask)
                                        {
                                            Utils.originalEmailsAsTask.Add(oRow["EntryID"]);
                                        }
                                    }

                                    lastID = oRow["EntryID"];

                                    Marshal.ReleaseComObject(mailItem);
                                    Marshal.ReleaseComObject(inFolder);
                                }
                                Marshal.ReleaseComObject(oItem);
                                Marshal.ReleaseComObject(oRow);
                            }
                            e.Result = listData;
                        }
                    }
                }

                if (selectedMailm != null) Marshal.ReleaseComObject(selectedMailm);
                if (mailItem != null) Marshal.ReleaseComObject(mailItem);
                if (inFolder != null) Marshal.ReleaseComObject(inFolder);
                if (oRow != null) Marshal.ReleaseComObject(oRow);
                if (oTable != null) Marshal.ReleaseComObject(oTable);
                if (oItem != null) Marshal.ReleaseComObject(oItem);
                if (conv != null) Marshal.ReleaseComObject(conv);

            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (selectedMailm != null) Marshal.ReleaseComObject(selectedMailm);
                if (mailItem != null) Marshal.ReleaseComObject(mailItem);
                if (inFolder != null) Marshal.ReleaseComObject(inFolder);
                if (oRow != null) Marshal.ReleaseComObject(oRow);
                if (oTable != null) Marshal.ReleaseComObject(oTable);
                if (oItem != null) Marshal.ReleaseComObject(oItem);
                if (conv != null) Marshal.ReleaseComObject(conv);
            }
        }


        /// <summary>
        /// Calcula la cantidad de adjuntos que tiene un item de Outlook sin considerar las imágenes
        /// que puedieran estar como cuerpo del item
        /// </summary>
        /// <param name="oItem">Es el Outlook item que puede ser un mailItem, appoitment, meeting.</param>
        /// <returns>Cantidad de adjuntos</returns>
        private int _cantidadAdjuntos(object oItem)
        {
            Outlook.MailItem mailItem = null;
            int cant = 0;

            try
            {
                if (oItem != null)
                {
                    if (oItem is Outlook.MailItem)
                    {
                        mailItem = oItem as Outlook.MailItem;

                        if (mailItem.Attachments.Count > 0)
                        {
                            // get attachments
                            foreach (Outlook.Attachment attachment in mailItem.Attachments)
                            {
                                var flags = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37140003");

                                //To ignore embedded attachments -
                                if (flags != 4)
                                {
                                    // As per present understanding - If rtF mail attachment comes here - and the embeded image is treated as attachment then Type value is 6 and ignore it
                                    if ((int)attachment.Type != 6)
                                    {
                                        cant++;
                                        //MailAttachment mailAttachment = new MailAttachment { Name = attachment.FileName };
                                        //mail.Attachments.Add(mailAttachment);
                                    }

                                }

                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }


            return cant;

        }


        private void listViewConversation_Resize(object sender, EventArgs e)
        {
            //this.Height = 115;
            //this.OutlookFormRegion.Reflow();

            listViewConversation.Width = this.Width - 10;
            listViewConversation.Left = 5;
            listViewConversation.Height = this.Height - 4;
            listViewConversation.Top = 4;

        }

        private void listViewConversation_DoubleClick(object sender, EventArgs e)
        {
            Outlook.MailItem selectedMailm = null;

            // user clicked an item of listview control
            if (listViewConversation.SelectedItems.Count == 1)
            {
                try
                {
                    ListViewItem lvItem = listViewConversation.SelectedItems[0];
                    selectedMailm = Globals.ThisAddIn.Application.Session.GetItemFromID(lvItem.Text);
                    selectedMailm.Display(false);
                    if (lvItem != null) Marshal.ReleaseComObject(lvItem);
                }
                catch (Exception)
                {

                }
                
            }
            
            if (selectedMailm != null) Marshal.ReleaseComObject(selectedMailm);
        }
    }
}
