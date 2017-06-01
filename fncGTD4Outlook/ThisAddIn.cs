using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;
using fncGTD4Outlook.Comun;
using System.Windows.Forms;
using Gma.System.MouseKeyHook;

namespace fncGTD4Outlook
{
    public partial class ThisAddIn
    {
        private IKeyboardMouseEvents m_Events;

        private void SubscribeGlobal()
        {
            Unsubscribe();
            Subscribe(Hook.GlobalEvents());
        }

        private void SubscribeApplication()
        {
            Unsubscribe();
            Subscribe(Hook.AppEvents());
        }

        private void Unsubscribe()
        {
            if (m_Events == null) return;
            m_Events.KeyDown -= OnKeyDown;
            m_Events.KeyUp -= OnKeyUp;
            m_Events.KeyPress -= HookManager_KeyPress;

            m_Events.Dispose();
            m_Events = null;
        }

        private void Subscribe(IKeyboardMouseEvents events)
        {
            m_Events = events;
            m_Events.KeyDown += OnKeyDown;
            m_Events.KeyUp += OnKeyUp;
            m_Events.KeyPress += HookManager_KeyPress;
        }

        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            //MessageBox.Show(string.Format("KeyDown  \t\t {0}\n", e.KeyCode));

            if (e.Control && e.KeyCode == Keys.Enter)
            {
                // get active Window
                object activeWindow = Globals.ThisAddIn.Application.ActiveWindow();
                Outlook.Inspector inspector = null;
                Outlook.MailItem currentMail = null;
                try
                {
                    //if (activeWindow is Outlook.Explorer)
                    //{

                    //}
                    if (activeWindow is Outlook.Inspector)
                    {
                        // its an inspector window
                        inspector = Globals.ThisAddIn.Application.ActiveInspector();
                        currentMail = inspector.CurrentItem as Outlook.MailItem;
                        if (currentMail != null)
                        {
                            if (currentMail.Sent)
                            {
                                frmDelegar f = new frmDelegar();
                                f.ShowDialog();
                                f.Dispose();
                                f = null;
                            }
                            else
                            {
                                frmDelegarEnviar f = new frmDelegarEnviar();
                                f.Show();
                                f = null;
                            }
                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    if (activeWindow != null) Marshal.ReleaseComObject(activeWindow);
                    if (inspector != null) Marshal.ReleaseComObject(inspector);
                    if (currentMail != null) Marshal.ReleaseComObject(currentMail);
                }
            }
        }

        private void OnKeyUp(object sender, KeyEventArgs e)
        {
            //MessageBox.Show(string.Format("KeyUp  \t\t {0}\n", e.KeyCode));
        }

        private void HookManager_KeyPress(object sender, KeyPressEventArgs e)
        {
            //MessageBox.Show(string.Format("KeyPress \t\t {0}\n", e.KeyChar));
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            SubscribeApplication();

            string folderArchivar = Constants.folderArchivar;
            string folderDelegar = Constants.folderDelegar;
            string folderDiferir = Constants.folderDiferir;
            string folderConservar = Constants.folderConservar;
            string folderReferencia = Constants.folderReferencia;
            string folderRecurrente = Constants.folderRecurrente;

            bool folderArchivarExiste = false;
            bool folderDelegarExiste = false;
            bool folderDiferirExiste = false;
            bool folderConservarExiste = false;
            bool folderReferenciaExiste = false;
            bool folderRecurrenteExiste = false;

            // revisamos si los folders existen, sino los creamos
            Outlook.Folder root = null;
            Outlook.Folders childFolders = null;
            try
            {
                root = Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
                childFolders = root.Folders;
                if (childFolders.Count > 0)
                {
                    foreach (Outlook.Folder childFolder in childFolders)
                    {
                        if (childFolder.Name == folderArchivar)
                        {
                            folderArchivarExiste = true;
                        }
                        else if (childFolder.Name == folderDelegar)
                        {
                            folderDelegarExiste = true;
                        }
                        else if (childFolder.Name == folderDiferir)
                        {
                            folderDiferirExiste = true;
                        }
                        else if (childFolder.Name == folderConservar)
                        {
                            folderConservarExiste = true;
                        }
                        else if (childFolder.Name == folderReferencia)
                        {
                            folderReferenciaExiste = true;
                        }
                        else if (childFolder.Name == folderRecurrente)
                        {
                            folderRecurrenteExiste = true;
                        }

                        if (childFolder != null) Marshal.ReleaseComObject(childFolder);
                    }
                }

                if (!folderArchivarExiste)
                {
                    Outlook.MAPIFolder objfolder = root.Folders.Add(folderArchivar);
                    if (objfolder != null) Marshal.ReleaseComObject(objfolder);
                }
                if (!folderDelegarExiste)
                {
                    Outlook.MAPIFolder objfolder = root.Folders.Add(folderDelegar);
                    if (objfolder != null) Marshal.ReleaseComObject(objfolder);
                }
                if (!folderDiferirExiste)
                {
                    Outlook.MAPIFolder objfolder = root.Folders.Add(folderDiferir);
                    if (objfolder != null) Marshal.ReleaseComObject(objfolder);
                }
                if (!folderConservarExiste)
                {
                    Outlook.MAPIFolder objfolder = root.Folders.Add(folderConservar);
                    if (objfolder != null) Marshal.ReleaseComObject(objfolder);
                }
                if (!folderReferenciaExiste)
                {
                    Outlook.MAPIFolder objfolder = root.Folders.Add(folderReferencia);
                    if (objfolder != null) Marshal.ReleaseComObject(objfolder);
                }
                if (!folderRecurrenteExiste)
                {
                    Outlook.MAPIFolder objfolder = root.Folders.Add(folderRecurrente);
                    if (objfolder != null) Marshal.ReleaseComObject(objfolder);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            finally
            {
                if (root != null) Marshal.ReleaseComObject(root);
                if (childFolders != null) Marshal.ReleaseComObject(childFolders);
            }

            this.Application.NewMailEx += Application_NewMailEx;


        }

        private void Application_NewMailEx(string EntryIDCollection)
        {
            object obj = null;
            Outlook.MailItem email = null;
            Outlook.UserProperties itemProps = null;
            Outlook.UserProperty newProp = null;

            try
            {
                string[] ids = EntryIDCollection.Split(',');
                for (int i = 0; i < ids.Length; i++)
                {
                    obj = null;
                    try
                    {
                        obj = Application.Session.GetItemFromID(ids[i], Type.Missing);
                        if (obj is Outlook.MailItem)
                        {
                            email = obj as Outlook.MailItem;
                            itemProps = email.UserProperties;

                            newProp = itemProps.Find("MarkAsTask");
                            if (newProp != null) email.MarkAsTask(newProp.Value);
                            else return;

                            newProp = itemProps.Find("TaskDueDate");
                            if (newProp != null) email.TaskDueDate = newProp.Value;

                            newProp = itemProps.Find("ReminderSet");
                            if (newProp != null) email.ReminderSet = newProp.Value;

                            newProp = itemProps.Find("ReminderTime");
                            if (newProp != null) email.ReminderTime = newProp.Value;

                            //email.BillingInformation = newProp.Value;

                            email.Save();

                            Utils.MoverEmail(email, Constants.folderDelegar);

                        }
                    }
                    catch (Exception)
                    {

                    }
                    finally
                    {
                        if (obj != null) Marshal.ReleaseComObject(obj);
                    }
                }
            }
            catch (Exception)
            {

            }
            finally
            {
                if (email != null) Marshal.ReleaseComObject(email);
                if (newProp != null) Marshal.ReleaseComObject(newProp);
                if (itemProps != null) Marshal.ReleaseComObject(itemProps);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Unsubscribe();

            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
