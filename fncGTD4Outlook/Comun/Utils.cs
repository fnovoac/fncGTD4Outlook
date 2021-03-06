﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace fncGTD4Outlook.Comun
{
    public static class Utils
    {
        public static List<string> originalEmailsAsTask;


        // Returns Folder object based on folder path
        public static Outlook.MAPIFolder GetFolderByName(string folderName)
        {
            // revisamos si los folders existen, sino los creamos
            Outlook.Folder root = null;
            Outlook.Folders childFolders = null;
            Outlook.MAPIFolder objfolder = null;

            try
            {
                // revisamos si los folders existen, sino los creamos
                root = Globals.ThisAddIn.Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
                childFolders = root.Folders;

                if (childFolders.Count > 0)
                {
                    foreach (Outlook.Folder childFolder in childFolders)
                    {
                        if (childFolder.Name == folderName)
                        {
                            objfolder = childFolder;
                            //if (childFolder != null) Marshal.ReleaseComObject(childFolder);
                            break;
                        }

                        //if (childFolder != null) Marshal.ReleaseComObject(childFolder);
                    }
                }
                if (root != null) Marshal.ReleaseComObject(root);
                if (childFolders != null) Marshal.ReleaseComObject(childFolders);
                return objfolder;

            }
            catch
            {
                if (root != null) Marshal.ReleaseComObject(root);
                if (childFolders != null) Marshal.ReleaseComObject(childFolders);
                if (objfolder != null) Marshal.ReleaseComObject(objfolder);
                return null;
            }
        }


        /// <summary>
        /// Obtiene un listado de todos los emails, appointments y meetings seleccionados
        /// </summary>
        /// <returns>Lista de objetos</returns>
        public static List<object> GetOutlookItems()
        {
            List<object> selObjects = new List<object>();

            object activeWindow = null;
            Outlook.Explorer explorer = null;
            Outlook.Inspector inspector = null;
            Object selObject = null;

            try
            {
                // get active Window
                activeWindow = Globals.ThisAddIn.Application.ActiveWindow();
                if (activeWindow is Outlook.Explorer)
                {
                    explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                    if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
                    {
                        for (int i = 1; i <= Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count; i++)
                        {
                            selObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[i];
                            if (selObject is Outlook.MailItem)
                            {
                                Outlook.MailItem item = (selObject as Outlook.MailItem);
                                selObjects.Add(item);
                            }

                            if (selObject is Outlook.AppointmentItem)
                            {
                                Outlook.AppointmentItem item = (selObject as Outlook.AppointmentItem);
                                selObjects.Add(item);
                            }

                            if (selObject is Outlook.MeetingItem)
                            {
                                Outlook.MeetingItem item = (selObject as Outlook.MeetingItem);
                                selObjects.Add(item);
                            }
                        }
                    }
                    if (explorer != null) Marshal.ReleaseComObject(explorer);
                }
                if (activeWindow is Outlook.Inspector)
                {
                    inspector = Globals.ThisAddIn.Application.ActiveInspector();
                    selObject = inspector.CurrentItem;
                    if (selObject is Outlook.MailItem)
                    {
                        Outlook.MailItem item = (selObject as Outlook.MailItem);
                        selObjects.Add(item);
                    }

                    if (selObject is Outlook.AppointmentItem)
                    {
                        Outlook.AppointmentItem item = (selObject as Outlook.AppointmentItem);
                        selObjects.Add(item);
                    }

                    if (selObject is Outlook.MeetingItem)
                    {
                        Outlook.MeetingItem item = (selObject as Outlook.MeetingItem);
                        selObjects.Add(item);
                    }
                    if (inspector != null) Marshal.ReleaseComObject(inspector);
                }
                if (activeWindow != null) Marshal.ReleaseComObject(activeWindow);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (activeWindow != null) Marshal.ReleaseComObject(activeWindow);
                if (explorer != null) Marshal.ReleaseComObject(explorer);
                if (inspector != null) Marshal.ReleaseComObject(inspector);
                if (selObject != null) Marshal.ReleaseComObject(selObject);
            }

            return selObjects;
        }

        /// <summary>
        /// Obtiene el email, appointment o meeting seleccionado
        /// </summary>
        /// <returns>objeto</returns>
        public static object GetOutlookItem()
        {
            object oItem = null;
            try
            {
                oItem = GetOutlookItems()[0];
            }
            catch (Exception)
            {
                throw;
            }

            return oItem;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static List<Outlook.AppointmentItem> GetAppointmentItems()
        {
            List<Outlook.AppointmentItem> appItems = new List<Outlook.AppointmentItem>();

            try
            {

                Outlook.Inspector actInspector = Globals.ThisAddIn.Application.ActiveInspector();
                if (actInspector == null)
                {
                    Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                    if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
                    {
                        for (int i = 1; i <= Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count; i++)
                        {
                            Object selObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[i];
                            if (selObject is Outlook.AppointmentItem)
                            {
                                Outlook.AppointmentItem mailItem = (selObject as Outlook.AppointmentItem);
                                appItems.Add(mailItem);
                            }
                        }
                    }
                    if (explorer != null) Marshal.ReleaseComObject(explorer);
                }
                else
                {
                    Object selObject = actInspector.CurrentItem;
                    if (selObject is Outlook.AppointmentItem)
                    {
                        Outlook.AppointmentItem mailItem = (selObject as Outlook.AppointmentItem);
                        appItems.Add(mailItem);
                    }
                    if (actInspector != null) Marshal.ReleaseComObject(actInspector);
                }
            }
            catch (Exception)
            {
                throw;
            }

            return appItems;

        }


        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static Outlook.AppointmentItem GetAppointmentItem()
        {
            Outlook.AppointmentItem appItem = null;
            try
            {
                appItem = GetAppointmentItems()[0];
            }
            catch (Exception)
            {
                throw;
            }

            return appItem;
        }


        /// <summary>
        /// Función que devuelve un arreglo con los mails seleccionados
        /// </summary>
        /// <returns>Mails seleccionados</returns>
        public static List<Outlook.MailItem> GetMailItems()
        {
            List<Outlook.MailItem> mails = new List<Outlook.MailItem>();

            Outlook.Inspector actInspector = Globals.ThisAddIn.Application.ActiveInspector();
            if (actInspector == null)
            {
                Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                try
                {
                    if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
                    {
                        for (int i = 1; i <= Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count; i++)
                        {
                            Object selObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[i];
                            if (selObject is Outlook.MailItem)
                            {
                                /*
                                 * Outlook.ContactItem
                                 * Outlook.AppointmentItem
                                 * Outlook.TaskItem
                                 * Outlook.MeetingItem
                                 */
                                Outlook.MailItem mailItem = (selObject as Outlook.MailItem);
                                mails.Add(mailItem);
                                //if (mailItem != null) Marshal.ReleaseComObject(mailItem);
                            }
                            //if (selObject != null) Marshal.ReleaseComObject(selObject);
                        }
                    }
                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    if (explorer != null) Marshal.ReleaseComObject(explorer);
                }
            }
            else
            {
                try
                {
                    Object selObject = actInspector.CurrentItem;
                    if (selObject is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem = (selObject as Outlook.MailItem);
                        mails.Add(mailItem);
                        //if (mailItem != null) Marshal.ReleaseComObject(mailItem);
                    }
                    //if (selObject != null) Marshal.ReleaseComObject(selObject);
                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    if (actInspector != null) Marshal.ReleaseComObject(actInspector);
                }
            }

            return mails;

        }

        public static Outlook.MailItem GetMailItem()
        {
            Outlook.MailItem email = null;
            try
            {
                email = GetMailItems()[0];
            }
            catch (Exception)
            {
                throw;
            }

            return email;
        }

        public static List<Outlook.ContactItem> GetListOfContacts (bool incluirSugeridos = false)
        {
            List<Outlook.ContactItem> contactItemsList = null;
            Outlook.Items folderItems = null;
            Outlook.MAPIFolder folderSuggestedContacts = null;
            Outlook.NameSpace ns = null;
            Outlook.MAPIFolder folderContacts = null;
            object itemObj = null;

            try
            {
                contactItemsList = new List<Outlook.ContactItem>();
                ns = Globals.ThisAddIn.Application.GetNamespace("MAPI");
                // getting items from the Contacts folder in Outlook
                folderContacts = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                folderItems = folderContacts.Items;
                for (int i = 1; folderItems.Count >= i; i++)
                {
                    itemObj = folderItems[i];
                    if (itemObj is Outlook.ContactItem)
                        contactItemsList.Add(itemObj as Outlook.ContactItem);
                    else
                        Marshal.ReleaseComObject(itemObj);
                }
                Marshal.ReleaseComObject(folderItems);
                folderItems = null;
                // getting items from the Suggested Contacts folder in Outlook
                if (incluirSugeridos)
                {
                    folderSuggestedContacts = ns.GetDefaultFolder(
                                          Outlook.OlDefaultFolders.olFolderSuggestedContacts);
                    folderItems = folderSuggestedContacts.Items;
                    for (int i = 1; folderItems.Count >= i; i++)
                    {
                        itemObj = folderItems[i];
                        if (itemObj is Outlook.ContactItem)
                            contactItemsList.Add(itemObj as Outlook.ContactItem);
                        else
                            Marshal.ReleaseComObject(itemObj);
                    }
                }
                
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (folderItems != null) Marshal.ReleaseComObject(folderItems);
                if (folderContacts != null) Marshal.ReleaseComObject(folderContacts);
                if (folderSuggestedContacts != null) Marshal.ReleaseComObject(folderSuggestedContacts);
                if (ns != null) Marshal.ReleaseComObject(ns);
            }
            return contactItemsList;

        }


        public static void ArchivarOutlookItems()
        {
            Outlook.MAPIFolder objfolder = null;
            object activeWindow = null;
            Outlook.Explorer explorer = null;
            Outlook.Inspector inspector = null;
            //Object selObject = null;

            try
            {
                //obtenemos el folder donde moveremos el email (debe existir -> ver ThisAddIn.cs)
                objfolder = Utils.GetFolderByName(Constants.folderArchivar);

                if (objfolder == null)
                    objfolder = Globals.ThisAddIn.Application.Session.DefaultStore.GetRootFolder().Folders.Add(Constants.folderArchivar);

                // get active Window
                activeWindow = Globals.ThisAddIn.Application.ActiveWindow();
                if (activeWindow is Outlook.Explorer)
                {
                    explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                    if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
                    {
                        foreach (object selObject in Globals.ThisAddIn.Application.ActiveExplorer().Selection)
                        {
                            if (selObject is Outlook.MailItem)
                            {
                                Outlook.MailItem item = (selObject as Outlook.MailItem);
                                item.UnRead = false;
                                item.Move(objfolder);
                            }

                            if (selObject is Outlook.AppointmentItem)
                            {
                                Outlook.AppointmentItem item = (selObject as Outlook.AppointmentItem);
                                item.UnRead = false;
                                item.Move(objfolder);
                            }

                            if (selObject is Outlook.MeetingItem)
                            {
                                Outlook.MeetingItem item = (selObject as Outlook.MeetingItem);
                                item.UnRead = false;
                                item.Move(objfolder);
                            }
                        }
                    }
                    if (explorer != null) Marshal.ReleaseComObject(explorer);
                }
                if (activeWindow is Outlook.Inspector)
                {
                    inspector = Globals.ThisAddIn.Application.ActiveInspector();
                    object selObject = inspector.CurrentItem;
                    if (selObject is Outlook.MailItem)
                    {
                        Outlook.MailItem item = (selObject as Outlook.MailItem);
                        item.UnRead = false;
                        item.Move(objfolder);
                    }

                    if (selObject is Outlook.AppointmentItem)
                    {
                        Outlook.AppointmentItem item = (selObject as Outlook.AppointmentItem);
                        item.UnRead = false;
                        item.Move(objfolder);
                    }

                    if (selObject is Outlook.MeetingItem)
                    {
                        Outlook.MeetingItem item = (selObject as Outlook.MeetingItem);
                        item.UnRead = false;
                        item.Move(objfolder);
                    }
                    if (inspector != null) Marshal.ReleaseComObject(inspector);
                    if (selObject != null) Marshal.ReleaseComObject(selObject);
                }
                if (activeWindow != null) Marshal.ReleaseComObject(activeWindow);
                
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (objfolder != null) Marshal.ReleaseComObject(objfolder);
                if (activeWindow != null) Marshal.ReleaseComObject(activeWindow);
                if (explorer != null) Marshal.ReleaseComObject(explorer);
                if (inspector != null) Marshal.ReleaseComObject(inspector);
            }
        }

        /// <summary>
        /// Mueve el email seleccinado a una carpeta destino
        /// La carpeta debe estar el mismo nivel que Inbox (root)
        /// OJO: Move un solo email
        /// </summary>
        /// <param name="carpetaDestino">Nombre de la carpeta destino</param>
        public static void MoverEmail(Outlook.MailItem email, string carpetaDestino)
        {
            //obtenemos el folder donde moveremos el email (debe existir -> ver ThisAddIn.cs)
            Outlook.MAPIFolder objfolder = Utils.GetFolderByName(carpetaDestino);
            try
            {
                if (objfolder == null)
                    objfolder = Globals.ThisAddIn.Application.Session.DefaultStore.GetRootFolder().Folders.Add(Constants.folderConservar);

                if (email != null)
                {
                    email.UnRead = false;
                    email.Move(objfolder);
                    if (email != null) Marshal.ReleaseComObject(email);
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (objfolder != null) Marshal.ReleaseComObject(objfolder);
            }

        }

        /// <summary>
        /// Mueve el email o emails seleccinados a una carpeta destino
        /// La carpeta debe estar el mismo nivel que Inbox (root)
        /// </summary>
        /// <param name="carpetaDestino">Nombre de la carpeta destino</param>
        public static void MoverEmail(List<Outlook.MailItem> emails, string carpetaDestino)
        {
            //obtenemos el folder donde moveremos el email (debe existir -> ver ThisAddIn.cs)
            Outlook.MAPIFolder objfolder = Utils.GetFolderByName(carpetaDestino);

            try
            {
                if (objfolder == null)
                    objfolder = Globals.ThisAddIn.Application.Session.DefaultStore.GetRootFolder().Folders.Add(Constants.folderConservar);

                for (int i = 0; i < emails.Count; i++)
                {
                    if (emails[i] != null)
                    {
                        emails[i].UnRead = false;
                        emails[i].Move(objfolder);
                    }
                }

            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (objfolder != null) Marshal.ReleaseComObject(objfolder);
                for (int i = 0; i < emails.Count; i++)
                {
                    if (emails[i] != null) Marshal.ReleaseComObject(emails[i]);
                }
            }
        }
        
        public static string GetFirstReceiverFromTo(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return string.Empty;
            }
            else
            {
                var i = input.IndexOf(Constants.emailDelimiter);
                return i == -1 ? input : input.Substring(0, i);
            }
            
        }

        public static string SubstringAfter(string source, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return source;
            }
            CompareInfo compareInfo = CultureInfo.InvariantCulture.CompareInfo;
            int index = compareInfo.IndexOf(source, value, CompareOptions.Ordinal);
            if (index < 0)
            {
                //No such substring
                return source;
                //return string.Empty;
            }
            return source.Substring(index + value.Length);
        }

        public static string SubstringBefore(string source, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return value;
            }
            CompareInfo compareInfo = CultureInfo.InvariantCulture.CompareInfo;
            int index = compareInfo.IndexOf(source, value, CompareOptions.Ordinal);
            if (index < 0)
            {
                //No such substring
                return source;
                //return string.Empty;
            }
            return source.Substring(0, index);
        }

        public static void MostrarDelegarForm()
        {
            // get active Window
            object activeWindow = Globals.ThisAddIn.Application.ActiveWindow();
            Outlook.Inspector inspector = null;
            Outlook.MailItem currentMail = null;
            try
            {
                if (activeWindow is Outlook.Explorer)
                {
                    //NOTA: a pesar de ejecutarse, queda el ENTER como remanente y abre el email si cancelamos el form
                    //mejor lo desactivo hasta encontrar otra forma de cancelar el ENTER luego de capturar la combinacion de teclas

                    //frmDelegar f = new frmDelegar();
                    //f.ShowDialog();
                    //f.Dispose();
                    //f = null;
                }
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
}
