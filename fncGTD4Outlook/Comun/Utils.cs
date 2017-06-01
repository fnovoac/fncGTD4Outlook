using System;
using System.Collections.Generic;
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
        /// 
        /// </summary>
        /// <returns></returns>
        public static List<Outlook.AppointmentItem> getAppointmentItems()
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
            catch (Exception ex)
            {
                throw;
            }

            return appItems;

        }


        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static Outlook.AppointmentItem getAppointmentItem()
        {
            Outlook.AppointmentItem appItem = null;
            try
            {
                appItem = getAppointmentItems()[0];
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
        public static List<Outlook.MailItem> getMailItems()
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

        public static Outlook.MailItem getMailItem()
        {
            Outlook.MailItem email = null;
            try
            {
                email = getMailItems()[0];
            }
            catch (Exception)
            {
                throw;
            }

            return email;
        }

        public static List<Outlook.ContactItem> getListOfContacts (bool incluirSugeridos = false)
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
    }
}
