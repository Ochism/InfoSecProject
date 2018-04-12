using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace SpamClassifier
{
    public partial class ThisAddIn
    {
        // Ensure GC does not free needed memory
        Outlook.Inspectors inspectors;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            this.Application.NewMail += new Outlook.ApplicationEvents_11_NewMailEventHandler(NewMailMethod);
            //inspectors = this.Application.Inspectors;
            //inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            // must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }
        void NewMailMethod()
        {
            Outlook.MAPIFolder inBox = Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.MAPIFolder junkBox = Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk);
            Outlook.Items items = (Outlook.Items)inBox.Items;
            Outlook.MailItem moveMail = null;
            items.Restrict("[UnRead] = true");
            Outlook.MAPIFolder destFolder = junkBox;
            foreach (object eMail in items)
            {
                try
                {
                    moveMail = eMail as Outlook.MailItem;
                    if (moveMail != null)
                    {
                        string titleSubject = (string)moveMail.Subject;
                        moveMail.Move(destFolder);
                        moveMail.Subject = "This  email is evil" + moveMail.Subject;
                        moveMail.Body = "This email is evil" + moveMail.Body;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID != null)
                {
                    //api call

                    //if(its evil)
                    mailItem.Move(this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk));
                    mailItem.Subject = "This  email is evil" + mailItem.Subject;
                    mailItem.Body = "This email is evil" + mailItem.Body;
                }
            }
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
