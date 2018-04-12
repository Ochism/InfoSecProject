using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using IBM.WatsonDeveloperCloud.NaturalLanguageClassifier.v1;
using IBM.WatsonDeveloperCloud.NaturalLanguageClassifier.v1.Model;
using System.Net;
using System.Diagnostics;

namespace SpamClassifier
{
    public partial class ThisAddIn
    {
        // Global constants and variables
        private string username = "a477516a-4cdf-4080-93bc-064265ec1643";
        private string password = "4JnCcEcxFDjM";
        private string subjectClassifierID = "2fc15ax329-nlc-819";
        private string bodyClassifierID = "ab2c7bx342-nlc-368";
        NaturalLanguageClassifierService _naturalLanguageClassifierService;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // For Windows 7 and later
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            // Set the credentials
            _naturalLanguageClassifierService = new NaturalLanguageClassifierService();
            _naturalLanguageClassifierService.SetCredential(username, password);

            this.Application.NewMail += new Outlook.ApplicationEvents_11_NewMailEventHandler(NewMailMethod);

            //Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)this.Application.
            //    ActiveExplorer().Session.GetDefaultFolder
            //    (Outlook.OlDefaultFolders.olFolderInbox);
            
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        private void ClassifyIncomingMail(object eMail, Outlook.MAPIFolder watsonSpamFolder)
        {
            Outlook.MailItem moveMail = null;
            try
            {
                moveMail = eMail as Outlook.MailItem;
                if (moveMail != null)
                {
                    // TODO: Incorporate both subject and body for our text field
                    ClassifyInput classifyInput = new ClassifyInput
                    {
                        Text = moveMail.Subject
                    };

                    Classification classifyResult = _naturalLanguageClassifierService.Classify(subjectClassifierID, classifyInput);
                    Debug.WriteLine("TOP CLASS: " + classifyResult.TopClass);
                    if (classifyResult.TopClass == "spam")
                    {
                        // TODO: Include messages in spam subject/body
                        moveMail.Move(watsonSpamFolder);
                        moveMail.Body = classifyResult.TopClass + moveMail.Body;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        void NewMailMethod()
        {
            // Declare our inbox and junk folder
            Outlook.MAPIFolder inBox = (Outlook.MAPIFolder) this.Application.
                ActiveExplorer().Session.GetDefaultFolder
                (Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.MAPIFolder root = inBox.Parent;
            Outlook.MAPIFolder watsonSpamFolder = root.Folders["WatsonSpam"];
            Outlook.MAPIFolder junkFolder = Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk);

            Outlook.Items inboxItems = inBox.Items;
            Outlook.Items junkItems = junkFolder.Items;
            inboxItems.Restrict("[UnRead] = true");
            junkItems.Restrict("[UnRead] = true");

            foreach (object eMail in inboxItems)
            {
                ClassifyIncomingMail(eMail, watsonSpamFolder);
            }

            foreach (object eMail in junkItems)
            {
                ClassifyIncomingMail(eMail, watsonSpamFolder);
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
