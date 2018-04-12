using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using IBM.WatsonDeveloperCloud.NaturalLanguageClassifier.v1;
using IBM.WatsonDeveloperCloud.NaturalLanguageClassifier.v1.Model;
using System.Net;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;

namespace SpamClassifier
{
    public partial class ThisAddIn
    {
        // Global constants and variables
        private string subUsername = "a477516a-4cdf-4080-93bc-064265ec1643";
        private string subPassword = "4JnCcEcxFDjM";
        private string subClassifierID = "2fc15ax329-nlc-819";
        private string bodyUsername = "cd32418e-01b1-478e-9c24-a46a0767a0c7";
        private string bodyPassword = "AXISL3obSiSo";
        private string bodyClassifierID = "ab2c7bx342-nlc-368";
        NaturalLanguageClassifierService _subClassifier;
        NaturalLanguageClassifierService _bodyClassifier;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // For Windows 7 and later
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            // Create and set credentials for both classifiers
            _subClassifier = new NaturalLanguageClassifierService();
            _subClassifier.SetCredential(subUsername, subPassword);
            _bodyClassifier = new NaturalLanguageClassifierService();
            _bodyClassifier.SetCredential(bodyUsername, bodyPassword);

            EnsureFolderExists("WatsonSpam");

            this.Application.NewMail += new Outlook.ApplicationEvents_11_NewMailEventHandler(NewMailMethod);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }
        /*
         * Creates the specified folder if it doesn't exist
         * 
         *@param foldername
         *      the name of the folder to ensure exists
         */
        private void EnsureFolderExists(string foldername)
        {
            Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)this.Application.
                ActiveExplorer().Session.GetDefaultFolder
                (Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.MAPIFolder root = inBox.Parent;

            bool spamFolderExists = false;
            foreach (Outlook.MAPIFolder f in root.Folders)
            {
                spamFolderExists |= f.Name == "WatsonSpam";
            }
            if (!spamFolderExists)
            {
                root.Folders.Add("WatsonSpam");
            }
        }
        
        /**
         * Moves incoming mail depending on whether or not it is spam.
         * @param eMail
         *      email object to be analyzed
         * @param watsonSpamFolder
         *      folder to move spam emails to
         **/
        private void MoveIncomingMail(object eMail, Outlook.MAPIFolder watsonSpamFolder)
        {
            Outlook.MailItem moveMail = null;
            try
            {
                moveMail = eMail as Outlook.MailItem;
                if (moveMail != null)
                {
                    // Classify email
                    string classification = ClassifyMail(moveMail);
                    
                    // Move email if classified as spam
                    if (classification == "spam")
                    {
                        // TODO: Include messages in spam subject/body
                        moveMail.Move(watsonSpamFolder);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /**
         * Classifies email based on weighted sum of subject and body
         * classification confidences.
         * 
         * @param moveMail
         *      candidate email to classify
         * 
         * @returns classification of given email
         **/
        private string ClassifyMail(Outlook.MailItem moveMail)
        {
            string classification;
            double subConfWeight = .35;
            double bodyConfWeight = .65;
            double confLimit = .88;

            // Classify subject
            ClassifyInput classifySubjectInput = new ClassifyInput
            {
                Text = moveMail.Subject
            };

            // Get top class and weighted confidence of subject
            Classification classifySubjectResult = _subClassifier.Classify(subClassifierID, classifySubjectInput);
            string subClass = classifySubjectResult.TopClass;
            double subConf = (double) classifySubjectResult.Classes[0].Confidence * subConfWeight;

            Dictionary<string, List<double>> bodyDict = new Dictionary<string, List<double>>();
            List<double> spamList = new List<double>();
            List<double> notSpamList = new List<double>();
            bodyDict.Add("spam", spamList);
            bodyDict.Add("not spam", notSpamList);

            // Break subject into manageable chunks to classify
            string cleanedBody = moveMail.Body.Replace("\n", "").Replace("\t", "").Replace("\r", "");
            IList<string> bodyChunks = ChunkBody(cleanedBody, 1000);
            foreach(string chunk in bodyChunks)
            {
                string cleanedChunk = chunk;
                // Classify chunk of body text
                ClassifyInput classifyChunkInput = new ClassifyInput
                {
                    Text = chunk
                };

                // Get top class of body chunk and add it and its confidence to bodyDict
                Classification classifyChunkResult = _bodyClassifier.Classify(bodyClassifierID, classifyChunkInput);
                string topChunkClass = classifyChunkResult.TopClass;
                double chunkConf = (double)classifyChunkResult.Classes[0].Confidence;
                bodyDict[topChunkClass].Add(chunkConf);
            }
            // Determine top classification of body and take average weighted confidence of chunks
            string bodyClass = bodyDict["spam"].Count > bodyDict["not spam"].Count ? "spam" : "not spam";
            List<double> bodyConfList = bodyDict[bodyClass];
            double bodyConf = bodyConfList.Average() * bodyConfWeight;

            // Combine classes and weighted confidences to determine final classification
            if(subClass == bodyClass)
            {
                double totalConf = subConf + bodyConf;
                classification = totalConf >= confLimit ? subClass : classifySubjectResult.Classes[1].ClassName;
            }
            else
            {
                classification = subConf > bodyConf ? subClass : bodyClass;
            }

            return classification;
        }

        /**
         * Breaks text into list of evenly-sized chunks based on string length.
         * 
         * @param text
         *      text to be broken up into chunks
         * @param chunkSize
         *      size of chunks to be returned
         * 
         * @returns list of chunks with max size = chunkSize
         **/
        private IList<string> ChunkBody(string text, int chunkSize)
        {
            List<string> chunks = new List<string>();
            int offset = 0;
            while (offset < text.Length)
            {
                int size = Math.Min(chunkSize, text.Length - offset);
                chunks.Add(text.Substring(offset, size));
                offset += size;
            }
            return chunks;
        }

        /**
         * Fires whenever account receives a new email.
         **/
        void NewMailMethod()
        {
            // Declare folders to use in mail management
            Outlook.MAPIFolder inBox = (Outlook.MAPIFolder) this.Application.
                ActiveExplorer().Session.GetDefaultFolder
                (Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.MAPIFolder root = inBox.Parent;
            Outlook.MAPIFolder watsonSpamFolder = root.Folders["WatsonSpam"];
            Outlook.MAPIFolder junkFolder = Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk);

            Outlook.Items junkItems = junkFolder.Items;
            junkItems.Restrict("[UnRead] = true");

            // Move mail already classified as junk back to inbox
            foreach (object eMail in junkItems)
            {
                Outlook.MailItem moveMail = null;
                try
                {
                    moveMail = eMail as Outlook.MailItem;
                    if (moveMail != null)
                    {
                        moveMail.Move(inBox);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            // Classify mail in inbox with Watson
            Outlook.Items inboxItems = inBox.Items;
            inboxItems.Restrict("[UnRead] = true");
            foreach (object eMail in inboxItems)
            {
                MoveIncomingMail(eMail, watsonSpamFolder);
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
