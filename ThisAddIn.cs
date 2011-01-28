using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;


namespace WarnOnCrossSendOutlookAddIn
{
    public partial class ThisAddIn
    {

        private string[] accountDomains;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            identifyMultipleEmailDomains();
            this.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        }

        void Application_ItemSend(object Item, ref bool Cancel)
        {
            if (isCrossSendEmail(Item))
            {
                if (MessageBox.Show("Your attemping to mail from one Domain to users in another.\n" +
                    "Do you want to send any way?",
                    "Cross Send Warning!", MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.No) 
                {
                    Cancel = true;
                }
            }

        }



        private bool isCrossSendEmail(object Item)
        {
            if (accountDomains != null)
            {
                var mailItem = Item as Outlook.MailItem;
                if (mailItem != null &&
                    mailItem.SendUsingAccount != null &&
                    mailItem.SendUsingAccount.SmtpAddress != null)
                {
                    //Identify the account that user is sending from
                    string senderDomain = getDomainStringFromAddress(mailItem.SendUsingAccount.SmtpAddress);

                    var otherDomains = (from i in accountDomains
                                       where !i.Equals(senderDomain)
                                       select i).ToList();
                    
                    if (otherDomains.Count > 0)
                    foreach(Outlook.Recipient recipient in mailItem.Recipients)
                    {
                        string recipientDomain = getDomainStringFromAddress(recipient.Address);
                        if (string.IsNullOrWhiteSpace(recipientDomain))
                        {
                            var exchangeUser = recipient.AddressEntry.GetExchangeUser();
                            if (exchangeUser != null)
                            {
                                recipientDomain = getDomainStringFromAddress(
                                    exchangeUser.PrimarySmtpAddress);
                            }
                            else
                            {
                                var exchangeDL = recipient.AddressEntry.GetExchangeDistributionList();
                                if (exchangeDL != null)
                                {
                                    recipientDomain = getDomainStringFromAddress(
                                        exchangeDL.PrimarySmtpAddress);
                                }
                            }
                        }

                        if (!string.IsNullOrEmpty(recipientDomain))
                        {
                            if (otherDomains.Exists(j => j == recipientDomain))
                            {
                                return true;
                            }   
                        }
                    } //end foreach recipient
                }
            }
            return false;
        }

        private void identifyMultipleEmailDomains()
        {
            List<string> accountDomainList = new List<string>();
            foreach (Outlook.Account account in this.Application.Session.Accounts)
            {
                if (account != null && !string.IsNullOrWhiteSpace(account.SmtpAddress))
                {
                    string accountDomain = getDomainStringFromAddress(account.SmtpAddress);
                    if (!string.IsNullOrEmpty(accountDomain))
                        accountDomainList.Add(accountDomain);
                }
            }

            if (accountDomainList.Count > 1)
            {
                accountDomains = accountDomainList.ToArray();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private static string getDomainStringFromAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address) ||
                !address.Contains("@"))
                return string.Empty;

            return address.Substring(address.IndexOf('@') + 1);
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
