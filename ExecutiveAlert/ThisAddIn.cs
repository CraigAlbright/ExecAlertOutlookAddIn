using System;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace ExecutiveAlert
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace _outlookNameSpace;
        Outlook.MAPIFolder _inbox;
        Outlook.Items _items;
        private Outlook.MAPIFolder _contacts;
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _outlookNameSpace = Application.GetNamespace("MAPI");
            _inbox = _outlookNameSpace.GetDefaultFolder(
                Outlook.OlDefaultFolders.olFolderInbox);


            _contacts = _outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            _items = _inbox.Items;
            _items.ItemAdd +=
                items_ItemAdd;
            AddContactsToItems();
        }

        private void AddContactsToItems()
        {
            foreach (var item in _items)
            {
                var mitem = item as Outlook.MailItem;
                if (mitem == null)
                    return;
                var contactAddress = mitem.Sender.Address;
                var isCapSpireAddress = mitem.Sender.Address.Contains("@capspire.com");
                if (!isCapSpireAddress)
                    continue;
                var firstName = mitem.Sender.Address.Substring(0, mitem.Sender.Address.IndexOf('.'));
                var lastName = mitem.Sender.Address.Substring(mitem.Sender.Address.IndexOf('.') + 1,
                    mitem.Sender.Address.IndexOf('@') - mitem.Sender.Address.IndexOf('.') - 1);
                var contact =
                    (Outlook.ContactItem) _contacts.Items.Find(String.Format("[FirstName]='{0}' and "
                                                                             + "[LastName]='{1}'", firstName, lastName));
                if (contact == null)
                {
                    CreateContact(firstName, lastName, contactAddress);
                }

            }
        }

        private void CreateContact(string firstname, string lastname, string contactAddress)
        {
                Outlook.ContactItem newContact = Application.CreateItem(Outlook.OlItemType.olContactItem);
                newContact.Email1Address = contactAddress;
                newContact.FirstName = firstname;
                newContact.LastName = lastname;
                newContact.FullName = string.Concat(firstname, " ", lastname);
                newContact.CompanyName = "capSpire";
                newContact.Save();
        }

        private void items_ItemAdd(object item)
        {
            var mailItem = (Outlook.MailItem) item;
            if(mailItem == null)
                return;
            var execs = Strings.ExecMails.Split(',');
            var me = Strings.MyMail;
            bool isMe = mailItem.Sender.Address.Contains(me);

            var isExecutive = false;
            foreach (var exec in execs.Where(exec => mailItem.Sender.Address.Contains(exec)))
            {
                isExecutive = true;
            }
            if (isExecutive)
            {
                mailItem.FlagStatus = Outlook.OlFlagStatus.olFlagMarked;
                mailItem.FlagIcon = Outlook.OlFlagIcon.olBlueFlagIcon;
                mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                string newBody = "<Font Color=Green>";
                newBody += mailItem.Body + "</Font>";
                mailItem.HTMLBody = newBody;
            }
            if (isMe)
            {
                mailItem.FlagStatus = Outlook.OlFlagStatus.olFlagMarked;
                mailItem.FlagIcon = Outlook.OlFlagIcon.olBlueFlagIcon;
                mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                string newBody = "<Font Color=#484848; Face=Calibri>";
                newBody += mailItem.Body + "</Font>";
                mailItem.HTMLBody = newBody;
            }

        }

        private void NewMailInspector(Outlook.Inspector inspector)
        {
            var mailItem = inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "This text was added by using code";
                    mailItem.Body = "This text was added by using code";
                }

            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
