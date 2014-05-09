using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using stdole;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace ExecutiveAlert
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace _outlookNameSpace;
        Outlook.MAPIFolder _inbox;
        Outlook.Items _items;
        Outlook.MAPIFolder _contacts; 
        Office.CommandBar _newToolBar;
        Office.CommandBarButton _firstButton;
        Office.CommandBarButton _secondButton;
        Outlook.Explorers _selectExplorers;
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            //Executive flag addin
            _outlookNameSpace = Application.GetNamespace("MAPI");
            _inbox = _outlookNameSpace.GetDefaultFolder(
                Outlook.OlDefaultFolders.olFolderInbox);

            //Generate contacts from email I have received from capSpire
            _contacts = _outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            _items = _inbox.Items;
            _items.ItemAdd +=
                items_ItemAdd;
            AddContactsToItems();

            _selectExplorers = Application.Explorers;
            _selectExplorers.NewExplorer += newExplorer_Event;
            AddToolbar();
        }

        private void newExplorer_Event(Outlook.Explorer explorer)
        {
            
        }

        private void AddToolbar()
        {

            if (_newToolBar == null)
            {
                var cmdBars =
                    Application.ActiveExplorer().CommandBars;
                _newToolBar = cmdBars.Add("Create Contacts",
                    Office.MsoBarPosition.msoBarTop, false, true);
            }
            try
            {
                var createCapContacts =
                    (Office.CommandBarButton)_newToolBar.Controls
                    .Add(1, missing, missing, missing, missing);
                createCapContacts.Style = Office
                    .MsoButtonStyle.msoButtonIconAndCaption;
                createCapContacts.Caption = "Create capSpire Contacts";

                Image logo = Image.FromFile(
                        @"C:\Users\Craig\documents\visual studio 2013\Projects\ExecutiveAlert\ExecutiveAlert\Images\CapLogo.png");
                
                createCapContacts.Picture = PictureConverter.ImageToPictureDisp(logo);
                
                createCapContacts.Width = 130;
                createCapContacts.Height = 130;
                createCapContacts.Tag = "createContacts";
                if (_firstButton == null)
                {
                    _firstButton = createCapContacts;
                    _firstButton.Click += ButtonClick;
                }
                //Todo: second button idea
                var button2 = (Office.CommandBarButton)_newToolBar.Controls.Add(1);
                button2.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                button2.Caption = "Button 2";
                button2.Tag = "Button2";
                _newToolBar.Visible = true;
                if (_secondButton == null)
                {
                    _secondButton = button2;
                    _secondButton.Click += ButtonClick;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonClick(Office.CommandBarButton ctrl,
                ref bool cancel)
        {
            if(ctrl.Caption =="Create capSpire Contacts")
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
