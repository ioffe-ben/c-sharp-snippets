using System;
using System.Windows.Forms;
// Add EWS nuget reference
using Microsoft.Exchange.WebServices.Data;

namespace EmailFolderRedirect
{
    public partial class fMain : Form
    {
        ExchangeService exchange = null;

        public fMain()
        {
            InitializeComponent();
            // Initiate connection to EWS
            ConnectToExchangeServer();
        }
        
        // Function for AutodiscoverUrl
        static bool RedirectionCallback(string url)
        {
            return url.ToLower().StartsWith("https://");
        }

        public void ConnectToExchangeServer()
        {
            try
            {
                // Set a new exchange instance
                exchange = new ExchangeService();
                exchange.Credentials = new WebCredentials("USERNAME", "PASSWORD", "DOMAIN");
                exchange.AutodiscoverUrl("EMAIL@ADDRESS.COM", RedirectionCallback);  
                // Set and start a 1 min timer
                Timer timer = new Timer { Interval = 60000 };    
                timer.Enabled = true;
                // Call MainLogic function every 1 min/timer tick
                timer.Tick += new System.EventHandler(MainLogic);  
            }
            catch (Exception ex){ }
        }
       
        public static void MainLogic(object source, EventArgs e)    
        {
            // Define 1st mailbox with "Inbox" folder
            Mailbox firstMailbox = new Mailbox("EMAIL_1@ADDRESS.COM");
            FolderId firstMailboxFolder = new FolderId(WellKnownFolderName.Inbox, firstMailbox);
            Folder BindFirstMailboxFolder = Folder.Bind(exchange, firstMailboxFolder);

            // Define 2nd mailbox with "Drafts" folder
            Mailbox secondMailbox = new Mailbox("EMAIL_2@ADDRESS.COM");
            FindItemsResults<Item> secondMailbox_findResults = exchange.FindItems(new FolderId(WellKnownFolderName.Drafts, secondMailbox), new ItemView(100));
            
            // Check the 2nd mailbox "Drafts" folder if there are any emails there
            if (secondMailbox_findResults.Items.Count > 0)
            {
                // Loop through 2nd mailbox "Drafts" folder emails
                foreach (Item item in secondMailbox_findResults.Items)
                {
                    item.Load();
                    ItemId ItemToMoveId = new ItemId(item.Id.UniqueId.ToString());
                    Item ItemToMove = Item.Bind(exchange, ItemToMoveId);
                    // Move selected email from the 2nd mailbox "Drafts" folder to the 1st mailbox "Inbox" folder
                    ItemToMove.Move(BindFirstMailboxFolder.Id);
                }
            }
        } 
    }
}
