using System;
using System.Windows.Forms;
using Microsoft.Exchange.WebServices.Data;

namespace EmailFolderRedirect
{
    public partial class fMain : Form
    {
        ExchangeService exchange = null;

        public fMain()
        {
            InitializeComponent();
            ConnectToExchangeServer();
        }

        static bool RedirectionCallback(string url)
        {
            return url.ToLower().StartsWith("https://");
        }

        public void ConnectToExchangeServer()
        {
            try
            {
                exchange = new ExchangeService();
                exchange.Credentials = new WebCredentials("USERNAME", "PASSWORD", "DOMAIN");
                exchange.AutodiscoverUrl("EMAIL@ADDRESS.COM", RedirectionCallback);             
                Timer timer = new Timer { Interval = 60000 };    
                timer.Enabled = true;    
                timer.Tick += new System.EventHandler(MainLogic);  
            }
            catch (Exception ex){ }
        }
       
        public static void MainLogic(object source, EventArgs e)    
        {
            Mailbox firstMailbox = new Mailbox("EMAIL_1@ADDRESS.COM");
            FolderId firstMailboxFolder = new FolderId(WellKnownFolderName.Inbox, firstMailbox);
            Folder BindFirstMailboxFolder = Folder.Bind(exchange, firstMailboxFolder);

            Mailbox secondMailbox = new Mailbox("EMAIL_2@ADDRESS.COM");
            FindItemsResults<Item> secondMailbox_findResults = exchange.FindItems(new FolderId(WellKnownFolderName.Inbox, secondMailbox), new ItemView(10));

            if (secondMailbox_findResults.Items.Count > 0)
            {
                foreach (Item item in secondMailbox_findResults.Items)
                {
                    item.Load();
                    ItemId ItemToMoveId = new ItemId(item.Id.UniqueId.ToString());
                    Item ItemToMove = Item.Bind(exchange, ItemToMoveId);
                    ItemToMove.Move(BindFirstMailboxFolder.Id);
                }
            }
        } 
    }
}
