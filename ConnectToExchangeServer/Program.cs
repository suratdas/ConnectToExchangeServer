using Microsoft.Exchange.WebServices.Data;
using System;
using System.IO;
using System.Net;
using System.Threading;

namespace ConnectToExchangeServer
    {

    enum RequestType
        {
        Poll,
        Send
        }

    class Program
        {
        static void Main(string[] args)
            {
            Thread t = null;
            var reqType = new RequestType();
            try
                {
                if (args[0].Contains("poll"))
                    {
                    reqType = RequestType.Poll;
                    Console.WriteLine("\n\tTrying to connect to Exchange server.");
                    t = new Thread(ShowProgress);
                    t.Start();
                    }
                else if (args[0].Contains("send"))
                    reqType = RequestType.Send;
                else
                    {
                    Console.WriteLine("\n\tAccepted arguments are 'poll' or 'send'");
                    Environment.Exit(1);
                    }
                }
            catch
                {
                Console.WriteLine("\n\tYou have not provided any argument. Give either 'poll' or 'send' as argument.");
                Environment.Exit(1);
                }

            ExchangeService service = new ExchangeService();
            service.Credentials = new WebCredentials("<username>", "<password>", "<domain>");
            service.AutodiscoverUrl("<full email address>", RedirectionUrlValidationCallback);

            if (reqType == RequestType.Poll)
                {
                if (t != null)
                    t.Abort();
                Console.WriteLine("\n\n\tConnected. Polling for a matching email started. Hit Ctrl+C to quit.");
                while (true)
                    {
                    Thread.Sleep(10000);
                    FolderView folderView = new FolderView(int.MaxValue);
                    FindFoldersResults findFolderResults = service.FindFolders(WellKnownFolderName.Inbox, folderView);
                    foreach (Folder folder in findFolderResults)
                        {

                        if (folder.DisplayName == "<Folder name>" && folder.UnreadCount > 0)
                            {
                            SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
                            ItemView view = new ItemView(1);

                            // Fire the query for the unread items.
                            // This method call results in a FindItem call to EWS.
                            FindItemsResults<Item> findResults = service.FindItems(folder.Id, sf, view);
                            var email = (EmailMessage)findResults.Items[0];

                            EmailMessage message = EmailMessage.Bind(service, email.Id, new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.Attachments));
                            mailFrom = message.From.Address;
                            //Console.WriteLine("Name:" + email.From.Name);
                            string subject = email.Subject;

                            Console.WriteLine("\n\tEmail received from : " + mailFrom + " with subject \"" + subject + " at " + DateTime.Now);
                            
                            folder.MarkAllItemsAsRead(true);
                            //Perform the action you want. Example : Go to the desired URL
                            var client = new WebClient();
                            try{
                            client.OpenRead("http://google.com");
                            }
                            catch(Exception e){}
                            }
                        }
                    }
                }
            else if (reqType == RequestType.Send)
                {
                SendEmailWithReport(service, "<To email address>");
                }
            }

        private static void SendEmailWithReport(ExchangeService service, string recipient)
            {
            EmailMessage email = new EmailMessage(service);
            EmailAddress to = new EmailAddress();
            MessageBody body = new MessageBody();
            body.BodyType = BodyType.HTML;
            body.Text = "sample body.";
            var subject = "Sample subject " + DateTime.Now;

            to.Address = recipient;
            email.ToRecipients.Add(to);
            email.Subject = subject;

            email.Body = body + "\r\n" + File.ReadAllText(@"C:\sample.html");

            email.Send();
            }

        private static bool RedirectionUrlValidationCallback(String redirectionUrl)
            {
            bool redirectionValidated = false;
            if (redirectionUrl.Equals(
                "https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml"))
                redirectionValidated = true;
            return redirectionValidated;
            }

        private static void ShowProgress()
            {
            Console.Write("\n\t");
            while (true)
                {
                Console.Write("#");
                Thread.Sleep(500);
                }
            }

        }

    }
