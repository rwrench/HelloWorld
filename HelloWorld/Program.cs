using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Autodiscover;

namespace HelloWorld
{
    class Program
    {
        static void Main(string[] args)
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            service.Credentials = new WebCredentials("SP", "password","dpsol");
            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;
            try
            { 
                service.AutodiscoverUrl("salespoint@dpsol.com", RedirectionUrlValidationCallback);
            }
                catch (AutodiscoverRemoteException ex)
            {
                Console.WriteLine("Exception thrown: " + ex.Error.Message);
            }

            ItemsFromMail[] items = GetUnreadMailFromInbox(service);
            
        }
      
        private static void SendSampleMail(ExchangeService service,string emailAddr)
        {
            EmailMessage email = new EmailMessage(service);
            email.ToRecipients.Add(emailAddr);
            email.Subject = "HelloWorld";
            email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API");
            email.Send();
        }

        private static ItemsFromMail[] GetUnreadMailFromInbox(ExchangeService service)
        {
            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, new ItemView(128));
            PropertySet itempropertyset = new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.From, EmailMessageSchema.ToRecipients);
            itempropertyset.RequestedBodyType = BodyType.Text;
            ServiceResponseCollection<GetItemResponse> items =
                service.BindToItems(findResults.Select(item => item.Id), itempropertyset);
            return items.Select(item => {
                return new ItemsFromMail()
                {
                    From = ((Microsoft.Exchange.WebServices.Data.EmailAddress)item.Item[EmailMessageSchema.From]).Address,
                    Recipients = ((Microsoft.Exchange.WebServices.Data.EmailAddressCollection)item.Item[EmailMessageSchema.ToRecipients]).Select(recipient => recipient.Address).ToArray(),
                    Subject = item.Item.Subject,
                    Body = item.Item.Body.ToString(),
                    MsgDate = item.Item.DateTimeSent,
                    EntryId = item.Item.Id.UniqueId
                };
            }).ToArray();
        }
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);
            

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }

    public class ItemsFromMail
    {
        public string From;
        public string[] Recipients;
        public string Subject;
        public string Body;
        public DateTime MsgDate;
        public string EntryId;
    }
}
