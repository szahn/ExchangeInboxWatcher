using System;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeInboxWatcher
{
    class ExchangeEmailStore
    {
        private readonly ExchangeService service = new ExchangeService();
        private string email;
        private string password;
        private string domain;

        public ExchangeEmailStore(string email, string password, string domain)
        {
            this.email = email;
            this.password = password;
            this.domain = domain;
        }

        public void AutoDiscover()
        {
            service.UseDefaultCredentials = false;
            service.Credentials = new WebCredentials(email, password, domain);
            service.AutodiscoverUrl(email, (s) => true);            
        }

        public DateTime GetLastEmailReceiveDate()
        {
            var view = new ItemView(1);
            var items = service.FindItems(WellKnownFolderName.Inbox, view);
            if (!items.Any())
            {
                return DateTime.MinValue;
            }

            var emailItem = items.FirstOrDefault();
            var date = emailItem.DateTimeReceived;
            return date;
        }

        public void SendEmail(string recipient, string body, string subject)
        {
            var message = new EmailMessage(service) {Subject = subject, Body = new TextBody(body)};
            message.ToRecipients.Add(recipient);
            message.Send();
        }
    }
}
