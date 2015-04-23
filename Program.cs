using System;
using System.Diagnostics;
using System.Net.Mime;

namespace ExchangeInboxWatcher
{
    class Program
    {
        private const string ProgramName = "ExchangeInboxWatcher";

        static void Main(string[] args)
        {
            try
            {
                var options = Options.Parse(args);
                var emailStore = new ExchangeEmailStore(options.EmailAddress, options.AccountPassword, options.AccountDomain);
                emailStore.AutoDiscover();
                var lastRecievedDate = emailStore.GetLastEmailReceiveDate();
                validateReceiveDate(lastRecievedDate, options, emailStore);
            }
            catch (Exception ex)
            {
                if (!EventLog.SourceExists(ProgramName))
                {
                    EventLog.CreateEventSource(ProgramName, "Application");
                }

                EventLog.WriteEntry(ProgramName, ex.Message,
                    EventLogEntryType.Error);             
            }
        }

        private static void validateReceiveDate(DateTime lastRecievedDate, Options options, ExchangeEmailStore emailStore)
        {
            if (lastRecievedDate == DateTime.MinValue)
            {
                emailStore.SendEmail(options.AlertRecipient, string.Format("The inbox for the account {0} is empty.", options.EmailAddress), 
                    "Email Account Exceeded Maximum Idle Time");
            }
            else
            {
                var timespan = DateTime.Now - lastRecievedDate;
                if (timespan.TotalMinutes > options.MaxIdleTimeMinutes)
                {
                    var messageBody = BuildAlertMessage(options, timespan, lastRecievedDate);
                    emailStore.SendEmail(options.AlertRecipient, messageBody, "Email Account Exceeded Maximum Idle Time");
                }
            }
        }

        private static string BuildAlertMessage(Options options, TimeSpan timespan, DateTime lastRecievedDate)
        {
            return string.Format("The last email that was recieved for the account {0} was {1} minutes ago at {2} which exceeds the maximum idle time of {3} minutes.", 
                options.EmailAddress, (int)timespan.TotalMinutes, lastRecievedDate, options.MaxIdleTimeMinutes);
        }
    }
}
