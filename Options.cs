using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExchangeInboxWatcher
{
    class Options
    {
        public string EmailAddress { get; private set; }
        public string AccountPassword { get; private set; }
        public string AccountDomain { get; private set; }

        public double MaxIdleTimeMinutes { get; private set; }

        public string AlertRecipient { get; private set; }

        public static Options Parse(string[] args)
        {
            if (args.Length != 5)
            {
                throw new Exception("Invalid number of arguments");
            }

            return new Options
            {
                EmailAddress = args[0],
                AccountPassword= args[1],
                AccountDomain = args[2],
                MaxIdleTimeMinutes= double.Parse(args[3]),
                AlertRecipient = args[4]
            };

        }
    }
}
