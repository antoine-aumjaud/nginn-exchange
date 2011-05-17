using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeIntegration.Interfaces
{
    public class AddSubscription
    {
        public string AccountName { get; set; }
        public string SubscriptionAlias { get; set; }
        public string[] FolderIds { get; set; }
        public string[] EventTypes { get; set; }
    }

    public class SubscriptionCreated
    {
        public string SubscriptionAlias { get; set; }
        public string SubscriptionId { get; set; }
    }

    public class RemoveSubscription
    {
        public string SubscriptionAlias { get; set; }
        public string SubscriptionId { get; set; }
    }
}
