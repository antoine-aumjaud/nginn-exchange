using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeIntegration.Interfaces
{
    /// <summary>
    /// Add a new event subscription for specified account's folders
    /// </summary>
    public class AddSubscription
    {
        /// <summary>
        /// Account name (SMTP address)
        /// </summary>
        public string AccountName { get; set; }
        /// <summary>
        /// User-defined subscription alias. Should be unique per account, or globally unique.
        /// </summary>
        public string SubscriptionAlias { get; set; }
        /// <summary>
        /// List of folder Ids (either folder unique Ids or WellKnownFolderNames like Inbox, Calendar, Tasks...)
        /// </summary>
        public string[] FolderIds { get; set; }
        /// <summary>
        /// Event types to monitor for (Created, Modified, Deleted, NewMail, ...)
        /// </summary>
        public string[] EventTypes { get; set; }
    }

    /// <summary>
    /// Information that a new subscription has been created
    /// </summary>
    public class SubscriptionCreated
    {
        public string AccountName { get; set; }
        public string SubscriptionAlias { get; set; }
        public string SubscriptionId { get; set; }
    }

    /// <summary>
    /// Remove a subscription
    /// Fill either the AccountName and SubscriptionAlias or SubscriptionId
    /// </summary>
    public class RemoveSubscription
    {
        public string AccountName { get; set; }
        public string SubscriptionAlias { get; set; }
        public string SubscriptionId { get; set; }
    }
}
