using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeIntegration.Interfaces
{
    /// <summary>
    /// Send an email message
    /// </summary>
    public class SendEmailMessage : CreateItemMessage
    {
        public bool DeliveryReceipt { get; set; }
        public bool ReadReceipt { get; set; }
        public string ReplyToItemId { get; set; }
        public string ForwardItemId { get; set; }
    }

    /// <summary>
    /// Simple reply - for more advanced options use SendEmailMessage
    /// </summary>
    public class ReplyToMessage : ExchangeOperationMessage
    {
        public string ReplyToItemId { get; set; }
        public string Body { get; set; }
        public bool ReplyAll { get; set; }
    }
}
