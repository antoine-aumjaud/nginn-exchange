using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeIntegration.Interfaces
{
    public class ReceivedEmailMessage
    {
        public string From { get; set; }
        public string[] To { get; set; }
        public string[] Cc { get; set; }
        public string[] ReplyTo { get; set; }
        public string ItemId { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string[] AttachmentFiles { get; set; }
        public string InternetMessageId { get; set; }
        public string InReplyTo { get; set; }
        public DateTime ReceivedDate { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime SentDate { get; set; }
    }
}
