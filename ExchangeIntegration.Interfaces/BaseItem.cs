using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeIntegration.Interfaces
{
    public class BaseItem
    {
        public string ItemId { get; set; }
        public string ItemClass { get; set; }
        public string Subject { get; set; }
        public string InternetMessageId { get; set; }
        public string CorrelationId { get; set; }
        public string[] AttachmentFiles { get; set; }
        public bool DeleteAttachments { get; set; }

    }
}
