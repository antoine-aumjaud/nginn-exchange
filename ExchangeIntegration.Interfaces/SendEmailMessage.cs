using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeIntegration.Interfaces
{
    public class SendEmailMessage : CreateItemMessage
    {
        public bool DeliveryReceipt { get; set; }
        public bool ReadReceipt { get; set; }

    }
}
