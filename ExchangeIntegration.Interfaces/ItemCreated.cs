using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeIntegration.Interfaces
{
    public class ItemCreated
    {
        public string CorrelationId { get; set; }
        public string UniqueId { get; set; }
        public string InternetMessageId { get; set; }

        public override string ToString()
        {
            return string.Format("ItemCreated:{0}|{1}", UniqueId, CorrelationId);
        }
    }
}
