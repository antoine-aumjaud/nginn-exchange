using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeIntegration.Interfaces
{
    public class DeleteItem
    {
        public string CorrelationId { get; set; }
        public string ItemId { get; set; }
    }
}
