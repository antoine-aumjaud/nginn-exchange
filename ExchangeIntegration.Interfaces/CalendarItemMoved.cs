using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeIntegration.Interfaces
{
    public class CalendarItemMoved
    {
        public string ItemId { get; set; }
        public string ChangeKey { get; set; }
        public string CorrelationId { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
    }
}
