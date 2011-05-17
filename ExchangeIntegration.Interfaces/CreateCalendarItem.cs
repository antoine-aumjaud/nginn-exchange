using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeIntegration.Interfaces
{
    public class CreateCalendarItem : CreateItemMessage
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string[] Recipients { get; set; }
    }
}
