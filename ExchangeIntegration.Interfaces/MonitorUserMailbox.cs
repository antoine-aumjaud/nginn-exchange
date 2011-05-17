using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeIntegration.Interfaces
{
    /// <summary>
    /// Receive event notifications from user's mailbox
    /// </summary>
    public class MonitorUserMailbox : ExchangeOperationMessage
    {
        public bool MonitorCalendarItems { get; set; }
        public bool MonitorIncomingMail { get; set; }
        public bool MonitorTaskCompletion { get; set; }
        public DateTime ExpirationDate { get; set; }
    }
}
