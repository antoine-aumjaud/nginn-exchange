using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using System.Runtime.Serialization;

namespace ExchangeIntegration.Service
{
    public enum ExchangeEventType
    {
        Status = 0,
        NewMail = 1,
        Deleted = 2,
        Modified = 3,
        Moved = 4,
        Copied = 5,
        Created = 6,
        FreeBusyChanged = 7
    }

    [DataContract]
    public class BaseExchangeEvent
    {
        [DataMember(Order=1)]
        public ExchangeEventType EventType { get; set; }
        [DataMember(Order=2)]
        public DateTime TimeStamp { get; set; }
        [DataMember(Order=3)]
        public string ItemId { get; set; }
        [DataMember(Order=4)]
        public string FolderId { get; set; }
        [DataMember(Order=5)]
        public string Watermark { get; set; }
    }

    [DataContract]
    public class SubscriptionEventNotification
    {
        public SubscriptionEventNotification()
        {
            Events = new List<BaseExchangeEvent>();
        }
        [DataMember(Order=1)]
        public string SubscriptionId { get; set; }
        [IgnoreDataMember]
        public string Watermark
        {
            get
            {
                return Events.Count == 0 ? PreviousWatermark : Events[Events.Count - 1].Watermark;
            }
        }

        [DataMember(Order=3)]
        public string PreviousWatermark { get; set; }
        [DataMember(Order=4)]
        public bool MoreEvents { get; set; }
        [DataMember(Order=5)]
        public bool IsError { get; set; }
        [DataMember(Order=6)]
        public List<BaseExchangeEvent> Events { get; set; }
        /// <summary>
        /// Endpoint where notification messages
        /// should be sent. 
        /// </summary>
        [IgnoreDataMember]
        public string TargetEndpoint { get; set; }
    }
}
