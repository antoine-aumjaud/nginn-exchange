using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeIntegration.Service.Dao
{
    public class PushSubscriptionStatus
    {
        public virtual int Id { get; set; }
        public virtual string Alias { get; set; }
        public virtual string SubscriptionId { get; set; }
        public virtual string Watermark { get; set; }
        public virtual DateTime CreatedDate { get; set; }
        public virtual DateTime LastUpdate { get; set; }
        public virtual DateTime LastEventReceived { get; set; }
        public virtual DateTime ExpectedNextUpdate { get; set; }
        public virtual bool Active { get; set; }
        public virtual string Account { get; set; }
    }
}
