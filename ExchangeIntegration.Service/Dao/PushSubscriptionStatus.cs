using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExchangeIntegration.Interfaces;
using Newtonsoft.Json;

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
        /// <summary>
        /// Status update frequency, in minutes.
        /// </summary>
        public virtual int StatusUpdateFreqMinutes { get; set; }
        public virtual bool Active { get; set; }
        /// <summary>
        /// exchange account name
        /// </summary>
        public virtual string Account { get; set; }
        /// <summary>
        /// message bus endpoint of the event recipient
        /// </summary>
        public virtual string RecipientEndpoint { get; set; }
        public virtual string SubscriptionRequestJson { get; set; }

        public virtual AddSubscription SubscriptionRequest
        {
            get
            {
                return string.IsNullOrEmpty(SubscriptionRequestJson) ? null : JsonConvert.DeserializeObject<AddSubscription>(SubscriptionRequestJson);
            }
            set
            {
                SubscriptionRequestJson = JsonConvert.SerializeObject(value);
            }
        }
    }
}
