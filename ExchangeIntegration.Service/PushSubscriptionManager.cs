using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NLog;
using Microsoft.Exchange.WebServices.Data;
using System.Data;
using ExchangeIntegration.Interfaces;
using NHibernate;
using ExchangeIntegration.Service.Dao;

namespace ExchangeIntegration.Service
{
    /// <summary>
    /// Manages a set of exchange push subscription.
    /// Each subscription status is stored in a database table.
    /// </summary>
    public class PushSubscriptionManager : ISubscriptionManager
    {
        /// <summary>
        /// Exchange web services url
        /// </summary>
        public string EWSUrl { get; set; }

        
        public ISessionFactory SessionFactory { get; set; }
        public IExchangeConnect ExchangeConnect { get; set; }
        public ExchangeEventReceiver EventReceiver { get; set; }
        /// <summary>
        /// Notification status update frequency, in minutes
        /// </summary>
        public int StatusNotificationFreqMinutes { get; set; }

        /// <summary>
        /// Push notification callback url
        /// </summary>
        public string PushNotificationUrl { get; set; }

        private Logger log = LogManager.GetCurrentClassLogger();

        public PushSubscriptionManager()
        {
            StatusNotificationFreqMinutes = 10;
        }

        
        protected void Sub()
        {
            var es = ExchangeConnect.Connect();
            //es.SubscribeToPushNotifications(
            //PushSubscription ps = es.SubscribeToPushNotifications
            
        }

        void AddOrRecreateSubscription()
        {
            
        }

        protected void RecreateSubscription(ExchangeService es, ISession s, PushSubscriptionStatus st)
        {
            
        }

        public bool HandleSubscriptionNotification(SubscriptionEventNotification n)
        {
            using (var s = SessionFactory.OpenSession())
            {
                var sl = s.CreateQuery("from PushSubscriptionStatus where SubscriptionId=:sid")
                    .SetString("sid", n.SubscriptionId).List<PushSubscriptionStatus>();
                if (sl.Count == 0)
                {
                    log.Info("Subscription {0} is inactive - ending");
                    return false;
                }
                var ps = sl[0];
                if (ps.Active == false)
                {
                    log.Info("Subscription {0} is inactive - ending");
                    return false;
                }
                ps.LastUpdate = DateTime.Now;
                ps.Watermark = n.Watermark;
                if (n.Events.Count > 0)
                {
                    ps.LastEventReceived = n.Events[n.Events.Count - 1].TimeStamp;
                }
                ps.ExpectedNextUpdate = DateTime.Now.AddMinutes(2 * StatusNotificationFreqMinutes);
                s.Update(ps);
                s.Flush();
                EventReceiver.HandleIncomingEvents(n);
            }
            return true;
        }

        public string AddSubscription(AddSubscription a, string receiverMessageEndpoint)
        {
            using (var s = SessionFactory.OpenSession())
            {
                var es = ExchangeConnect.ConnectAndImpersonate(a.AccountName);
                
               
                var sl = s.CreateQuery("from PushSubscriptionStatus where Account = :account and Alias = :alias")
                    .SetString("account", a.AccountName)
                    .SetString("alias", a.SubscriptionAlias)
                    .List<PushSubscriptionStatus>();

                if (sl.Count > 0)
                {
                    var s0 = sl[0];
                    if (s0.Active)
                    {
                        log.Warn("Subscription {0}/{1} already active. Subscription Id: {2}", s0.Account, s0.Alias, s0.SubscriptionId);
                        //RecreateSubscription(s0);
                        return s0.SubscriptionId;
                    }
                }

                List<FolderId> lst = new List<FolderId>();
                foreach(string fold in a.FolderIds)
                {
                    WellKnownFolderName f;
                        
                    if (Enum.TryParse<WellKnownFolderName>(fold, out f))
                        lst.Add(f);
                    else
                        lst.Add(new FolderId(fold));
                }
                List<EventType> evs = new List<EventType>();
                foreach(string evtype in a.EventTypes)
                {
                    EventType ev;
                    if (Enum.TryParse<EventType>(evtype, out ev)) 
                        evs.Add(ev);
                    else
                        throw new Exception("Unknown event type: " + evtype);
                }
                string url = PushNotificationUrl;
                if (!string.IsNullOrEmpty(receiverMessageEndpoint))
                {
                    if (url.IndexOf('?') < 0)
                        url += "?endp=";
                    else
                        url += "&endp=";
                    url += Uri.EscapeDataString(receiverMessageEndpoint);
                }
                PushSubscription r = es.SubscribeToPushNotifications(lst, new Uri(url), StatusNotificationFreqMinutes, null, evs.ToArray());
                log.Info("Push subscription created for {0}, url: {1}, Subscription Id: {2}", a.AccountName, url, r.Id);
                var ps = new Dao.PushSubscriptionStatus();
                ps.Account = a.AccountName;
                ps.Active = true;
                ps.SubscriptionId = r.Id;
                ps.Watermark = r.Watermark;
                ps.Alias = a.SubscriptionAlias;
                ps.CreatedDate = DateTime.Now;
                ps.ExpectedNextUpdate = DateTime.Now.AddMinutes(2 * StatusNotificationFreqMinutes);
                ps.LastEventReceived = DateTime.Now;
                ps.LastUpdate = DateTime.Now;
                s.Save(ps);
                s.Flush();
                return r.Id;
            }
        }
    }
}
