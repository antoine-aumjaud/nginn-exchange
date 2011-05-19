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
using Newtonsoft.Json;
using System.Timers;
using NGinnBPM.MessageBus;
using System.Transactions;


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
        public IMessageBus MessageBus { get; set; }
        /// <summary>
        /// Notification status update frequency, in minutes
        /// </summary>
        public int StatusNotificationFreqMinutes { get; set; }

        /// <summary>
        /// Push notification callback url
        /// </summary>
        public string PushNotificationUrl { get; set; }

        private Logger log = LogManager.GetCurrentClassLogger();
        private Timer _refreshTimer;

        public PushSubscriptionManager()
        {
            StatusNotificationFreqMinutes = 10;
            _refreshTimer = new Timer();
            _refreshTimer.AutoReset = false;
            _refreshTimer.Interval = 30000.0;
            _refreshTimer.Elapsed += new ElapsedEventHandler(_refreshTimer_Elapsed);
            _refreshTimer.Start();
        }

        void _refreshTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                RecreateDeadSubscriptions();
            }
            catch (Exception ex)
            {
                log.Error("Error: {0}", ex);
            }
            _refreshTimer.Start();
        }

        protected void RecreateDeadSubscriptions()
        {
            using (var s = SessionFactory.OpenSession())
            {
                var l = s.CreateQuery("from PushSubscriptionStatus where Active = true and ExpectedNextUpdate < :deadline")
                    .SetDateTime("deadline", DateTime.Now.AddMinutes(-5))
                    .List<PushSubscriptionStatus>();
                
                foreach (var ps in l)
                {
                    try
                    {
                        log.Info("Recreating subscription {0}", ps.Id);
                        RecreateSubscription(ps.Id);
                    }
                    catch (Exception ex)
                    {
                        log.Error("Error recreating subscription {0}", ex);
                    }
                }
            }
        }

        

        internal PushSubscription SubscribeToPushNotifications(ExchangeService es, AddSubscription a, string watermark)
        {
            List<FolderId> lst = new List<FolderId>();
            foreach (string fold in a.FolderIds)
            {
                WellKnownFolderName f;
                if (Enum.TryParse<WellKnownFolderName>(fold, out f))
                    lst.Add(f);
                else
                    lst.Add(new FolderId(fold));
            }
            List<EventType> evs = new List<EventType>();
            foreach (string evtype in a.EventTypes)
            {
                EventType ev;
                if (Enum.TryParse<EventType>(evtype, out ev))
                    evs.Add(ev);
                else
                    throw new Exception("Unknown event type: " + evtype);
            }
            string url = PushNotificationUrl;
            PushSubscription r = es.SubscribeToPushNotifications(lst, new Uri(url), StatusNotificationFreqMinutes, watermark, evs.ToArray());
            return r;
        }

        protected void RecreateSubscription(int id)
        {
            using (var s = SessionFactory.OpenSession())
            {
                var ps = s.Load<PushSubscriptionStatus>(id);
                log.Info("Trying to recreate subscription {0} ({1})", ps.Id, ps.SubscriptionId);
                if (ps.SubscriptionRequest == null)
                {
                    log.Info("Can't recreate subscription - no request");
                    return;
                }
                var a = ps.SubscriptionRequest;
                var es = ExchangeConnect.ConnectAndImpersonate(a.AccountName);
                var r = SubscribeToPushNotifications(es, a, ps.Watermark);
                if (!ps.Active)
                {
                    ps.Active = true;
                    log.Info("Subscription {0} changed to active", ps.Id);
                }
                ps.LastUpdate = DateTime.Now;
                var prevId = ps.SubscriptionId;
                ps.Watermark = r.Watermark;
                ps.SubscriptionId = r.Id;
                ps.StatusUpdateFreqMinutes = this.StatusNotificationFreqMinutes;
                ps.ExpectedNextUpdate = DateTime.Now.AddMinutes(2 * ps.StatusUpdateFreqMinutes);
                if (!string.IsNullOrEmpty(ps.RecipientEndpoint))
                {
                    MessageBus.NewMessage(new SubscriptionRecreated
                    {
                        AccountName = a.AccountName,
                        NewSubscriptionId = ps.SubscriptionId,
                        SubscriptionId = prevId,
                        SubscriptionAlias = ps.Alias
                    }).Send(ps.RecipientEndpoint);
                }
                s.Update(ps);
                s.Flush();
            }
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
                n.TargetEndpoint = ps.RecipientEndpoint;
                if (n.Events.Count > 0)
                {
                    ps.LastEventReceived = n.Events[n.Events.Count - 1].TimeStamp;
                }
                ps.ExpectedNextUpdate = DateTime.Now.AddMinutes(2 * ps.StatusUpdateFreqMinutes);
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
                
                var r = SubscribeToPushNotifications(es, a, null);
                log.Info("Push subscription created for {0}, url: {1}, Subscription Id: {2}", a.AccountName, PushNotificationUrl, r.Id);
                var ps = new Dao.PushSubscriptionStatus();
                ps.Account = a.AccountName;
                ps.Active = true;
                ps.SubscriptionId = r.Id;
                ps.Watermark = r.Watermark;
                ps.Alias = a.SubscriptionAlias;
                ps.CreatedDate = DateTime.Now;
                ps.StatusUpdateFreqMinutes = StatusNotificationFreqMinutes;
                ps.ExpectedNextUpdate = DateTime.Now.AddMinutes(2 * ps.StatusUpdateFreqMinutes);
                ps.LastEventReceived = DateTime.Now;
                ps.LastUpdate = DateTime.Now;
                ps.RecipientEndpoint = receiverMessageEndpoint;
                ps.SubscriptionRequestJson = JsonConvert.SerializeObject(a);
                s.Save(ps);
                s.Flush();
                return r.Id;
            }
        }
    }
}
