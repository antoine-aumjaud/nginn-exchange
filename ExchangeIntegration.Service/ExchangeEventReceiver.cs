using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NLog;
using Microsoft.Exchange.WebServices.Data;
using EWSData = Microsoft.Exchange.EWSDataTypes;

using NGinnBPM.MessageBus;
using ExchangeIntegration.Interfaces;
using System.IO;
using System.Threading;

namespace ExchangeIntegration.Service
{
    public class ExchangeEventReceiver : NGinnBPM.MessageBus.Impl.IStartableService
    {
        public string EWSUrl { get; set; }
        public IMessageBus MessageBus { get; set; }
        public IExchangeIntegrationService ExchangeIntegration { get; set; }
        public string EWSUser { get; set; }
        public string EWSPassword { get; set; }
        public TimeSpan PollingInterval { get; set; }
        public int SubscriptionTimeout { get; set; }
        public string Name { get; set; }

        private static Logger log = LogManager.GetCurrentClassLogger();
        private AutoResetEvent _stopEvent = new AutoResetEvent(false);
        private Thread _pollingThread = null;
        private bool _stop = false;

        public ExchangeEventReceiver()
        {
            PollingInterval = TimeSpan.FromSeconds(10);
            SubscriptionTimeout = 5;
        }

        public bool IsRunning
        {
            get { return _pollingThread != null; }
        }

        public void Start()
        {
            
            lock (this)
            {
                if (IsRunning)
                    return;
                log.Info("Starting event polling thread");
                _stop = false;
                //_pollingThread = new Thread(new ThreadStart(this.EventPollingThread));
                //_pollingThread.Start();
            }
        }

        private void EventPollingThread()
        {
            log.Info("Starting pull subscription");
            var es = Connect();
            
            FolderId[] folders = new FolderId[] {
                WellKnownFolderName.Calendar,
                WellKnownFolderName.Inbox,
                WellKnownFolderName.Tasks
            };
            EventType[] events = new EventType[] {
                EventType.NewMail,
                EventType.Created,
                EventType.Deleted,
                //EventType.FreeBusyChanged,
                EventType.Modified,
                EventType.Moved,
                EventType.Copied
            };
            PullSubscription ps = es.SubscribeToPullNotifications(folders, SubscriptionTimeout, Watermark, events);
            log.Info("Pull subscription created: {0}. Watermark: {1}. Events available: {2}", ps.Id, ps.Watermark, ps.MoreEventsAvailable);
            
            while (!_stop)
            {
                GetEventsResults r = ps.GetEvents();
                log.Debug("Got {0} events, more available: {1}", r.AllEvents.Count, ps.MoreEventsAvailable);
                foreach (NotificationEvent ev in r.AllEvents)
                {
                    if (_stop) break;
                    ProcessEvent(es, ev);
                    Watermark = ps.Watermark;
                }
                Watermark = ps.Watermark;
                if (ps.MoreEventsAvailable.HasValue && ps.MoreEventsAvailable.Value == false)
                {
                    _stopEvent.WaitOne(PollingInterval);
                }
            }
            ps.Unsubscribe();
        }

        protected void ProcessEvent(ExchangeService es, NotificationEvent ev)
        {
            if (ev is ItemEvent)
            {
                ItemEvent ie = (ItemEvent)ev;
                Item it = Item.Bind(es, ie.ItemId);
                log.Info("Processing item event {0} ({1}) Timestamp: {2}", ie.GetType().Name, ie.EventType, ie.TimeStamp);
                DumpItemInfo(es, it);
                if (ie.EventType == EventType.Moved)
                {
                    ProcessItemMoved(ie);
                }
            }
            else
            {
                log.Info("Processing event {0} {1}", ev.GetType().Name, ev.EventType);
            }
        }

        protected void ProcessItemMoved(ItemEvent ev)
        {
            ExchangeService es = Connect();
            Item it = Item.Bind(es, ev.ItemId);
            log.Info("\n\nMOVED item: {0}: {1}\n\n", it.GetType().Name, it.Subject);
        }

        private void DumpItemInfo(ExchangeService es, Item it)
        {
            log.Info("\n----------------------------------");
            log.Info("Item {0}: {1}, Received: {2}, To: {3}, ItemClass: {4}, InReplyTo: {5}", it.GetType().Name, it.Subject, it.DateTimeReceived, it.DisplayTo, it.ItemClass, it.InReplyTo);
            if (it.InternetMessageHeaders != null)
            {
                foreach (var h in it.InternetMessageHeaders)
                {
                    log.Info("      {0}: {1}", h.Name, h.Value);
                }
            }
            if (it is EmailMessage)
            {
                EmailMessage em = (EmailMessage)it;
                
                //em.IsReminderSet = false;
                //em.IsResponseRequested = true;
                //em.IsAssociated = true;
                //em.IsDeliveryReceiptRequested = true;
                
                log.Info("Email. Internet message id: {0}", em.InternetMessageId);
                if (!string.IsNullOrEmpty(em.InReplyTo))
                {
                    SearchFilter sf = new SearchFilter.IsEqualTo(EmailMessageSchema.InternetMessageId, em.InReplyTo);
                    FindItemsResults<Item> r = es.FindItems(it.ParentFolderId, sf, new ItemView(10));
                    log.Info("Found {0} original messages", r.Items.Count);
                    if (r.Items.Count > 0)
                    {
                        var it2 = r.Items[0];
                        log.Info("Original message: {0}", it2.Subject);
                    }
                }

            }
            else if (it is MeetingRequest)
            {
            }
            log.Info("-------------------------------------\n\n");
        }

        public void Stop()
        {
            lock (this)
            {
                if (_pollingThread != null)
                {
                    _stop = true;
                    _stopEvent.Set();
                    _pollingThread.Join(30000);
                    _pollingThread = null;
                }
            }
        }

        private string _wmark;
        private string Watermark
        {
            get
            {
                if (_wmark != null)
                    return _wmark;
                string path = Path.Combine(Path.GetDirectoryName(typeof(ExchangeIntegrationService).Assembly.Location), "watermark.txt");
                if (File.Exists(path))
                {
                    using (StreamReader sr = new StreamReader(path))
                    {
                        _wmark = sr.ReadToEnd();
                    }
                }
                return _wmark;
            }
            set
            {
                lock (this)
                {
                    if (value != null && value != _wmark)
                    {
                        string path = Path.Combine(Path.GetDirectoryName(typeof(ExchangeIntegrationService).Assembly.Location), "watermark.txt");
                        using (StreamWriter sw = new StreamWriter(path, false))
                        {
                            sw.Write(_wmark);
                        }
                    }
                    _wmark = value;
                    log.Debug("Watermark is {0}", _wmark);
                }
            }
        }

        private ExchangeService Connect()
        {
            ExchangeService es = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            es.Url = new Uri(EWSUrl);
            if (string.IsNullOrEmpty(EWSUser))
            {
                es.UseDefaultCredentials = true;
            }
            else
            {
                es.Credentials = new System.Net.NetworkCredential(EWSUser, EWSPassword);
            }
            
            return es;
        }


        public void HandleIncomingEvents(SubscriptionEventNotification n)
        {
            var es = Connect();
            foreach (var ev in n.Events)
            {
                if (!string.IsNullOrEmpty(ev.ItemId))
                {
                    HandleItemEvent(ev, n, es);
                }
                else if (!string.IsNullOrEmpty(ev.FolderId))
                {
                }
                else
                {
                }
                
            }
        }

        protected void HandleItemEvent(BaseExchangeEvent ev, SubscriptionEventNotification n, ExchangeService es)
        {
            Item it = Item.Bind(es, new ItemId(ev.ItemId));
            log.Info("Item event: {0}. Item class: {1} ({3}), Id: {2}", ev.EventType, it.ItemClass, it.Id, it.GetType().Name);
            ExtendedPropertyDefinition epd = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "X-CorrelationId", MapiPropertyType.String);
            it.Load(new PropertySet(epd));
            foreach (ExtendedProperty ep in it.ExtendedProperties)
            {
                log.Info("PROP: {0}={1}", ep.PropertyDefinition.Name, ep.Value);
            }
            List<object> messages = new List<object>();
            if (ev.EventType == ExchangeEventType.Created)
            {
                /*messages.Add(new ItemCreated
                {
                    UniqueId = it.Id.UniqueId
                });*/
            }
            else if (ev.EventType == ExchangeEventType.Modified)
            {
                messages.Add(new CalendarItemMoved
                {
                    ItemId = it.Id.UniqueId
                });
            }
            else if (ev.EventType == ExchangeEventType.NewMail)
            {
                messages.Add(new ItemCreated
                {
                    UniqueId = it.Id.UniqueId
                });
            }
            if (messages.Count > 0)
            {
                MessageBus.NewMessage().SetBatch(messages.ToArray())
                    .Send(string.IsNullOrEmpty(n.TargetEndpoint) ? MessageBus.Endpoint : n.TargetEndpoint);
            }
        }
        
    }
}
