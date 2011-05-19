using System;
using System.Collections.Generic;
using SC = System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using Castle.Windsor;
using Castle.MicroKernel.Registration;
using NLog;
using NGinnBPM.MessageBus;
using NGinnBPM.MessageBus.Windsor;
using System.Configuration;
using NHibernate;
using System.ServiceModel;

namespace ExchangeIntegration.Service
{
    public partial class Service1 : ServiceBase
    {
        private IWindsorContainer _container = null;
        private ServiceHost _pushReceiverHost = null;

        private Logger log = LogManager.GetCurrentClassLogger();

        public Service1()
        {
            InitializeComponent();
        
        }

        public void Start(string[] args)
        {
            this.OnStart(args);
        }


        protected override void OnStart(string[] args)
        {
            OnStop();
            BuildConfiguration();
        }

        protected void BuildConfiguration()
        {
            var httplistener = ConfigurationManager.AppSettings["NGinnMessageBus.HttpReceiver"];
            var pushReceiverUrl = ConfigurationManager.AppSettings["WcfPushNotificationReceiverUrl"];

            _container = new WindsorContainer();
            log.Debug("Configuring NH");
            NHibernate.Cfg.Configuration ncf = new NHibernate.Cfg.Configuration();
            ncf.AddAssembly(typeof(PushSubscriptionManager).Assembly);
            ncf.Configure();
            var sf = ncf.BuildSessionFactory();
            _container.Register(Component.For<ISessionFactory>().Instance(sf));

            log.Debug("Configuring message bus");
            _container.Register(Component.For<IExchangeIntegrationService, IExchangeConnect>()
                .ImplementedBy<ExchangeIntegrationService>().LifeStyle.Singleton
                .DependsOn(new
                {
                    EWSUser = ConfigurationManager.AppSettings["EWSUser"],
                    EWSPassword = ConfigurationManager.AppSettings["EWSPassword"],
                    EWSUrl = ConfigurationManager.AppSettings["EWSUrl"]
                }));
            _container.Register(Component.For<NGinnBPM.MessageBus.Impl.IStartableService, ExchangeEventReceiver>()
                .ImplementedBy<ExchangeEventReceiver>().LifeStyle.Singleton
                .DependsOn(new
                {
                    EWSUser = ConfigurationManager.AppSettings["EWSUser"],
                    EWSPassword = ConfigurationManager.AppSettings["EWSPassword"],
                    EWSUrl = ConfigurationManager.AppSettings["EWSUrl"],
                    Name = "atmosfera_recv"
                }));
            _container.Register(Component.For<ISubscriptionManager, PushSubscriptionManager>()
                .ImplementedBy<PushSubscriptionManager>().LifeStyle.Singleton
                .DependsOn(new
                {
                    //PushNotificationUrl = "http://192.168.21.1:9019/PushNotification",
                    PushNotificationUrl = "http://192.168.21.1:9019/push/",
                    StatusNotificationFreqMinutes = 1
                }));
            MessageBusConfigurator c = MessageBusConfigurator.Begin(_container);
            foreach (string s in ConfigurationManager.AppSettings.Keys)
            {
                var prf = "NGinnMessageBus.ConnectionString.";
                if (s.StartsWith(prf))
                    c.AddConnectionString(s.Substring(prf.Length), ConfigurationManager.AppSettings[s]);
            }

            if (!string.IsNullOrEmpty(httplistener))
                c.ConfigureHttpReceiver(httplistener);

            c.SetEndpoint(ConfigurationManager.AppSettings["NGinnMessageBus.Endpoint"])
                .AddMessageHandlersFromAssembly(typeof(Service1).Assembly)
                .UseSqlSubscriptions()
                .SetMessageRetentionPeriod(TimeSpan.FromDays(10))
                .FinishConfiguration()
                .RegisterHttpHandlersFromAssembly(typeof(Service1).Assembly)
                .StartMessageBus();
            log.Info("Message bus started");

            _container.Resolve<ISubscriptionManager>();

            /*if (!string.IsNullOrEmpty(pushReceiverUrl))
            {
                log.Info("Creating WCF push receiver host at {0}", pushReceiverUrl);
                _container.Register(Component.For<WcfExchangePushNotificationReceiver>()
                    .ImplementedBy<WcfExchangePushNotificationReceiver>());
                _pushReceiverHost = new ServiceHost(_container.Resolve<WcfExchangePushNotificationReceiver>(), new Uri(pushReceiverUrl));
                _pushReceiverHost.Open();
            }*/
        }

        protected override void OnStop()
        {
            if (_pushReceiverHost != null)
            {
                _pushReceiverHost.Close();
                _pushReceiverHost = null;
            }
            if (_container != null)
            {
                foreach (NGinnBPM.MessageBus.Impl.IStartableService srv in _container.ResolveAll<NGinnBPM.MessageBus.Impl.IStartableService>())
                {
                    srv.Stop();
                }
                _container.Dispose();
                _container = null;
            }
        }
    }
}
