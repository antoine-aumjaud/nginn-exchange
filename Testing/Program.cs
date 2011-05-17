using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NLog;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using NGinnBPM.MessageBus;
using NGinnBPM.MessageBus.Windsor;
using Castle.Windsor;
using Castle.MicroKernel.Registration;
using System.Configuration;
using ExchangeIntegration.Interfaces;

namespace Testing
{
    class Program
    {
        static Logger log = LogManager.GetCurrentClassLogger();
        static WindsorContainer _wc;

        static void Main(string[] args)
        {
            try
            {
                NLog.Config.SimpleConfigurator.ConfigureForConsoleLogging();
                ConfigureMessageBus();
                //TestSendEmail();
                //TestCreateItem();
                //TestGetItems();
                //TestRaw();
                TestSubscribe();
            }
            catch (Exception ex)
            {
                log.Error("Error: {0}", ex);
            }
            Console.ReadLine();
        }

        static ExchangeService Connect()
        {
            log.Info("Connect");
            ExchangeService es = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            es.Url = new Uri("http://vmex03/EWS/exchange.asmx");
            //es.UseDefaultCredentials = false;
            
            //es.Credentials = new WebCredentials("rafalg", "1qaz@WSX", "COGIT2.PL");
            es.Credentials = new NetworkCredential("rafalg", "1qazXSW@");
            
            return es;
        }

        static void TestGetItems()
        {
            ExchangeService es = Connect();
            var res = es.FindItems(WellKnownFolderName.Inbox, new ItemView(50));
            log.Info("Test get items finished");
        }

        static void TestRaw()
        {
            WebRequest wr = WebRequest.Create("http://vmex03/EWS/Services.wsdl");
            //wr.AuthenticationLevel = System.Net.Security.AuthenticationLevel.MutualAuthRequested;
            wr.Credentials = new NetworkCredential("rafalg", "1qazXSW@");
            
            
            WebResponse resp = wr.GetResponse();
            log.Info("Ret: {0}", resp.ContentLength);
        }

        static void ConfigureMessageBus()
        {
            _wc = new WindsorContainer();
            var c = MessageBusConfigurator.Begin(_wc);
            foreach (string s in ConfigurationManager.AppSettings.Keys)
            {
                var prf = "NGinnMessageBus.ConnectionString.";
                if (s.StartsWith(prf))
                    c.AddConnectionString(s.Substring(prf.Length), ConfigurationManager.AppSettings[s]);
            }
            c.SetEndpoint(ConfigurationManager.AppSettings["NGinnMessageBus.Endpoint"])
                .FinishConfiguration()
                .StartMessageBus();
            log.Info("Message bus started");
        }

        static void TestCreateItem()
        {
            IMessageBus mb = _wc.Resolve<IMessageBus>();
            mb.NewMessage(new CreateCalendarItem {
                StartDate = DateTime.Now.AddHours(1),
                EndDate = DateTime.Now.AddHours(2),
                Subject = "nr 2",
                ImpersonateUser = "rafalg@cogit2.pl",
                CorrelationId = "3482349234",
                Body="<b>A tu inne <i>zdarzonko</i></b>",
                Recipients = new string[] { "rafalg@cogit2.pl", "atmosfera@cogit2.pl" }})
                .Send("sql://exch/MQ_ExchangeIntegration");
        }

        static void TestSendEmail()
        {
            IMessageBus mb = _wc.Resolve<IMessageBus>();
            mb.NewMessage(new SendEmailMessage
            {
                //ImpersonateUser = "rafalg@cogit2.pl",
                Subject = "powiadomienie 324",
                Recipients = new string[] { "rafalg@cogit2.pl" },
                CorrelationId = Guid.NewGuid().ToString(),
                Body = "<html><h1>MOJA WIADOMOSC</h1> <b>jest straszna</b></html>",
                SaveMode = SaveModes.SendOnly
            }).Send("sql://exch/MQ_ExchangeIntegration");
        }

        static void TestSubscribe()
        {
            IMessageBus mb = _wc.Resolve<IMessageBus>();
            mb.NewMessage(new AddSubscription
            {
                AccountName = "rafalg@cogit2.pl",
                EventTypes = new string[] {"Created", "Modified", "Deleted"},
                FolderIds = new string[] {"Calendar", "Tasks"},
                SubscriptionAlias = "rafalg_S1"
            }).Send("sql://exch/MQ_ExchangeIntegration");
        }
    }

    
}
