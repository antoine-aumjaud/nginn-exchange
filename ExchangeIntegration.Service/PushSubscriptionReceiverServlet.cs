using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NLog;
using NGinnBPM.MessageBus.Impl.HttpService;
using System.IO;
using System.Xml;
using System.Runtime.Serialization;
using System.Xml.Xsl;
using System.IO;

namespace ExchangeIntegration.Service
{
    [UrlPattern(@"^/push/")]
    public class PushSubscriptionReceiverServlet : NGinnBPM.MessageBus.Impl.HttpService.Servlet
    {
        public PushSubscriptionManager SubscriptionManager { get; set; }
        private Logger log = LogManager.GetCurrentClassLogger();
        private static XslCompiledTransform _tran;
        private DataContractSerializer _ser = new DataContractSerializer(typeof(SubscriptionEventNotification));
        private XslCompiledTransform GetNotificationTransform()
        {
            if (_tran == null)
            {
                _tran = new XslCompiledTransform();
                using (Stream stm = typeof(PushSubscriptionReceiverServlet).Assembly.GetManifestResourceStream("ExchangeIntegration.Service.PushNotificationTransform.xslt"))
                {
                    _tran.Load(XmlReader.Create(stm));
                }
            }
            return _tran;
        }

        protected override void OnRequest(System.IO.TextReader input, System.IO.TextWriter output, RequestContext ctx)
        {
            string inputXml = input.ReadToEnd();
            log.Info("input {1}: {0}", inputXml, ctx.Request.RawUrl);
            try
            {
                ctx.Response.ContentEncoding = Encoding.GetEncoding("utf-8");
                ctx.Response.AddHeader("Content-Type", "text/xml; charset=utf-8");

                if (ctx.Request.HttpMethod == "GET")
                {
                    SubscriptionEventNotification notif = new SubscriptionEventNotification();
                    notif.SubscriptionId = "asiodfioasodfasodf";
                    notif.Events.Add(new BaseExchangeEvent { ItemId = "89234829349234", EventType = ExchangeEventType.Modified, FolderId = "89923492349234", Watermark = "9234l23ll234234", TimeStamp = DateTime.Now });
                    var sw1 = new StringWriter();
                    var xw1 = XmlWriter.Create(sw1);
                    _ser.WriteObject(xw1, notif);
                    xw1.Flush();
                    output.Write(sw1.ToString());
                    return;
                }

                StringWriter sw = new StringWriter();
                XmlWriterSettings s = new XmlWriterSettings();
                s.Encoding = Encoding.UTF8;
                s.OmitXmlDeclaration = true;
                XmlWriter xw = XmlWriter.Create(sw, s);
                GetNotificationTransform().Transform(XmlReader.Create(new StringReader(inputXml)), xw);
                xw.Flush();
                log.Info("deserialize: {0}", sw.ToString());

                SubscriptionEventNotification n = (SubscriptionEventNotification) _ser.ReadObject(XmlReader.Create(new StringReader(sw.ToString())));
                n.TargetEndpoint = ctx.Request.QueryString["endp"];
                bool b = SubscriptionManager.HandleSubscriptionNotification(n);
                if (!b)
                    log.Info("Unsubscribing subscription {0}", n.SubscriptionId);
                string retXml = string.Format("<s:Envelope xmlns:s=\"http://schemas.xmlsoap.org/soap/envelope/\"><s:Body xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\"><SendNotificationResult xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\"><SubscriptionStatus>{0}</SubscriptionStatus></SendNotificationResult></s:Body></s:Envelope>", b ? "OK" : "Unsubscribe");
                output.Write(retXml);
            }
            catch (Exception ex)
            {
                log.Error("Error processing request: {0}. Input: {1}", ex, inputXml);
                throw;
            }
        }


    }
}
