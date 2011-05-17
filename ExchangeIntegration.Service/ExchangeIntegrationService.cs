using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NLog;
using Microsoft.Exchange.WebServices.Data;
using NGinnBPM.MessageBus;
using ExchangeIntegration.Interfaces;
using System.IO;
using System.Threading;

namespace ExchangeIntegration.Service
{
    public class ExchangeIntegrationService : IExchangeIntegrationService, IExchangeConnect
    {
        public string EWSUrl { get; set; }
        public IMessageBus MessageBus { get; set; }
        public string EWSUser { get; set; }
        public string EWSPassword { get; set; }

        private static Logger log = LogManager.GetCurrentClassLogger();

        public ExchangeIntegrationService()
        {
        }


        public ExchangeService Connect()
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

       

        public ItemCreated CreateCalendarItem(CreateCalendarItem message)
        {
            ExchangeService es = Connect();
            es.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, message.ImpersonateUser);
            Appointment app = new Appointment(es);
            log.Info("Created appointment, item Id: {0}", app.Id);
            foreach (string r in message.Recipients)
                app.RequiredAttendees.Add(r);
            
            app.Subject = message.Subject;
            app.Body = new MessageBody(BodyType.HTML, message.Body);
            app.Start = message.StartDate;
            app.End = message.EndDate;
            
            app.Save(SendInvitationsMode.SendToAllAndSaveCopy);
            
            log.Info("Saved calendar item: {0}", app.Id);
            return new ItemCreated
            {
                UniqueId = app.Id.UniqueId,
                InternetMessageId = null,
                CorrelationId = message.CorrelationId
            };
        }

        public void CreateTask()
        {
            ExchangeService es = Connect();
            Task tsk = new Task(es);
            //tsk.AllowedResponseActions = ResponseActions.Accept | ResponseActions.Decline;
            
        }


        public ItemCreated SendEmail(SendEmailMessage msg)
        {
            ExchangeService es = Connect();
            if (!string.IsNullOrEmpty(msg.ImpersonateUser)) es.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, msg.ImpersonateUser);

            EmailMessage em = new EmailMessage(es);
            em.Subject = msg.Subject;
            foreach(string s in msg.Recipients)
                em.ToRecipients.Add(s);
            em.Body = msg.Body;
            ExtendedPropertyDefinition epd = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "X-CorrelationId", MapiPropertyType.String);
            em.SetExtendedProperty(epd, msg.CorrelationId);
            em.IsDeliveryReceiptRequested = true;
            
            em.SendAndSaveCopy();
            log.Info("Sent email to {2} ({3}). Item id: {0}. Iternet Id: {1}", em.Id, em.InternetMessageId, em.DisplayTo, em.Subject);
            return new ItemCreated
            {
                UniqueId = em.Id.UniqueId,
                InternetMessageId = em.InternetMessageId,
                CorrelationId = msg.CorrelationId
            };
        }


        public ItemCreated CreateTask(CreateTask msg)
        {
            ExchangeService es = Connect();
            if (!string.IsNullOrEmpty(msg.ImpersonateUser)) es.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, msg.ImpersonateUser);

            Task tsk = new Task(es);
            tsk.Body = msg.Body;
            tsk.Subject = msg.Subject;
            tsk.DueDate = msg.Deadline;
            tsk.Save();
            return new ItemCreated
            {
                UniqueId = tsk.Id.UniqueId,
                InternetMessageId = null,
                CorrelationId = msg.CorrelationId
            };
        }


        public void DeleteItem(string itemId)
        {
            ExchangeService es = Connect();
            //if (!string.IsNullOrEmpty(msg.ImpersonateUser)) es.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, msg.ImpersonateUser);
            //es.DeleteItems(new ItemId[] { new ItemId(itemId) }, DeleteMode.SoftDelete, SendCancellationsMode.SendToNone);

            throw new NotImplementedException();
        }

        

        public ExchangeService ConnectAndImpersonate(string userEmail)
        {
            var es = Connect();
            es.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, userEmail);
            return es;
        }
    }
}
