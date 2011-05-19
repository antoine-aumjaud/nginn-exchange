using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NGinnBPM.MessageBus;
using ExchangeIntegration.Interfaces;

namespace ExchangeIntegration.Service
{
    public class MessageHandler : 
        IMessageConsumer<CreateCalendarItem>,
        IMessageConsumer<SendEmailMessage>,
        IMessageConsumer<AddSubscription>,
        IMessageConsumer<ReplyToMessage>
    {
        public IMessageBus MessageBus { get; set; }
        public IExchangeIntegrationService Exchange { get; set; }
        public ISubscriptionManager SubscriptionManager { get; set; }

        public void Handle(CreateCalendarItem message)
        {
            var ret = Exchange.CreateCalendarItem(message);
            MessageBus.Reply(ret);
        }

        public void Handle(SendEmailMessage message)
        {
            var ret = Exchange.SendEmail(message);
            MessageBus.Reply(ret);
        }

        public void Handle(AddSubscription message)
        {
            var id = SubscriptionManager.AddSubscription(message, MessageBus.CurrentMessageInfo.Sender);
            SubscriptionCreated sc = new SubscriptionCreated
            {
                SubscriptionAlias = message.SubscriptionAlias,
                SubscriptionId = id
            };
            MessageBus.Reply(sc);
        }

        public void Handle(ReplyToMessage message)
        {
            this.Exchange.ReplyToMessage(message);
        }
    }
}
