using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NGinnBPM.MessageBus;
using NLog;
using ExchangeIntegration.Interfaces;

namespace Testing
{
    public class ItemCreatedHandler : IMessageConsumer<ItemCreated>
    {
        private Logger log = LogManager.GetCurrentClassLogger();
        
        public IMessageBus MessageBus { get; set; }

        public void Handle(ItemCreated message)
        {
            log.Info("Item created: {0}", message.UniqueId);
            MessageBus.Reply(new ReplyToMessage
            {
                ReplyToItemId = message.UniqueId,
                ReplyAll = true,
                Body = "<html><h3>A to moja odpowiedź</h3></html>"
                
            });
        }
    }
}
