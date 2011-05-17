using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExchangeIntegration.Interfaces;

namespace ExchangeIntegration.Service
{
    public interface ISubscriptionManager
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="a"></param>
        /// <param name="receiverBusEndpoint"></param>
        /// <returns>Subscription ID</returns>
        string AddSubscription(AddSubscription a, string receiverBusEndpoint);
    }
}
