using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExchangeIntegration.Interfaces;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeIntegration.Service
{
    public interface IExchangeIntegrationService
    {
        ItemCreated CreateCalendarItem(CreateCalendarItem item);
        ItemCreated SendEmail(SendEmailMessage msg);
        ItemCreated CreateTask(CreateTask msg);
        void DeleteItem(string itemId);
        void ReplyToMessage(ReplyToMessage msg);
    }

    public interface IExchangeConnect
    {
        ExchangeService Connect();
        ExchangeService ConnectAndImpersonate(string userEmail);
    }
}
