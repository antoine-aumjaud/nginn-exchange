using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeIntegration.Service
{
    public class ExchangeConnect : IExchangeConnect
    {
        public string User { get; set; }
        public string Password { get; set; }
        public string EWSUrl { get; set; }

        public Microsoft.Exchange.WebServices.Data.ExchangeService Connect()
        {
            ExchangeService es = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            es.Url = new Uri(EWSUrl);
            if (string.IsNullOrEmpty(User))
            {
                es.UseDefaultCredentials = true;
            }
            else
            {
                es.Credentials = new System.Net.NetworkCredential(User, Password);
            }

            return es;
        }

        public Microsoft.Exchange.WebServices.Data.ExchangeService ConnectAndImpersonate(string userEmail)
        {
            var es = Connect();
            if (!string.IsNullOrEmpty(userEmail)) es.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, userEmail);
            return es;
        }
    }
}
