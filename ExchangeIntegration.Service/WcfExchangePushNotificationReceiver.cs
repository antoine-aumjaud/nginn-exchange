using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.EWSDataTypes;
using NLog;
using System.ServiceModel;

namespace ExchangeIntegration.Service
{
    /// <summary>
    /// Push notification receiver.
    /// </summary>
    [ServiceBehavior(InstanceContextMode = InstanceContextMode.Single)]
    public class WcfExchangePushNotificationReceiver : NotificationServicePortType
    {
        private Logger log = LogManager.GetCurrentClassLogger();

        public ExchangeEventReceiver EventReceiver { get; set; }

        public SendNotificationResponse SendNotification(SendNotificationRequest request)
        {
            log.Info("Push notification."); 
            SendNotificationResponse resp = new SendNotificationResponse();
            resp.SendNotificationResult = new SendNotificationResultType();
            resp.SendNotificationResult.SubscriptionStatus = SubscriptionStatusType.OK;

            // Get the push notifications.
            ResponseMessageType[] rmta = request.SendNotification.ResponseMessages.Items;


            foreach (ResponseMessageType rmt in rmta)
            {
                if (rmt.ResponseCode != ResponseCodeType.NoError)
                {
                    log.Info("Push notification error - discontinue the subscription (TODO)");
                    resp.SendNotificationResult.SubscriptionStatus = SubscriptionStatusType.Unsubscribe;
                    return resp;
                }

                // Cast to the correct response message type.
                SendNotificationResponseMessageType snrmt = rmt as SendNotificationResponseMessageType;

                var notification = snrmt.Notification;
                var subscriptionId = notification.SubscriptionId;
                var watermark = notification.PreviousWatermark;
                log.Info("Push notification subscription: {0}, watermark: {1}. Events: {2}", notification.SubscriptionId, notification.PreviousWatermark, notification.Items.Length);

                // Cast the notification event to the proper type
                foreach (BaseNotificationEventType bnet in notification.Items)
                {
                    log.Info("Event: {0}", bnet.ToString());
                }
                
                // If there are more events ready to be sent, then continue the subscription.
                if (notification.MoreEvents)
                {
                }
                //EventReceiver.ProcessNotification(notification);
            }

            return resp;
        }
    }
}
