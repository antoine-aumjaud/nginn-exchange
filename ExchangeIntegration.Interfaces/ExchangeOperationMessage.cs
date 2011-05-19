using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeIntegration.Interfaces
{
    /// <summary>
    /// Klasa bazowa dla operacji na exchange
    /// </summary>
    public class ExchangeOperationMessage
    {
        /// <summary>
        /// Użytkownik w imieniu którego wykonujemy operację
        /// </summary>
        public string ImpersonateUser { get; set; }
        public string CorrelationId { get; set; }

        public override string ToString()
        {
            return string.Format("{0}:{1}", GetType().Name, CorrelationId);
        }
    }

    public enum SaveModes
    {
        /// <summary>
        /// Save message only.
        /// </summary>
        SaveOnly = 1,
        /// <summary>
        /// Only send a message. Will not return ItemId of the message.
        /// </summary>
        SendOnly = 2,
        /// <summary>
        /// Will send a message and save a copy, but will not
        /// return ItemId of newly sent message
        /// </summary>
        SaveAndSend = 3,
        /// <summary>
        /// This will save message, send it, and return 
        /// its ItemId
        /// </summary>
        SaveSendAndReturnId = 4
    }

    public class CreateItemMessage : ExchangeOperationMessage
    {
        public CreateItemMessage()
        {
            SaveMode = SaveModes.SaveAndSend;
        }

        public SaveModes SaveMode { get; set; }
        public string WellKnownFolderName { get; set; }
        public string[] Recipients { get; set; }
        public string Subject { get; set; }
        /// <summary>
        /// Message body text. For HTML messages the body must start with &lt;html&gt; tag.
        /// </summary>
        public string Body { get; set; }
        public DateTime? ReminderDate { get; set; }
        /// <summary>
        /// List of attachment file paths
        /// </summary>
        public string[] AttachmentFiles { get; set; }
        /// <summary>
        /// Delete attachment files after sending
        /// the message
        /// </summary>
        public bool DeleteAttachments { get; set; }
    }
}
