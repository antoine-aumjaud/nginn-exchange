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
    }

    public enum SaveModes
    {
        SaveAndSend = 0,
        SaveOnly = 1,
        SendOnly = 2
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
        public string Body { get; set; }
    }
}
