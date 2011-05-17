using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeIntegration.Interfaces
{
    public class CreateTask : CreateItemMessage
    {
        public DateTime? Deadline { get; set; }
    }
}
