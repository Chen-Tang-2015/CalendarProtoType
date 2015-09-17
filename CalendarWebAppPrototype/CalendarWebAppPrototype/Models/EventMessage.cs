using System.Collections.Generic;

namespace CalendarWebAppPrototype.Models
{
    public class EventMessage
    {
        [Newtonsoft.Json.JsonProperty("@odata.type")]
        public string Type
        {
            get; set;
        }

        public SenderClass Sender;

        public class EmailAddress
        {
            public string Address { get; set; }
            public string Name { get; set; }
        }

        public class SenderClass
        {
            public EmailAddress EmailAddress { get; set; }
        }

        public string Subject { get; set; }
        public string BodyPreview { get; set; }
        public string Id { get; set; } 
    }
}