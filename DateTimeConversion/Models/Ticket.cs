using System;

namespace DateTimeConversion.Models
{
    public class Ticket
    {
        public int TicketId { get; set; }
        public string CreatedDateString { get; set; }
        public string ModifiedDateString { get; set; }
        public string ClosedDateString { get; set; }
        public string DbScript { get; set; }

        public DateTime CreatedDateTimeUtc { get; set; }
        public DateTime ModifiedDateTimeUtc { get; set; }
        public DateTime CloseDateTimeUtc { get; set; }
    }
}
