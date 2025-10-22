using System;
using System.Collections.Generic;

namespace RedInk.OutlookAddIn.Models
{
    public sealed class MailItemContext
    {
        public string EntryId { get; set; }

        public string Subject { get; set; }

        public string Body { get; set; }

        public string NormalizedBody { get; set; }

        public MailParticipant Sender { get; set; }

        public IList<MailParticipant> To { get; set; } = new List<MailParticipant>();

        public IList<MailParticipant> Cc { get; set; } = new List<MailParticipant>();

        public IList<MailParticipant> Bcc { get; set; } = new List<MailParticipant>();

        public bool HasAttachments { get; set; }

        public DateTime? ReceivedTime { get; set; }

        public DateTime? SentOn { get; set; }

        public string ConversationId { get; set; }

        public string Language { get; set; }
    }
}
