using System.Collections.Generic;
using RedInk.OutlookAddIn.Models;

namespace RedInk.OutlookAddIn.Services
{
    public sealed class LlmRequestPayload
    {
        public string Intent { get; set; }

        public MailItemContext Mail { get; set; }

        public IDictionary<string, string> Metadata { get; set; } = new Dictionary<string, string>();
    }

    public sealed class LlmResponsePayload
    {
        public string Content { get; set; }

        public string Title { get; set; }

        public string SuggestedSubject { get; set; }
    }
}
