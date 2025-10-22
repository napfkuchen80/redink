using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RedInk.OutlookAddIn.Models
{
    public enum MailParticipantRole
    {
        Sender,
        To,
        Cc,
        Bcc
    }

    public sealed class MailParticipant
    {
        public string DisplayName { get; set; }

        public string Address { get; set; }

        public MailParticipantRole Role { get; set; }

        public static MailParticipant FromSender(Outlook.MailItem mailItem)
        {
            if (mailItem == null)
            {
                return null;
            }

            return new MailParticipant
            {
                DisplayName = mailItem.SenderName,
                Address = mailItem.SenderEmailAddress,
                Role = MailParticipantRole.Sender
            };
        }

        public static IList<MailParticipant> FromRecipients(Outlook.Recipients recipients, Outlook.OlMailRecipientType type)
        {
            var result = new List<MailParticipant>();
            if (recipients == null)
            {
                return result;
            }

            foreach (Outlook.Recipient recipient in recipients)
            {
                if (recipient == null)
                {
                    continue;
                }

                if (recipient.Type != (int)type)
                {
                    continue;
                }

                result.Add(new MailParticipant
                {
                    DisplayName = recipient.Name,
                    Address = GetSmtpAddress(recipient),
                    Role = ConvertRole(type)
                });
            }

            return result;
        }

        private static MailParticipantRole ConvertRole(Outlook.OlMailRecipientType type)
        {
            switch (type)
            {
                case Outlook.OlMailRecipientType.olTo:
                    return MailParticipantRole.To;
                case Outlook.OlMailRecipientType.olCC:
                    return MailParticipantRole.Cc;
                case Outlook.OlMailRecipientType.olBCC:
                    return MailParticipantRole.Bcc;
                default:
                    return MailParticipantRole.To;
            }
        }

        private static string GetSmtpAddress(Outlook.Recipient recipient)
        {
            try
            {
                var addressEntry = recipient.AddressEntry;
                if (addressEntry == null)
                {
                    return recipient.Address;
                }

                if (addressEntry.Type == "EX")
                {
                    var exchUser = addressEntry.GetExchangeUser();
                    if (exchUser != null)
                    {
                        return exchUser.PrimarySmtpAddress;
                    }

                    var exchDistList = addressEntry.GetExchangeDistributionList();
                    if (exchDistList != null)
                    {
                        return exchDistList.PrimarySmtpAddress;
                    }
                }

                return addressEntry.Address;
            }
            catch
            {
                return recipient.Address;
            }
        }
    }
}
