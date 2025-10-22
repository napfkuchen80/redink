using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Tools;
using Outlook = Microsoft.Office.Interop.Outlook;
using RedInk.OutlookAddIn.Configuration;
using RedInk.OutlookAddIn.Models;
using RedInk.OutlookAddIn.Services;
using RedInk.OutlookAddIn.TaskPane;

namespace RedInk.OutlookAddIn
{
    public sealed partial class ThisAddIn
    {
        private CustomTaskPane _assistantPane;
        private AssistantTaskPaneControl _assistantControl;
        private LlmGatewayClient _gatewayClient;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Globals.ThisAddIn = this;
            Globals.Factory ??= Factory;

            var options = AddInConfiguration.Load();
            _gatewayClient = new LlmGatewayClient(options);

            _assistantControl = new AssistantTaskPaneControl();
            _assistantPane = CustomTaskPanes.Add(_assistantControl, "Red Ink Assistent");
            _assistantPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            _assistantPane.Width = 420;
            _assistantPane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _assistantPane?.Dispose();
            _gatewayClient?.Dispose();
        }

        internal async Task ExecuteIntentAsync(LlmIntent intent, CancellationToken cancellationToken = default)
        {
            if (_assistantControl == null)
            {
                return;
            }

            var intentLabel = intent switch
            {
                LlmIntent.GenerateReply => "Antwort generieren",
                LlmIntent.Summarize => "Zusammenfassen",
                LlmIntent.ReviewDraft => "Entwurf prüfen",
                _ => intent.ToString()
            };

            _assistantControl.ShowStatus($"{intentLabel} – wird vorbereitet...");

            var context = GetCurrentMailItem();
            if (context == null)
            {
                _assistantControl.ShowStatus("Bitte wählen Sie eine E-Mail aus oder öffnen Sie einen Entwurf.");
                return;
            }

            try
            {
                var response = await _gatewayClient.ExecuteIntentAsync(intent, context, cancellationToken).ConfigureAwait(false);
                _assistantControl.DisplayResponse(intentLabel, response.Content);
            }
            catch (OperationCanceledException)
            {
                _assistantControl.ShowStatus("Anfrage wurde abgebrochen.");
            }
            catch (Exception ex)
            {
                _assistantControl.ShowStatus($"Fehler: {ex.Message}");
            }
        }

        internal MailItemContext GetCurrentMailItem()
        {
            Outlook.MailItem mailItem = null;

            var inspector = Application?.ActiveInspector();
            if (inspector?.CurrentItem is Outlook.MailItem inspectorMail)
            {
                mailItem = inspectorMail;
            }
            else
            {
                var explorer = Application?.ActiveExplorer();
                var selection = explorer?.Selection;
                if (selection != null && selection.Count > 0)
                {
                    mailItem = selection[1] as Outlook.MailItem;
                }
            }

            if (mailItem == null)
            {
                return null;
            }

            return new MailItemContext
            {
                EntryId = mailItem.EntryID,
                Subject = mailItem.Subject,
                Body = mailItem.Body,
                NormalizedBody = TryGetNormalizedBody(mailItem),
                To = MailParticipant.FromRecipients(mailItem.Recipients, Outlook.OlMailRecipientType.olTo),
                Cc = MailParticipant.FromRecipients(mailItem.Recipients, Outlook.OlMailRecipientType.olCC),
                Bcc = MailParticipant.FromRecipients(mailItem.Recipients, Outlook.OlMailRecipientType.olBCC),
                Sender = MailParticipant.FromSender(mailItem),
                HasAttachments = mailItem.Attachments?.Count > 0,
                ReceivedTime = mailItem.ReceivedTime,
                SentOn = mailItem.SentOn,
                ConversationId = mailItem.ConversationID,
                Language = mailItem.Language
            };
        }

        private static string TryGetNormalizedBody(Outlook.MailItem mailItem)
        {
            return mailItem?.Body;
        }
    }
}
