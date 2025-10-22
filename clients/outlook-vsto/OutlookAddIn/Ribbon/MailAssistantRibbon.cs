using System.Threading.Tasks;
using Microsoft.Office.Tools.Ribbon;
using RedInk.OutlookAddIn.Services;

namespace RedInk.OutlookAddIn.Ribbon
{
    public partial class MailAssistantRibbon
    {
        private async void GenerateReplyButton_Click(object sender, RibbonControlEventArgs e)
        {
            await ExecuteAsync(LlmIntent.GenerateReply);
        }

        private async void SummarizeButton_Click(object sender, RibbonControlEventArgs e)
        {
            await ExecuteAsync(LlmIntent.Summarize);
        }

        private async void ReviewDraftButton_Click(object sender, RibbonControlEventArgs e)
        {
            await ExecuteAsync(LlmIntent.ReviewDraft);
        }

        private static Task ExecuteAsync(LlmIntent intent)
        {
            return Globals.ThisAddIn?.ExecuteIntentAsync(intent) ?? Task.CompletedTask;
        }
    }

    partial class ThisRibbonCollection
    {
        internal MailAssistantRibbon MailAssistantRibbon => GetRibbon<MailAssistantRibbon>();
    }
}
