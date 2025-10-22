using System;
using Microsoft.Office.Tools.Ribbon;

namespace RedInk.OutlookAddIn.Ribbon
{
    partial class MailAssistantRibbon : RibbonBase
    {
        private RibbonTab _assistantTab;
        private RibbonGroup _assistantGroup;
        internal RibbonButton GenerateReplyButton;
        internal RibbonButton SummarizeButton;
        internal RibbonButton ReviewDraftButton;

        public MailAssistantRibbon()
            : base((Globals.Factory ?? throw new InvalidOperationException("Ribbon factory wurde nicht initialisiert.")).GetRibbonFactory())
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            _assistantTab = Factory.CreateRibbonTab();
            _assistantGroup = Factory.CreateRibbonGroup();
            GenerateReplyButton = Factory.CreateRibbonButton();
            SummarizeButton = Factory.CreateRibbonButton();
            ReviewDraftButton = Factory.CreateRibbonButton();

            _assistantTab.Label = "Red Ink";
            _assistantTab.Groups.Add(_assistantGroup);

            _assistantGroup.Label = "KI-Assistent";
            _assistantGroup.Items.Add(GenerateReplyButton);
            _assistantGroup.Items.Add(SummarizeButton);
            _assistantGroup.Items.Add(ReviewDraftButton);

            GenerateReplyButton.Label = "Antwort generieren";
            GenerateReplyButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
            GenerateReplyButton.ShowImage = true;
            GenerateReplyButton.Click += GenerateReplyButton_Click;

            SummarizeButton.Label = "Zusammenfassen";
            SummarizeButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
            SummarizeButton.ShowImage = true;
            SummarizeButton.Click += SummarizeButton_Click;

            ReviewDraftButton.Label = "Entwurf prüfen";
            ReviewDraftButton.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
            ReviewDraftButton.ShowImage = true;
            ReviewDraftButton.Click += ReviewDraftButton_Click;

            Tabs.Add(_assistantTab);
        }
    }
}
