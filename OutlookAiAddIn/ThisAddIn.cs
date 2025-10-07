using System;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;

namespace OutlookAiAddIn
{
    public enum AiInteractionMode
    {
        SuggestedReply,
        ImproveDraft,
        Proofread
    }

    public partial class ThisAddIn
    {
        internal CustomTaskPane AiTaskPaneHost { get; private set; }
        internal UI.AiTaskPane AiTaskPaneControl { get; private set; }
        internal Services.OpenAIService OpenAIService { get; private set; }
        internal Services.OutlookContextService OutlookContextService { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Globals.ThisAddIn = this;

            OpenAIService = new Services.OpenAIService();
            OutlookContextService = new Services.OutlookContextService(this.Application);

            AiTaskPaneControl = new UI.AiTaskPane(OpenAIService, OutlookContextService);
            AiTaskPaneHost = this.CustomTaskPanes.Add(AiTaskPaneControl, "Assistente AI");
            AiTaskPaneHost.Visible = false;
            AiTaskPaneHost.Width = 420;

            this.Application.ItemSend += Application_ItemSend;
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            AiTaskPaneControl?.HandleItemSend(item, ref cancel);
        }

        internal void ShowTaskPane(bool visible)
        {
            if (AiTaskPaneHost == null)
            {
                return;
            }

            AiTaskPaneHost.Visible = visible;
            if (visible)
            {
                AiTaskPaneControl?.LoadContext();
            }
        }

        internal void TogglePane()
        {
            if (AiTaskPaneHost == null)
            {
                return;
            }

            ShowTaskPane(!AiTaskPaneHost.Visible);
        }

        internal void TriggerMode(AiInteractionMode mode)
        {
            ShowTaskPane(true);
            AiTaskPaneControl?.TriggerMode(mode);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            this.Application.ItemSend -= Application_ItemSend;

            if (AiTaskPaneHost != null)
            {
                this.CustomTaskPanes.Remove(AiTaskPaneHost);
                AiTaskPaneHost = null;
            }

            AiTaskPaneControl?.Dispose();
            OpenAIService?.Dispose();
            OpenAIService = null;
            AiTaskPaneControl = null;
            OutlookContextService = null;
            Globals.ThisAddIn = null;
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon.AiRibbon();
        }
    }
}
