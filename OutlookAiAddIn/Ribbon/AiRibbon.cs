using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace OutlookAiAddIn.Ribbon
{
    [ComVisible(true)]
    public class AiRibbon : IRibbonExtensibility
    {
        private IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OutlookAiAddIn.Ribbon.AiRibbon.xml");
        }

        public void OnLoad(IRibbonUI ribbonUi)
        {
            _ribbon = ribbonUi;
        }

        public void OnTogglePane(IRibbonControl control)
        {
            Globals.ThisAddIn?.TogglePane();
        }

        public void OnSuggestReply(IRibbonControl control)
        {
            Globals.ThisAddIn?.TriggerMode(AiInteractionMode.SuggestedReply);
        }

        public void OnImproveDraft(IRibbonControl control)
        {
            Globals.ThisAddIn?.TriggerMode(AiInteractionMode.ImproveDraft);
        }

        public void OnProofread(IRibbonControl control)
        {
            Globals.ThisAddIn?.TriggerMode(AiInteractionMode.Proofread);
        }

        private static string GetResourceText(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var resourceStream = assembly.GetManifestResourceStream(resourceName))
            {
                if (resourceStream == null)
                {
                    throw new InvalidOperationException($"Resource '{resourceName}' non trovata");
                }

                using (var reader = new StreamReader(resourceStream))
                {
                    return reader.ReadToEnd();
                }
            }
        }
    }
}
