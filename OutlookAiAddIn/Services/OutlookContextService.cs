using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace OutlookAiAddIn.Services
{
    internal sealed class OutlookContextService
    {
        private readonly Outlook.Application _application;

        public OutlookContextService(Outlook.Application application)
        {
            _application = application;
        }

        public Outlook.MailItem GetActiveMailItem()
        {
            var inspector = _application.ActiveInspector();
            if (inspector?.CurrentItem is Outlook.MailItem inspectorMail)
            {
                return inspectorMail;
            }

            var explorer = _application.ActiveExplorer();
            if (explorer?.Selection != null && explorer.Selection.Count > 0)
            {
                return explorer.Selection[1] as Outlook.MailItem;
            }

            return null;
        }

        public string GetActiveMailBody()
        {
            var mailItem = GetActiveMailItem();
            return mailItem == null ? string.Empty : GetMailBody(mailItem, preferHtml: false);
        }

        public string GetMailBody(Outlook.MailItem mailItem, bool preferHtml)
        {
            if (mailItem == null)
            {
                return string.Empty;
            }

            if (preferHtml && mailItem.BodyFormat == Outlook.OlBodyFormat.olFormatHTML)
            {
                return mailItem.HTMLBody ?? string.Empty;
            }

            return mailItem.Body ?? string.Empty;
        }

        public bool TryInsertAtCursor(string text)
        {
            var inspector = _application.ActiveInspector();
            if (inspector == null || string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            if (inspector.EditorType == Outlook.OlEditorType.olEditorWord && inspector.WordEditor is Word.Document document)
            {
                var selection = document.Application.Selection;
                selection.TypeText(text);
                return true;
            }

            if (inspector.CurrentItem is Outlook.MailItem mailItem)
            {
                mailItem.Body += Environment.NewLine + text;
                mailItem.Save();
                return true;
            }

            return false;
        }

        public bool TryOverwriteDraft(string text, bool html)
        {
            var mailItem = GetActiveMailItem();
            if (mailItem == null)
            {
                return false;
            }

            if (html && mailItem.BodyFormat == Outlook.OlBodyFormat.olFormatHTML)
            {
                mailItem.HTMLBody = text;
            }
            else
            {
                mailItem.Body = text;
            }

            mailItem.Save();
            return true;
        }

        public Outlook.MailItem EnsureReplyWindow()
        {
            var mailItem = GetActiveMailItem();
            if (mailItem == null)
            {
                return null;
            }

            if (IsDraftEditable(mailItem))
            {
                mailItem.Display();
                return mailItem;
            }

            var reply = mailItem.Reply();
            reply.Display();
            return reply;
        }

        public static bool IsDraftEditable(Outlook.MailItem mailItem)
        {
            return mailItem != null && !mailItem.Sent;
        }
    }
}
