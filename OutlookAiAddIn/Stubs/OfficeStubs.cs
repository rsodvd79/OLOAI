#define OFFICE_STUBS
#if OFFICE_STUBS
using System;
using System.Collections.Generic;
using System.Windows.Forms;

// Minimal VSTO/Office/Interop stubs to allow compilation without the Microsoft Office Tools runtime.

namespace Microsoft.Office.Core
{
    public interface IRibbonExtensibility
    {
        string GetCustomUI(string ribbonID);
    }

    public interface IRibbonUI { void Invalidate(); }

    public interface IRibbonControl { string Id { get; } }
}

namespace Microsoft.Office.Tools
{
    public class CustomTaskPane
    {
        internal CustomTaskPane(UserControl control, string title)
        {
            Control = control;
            Title = title;
        }

        public UserControl Control { get; }
        public string Title { get; }
        public bool Visible { get; set; }
        public int Width { get; set; }
    }

    public class CustomTaskPaneCollection : List<CustomTaskPane>
    {
        public CustomTaskPane Add(UserControl control, string title)
        {
            var pane = new CustomTaskPane(control, title);
            Add(pane);
            return pane;
        }

        public new void Remove(CustomTaskPane pane)
        {
            base.Remove(pane);
        }
    }
}

namespace Microsoft.Office.Interop.Word
{
    public class Application
    {
        public Selection Selection { get; } = new Selection();
    }

    public class Document
    {
        public Application Application { get; } = new Application();
        public string BodyText { get; set; } = string.Empty;
    }

    public class Selection
    {
        public string Buffer { get; private set; } = string.Empty;
        public void TypeText(string text)
        {
            Buffer += text;
        }
    }
}

namespace Microsoft.Office.Interop.Outlook
{
    using Word = Microsoft.Office.Interop.Word;

    public delegate void ApplicationItemSendEventHandler(object item, ref bool cancel);

    public enum OlEditorType { olEditorWord = 1, olEditorText = 2 }
    public enum OlBodyFormat { olFormatUnspecified = 0, olFormatPlain, olFormatHTML, olFormatRichText }

    public class Application
    {
        private readonly Inspector _inspector = new Inspector();
        private readonly Explorer _explorer = new Explorer();

        public Inspector ActiveInspector() => _inspector;
        public Explorer ActiveExplorer() => _explorer;

        public event ApplicationItemSendEventHandler ItemSend;
        public void RaiseItemSend(object item)
        {
            bool cancel = false;
            ItemSend?.Invoke(item, ref cancel);
        }
    }

    public class Inspector
    {
        public object CurrentItem { get; set; } = new MailItem();
        public OlEditorType EditorType { get; set; } = OlEditorType.olEditorWord;
        public object WordEditor { get; set; } = new Word.Document();
        public void Display() { }
    }

    public class Explorer
    {
        public Selection Selection { get; } = new Selection();
    }

    public class Selection : List<object>
    {
        public int Count => base.Count;
        public object this[int index] => base[index - 1];
    }

    public class MailItem
    {
        public string Body { get; set; } = string.Empty;
        public string HTMLBody { get; set; } = string.Empty;
        public OlBodyFormat BodyFormat { get; set; } = OlBodyFormat.olFormatHTML;
        public bool Sent { get; set; }

        public void Save() { }
        public void Display() { }
        public MailItem Reply() => new MailItem();
    }
}

namespace Microsoft.Office.Tools.Outlook
{
    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.Outlook;
    using Microsoft.Office.Tools;

    public abstract class OutlookAddIn
    {
        protected OutlookAddIn()
        {
            Application = new Application();
            CustomTaskPanes = new CustomTaskPaneCollection();
        }

        public Application Application { get; }

        public CustomTaskPaneCollection CustomTaskPanes { get; }

        public event EventHandler Startup;

        public event EventHandler Shutdown;

        protected virtual IRibbonExtensibility CreateRibbonExtensibilityObject() => null;

        protected void OnStartup(EventArgs e) => Startup?.Invoke(this, e);

        protected void OnShutdown(EventArgs e) => Shutdown?.Invoke(this, e);
    }
}
#endif
