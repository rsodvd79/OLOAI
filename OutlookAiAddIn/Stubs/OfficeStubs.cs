#if OFFICE_STUBS
using System;
using System.Collections.Generic;
using System.Windows.Forms;

// Minimal VSTO stubs to allow compilation without the Microsoft Office Tools runtime.

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
