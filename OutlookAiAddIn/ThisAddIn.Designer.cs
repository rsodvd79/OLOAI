using System;
using Microsoft.Office.Tools.Outlook;

namespace OutlookAiAddIn
{
    public partial class ThisAddIn : OutlookAddIn
    {
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }
}
