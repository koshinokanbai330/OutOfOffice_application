using System;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutOfOfficeAddin
{
    public partial class ThisAddIn
    {
        /// <summary>
        /// Singleton reference so that UI classes (e.g. TaskPaneHost) can access
        /// the add-in without depending on the VSTO-generated Globals class.
        /// </summary>
        internal static ThisAddIn Current { get; private set; }

        private CustomTaskPane _taskPane;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Current = this;
            var host = new UI.TaskPaneHost();
            _taskPane = CustomTaskPanes.Add(host, "Out of Office");
            _taskPane.Width = 420;
            _taskPane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Current = null;
        }

        /// <summary>
        /// Returns the active Outlook Application instance.
        /// </summary>
        internal Outlook.Application OutlookApp => Application;

        #region VSTO generated code

        private void InternalStartup()
        {
            Startup += new EventHandler(ThisAddIn_Startup);
            Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
