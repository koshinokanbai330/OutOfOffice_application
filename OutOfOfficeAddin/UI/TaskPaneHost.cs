using System.Windows.Forms;
using System.Windows.Forms.Integration;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutOfOfficeAddin.UI
{
    /// <summary>
    /// Windows Forms UserControl that hosts the WPF <see cref="TaskPaneView"/>
    /// via <see cref="ElementHost"/>.  This is required because VSTO Custom Task Panes
    /// expect a <see cref="System.Windows.Forms.Control"/>.
    /// </summary>
    public class TaskPaneHost : UserControl
    {
        private readonly ElementHost _host;
        private readonly TaskPaneView _wpfView;
        private readonly TaskPaneViewModel _viewModel;

        /// <summary>
        /// Initialises the host using the static <see cref="ThisAddIn.Current"/> reference
        /// set during add-in startup.
        /// </summary>
        public TaskPaneHost()
        {
            var outlookApp = ThisAddIn.Current.OutlookApp;

            _viewModel = new TaskPaneViewModel(outlookApp);
            _wpfView = new TaskPaneView { DataContext = _viewModel };

            _host = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = _wpfView,
            };

            Controls.Add(_host);
        }
    }
}
