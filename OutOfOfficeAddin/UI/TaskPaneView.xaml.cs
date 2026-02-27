using System.Windows.Controls;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutOfOfficeAddin.UI
{
    /// <summary>
    /// Code-behind for the WPF task pane UserControl.
    /// The DataContext is set by <see cref="TaskPaneHost"/>.
    /// </summary>
    public partial class TaskPaneView : UserControl
    {
        public TaskPaneView()
        {
            InitializeComponent();
        }
    }
}
