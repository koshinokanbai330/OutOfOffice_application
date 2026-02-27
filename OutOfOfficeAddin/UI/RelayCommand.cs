using System.Windows.Input;

namespace OutOfOfficeAddin.UI
{
    /// <summary>
    /// Simple ICommand implementation backed by delegates (standard MVVM helper).
    /// </summary>
    public sealed class RelayCommand : ICommand
    {
        private readonly System.Action _execute;
        private readonly System.Func<bool> _canExecute;

        public RelayCommand(System.Action execute, System.Func<bool> canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public event System.EventHandler CanExecuteChanged
        {
            add { System.Windows.Input.CommandManager.RequerySuggested += value; }
            remove { System.Windows.Input.CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object parameter) => _canExecute == null || _canExecute();

        public void Execute(object parameter) => _execute();

        /// <summary>Raises CanExecuteChanged manually if needed.</summary>
        public void RaiseCanExecuteChanged()
            => System.Windows.Input.CommandManager.InvalidateRequerySuggested();
    }
}
