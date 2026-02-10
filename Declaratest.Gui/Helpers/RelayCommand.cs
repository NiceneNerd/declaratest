using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Input;

namespace Declaratest.Gui.Helpers;

public class RelayCommand : ICommand
{
    private readonly Func<object?, Task>? _executeAsync;
    private readonly Action<object?>? _execute;
    private readonly Predicate<object?>? _canExecute;
    private readonly Action<Exception>? _onError;

    public RelayCommand(Action execute)
        : this(_ => execute())
    {
    }

    public RelayCommand(Action<object?> execute, Predicate<object?>? canExecute = null)
    {
        _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        _canExecute = canExecute;
    }

    public RelayCommand(Func<object?, Task> executeAsync, Predicate<object?>? canExecute = null, Action<Exception>? onError = null)
    {
        _executeAsync = executeAsync ?? throw new ArgumentNullException(nameof(executeAsync));
        _canExecute = canExecute;
        _onError = onError;
    }

    public bool CanExecute(object? parameter) => _canExecute?.Invoke(parameter) ?? true;

    public void Execute(object? parameter)
    {
        if (_executeAsync != null)
        {
            _ = ExecuteAsync(parameter);
        }
        else
        {
            _execute?.Invoke(parameter);
        }
    }

    private async Task ExecuteAsync(object? parameter)
    {
        try
        {
            await _executeAsync!.Invoke(parameter);
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex);
            _onError?.Invoke(ex);
        }
    }

    public event EventHandler? CanExecuteChanged
    {
        add => CommandManager.RequerySuggested += value;
        remove => CommandManager.RequerySuggested -= value;
    }
}
