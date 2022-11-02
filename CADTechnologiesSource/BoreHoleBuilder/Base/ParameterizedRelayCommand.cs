using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Input;

namespace BoreExcavationBuilder.Base
{
    public class ParameterizedRelayCommand : ICommand
    {
        public event EventHandler CanExecuteChanged = (sender, e) => { };

        private Action<object> mAction;

        public ParameterizedRelayCommand(Action<object> action)
        {
            mAction = action;
        }

        //CanExecute is code that tells the button to be clikable.
        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            mAction(parameter);
        }
    }
}
