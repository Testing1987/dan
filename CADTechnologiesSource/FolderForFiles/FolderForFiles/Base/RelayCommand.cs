using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Input;

namespace FolderForFiles.Base
{
    public class RelayCommand : ICommand
    {
        public event EventHandler CanExecuteChanged = (sender, e) => { };

        private Action mAction;

        public RelayCommand(Action action)
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
            mAction();
        }
    }
}
