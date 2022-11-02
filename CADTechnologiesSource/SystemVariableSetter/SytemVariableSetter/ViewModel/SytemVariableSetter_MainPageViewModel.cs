using CADTechnologiesSource.All.Base;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace SystemVariableSetter.ViewModel
{
    public class SystemVariableSetter_MainPageViewModel
    {

        #region Singleton
        public static SystemVariableSetter_MainPageViewModel Instance = new SystemVariableSetter_MainPageViewModel();
        #endregion

        #region Public Properties

        public ObservableCollection<string> Drawings { get; set; }
        public string SystemVariable { get; set; }
        public string SystemVariableValue { get; set; }

        #endregion

        #region Public Commands

        public ICommand AddDrawingCommand { get; set; }
        public ICommand RemoveDrawingCommand { get; set; }
        public ICommand RemoveAllDrawingsCommand { get; set; }
        public ICommand OpenDrawingCommand { get; set; }
        public ICommand SetSystemVariableCommand { get; set; }
        public ICommand SaveListCommand { get; set; }
        public ICommand LoadListCommand { get; set; }

        #endregion

        public SystemVariableSetter_MainPageViewModel()
        {
            if (Drawings == null || Drawings.Count != 0)
            {
                Drawings = new ObservableCollection<string>();
                Drawings.Clear();
            }

            #region Commands

            AddDrawingCommand = new RelayCommand(() => AddDrawingsClick());
            RemoveDrawingCommand = new RelayCommand(() => RemoveDrawingClick());
            RemoveAllDrawingsCommand = new RelayCommand(() => RemoveAllDrawingsClick());
            OpenDrawingCommand = new RelayCommand(() => OpenDrawingClick());
            SetSystemVariableCommand = new RelayCommand(() => SetSystemVariableCommandClick());
            SaveListCommand = new RelayCommand(() => SaveListCommandClick());
            LoadListCommand = new RelayCommand(() => LoadListCommandClick());

            #endregion
        }

        private void AddDrawingsClick()
        {
            throw new NotImplementedException();
        }

        private void RemoveDrawingClick()
        {
            throw new NotImplementedException();
        }

        private void RemoveAllDrawingsClick()
        {
            throw new NotImplementedException();
        }

        private void OpenDrawingClick()
        {
            throw new NotImplementedException();
        }

        private void SetSystemVariableCommandClick()
        {
            throw new NotImplementedException();
        }

        private void SaveListCommandClick()
        {
            throw new NotImplementedException();
        }

        private void LoadListCommandClick()
        {
            throw new NotImplementedException();
        }
    }
}
