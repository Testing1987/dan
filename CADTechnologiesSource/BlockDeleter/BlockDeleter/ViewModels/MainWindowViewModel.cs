using BlockDeleter.Base;
using BlockDeleter.CoreLogic;
using CADTechnologiesSource.All.AutoCADHelpers;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;

namespace BlockDeleter.ViewModels
{
    public class MainWindowViewModel : BaseViewModel
    {
        #region Singleton
        public static MainWindowViewModel Instance = new MainWindowViewModel();
        #endregion

        #region Public Properties

        public string BlockName { get; set; }
        public string SelectedDrawing { get; set; }
        public ObservableCollection<string> AddedDrawings { get; set; }

        #endregion

        #region Public Commands

        public ICommand AddDrawingCommand { get; set; }
        public ICommand RemoveDrawingCommand { get; set; }
        public ICommand RemoveAllDrawingsCommand { get; set; }
        public ICommand OpenDrawingCommand { get; set; }
        public ICommand DeleteBlocksCommand { get; set; }

        #endregion

        #region Constructor

        public MainWindowViewModel()
        {
            if (AddedDrawings == null)
            {
                AddedDrawings = new ObservableCollection<string>();
            }

            AddDrawingCommand = new RelayCommand(() => AddDrawingsClick());
            RemoveDrawingCommand = new RelayCommand(() => RemoveDrawingClick());
            RemoveAllDrawingsCommand = new RelayCommand(() => RemoveAllDrawingsClick());
            OpenDrawingCommand = new RelayCommand(() => OpenDrawingClick());
            DeleteBlocksCommand = new RelayCommand(() => DeleteBlocksCommandClick());
        }

        private void AddDrawingsClick()
        {
            CADAccess ca = new CADAccess();

            foreach (var result in ca.AddtoDrawingList())
            {
                //make sure the added string isn't null and isn't already on the added drawings list...
                if (result != null && AddedDrawings.Any(p => p == result) == false)
                {
                    //if so, add the result to the list of added drawings.
                    AddedDrawings.Add(result);
                }
            }
        }

        private void RemoveDrawingClick()
        {
            if(!string.IsNullOrEmpty(SelectedDrawing))
            {
                AddedDrawings.Remove(SelectedDrawing);
            }
        }

        private void RemoveAllDrawingsClick()
        {
            if(AddedDrawings != null && AddedDrawings.Count > 0)
            {
                AddedDrawings.Clear();
            }
        }

        private void OpenDrawingClick()
        {
            try
            {
                if (!string.IsNullOrEmpty(SelectedDrawing))
                {
                    DatabaseHelpers dh = new DatabaseHelpers();
                    dh.OpenDrawing(SelectedDrawing);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void DeleteBlocksCommandClick()
        {
            if (AddedDrawings != null && AddedDrawings.Count != 0)
            {
                System.Windows.Forms.DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Added drawings will be modified and saved. " +
                    "This process cannot be undone. Continue?", "Save drawings?", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                if (dialogResult == System.Windows.Forms.DialogResult.Yes)
                {
                    try
                    {
                        if (AddedDrawings != null && AddedDrawings.Count != 0)
                        {
                            CADAccess da = new CADAccess();
                            foreach (var drawing in AddedDrawings)
                                da.DeleteBlocks(drawing, BlockName);
                        }
                        else
                        {
                            MessageBox.Show("You have not added any drawings for this operation.");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    MessageBox.Show("Done.");
                }
            }
        }

        #endregion
    }
}
