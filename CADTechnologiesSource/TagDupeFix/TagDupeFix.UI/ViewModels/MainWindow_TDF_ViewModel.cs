using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using CADTechnologiesSource.All.AutoCADHelpers;
using CADTechnologiesSource.All.Base;
using TagDupeFix.Core.CoreLogic;

namespace TagDupeFix.UI.ViewModels
{
    public class MainWindow_TDF_ViewModel : BaseViewModel
    {
        #region Singleton
        public static MainWindow_TDF_ViewModel Instance = new MainWindow_TDF_ViewModel();
        #endregion

        #region Public Properties

        public string FindText { get; set; }
        public string ReplaceText { get; set; }
        public string SelectedDrawing { get; set; }
        public string SelectedParameter { get; set; }
        public int ComboboxIndex { get; set; }
        public ObservableCollection<string> AddedDrawings { get; set; }
        public ICommand AddDrawingCommand { get; set; }
        public ICommand RemoveDrawingCommand { get; set; }
        public ICommand RemoveAllDrawingsCommand { get; set; }
        public ICommand ReplaceTextCommand { get; set; }
        public ICommand OpenDrawingCommand { get; set; }
        public ICommand ATTSYNCCommand { get; set; }

        #endregion

        public MainWindow_TDF_ViewModel()
        {
            //Define the commands
            AddDrawingCommand = new RelayCommand(() => AddDrawingClick());
            RemoveDrawingCommand = new RelayCommand(() => RemoveDrawingClick());
            RemoveAllDrawingsCommand = new RelayCommand(() => RemoveAllDrawingsClick());
            ReplaceTextCommand = new RelayCommand(() => ReplaceClick());
            OpenDrawingCommand = new RelayCommand(() => OpenDrawingClick());
            ATTSYNCCommand = new RelayCommand(() => ATTSYNCClick());

            //Ensure the List of Added Drawings is not null
            if (AddedDrawings == null)
            {
                AddedDrawings = new ObservableCollection<string>();
            }
        }

        #region Command Methods

        #region Add/Remove
        /// <summary>
        /// Adds drawings to the <see cref="AddedDrawings"/> List.
        /// </summary>
        private void AddDrawingClick()
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

        /// <summary>
        /// Removes drawings from the <see cref="AddedDrawings"/> List.
        /// </summary>
        private void RemoveDrawingClick()
        {
            AddedDrawings.Remove(SelectedDrawing);
        }

        /// <summary>
        /// Removes all drawings from the AddedDrawings List.
        /// </summary>
        private void RemoveAllDrawingsClick()
        {
            AddedDrawings.Clear();
        }
        #endregion

        #region Sync
        private void ATTSYNCClick()
        {
            if (AddedDrawings != null && AddedDrawings.Count != 0)
            {
                System.Windows.Forms.DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Added drawings will be modified and saved. This process cannot be undone. Continue?", "Save drawings?", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                if (dialogResult == System.Windows.Forms.DialogResult.Yes)
                {
                    try
                    {
                        if (AddedDrawings != null && AddedDrawings.Count != 0)
                        {
                            CADAccess ca = new CADAccess();
                            foreach (var drawing in AddedDrawings)
                                ca.ATTSYNCBlock(drawing);
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

        /// <summary>
        /// Runs Find/Replace on the <see cref="AddedDrawings"/>.
        /// </summary>
        private void ReplaceClick()
        {
            if (AddedDrawings != null && AddedDrawings.Count != 0)
            {
                System.Windows.Forms.DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Added drawings will be modified and saved. This process cannot be undone. Continue?", "Save drawings?", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                if (dialogResult == System.Windows.Forms.DialogResult.Yes)
                {
                    try
                    {
                        if (AddedDrawings != null && AddedDrawings.Count != 0)
                        {
                            CADAccess ca = new CADAccess();
                            foreach (var drawing in AddedDrawings)
                                ca.FindandFixTagDupes(drawing);
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

        #region Helpers
        /// <summary>
        /// Opens the selected drawing from the <see cref="AddedDrawings"/> List.
        /// </summary>
        private void OpenDrawingClick()
        {
            try
            {
                if (SelectedDrawing != null)
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
        #endregion
        #endregion

    }
}
