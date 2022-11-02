using CADTechnologiesSource.All.AutoCADHelpers;
using CADTechnologiesSource.All.Base;
using GFAR.Core.CoreLogic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace GFAR.Core.ViewModels
{
    public class MainWindow_GFAR_ViewModel : BaseViewModel
    {
        #region Singleton
        /// <summary>
        /// A single instance of <see cref="MainWindow_GFAR_ViewModel"/> to bind to.
        /// </summary>
        public static MainWindow_GFAR_ViewModel Instance = new MainWindow_GFAR_ViewModel();

        #endregion

        #region Public Properties

        public string FindText { get; set; }
        public string ReplaceText { get; set; }
        public string SelectedDrawing { get; set; }
        public string SelectedParameter { get; set; }
        public int ComboboxIndex { get; set; }
        public ObservableCollection<string> AddedDrawings { get; set; }
        public ObservableCollection<string> SearchParameters { get; set; }
        public ICommand AddDrawingCommand { get; set; }
        public ICommand RemoveDrawingCommand { get; set; }
        public ICommand RemoveAllDrawingsCommand { get; set; }
        public ICommand ReplaceTextCommand { get; set; }
        public ICommand OpenDrawingCommand { get; set; }
        public ICommand InfoCommand { get; set; }

        #endregion

        #region Constructor
        /// <summary>
        /// Default Constructor
        /// </summary>
        public MainWindow_GFAR_ViewModel()
        {
            //Define the commands
            AddDrawingCommand = new RelayCommand(() => AddDrawingClick());
            RemoveDrawingCommand = new RelayCommand(() => RemoveDrawingClick());
            RemoveAllDrawingsCommand = new RelayCommand(() => RemoveAllDrawingsClick());
            ReplaceTextCommand = new RelayCommand(() => ReplaceClick());
            OpenDrawingCommand = new RelayCommand(() => OpenDrawingClick());
            InfoCommand = new RelayCommand(() => InfoClick());

            //Ensure the List of Added Drawings is not null
            if (AddedDrawings == null)
            {
                AddedDrawings = new ObservableCollection<string>();
            }

            //Ensure the search combobox is not null
            if (SearchParameters == null)
            {
                SearchParameters = new ObservableCollection<string>();
                SearchParameters.Add("Model Space");
                SearchParameters.Add("Paper Space (All Layouts)");
                SearchParameters.Add("Both");
            }

            ComboboxIndex = 0;
        }



        #endregion

        #region Command Methods
        //Define the commands created in the constructor.

        #region Add/Remove
        /// <summary>
        /// Adds drawings to the <see cref="AddedDrawings"/> List.
        /// </summary>
        private void AddDrawingClick()
        {
            DataAccess da = new DataAccess();

            foreach (var result in da.AddtoDrawingList())
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

        /// <summary>
        /// Runs Find/Replace on the <see cref="AddedDrawings"/>.
        /// </summary>
        private void ReplaceClick()
        {
            if(AddedDrawings != null && AddedDrawings.Count != 0)
            {
                System.Windows.Forms.DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Added drawings will be modified and saved. This process cannot be undone. Continue?", "Save drawings?", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                if (dialogResult == System.Windows.Forms.DialogResult.Yes)
                {
                    try
                    {
                        if (AddedDrawings != null && AddedDrawings.Count != 0)
                        {
                            if (SelectedParameter != null)
                            {
                                DataAccess da = new DataAccess();
                                foreach (var drawing in AddedDrawings)
                                    da.FindAndReplaceText(drawing, FindText, ReplaceText, SelectedParameter);
                            }
                            else
                            {
                                MessageBox.Show("Please use the drop down menu to select a search parameter.");
                            }
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

        private void InfoClick()
        {
            MessageBox.Show("GFAR will search for the given text in all " +
                "of the added drawings text, mtext, mleaders, and block attributes  " +
                "and, if found, replace it with the given replacement text.");
        }

        #endregion

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
    }
}