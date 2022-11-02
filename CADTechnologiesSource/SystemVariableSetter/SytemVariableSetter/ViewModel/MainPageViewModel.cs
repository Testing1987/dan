using SystemVariableSetter.Base;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using SystemVariableSetter.CoreLogic;
using SystemVariableSetter.AutoCADHelpers;

namespace SystemVariableSetter.ViewModel
{
    public class MainPageViewModel : BaseViewModel
    {

        #region Singleton
        public static MainPageViewModel Instance = new MainPageViewModel();
        #endregion

        #region Public Properties

        public ObservableCollection<string> Drawings { get; set; }
        public string SelectedDrawing { get; set; }

        public string SystemVariable { get; set; }
        public int SystemVariableValue { get; set; }

        public string DrawingListName { get; set; }


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

        public MainPageViewModel()
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
            CADAccess ca = new CADAccess();

            foreach (var result in ca.AddtoDrawingList())
            {
                //make sure the added string isn't null and isn't already on the added drawings list...
                if (result != null && Drawings.Any(p => p == result) == false)
                {
                    //if so, add the result to the list of added drawings.
                    Drawings.Add(result);
                }
            }
        }

        private void RemoveDrawingClick()
        {
            if (!string.IsNullOrEmpty(SelectedDrawing))
            {
                Drawings.Remove(SelectedDrawing);
            }
        }

        private void RemoveAllDrawingsClick()
        {
            if (Drawings != null && Drawings.Count > 0)
            {
                Drawings.Clear();
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
                System.Windows.MessageBox.Show(ex.Message);
            }
        }

        private void SetSystemVariableCommandClick()
        {
            if (Drawings != null && Drawings.Count != 0)
            {
                System.Windows.Forms.DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Added drawings will be modified and saved. " +
                    "This process cannot be undone. Continue?", "Save drawings?", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                if (dialogResult == System.Windows.Forms.DialogResult.Yes)
                {
                    CADAccess ca = new CADAccess();
                    foreach (var drawing in Drawings)
                    {
                        if (Drawings != null && Drawings.Count > 0)
                        {
                            try
                            {
                                ca.ModifySystemVariable(drawing, SystemVariable, SystemVariableValue);
                            }
                            catch (Exception ex)
                            {
                                System.Windows.MessageBox.Show(ex.Message);
                            }
                        }
                    }
                    System.Windows.Forms.MessageBox.Show("Done.", "Operation Complete",
                                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                }
            }
        }

        private void SaveListCommandClick()
        {
            try
            {
                if (Drawings.Count >= 1 && Drawings != null)
                {
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Text File (*.txt)|*.txt";
                    saveFileDialog.Title = "Save the current drawing list.";
                    saveFileDialog.RestoreDirectory = true;
                    Nullable<bool> result = saveFileDialog.ShowDialog();

                    if (result == true)
                    {
                        System.IO.File.WriteAllLines(saveFileDialog.FileName, Drawings);
                        DrawingListName = saveFileDialog.FileName;
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("The drawing list is currently empty. If you want to save the list, please add items to it first.",
                        "The drawing list is empty.",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }
        }

        private void LoadListCommandClick()
        {
            if (Drawings == null)
            {
                Drawings = new ObservableCollection<string>();
            }
            if (Drawings.Count > 0)
            {
                System.Windows.Forms.DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Loading a new list will clear out your current list. Do you wish to continue?",
                    "Clear current list?",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Question);
                switch (dialogResult)
                {
                    case System.Windows.Forms.DialogResult.Yes:
                        Drawings.Clear();
                        List<string> missingFiles = new List<string>();
                        OpenFileDialog openFileDialog = new OpenFileDialog();
                        openFileDialog.Multiselect = false;
                        openFileDialog.Filter = "Drawing List (*.txt)|*.txt";
                        Nullable<bool> result = openFileDialog.ShowDialog();
                        if (result == true)
                        {
                            try
                            {
                                using (var fileStream = System.IO.File.OpenRead(openFileDialog.FileName))
                                {
                                    using (var streamReader = new System.IO.StreamReader(fileStream.Name))
                                    {
                                        foreach (var line in System.IO.File.ReadAllLines(openFileDialog.FileName, System.Text.Encoding.GetEncoding(1250)))
                                        {
                                            if (System.IO.File.Exists(line) == true)
                                            {
                                                if (string.IsNullOrEmpty(line) == false && line.Contains(".dwg"))
                                                {
                                                    Drawings.Add(line);
                                                }
                                                else
                                                {
                                                    {
                                                        System.Windows.MessageBox.Show($"{line} cannot be found. Please verify that the drawing still exists and is accessible.",
                                                            "File Not Found",
                                                            System.Windows.MessageBoxButton.OK,
                                                            System.Windows.MessageBoxImage.Error);
                                                    }
                                                }
                                            }
                                        }
                                        DrawingListName = Path.GetFileName(openFileDialog.FileName);
                                    }
                                }
                            }
                            catch (Exception)
                            {
                                System.Windows.MessageBox.Show("A valid drawing list could not be determined. Please make sure the .txt file is accessible.",
                                    "Can't access drawing list.",
                                    System.Windows.MessageBoxButton.OK,
                                    System.Windows.MessageBoxImage.Error);
                            }
                        }
                        break;
                    case System.Windows.Forms.DialogResult.No:
                        break;
                    default:
                        break;
                }
            }
            else
            {
                List<string> missingFiles = new List<string>();
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Multiselect = false;
                openFileDialog.Filter = "Drawing List (*.txt)|*.txt";
                Nullable<bool> result = openFileDialog.ShowDialog();
                if (result == true)
                {
                    try
                    {
                        using (var fileStream = System.IO.File.OpenRead(openFileDialog.FileName))
                        {
                            using (var streamReader = new System.IO.StreamReader(fileStream.Name))
                            {
                                foreach (var line in System.IO.File.ReadAllLines(openFileDialog.FileName, System.Text.Encoding.GetEncoding(1250)))
                                {
                                    if (System.IO.File.Exists(line) == true)
                                    {
                                        if (string.IsNullOrEmpty(line) == false && line.Contains(".dwg"))
                                        {
                                            Drawings.Add(line);
                                        }
                                        else
                                        {
                                            {
                                                System.Windows.MessageBox.Show($"{line} cannot be found. Please verify that the drawing still exists and is accessible.", "File Not Found",
                                                    System.Windows.MessageBoxButton.OK,
                                                    System.Windows.MessageBoxImage.Error);
                                            }
                                        }
                                    }
                                }
                                DrawingListName = Path.GetFileName(openFileDialog.FileName);
                            }
                        }
                    }
                    catch (Exception)
                    {
                        System.Windows.MessageBox.Show("A valid drawing list could not be determined. Please make sure the .txt file is accessible.", "Can't access drawing list.",
                            System.Windows.MessageBoxButton.OK,
                            System.Windows.MessageBoxImage.Error);
                    }
                }
            }
        }
    }
}
