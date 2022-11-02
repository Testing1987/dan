using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using LayerManager.Helpers.AutoCADHelpers;
using LayerManager.Helpers.PropertyHelpers;
using LayerManager.CoreLogic;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows.Input;
using LayerManager.Base;

namespace LayerManager.ViewModels
{
    public class MainWindowViewModel : BaseViewModel
    {

        #region Singleton
        public static MainWindowViewModel Instance = new MainWindowViewModel();
        #endregion

        #region Public Properties

        public string DrawingListName { get; set; }

        public string LayerName { get; set; }
        public bool IsFrozen { get; set; }
        public bool Off { get; set; }
        public bool IsLocked { get; set; }
        public bool IsPlottable { get; set; }
        public bool IsXrefLayer { get; set; }


        public List<string> ColorComboBoxItems { get; set; }
        public bool AdjustColor { get; set; }
        public string SelectedColorType { get; set; }
        public bool ColorComboBoxIndex { get; set; }
        public string ColorValue { get; set; }


        public bool AdjustLineweight { get; set; }
        public List<string> Lineweights { get; set; }
        public string SelectedLineweight { get; set; }
        public int SelectedLineweightIndex { get; set; }


        public bool AdjustTransparency { get; set; }
        public string Transparency { get; set; }


        public bool AdjustLinetype { get; set; }
        public string Linetype { get; set; }


        public ObservableCollection<string> AddedDrawings { get; set; }
        public string SelectedDrawing { get; set; }


        //Add Linetype, Lineweight, Transparency, and Color

        #endregion

        #region Public Commands
        public ICommand AddDrawingCommand { get; set; }
        public ICommand RemoveDrawingCommand { get; set; }
        public ICommand RemoveAllDrawingsCommand { get; set; }
        public ICommand OpenDrawingCommand { get; set; }
        public ICommand EditLayersCommand { get; set; }
        public ICommand SaveListCommand { get; set; }
        public ICommand LoadListCommand { get; set; }
        #endregion

        public MainWindowViewModel()
        {
            #region Commands
            AddDrawingCommand = new RelayCommand(() => AddDrawingsClick());
            RemoveDrawingCommand = new RelayCommand(() => RemoveDrawingClick());
            RemoveAllDrawingsCommand = new RelayCommand(() => RemoveAllDrawingsClick());
            OpenDrawingCommand = new RelayCommand(() => OpenDrawingClick());
            EditLayersCommand = new RelayCommand(() => EditLayersCommandClick());
            SaveListCommand = new RelayCommand(() => SaveListCommandClick());
            LoadListCommand = new RelayCommand(() => LoadListCommandClick());
            #endregion

            #region List Population

            if (ColorComboBoxItems == null)
            {
                ColorComboBoxItems = new List<string>();
                ColorComboBoxItems.Add("Color Index");
                ColorComboBoxItems.Add("True Color");
            }

            if (Lineweights == null)
            {
                Lineweights = Enum.GetNames(typeof(LineWeight)).ToList();
                Lineweights.Remove("ByLayer");
                Lineweights.Remove("ByBlock");
                Lineweights.Remove("ByLineWeightDefault");
            }

            #endregion

            #region Null Checks
            if (AddedDrawings == null)
            {
                AddedDrawings = new ObservableCollection<string>();
            }
            #endregion

            #region Defaults
            IsFrozen = false;
            Off = false;
            IsLocked = false;
            IsPlottable = true;
            AdjustLineweight = false;
            AdjustTransparency = false;
            Transparency = "0";
            SelectedLineweightIndex = 26;
            DrawingListName = "Not Set";
            #endregion
        }

        #region List Control
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
            if (!string.IsNullOrEmpty(SelectedDrawing))
            {
                AddedDrawings.Remove(SelectedDrawing);
            }
        }

        private void RemoveAllDrawingsClick()
        {
            if (AddedDrawings != null && AddedDrawings.Count > 0)
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
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void LoadListCommandClick()
        {
            if (AddedDrawings == null)
            {
                AddedDrawings = new ObservableCollection<string>();
            }
            if (AddedDrawings.Count > 0)
            {
                System.Windows.Forms.DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Loading a new list will clear out your current list. Do you wish to continue?",
                    "Clear current list?",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Question);
                switch (dialogResult)
                {
                    case System.Windows.Forms.DialogResult.Yes:
                        AddedDrawings.Clear();
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
                                                    AddedDrawings.Add(line);
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
                                            AddedDrawings.Add(line);
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
                                DrawingListName =  Path.GetFileName(openFileDialog.FileName);
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

        private void SaveListCommandClick()
        {
            try
            {
                if (AddedDrawings.Count >= 1 && AddedDrawings != null)
                {
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Text File (*.txt)|*.txt";
                    saveFileDialog.Title = "Save the current drawing list.";
                    saveFileDialog.RestoreDirectory = true;
                    Nullable<bool> result = saveFileDialog.ShowDialog();

                    if (result == true)
                    {
                        System.IO.File.WriteAllLines(saveFileDialog.FileName, AddedDrawings);
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

        #endregion

        #region Command Methods

        private void EditLayersCommandClick()
        {
            if (string.IsNullOrEmpty(LayerName) == false)
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
                                int transparencyNumber;
                                bool transparencySuccess = int.TryParse(Transparency, out transparencyNumber);

                                if (SelectedLineweight == null)
                                {
                                    SelectedLineweight = "ByLineWeightDefault";
                                }
                                LineWeight newLineweight = (LineWeight)Enum.Parse(typeof(LineWeight), SelectedLineweight.ToString());


                                Autodesk.AutoCAD.Colors.Color newColor = new Autodesk.AutoCAD.Colors.Color();

                                if (string.IsNullOrEmpty(ColorValue) == false)
                                {
                                    if (AdjustColor)
                                    {
                                        switch (SelectedColorType)
                                        {
                                            case "Color Index":

                                                switch (ColorValue.ToLower())
                                                {
                                                    case "red":
                                                        newColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Red);
                                                        break;
                                                    case "yellow":
                                                        newColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Yellow);
                                                        break;
                                                    case "green":
                                                        newColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Green);
                                                        break;
                                                    case "cyan":
                                                        newColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Cyan);
                                                        break;
                                                    case "blue":
                                                        newColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Blue);
                                                        break;
                                                    case "magenta":
                                                        newColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_Magenta);
                                                        break;
                                                    case "white":
                                                        newColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, AcColors.AutoCAD_White_Black);
                                                        break;
                                                    default:
                                                        short shortColor = Convert.ToInt16(ColorValue);
                                                        newColor = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, shortColor);
                                                        break;
                                                }

                                                break;

                                            case "True Color":
                                                if (ColorValue.ToLower().Contains(",") == true)
                                                {
                                                    int idx1 = ColorValue.IndexOf(",", 0);
                                                    byte R = Convert.ToByte(ColorValue.Substring(0, idx1));
                                                    int idx2 = ColorValue.IndexOf(",", idx1 + 1);
                                                    byte G = Convert.ToByte(ColorValue.Substring(idx1 + 1, idx2 - idx1 - 1));
                                                    byte B = Convert.ToByte(ColorValue.Substring(idx2 + 1, ColorValue.Length - idx2 - 1));
                                                    newColor = Autodesk.AutoCAD.Colors.Color.FromRgb(R, G, B);
                                                }
                                                break;
                                            default:
                                                break;
                                        }
                                    }
                                }

                                if (transparencySuccess)
                                {
                                    CADAccess da = new CADAccess();
                                    foreach (var drawing in AddedDrawings)
                                        da.UpdateLayerProperties(drawing, LayerName.ToLower(), IsXrefLayer, IsFrozen, Off, IsLocked, IsPlottable, AdjustTransparency, transparencyNumber, AdjustLinetype, Linetype, AdjustLineweight, newLineweight, AdjustColor, newColor);
                                }
                                else
                                {
                                    System.Windows.Forms.MessageBox.Show($"Transparency has to be a number. You entered '{Transparency}', which cannot parse into a number. Please enter a number.", "Not a number",
                                        System.Windows.Forms.MessageBoxButtons.OK,
                                        System.Windows.Forms.MessageBoxIcon.Information);
                                    return;
                                }
                            }
                            else
                            {
                                System.Windows.Forms.MessageBox.Show("You have not added any drawings.", "No drawings added",
                                    System.Windows.Forms.MessageBoxButtons.OK,
                                    System.Windows.Forms.MessageBoxIcon.Information);
                            }
                        }
                        catch (System.Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show(ex.Message);
                        }
                        System.Windows.Forms.MessageBox.Show("Done.", "Layer Update Complete",
                                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    }
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("You must enter a layer name to edit.",
                    "No layer name.",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
        }

        #endregion
    }
}
