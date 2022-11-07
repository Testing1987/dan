using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using MessageBox = System.Windows.MessageBox;

namespace FileHelper
{
    public partial class MainViewModel : ObservableObject
    {

        [ObservableProperty]
        string searchPath;

        [ObservableProperty]
        ObservableCollection<string> files;

        public MainViewModel()
        {
            files = new ObservableCollection<string>();
            searchPath = "Select a folder";
        }

        [RelayCommand]
        public void GetPaths()
        {
            if (files != null)
            {
                if (files.Count > 0)
                {
                    files.Clear();
                }
                try
                {
                    //Have the user select a parent folder to save the txt file...
                    // Prepare a dummy string, this would appear in the dialog
                    string dummyFileName = "Start search here";

                    System.Windows.Forms.SaveFileDialog saveFIleDialog = new System.Windows.Forms.SaveFileDialog();

                    // Feed the dummy name to the save dialog
                    saveFIleDialog.FileName = dummyFileName;

                    if (saveFIleDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Now here's our save folder
                        string userPath = Path.GetDirectoryName(saveFIleDialog.FileName);
                        searchPath = userPath;
                        string[] result = Directory.GetFiles(searchPath, "*", SearchOption.AllDirectories);
                        {
                            foreach (string path in result)
                            {
                                if(path.Contains(".dwg"))
                                files.Add(path);
                            }
                        }

                        if (result != null)
                        {
                            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "found_files.txt");
                            System.IO.File.WriteAllLines(filePath, result);
                            //System.Diagnostics.Process.Start(filePath);
                        }
                        MessageBox.Show("I've saved a list of the files to your desktop.", "List saved.", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
