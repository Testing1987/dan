using FolderForFiles.Base;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;

namespace FolderForFiles.ViewModels
{
    public class MainWindowViewModel : BaseViewModel
    {
        public static MainWindowViewModel Instance = new MainWindowViewModel();

        #region Public Properties
        public string UserPath { get; set; }

        public string FindText { get; set; }
        public string ReplaceText { get; set; }

        public ICommand FindPathCommand { get; set; }
        public ICommand MakeFoldersCommand { get; set; }
        public ICommand MoveToFoldersCommand { get; set; }
        public ICommand FindReplaceCommand { get; set; }
        #endregion

        public MainWindowViewModel()
        {
            FindPathCommand = new RelayCommand(() => FindPathCommandClick());
            MakeFoldersCommand = new RelayCommand(() => MakeFoldersCommandClick());
            MoveToFoldersCommand = new RelayCommand(() => MoveToFoldersCommandClick());
            FindReplaceCommand = new RelayCommand(() => FindReplaceCommandClick());
            FindText = "";
            ReplaceText = "";
        }


        private void FindPathCommandClick()
        {
            try
            {
                // Prepare a dummy string, thos would appear in the dialog
                string dummyFileName = "Save Here";

                System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
                //use this filter to hide other files
                saveFileDialog.Filter = "Directory | directory";

                // Feed the dummy name to the save dialog
                saveFileDialog.FileName = dummyFileName;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Now here's our save folder
                    string savePath = Path.GetDirectoryName(saveFileDialog.FileName);
                    UserPath = savePath;
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }
        }

        private void MakeFoldersCommandClick()
        {
            try
            {
                string[] files1 = Directory.GetFiles(UserPath);
                foreach (string file in files1)
                {

                    //Make sure it's a PDF
                    if (Path.GetFullPath(file).Contains(".pdf"))
                    {
                        //Make Folders
                        string filename = Path.GetFullPath(file);
                        Directory.CreateDirectory(filename.Replace(".pdf", ""));
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }


            string[] files2 = Directory.GetFiles(UserPath);
            foreach (var file in files2)
            {
                if (file.Contains(".pdf"))
                {
                    string destination = UserPath + Path.DirectorySeparatorChar + Path.GetFileNameWithoutExtension(file) + Path.DirectorySeparatorChar + Path.GetFileName(file);
                    //if (Directory.Exists(destination))
                    //{
                    //}
                    File.Move(file, destination);
                }
            }
        }

        private void MoveToFoldersCommandClick()
        {
            string[] files = Directory.GetFiles(UserPath);
            foreach (var file in files)
            {
                if (file.Contains(".pdf"))
                {
                    string destination = UserPath + Path.DirectorySeparatorChar + Path.GetFileNameWithoutExtension(file) + Path.DirectorySeparatorChar + Path.GetFileName(file);
                    //if (Directory.Exists(destination))
                    //{
                    //}
                    File.Move(file, destination);
                }
            }
        }

        private void FindReplaceCommandClick()
        {
            string[] files = Directory.GetFiles(UserPath);
            foreach (var file in files)
            {
                if (file.Contains(FindText))
                {
                    string oldFileName = Path.GetFullPath(file);
                    string newFileName = Path.GetFullPath(file).ToString().Replace(FindText, ReplaceText);

                    File.Move(oldFileName, newFileName);
                }
            }
        }

    }
}
