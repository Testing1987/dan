using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace SystemVariableSetter.CoreLogic
{
    public class CADAccess
    {
        /// <summary>
        /// Creates a list of drawing files by asking the user to select them with an <see cref="OpenFileDialog"/> and returns that list.
        /// </summary>
        /// <returns></returns>
        public IList<string> AddtoDrawingList()
        {
            var items = new List<string>();

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Select Drawings";
            openFileDialog.Filter = "Drawing (*.dwg) | *.dwg";
            openFileDialog.Multiselect = true;
            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string file in openFileDialog.FileNames)
                {
                    items.Add(file);
                }
            }
            return items;
        }


        public void ModifySystemVariable(string drawing, string systemVariable, int systemVariableValue)
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;

            using (DocumentLock lock1 = thisDrawing.LockDocument())
            {
                using (Transaction thisTransaction = thisDrawing.TransactionManager.StartTransaction())
                {

                    if (File.Exists(drawing) && drawing != null)
                    {
                        try
                        {
                            using (Database database2 = new Database(false, true))
                            {
                                database2.ReadDwgFile(drawing, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                database2.CloseInput(true);
                                HostApplicationServices.WorkingDatabase = database2;

                                using (Transaction transaction2 = database2.TransactionManager.StartTransaction())
                                {
                                    try
                                    {
                                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable(systemVariable.ToUpper(),  systemVariableValue);
                                    }
                                    catch (System.Exception ex)
                                    {
                                        MessageBox.Show($"The system variable {systemVariable} does not exist, or you entered an invalid option for it's value. Please make sure you typed the variable name and it's value correctly.",
                                            "Incorrect Input",
                                            MessageBoxButton.OK,
                                            MessageBoxImage.Information);
                                        return;
                                    }

                                    //Commit the transaction and save the drawing.
                                    transaction2.Commit();
                                    database2.SaveAs(drawing, true, DwgVersion.Current, thisDrawing.Database.SecurityParameters);

                                    HostApplicationServices.WorkingDatabase = thisDrawing.Database;
                                }
                            }
                        }
                        catch (System.Exception)
                        {
                            MessageBox.Show($"Problem accessing database of {drawing}",
                                "The drawing could not be accessed.",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
                        }
                    }

                    thisTransaction.Commit();
                }
            }
        }
    }
}
