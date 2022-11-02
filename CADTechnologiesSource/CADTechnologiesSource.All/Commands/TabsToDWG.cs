using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Microsoft.Win32;
using System.IO;
using System.Windows;

namespace CADTechnologiesSource.All.Commands
{
    public class TabsToDWG
    {
        [CommandMethod("TABS_TO_DWG")]
        public void TabsToDwg()
        {
            MessageBox.Show("Do not attempt to use this command on a drawing that is hosted on Talon. Copy the drawing to your C drive first.",
                "Talon Bad", MessageBoxButton.OK,
                MessageBoxImage.Warning);

            string drawingPath = "";
            try
            {
                //Select a file
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Title = "Select .dwg";
                openFileDialog.Filter = "Dwg files|*.dwg";
                if (openFileDialog.ShowDialog() == true)
                {
                    drawingPath = openFileDialog.FileName;
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }


            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;

            if (string.IsNullOrEmpty(drawingPath) == false)
            {
                using (DocumentLock docLock = thisDrawing.LockDocument())
                {
                    using (Database sideDatabase = new Database(false, true))
                    {
                        //then use the ReadDwgFile method provided by the API to access a DWG file and get it's database. 
                        sideDatabase.ReadDwgFile(drawingPath, FileOpenMode.OpenForReadAndWriteNoShare, false, null);
                        sideDatabase.CloseInput(true);
                        //sideDatabase.ResolveXrefs(false, false);

                        string saveAsNew = Path.GetDirectoryName(drawingPath);

                        HostApplicationServices.WorkingDatabase = sideDatabase;

                        using (Transaction sideTransaction = sideDatabase.TransactionManager.StartTransaction())
                        {
                            DBDictionary layouts = (DBDictionary)sideTransaction.GetObject(sideDatabase.LayoutDictionaryId, OpenMode.ForRead);
                            LayoutManager layoutManager = LayoutManager.Current;


                            //save a drawing for each layout
                            foreach (DBDictionaryEntry layout in layouts)
                            {
                                Layout destinationLayout = (Layout)sideTransaction.GetObject(layout.Value, OpenMode.ForRead);
                                if (destinationLayout.LayoutName != "Model")
                                {
                                    string saveAsDirectory = Path.GetDirectoryName(sideDatabase.Filename);
                                    string saveAsLocation = $"{saveAsNew}" + $"\\{destinationLayout.LayoutName}" + ".dwg";
                                    sideDatabase.SaveAs(saveAsLocation, true, DwgVersion.Current, thisDrawing.Database.SecurityParameters);

                                    using (Database newDatabase = new Database(false, true))
                                    {
                                        //then use the ReadDwgFile method provided by the API to access a DWG file and get it's database. 
                                        newDatabase.ReadDwgFile(saveAsLocation, FileOpenMode.OpenForReadAndWriteNoShare, false, null);
                                        newDatabase.CloseInput(true);
                                        //newDatabase.ResolveXrefs(false, false);
                                        HostApplicationServices.WorkingDatabase = newDatabase;

                                        using (Transaction newTransaction = newDatabase.TransactionManager.StartTransaction())
                                        {
                                            DBDictionary newLayouts = (DBDictionary)newTransaction.GetObject(newDatabase.LayoutDictionaryId, OpenMode.ForRead);
                                            LayoutManager newLayoutManager = LayoutManager.Current;

                                            //Delete layouts that don't equal the drawing name
                                            foreach (DBDictionaryEntry newLayout in newLayouts)
                                            {
                                                Layout newDestinationLayout = (Layout)newTransaction.GetObject(newLayout.Value, OpenMode.ForRead);
                                                string drawingName = Path.GetFileNameWithoutExtension(saveAsLocation);
                                                if (newDestinationLayout.LayoutName != "Model" && newDestinationLayout.LayoutName != drawingName)
                                                {
                                                    newLayoutManager.DeleteLayout(newDestinationLayout.LayoutName);
                                                    newDatabase.SaveAs(saveAsLocation, true, DwgVersion.Current, thisDrawing.Database.SecurityParameters);
                                                }
                                            }
                                            newTransaction.Commit();
                                        }
                                        HostApplicationServices.WorkingDatabase = thisDatabase;
                                    }
                                }
                            }
                            sideTransaction.Commit();
                        }
                        HostApplicationServices.WorkingDatabase = thisDatabase;
                    }
                }
            }
            else
            {
                MessageBox.Show($"{drawingPath} was not a valid drawing or could not be found.", "File not found", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
        }
    }
}
