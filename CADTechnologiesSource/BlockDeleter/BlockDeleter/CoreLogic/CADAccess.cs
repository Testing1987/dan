using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using CADTechnologiesSource.All.AutoCADHelpers;
using CADTechnologiesSource.All.ColorHelpers;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace BlockDeleter.CoreLogic
{
    public class CADAccess
    {
        /// <summary>
        /// Creates a list of drawing files by asking the user to select them with an asf<see cref="OpenFileDialog"/> and returns that list.
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

        internal void DeleteBlocks(string drawing, string blockName)
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock lock1 = thisDrawing.LockDocument())
            {
                using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    if (File.Exists(drawing) && drawing != null)
                    {
                        try
                        {
                            using (Database database2 = new Database(false, true))
                            {
                                database2.ReadDwgFile(drawing, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                database2.CloseInput(true);
                                //database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);

                                HostApplicationServices.WorkingDatabase = database2;

                                using (Transaction transaction2 = database2.TransactionManager.StartTransaction())
                                {
                                    //Get the dictionary of layouts
                                    DBDictionary dBDictionary = transaction2.GetObject(database2.LayoutDictionaryId, OpenMode.ForWrite) as DBDictionary;

                                    //Access the block table and find the block.
                                    //BlockTable blockTable = transaction2.GetObject(database2.BlockTableId, OpenMode.ForRead) as BlockTable;
                                    //BlockTableRecord blockTableRecord = transaction2.GetObject(blockTable[blockName], OpenMode.ForRead) as BlockTableRecord;


                                    foreach (DBDictionaryEntry dBDictionaryEntry in dBDictionary)
                                    {
                                        Layout layout = dBDictionaryEntry.Value.GetObject(OpenMode.ForRead) as Layout;
                                        if (layout.LayoutName != "Model")
                                        {
                                            BlockTableRecord paperSpace = transaction2.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                                            foreach (ObjectId objectId in paperSpace)
                                            {
                                                BlockReference blockReference = transaction2.GetObject(objectId, OpenMode.ForWrite) as BlockReference;
                                                if (blockReference != null)
                                                {
                                                    if (blockReference.IsDynamicBlock == true)
                                                    {
                                                        string dynamicBlockName = blockReference.IsDynamicBlock ? ((BlockTableRecord)blockReference.DynamicBlockTableRecord.GetObject(OpenMode.ForWrite)).Name : blockReference.Name;
                                                        if (dynamicBlockName == blockName)
                                                        {
                                                            blockReference.Erase();
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if(blockReference.Name == blockName)
                                                        blockReference.Erase();
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    transaction2.Commit();
                                    database2.SaveAs(drawing, true, DwgVersion.Current, thisDrawing.Database.SecurityParameters);

                                    HostApplicationServices.WorkingDatabase = thisDrawing.Database;
                                }
                            }
                        }
                        catch (System.Exception)
                    {
                        MessageBox.Show("The drawing could not be accessed.", $"Problem accessing database of {drawing}", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }
    }
}
}














