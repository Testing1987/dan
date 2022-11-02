using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace TagDupeFix.Core.CoreLogic
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
            openFileDialog.Title = "Select The Target Drawings To Add To Add To The List";
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

        public void FindandFixTagDupes(string drawing)
        {
            try
            {
                Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (DocumentLock docLock = thisDrawing.LockDocument())
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
                                    database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);

                                    HostApplicationServices.WorkingDatabase = database2;

                                    using (Transaction trans2 = database2.TransactionManager.StartTransaction())
                                    {

                                        DBDictionary dBDictionary = trans2.GetObject(database2.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                                        foreach (DBDictionaryEntry dBDictionaryEntry in dBDictionary)
                                        {
                                            Layout layout = dBDictionaryEntry.Value.GetObject(OpenMode.ForRead) as Layout;
                                            if (layout.LayoutName != "Model")
                                            {
                                                BlockTable blockTable = trans2.GetObject(database2.BlockTableId, OpenMode.ForWrite) as BlockTable;
                                                int attributeIndex = 0;

                                                if (blockTable.Has("TC_UL_RV") == true)
                                                {
                                                    BlockTableRecord revBlock = trans2.GetObject(blockTable["TC_UL_RV"], OpenMode.ForWrite) as BlockTableRecord;
                                                    if (revBlock != null)
                                                    {
                                                        foreach (ObjectId objectId in revBlock)
                                                        {
                                                            Entity entity = trans2.GetObject(objectId, OpenMode.ForWrite) as Entity;
                                                            if (entity != null)
                                                            {
                                                                AttributeDefinition attributeDefinition = entity as AttributeDefinition;
                                                                if (attributeDefinition != null)
                                                                {
                                                                    attributeIndex++;
                                                                    if (attributeDefinition.Tag == "REV_DRAFTCHECK8" && attributeIndex >= 48)
                                                                    {
                                                                        attributeDefinition.Tag = "REV_DRAFTCHECK17";
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        trans2.Commit();
                                        //Document accessedDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.GetDocument(database2);
                                        //accessedDrawing.SendStringToExecute("ATTSYNC N TC_UL_RV" + "\r", false, false, true);
                                        database2.SaveAs(drawing, true, DwgVersion.Current, thisDrawing.Database.SecurityParameters);
                                    }
                                    HostApplicationServices.WorkingDatabase = thisDrawing.Database;
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void ATTSYNCBlock(string drawing)
        {
            try
            {
                Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (File.Exists(drawing) && drawing != null)
                {
                    try
                    {
                        Document accessedDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.Open(drawing, false);
                        DocumentLock accessedDocLock = accessedDrawing.LockDocument();
                        Database accessedDatabase = accessedDrawing.Database;
                        using (Transaction thisTransaction = accessedDatabase.TransactionManager.StartTransaction())
                        {
                            DBDictionary dBDictionary = thisTransaction.GetObject(accessedDatabase.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                            foreach (DBDictionaryEntry dBDictionaryEntry in dBDictionary)
                            {
                                Layout layout = dBDictionaryEntry.Value.GetObject(OpenMode.ForRead) as Layout;
                                if (layout.LayoutName != "Model")
                                {
                                    BlockTable blockTable = thisTransaction.GetObject(accessedDatabase.BlockTableId, OpenMode.ForWrite) as BlockTable;
                                    if (blockTable.Has("TC_UL_RV") == true)
                                    {
                                        BlockTableRecord revBlock = thisTransaction.GetObject(blockTable["TC_UL_RV"], OpenMode.ForWrite) as BlockTableRecord;
                                        if (revBlock != null)
                                        {
                                            accessedDrawing.SendStringToExecute("ATTSYNC N TC_UL_RV" + "\r", true, false, false);
                                        }
                                    }
                                }
                            }
                            thisTransaction.Commit();
                            accessedDrawing.Database.SaveAs(accessedDrawing.Name, true, DwgVersion.Current, accessedDrawing.Database.SecurityParameters);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
