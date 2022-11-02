using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using LayerManager.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LayerManager.Base
{
    public class Launcher
    {
        [CommandMethod("Name_Of_Command")]
        public void TheMethod()
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;
            Editor editor = thisDrawing.Editor;

            using (DocumentLock lock1 = thisDrawing.LockDocument())
            {
                using (Transaction thisTransaction = thisDrawing.TransactionManager.StartTransaction())
                {
                    #region Layouts
                    //Access Layouts
                    DBDictionary dBDictionary = thisTransaction.GetObject(thisDatabase.LayoutDictionaryId, OpenMode.ForWrite) as DBDictionary;

                    //Iterate through layouts, not model space.
                    foreach (DBDictionaryEntry dBDictionaryEntry in dBDictionary)
                    {
                        Layout layout = dBDictionaryEntry.Value.GetObject(OpenMode.ForRead) as Layout;
                        if (layout.LayoutName != "Model")
                        {
                            //Do here
                        }
                    }
                    #endregion

                    #region Block Table
                    //Access the block table
                    BlockTable blockTable = thisTransaction.GetObject(thisDatabase.BlockTableId, OpenMode.ForRead) as BlockTable;
                    //Access a specific block by name.
                    BlockTableRecord blockTableRecord = thisTransaction.GetObject(blockTable["Block_Name"], OpenMode.ForRead) as BlockTableRecord;
                    //Access Model Space
                    BlockTableRecord modelSpace = thisTransaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    #endregion

                    #region Layers
                    //Access the Layer Table
                    LayerTable thisLayerTable = thisTransaction.GetObject(thisDatabase.LayerTableId, OpenMode.ForWrite) as LayerTable;
                    //Get Layers and their linetypes.
                    foreach (ObjectId objectId in thisLayerTable)
                    {
                        LayerTableRecord layerTableRecord = thisTransaction.GetObject(objectId, OpenMode.ForWrite) as LayerTableRecord;
                        LinetypeTable linetypeTable = thisTransaction.GetObject(thisDatabase.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;
                        LinetypeTableRecord lineTypeRecord = thisTransaction.GetObject(layerTableRecord.LinetypeObjectId, OpenMode.ForWrite) as LinetypeTableRecord;
                    }
                    #endregion



                    thisTransaction.Commit();
                }
            }
        }
    }
}
