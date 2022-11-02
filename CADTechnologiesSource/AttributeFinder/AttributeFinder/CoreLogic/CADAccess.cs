using AttributeFinder.Models;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
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

namespace AttributeFinder.CoreLogic
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

        internal void PopulateBlocks(string drawing, string blockName, string attributeName, string matchBy, List<OwnersModel> owners)
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
                                    DBDictionary dBDictionary = transaction2.GetObject(database2.LayoutDictionaryId, OpenMode.ForWrite) as DBDictionary;
                                    string nameToUse = "";
                                    Handle blockHandle = new Handle();


                                    BlockTable blockTable = transaction2.GetObject(database2.BlockTableId, OpenMode.ForRead) as BlockTable;
                                    BlockTableRecord blockTableRecord = transaction2.GetObject(blockTable[blockName], OpenMode.ForRead) as BlockTableRecord;
                                    foreach (ObjectId objectId1 in blockTableRecord)
                                    {
                                        DBObject dBObject = objectId1.GetObject(OpenMode.ForRead);
                                        AttributeDefinition attributeDefinition = dBObject as AttributeDefinition;
                                        if (attributeDefinition != null && !attributeDefinition.Constant)
                                        {
                                            if (attributeDefinition.Tag == attributeName)
                                            {
                                                blockTableRecord.UpgradeOpen();
                                                attributeDefinition.UpgradeOpen();
                                                blockTableRecord.Explodable = true;

                                                #region SetLayer
                                                string layerName = "Property_Owners";
                                                Autodesk.AutoCAD.DatabaseServices.LayerTable layerTable;
                                                layerTable = (Autodesk.AutoCAD.DatabaseServices.LayerTable)transaction2.GetObject(database2.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                                if (layerTable.Has(layerName) == true)
                                                {
                                                    layerTable.UpgradeOpen();
                                                    LayerTableRecord newLayer = transaction2.GetObject(layerTable[layerName], OpenMode.ForWrite) as LayerTableRecord;
                                                    if (newLayer != null)
                                                    {
                                                        newLayer.IsFrozen = false;
                                                        newLayer.IsPlottable = true;
                                                        attributeDefinition.Layer = layerName;
                                                    }
                                                }

                                                if (layerTable.Has(layerName) == false)
                                                {
                                                    layerTable.UpgradeOpen();
                                                    LayerTableRecord new_layer = new Autodesk.AutoCAD.DatabaseServices.LayerTableRecord();
                                                    new_layer.Name = layerName;
                                                    new_layer.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 2);
                                                    new_layer.IsPlottable = true;
                                                    new_layer.IsFrozen = false;
                                                    layerTable.Add(new_layer);
                                                    transaction2.AddNewlyCreatedDBObject(new_layer, true);
                                                    attributeDefinition.Layer = layerName;
                                                }
                                                #endregion
                                            }
                                        }
                                    }
                                    SynchronizeBlockAttributes.SynchronizeAttributes(blockTableRecord);


                                    foreach (DBDictionaryEntry dBDictionaryEntry in dBDictionary)
                                    {
                                        Layout layout = dBDictionaryEntry.Value.GetObject(OpenMode.ForRead) as Layout;
                                        if (layout.LayoutName != "Model")
                                        {
                                            BlockTableRecord blockTableLayoutRecord = transaction2.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                                            foreach (var owner in owners)
                                            {
                                                foreach (ObjectId objectId in blockTableLayoutRecord)
                                                {
                                                    BlockReference blockReference = transaction2.GetObject(objectId, OpenMode.ForWrite) as BlockReference;
                                                    if (blockReference != null)
                                                    {
                                                        if (blockReference.IsDynamicBlock == true)
                                                        {
                                                            string dynamicBlockName = blockReference.IsDynamicBlock ? ((BlockTableRecord)blockReference.DynamicBlockTableRecord.GetObject(OpenMode.ForWrite)).Name : blockReference.Name;
                                                            if (dynamicBlockName == blockName)
                                                            {
                                                                AttributeCollection attributeCollection = blockReference.AttributeCollection;
                                                                foreach (ObjectId attributeReferenceID in attributeCollection)
                                                                {
                                                                    AttributeReference attributeReference = transaction2.GetObject(attributeReferenceID, OpenMode.ForWrite) as AttributeReference;
                                                                    if (attributeReference.Tag == matchBy)
                                                                    {
                                                                        if (attributeReference.TextString.Contains(owner.APN))
                                                                        {
                                                                            nameToUse = owner.OwnerName;
                                                                            blockHandle = blockReference.Handle;
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                foreach (ObjectId objectId2 in blockTableLayoutRecord)
                                                {
                                                    BlockReference blockReference = transaction2.GetObject(objectId2, OpenMode.ForWrite) as BlockReference;
                                                    if (blockReference != null)
                                                    {
                                                        if (blockReference.IsDynamicBlock == true)
                                                        {
                                                            //string dynamicBlockName = blockReference.IsDynamicBlock ? ((BlockTableRecord)blockReference.DynamicBlockTableRecord.GetObject(OpenMode.ForWrite)).Name : blockReference.Name;
                                                            if (blockReference.Handle == blockHandle)
                                                            {
                                                                AttributeCollection attributeCollection = blockReference.AttributeCollection;
                                                                foreach (ObjectId attributeReferenceID in attributeCollection)
                                                                {
                                                                    AttributeReference attributeReference = transaction2.GetObject(attributeReferenceID, OpenMode.ForWrite) as AttributeReference;
                                                                    if (attributeReference.Tag == attributeName)
                                                                    {
                                                                        if (nameToUse != "")
                                                                        {
                                                                            if (attributeReference.IsMTextAttribute)
                                                                            {
                                                                                MText newMtext = attributeReference.MTextAttribute;
                                                                                newMtext.Contents = nameToUse;
                                                                                attributeReference.MTextAttribute = newMtext;
                                                                            }
                                                                            else
                                                                            {
                                                                                attributeReference.TextString = nameToUse;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
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














