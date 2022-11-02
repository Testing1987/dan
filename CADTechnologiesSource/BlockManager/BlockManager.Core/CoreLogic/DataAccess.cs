using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Internal;
using BlockManager.Core.Models;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;

namespace BlockManager.Core.CoreLogic
{
    public class DataAccess
    {
        #region Drawing List Methods

        /// <summary>
        /// Creates a list of drawing files by asking the user to select them with an <see cref="OpenFileDialog"/> and returns that list.
        /// </summary>
        /// <returns></returns>
        public IList<TargetDrawingModelBAM> AddtoTargetDrawingList()
        {
            var items = new List<TargetDrawingModelBAM>();

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Select The Target Drawings";
            openFileDialog.Filter = "Drawing (*.dwg) | *.dwg";
            openFileDialog.Multiselect = true;
            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string file in openFileDialog.FileNames)
                {
                    TargetDrawingModelBAM targetDrawing = new TargetDrawingModelBAM();
                    targetDrawing.DrawingPath = file;
                    targetDrawing.TrimmedPath = Path.GetFileName(file);
                    items.Add(targetDrawing);
                }
            }
            return items;
        }

        #endregion

        #region AutoCAD Database Access
        /// <summary>
        /// Looks through the given drawing for block references with attributes and returns them in an <see cref="ObservableCollection{BlockModel}"/>
        /// </summary>
        /// <param name="drawingName"></param>
        /// <returns></returns>
        public ObservableCollection<BlockModel> GetAttributedBlocksFromDrawings(string drawingName)
        {
            ObservableCollection<BlockModel> result = new ObservableCollection<BlockModel>();

            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;

            try
            {
                using (DocumentLock documentLock = thisDrawing.LockDocument())
                {
                    using (Transaction transaction = thisDrawing.TransactionManager.StartTransaction())
                    {
                        if (File.Exists(drawingName) && drawingName != null)
                            try
                            {
                                using (Database database2 = new Database(false, true))
                                {
                                    database2.ReadDwgFile(drawingName, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                    database2.CloseInput(true);
                                    //database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);
                                    HostApplicationServices.WorkingDatabase = database2;

                                    using (Transaction transaction2 = database2.TransactionManager.StartTransaction())
                                    {

                                        #region Paper Space Layouts
                                        DBDictionary dBDictionary = transaction2.GetObject(database2.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                                        foreach (DBDictionaryEntry dBDictionaryEntry in dBDictionary)
                                        {
                                            Layout layout = dBDictionaryEntry.Value.GetObject(OpenMode.ForRead) as Layout;
                                            if (layout.LayoutName != "Model")
                                            {
                                                BlockTableRecord blockTablePaperSpace = transaction2.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                                                foreach (ObjectId objectId in blockTablePaperSpace)
                                                {
                                                    BlockReference blockReference = transaction2.GetObject(objectId, OpenMode.ForRead) as BlockReference;
                                                    if (blockReference != null)
                                                    {
                                                        BlockModel blockModel = new BlockModel();
                                                        blockModel.BlockName = blockReference.Name;
                                                        blockModel.BlockHandle = blockReference.Handle;
                                                        blockModel.BlockObjectId = blockReference.ObjectId;
                                                        blockModel.HostDrawing = drawingName;
                                                        blockModel.BlockLocation = $"Layout - {layout.LayoutName}";
                                                        blockModel.LayoutName = layout.LayoutName;
                                                        blockModel.CurrentSpace = "Paper";

                                                        AttributeCollection attributeCollection = blockReference.AttributeCollection;
                                                        if (attributeCollection != null && attributeCollection.Count > 0)
                                                        {
                                                            foreach (ObjectId attribute in attributeCollection)
                                                            {
                                                                AttributeReference attributeReference = transaction2.GetObject(attribute, OpenMode.ForRead) as AttributeReference;
                                                                AttributeModel attributeModel = new AttributeModel();
                                                                attributeModel.AttributeTag = attributeReference.Tag;
                                                                attributeModel.AttributeValue = attributeReference.TextString;

                                                                if (blockModel.Attributes == null)
                                                                {
                                                                    blockModel.Attributes = new ObservableCollection<AttributeModel>();
                                                                }
                                                                blockModel.Attributes.Add(attributeModel);
                                                            }
                                                            result.Add(blockModel);
                                                        }
                                                    }
                                                }
                                                //no need to commit or save since we're only here to read data.
                                            }
                                        }
                                        #endregion

                                        #region Model Space
                                        BlockTable blockTable = transaction2.GetObject(database2.BlockTableId, OpenMode.ForRead) as BlockTable;
                                        BlockTableRecord blockTableRecord = transaction2.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                                        foreach (ObjectId objectId in blockTableRecord)
                                        {
                                            BlockReference blockReference = transaction2.GetObject(objectId, OpenMode.ForRead) as BlockReference;
                                            if (blockReference != null)
                                            {
                                                BlockModel blockModel = new BlockModel();
                                                blockModel.BlockName = blockReference.Name;
                                                blockModel.BlockHandle = blockReference.Handle;
                                                blockModel.BlockObjectId = blockReference.ObjectId;
                                                blockModel.HostDrawing = drawingName;
                                                blockModel.BlockLocation = $"Model Space - {blockReference.Position}";
                                                blockModel.StringBlockLocation = $"Model Space - {blockReference.Position}";
                                                blockModel.CurrentSpace = "Model";
                                                blockModel.LayoutName = "Model";


                                                AttributeCollection attributeCollection = blockReference.AttributeCollection;
                                                if (attributeCollection != null && attributeCollection.Count > 0)
                                                {
                                                    foreach (ObjectId attribute in attributeCollection)
                                                    {
                                                        AttributeReference attributeReference = transaction2.GetObject(attribute, OpenMode.ForRead) as AttributeReference;
                                                        AttributeModel attributeModel = new AttributeModel();
                                                        attributeModel.AttributeTag = attributeReference.Tag;
                                                        attributeModel.AttributeValue = attributeReference.TextString;

                                                        if (blockModel.Attributes == null)
                                                        {
                                                            blockModel.Attributes = new ObservableCollection<AttributeModel>();
                                                        }
                                                        blockModel.Attributes.Add(attributeModel);
                                                    }
                                                    result.Add(blockModel);
                                                }
                                            }
                                        }
                                        #endregion
                                    }
                                    HostApplicationServices.WorkingDatabase = thisDrawing.Database;
                                }
                            }
                            catch (System.Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return result;
        }

        /// <summary>
        /// Passes in the new attribute information and updates the block inside the drawing.
        /// </summary>
        /// <param name="drawingName"></param>
        /// <param name="currentSpace"></param>
        /// <param name="blockHandle"></param>
        /// <param name="blockName"></param>
        /// <param name="attributeModels"></param>
        public void UpdateBlock(string drawingName, string currentSpace, Handle blockHandle, string blockName, ObservableCollection<AttributeModel> attributeModels)
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;

            using (DocumentLock documentLock = thisDrawing.LockDocument())
            {
                using (Transaction transaction1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    if (File.Exists(drawingName) && drawingName != null)
                    {
                        try
                        {
                            using (Database database2 = new Database(false, true))
                            {
                                database2.ReadDwgFile(drawingName, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                database2.CloseInput(true);
                                //otherDatabase.ResolveXrefs(useThreadEngine: true, doNewOnly: false);
                                HostApplicationServices.WorkingDatabase = database2;

                                using (Transaction transaction2 = database2.TransactionManager.StartTransaction())
                                {
                                    switch (currentSpace)
                                    {
                                        case "Paper":
                                            #region Paper Space All Layouts
                                            DBDictionary dBDictionary = transaction2.GetObject(database2.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                                            foreach (DBDictionaryEntry dBDictionaryEntry in dBDictionary)
                                            {
                                                Layout layout = dBDictionaryEntry.Value.GetObject(OpenMode.ForRead) as Layout;
                                                if (layout.LayoutName != "Model")
                                                {
                                                    BlockTableRecord blockTablePaperSpace = transaction2.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                                                    foreach (ObjectId objectId in blockTablePaperSpace)
                                                    {
                                                        BlockReference blockReference = transaction2.GetObject(objectId, OpenMode.ForWrite) as BlockReference;

                                                        //Make sure we have the right block reference
                                                        if (blockReference != null)
                                                        {
                                                            if (blockReference.Name == blockName && blockReference.ObjectId == objectId)
                                                            {
                                                                //Make sure the attribute collection has items in it
                                                                AttributeCollection attributeCollection = blockReference.AttributeCollection;
                                                                if (attributeCollection.Count > 0)
                                                                {
                                                                    foreach (ObjectId objectId2 in attributeCollection)
                                                                    {
                                                                        AttributeReference attribute = transaction2.GetObject(objectId2, OpenMode.ForWrite) as AttributeReference;
                                                                        foreach (var attributeModel in attributeModels)
                                                                        {
                                                                            if (attributeModel.AttributeTag == attribute.Tag)
                                                                            {
                                                                                attribute.TextString = attributeModel.AttributeValue;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion
                                            break;
                                        case "Model":
                                            #region Model Space
                                            BlockTable blockTable = transaction2.GetObject(database2.BlockTableId, OpenMode.ForWrite) as BlockTable;
                                            BlockTableRecord blockTableRecord = transaction2.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                                            foreach (ObjectId objectId in blockTableRecord)
                                            {
                                                BlockReference blockReference = transaction2.GetObject(objectId, OpenMode.ForWrite) as BlockReference;
                                                //Make sure we have the right block reference
                                                if (blockReference != null)
                                                {
                                                    if (blockReference.Name == blockName && blockReference.ObjectId == objectId)
                                                    {
                                                        //Make sure the attribute collection has items in it
                                                        AttributeCollection attributeCollection = blockReference.AttributeCollection;
                                                        if (attributeCollection.Count > 0)
                                                        {
                                                            foreach (ObjectId objectId2 in attributeCollection)
                                                            {
                                                                AttributeReference attribute = transaction2.GetObject(objectId2, OpenMode.ForWrite) as AttributeReference;
                                                                foreach (var attributeModel in attributeModels)
                                                                {
                                                                    if (attributeModel.AttributeTag == attribute.Tag)
                                                                    {
                                                                        attribute.TextString = attributeModel.AttributeValue;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion
                                            break;
                                        default:
                                            break;
                                    }
                                    transaction2.Commit();
                                    database2.SaveAs(drawingName, true, DwgVersion.Current, thisDrawing.Database.SecurityParameters);
                                }
                                HostApplicationServices.WorkingDatabase = thisDrawing.Database;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }
        }
        #endregion

        #region Block Insert Methods

        public string GetDrawingPath()
        {
            string output = null;

            //Get the user to select the file
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Title = "Choose WBlock";
                openFileDialog.Filter = "Dwg files|*.dwg";
                if (openFileDialog.ShowDialog() == true)
                {
                    output = openFileDialog.FileName;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return output;
        }

        public string GetCurrentDWGPath()
        {
            string output = null;


            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            output = Path.GetFullPath(thisDrawing.Name);


            return output;
        }

        public Point3d GetInsertionPointForBlock()
        {
            Point3d output = new Point3d(0, 0, 0);

            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor thisEditor = thisDrawing.Editor;

            Utils.SetFocusToDwgView();

            try
            {
                PromptPointOptions ppo = new PromptPointOptions("");
                ppo.Message = "Select insertion point for block.";
                ppo.AllowNone = false;

                PromptPointResult ppr = thisEditor.GetPoint(ppo);
                output = ppr.Value;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return output;
        }

        public void CloneBlockFromExternalDrawing(string blockName, string sourceFilePath, string destinationFilePath, bool placeInModelSpace, int layoutIndex, bool allLayouts, Point3d newBlockInsertPoint, bool newReferenceOnDuplicate, bool replaceDuplicate, bool doNothingDuplicate)
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;

            ObjectId blockId = ObjectId.Null;
            ObjectIdCollection objectIdCollection = new ObjectIdCollection();

            using (DocumentLock docLock = thisDrawing.LockDocument())
            {
                if (File.Exists(sourceFilePath))
                {
                    if (File.Exists(destinationFilePath))
                    {
                        using (Database sourceDatabase = new Database(false, true))
                        {
                            //then use the ReadDwgFile method provided by the API to access a DWG file and get it's database. 
                            sourceDatabase.ReadDwgFile(sourceFilePath, FileOpenMode.OpenForReadAndWriteNoShare, false, null);
                            sourceDatabase.CloseInput(true);
                            HostApplicationServices.WorkingDatabase = sourceDatabase;

                            //Get the object to be cloned from the source
                            using (Transaction sourceTransaction = sourceDatabase.TransactionManager.StartTransaction())
                            {
                                BlockTable sourceBlockTable = sourceTransaction.GetObject(sourceDatabase.BlockTableId, OpenMode.ForRead) as BlockTable;
                                if (sourceBlockTable.Has(blockName))
                                {
                                    foreach (ObjectId objectId in sourceBlockTable)
                                    {
                                        BlockTableRecord blockToClone = sourceTransaction.GetObject(objectId, OpenMode.ForRead) as BlockTableRecord;
                                        if (!(blockToClone.IsLayout || blockToClone.IsFromExternalReference || blockToClone.IsFromOverlayReference || blockToClone.IsDependent || blockToClone.IsAnonymous))
                                        {
                                            if (blockToClone.Name.ToLower() == blockName)
                                            {
                                                objectIdCollection.Add(objectId);
                                            }
                                        }
                                    }

                                }
                                else
                                {
                                    MessageBox.Show("There is no block in the source drawing with the given name. Double check your block name and try again.",
                                    "No block found",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Information);
                                    sourceTransaction.Abort();
                                    return;
                                }
                                sourceTransaction.Commit();
                            }

                            //Clone the object from the source to the destination
                            try
                            {
                                Database destinationDatabase = new Database(true, true);
                                destinationDatabase.ReadDwgFile(destinationFilePath, FileOpenMode.OpenForReadAndWriteNoShare, false, null);
                                destinationDatabase.CloseInput(true);
                                HostApplicationServices.WorkingDatabase = destinationDatabase;

                                //Clone the object
                                try
                                {
                                    if (replaceDuplicate == true)
                                    {
                                        try
                                        {
                                            IdMapping map = new IdMapping();
                                            sourceDatabase.WblockCloneObjects(objectIdCollection, destinationDatabase.BlockTableId, map, DuplicateRecordCloning.Replace, false);
                                        }
                                        catch (System.Exception ex)
                                        {
                                            if (ex.Message == "eHandleExists")
                                            {
                                                MessageBox.Show($"You attempted to insert an EXACT copy of an object that already exists in {destinationFilePath}."
                                                    + System.Environment.NewLine +
                                                       System.Environment.NewLine +

                                                    "This cannot be done because AutoCAD requires every object in the database to have a unique identifier called a Handle."
                                                    + System.Environment.NewLine +
                                                       System.Environment.NewLine +

                                                    "It is likely that you used this tool to copy the object over already."
                                                    + System.Environment.NewLine +
                                                       System.Environment.NewLine +

                                                    "You can insert a new block with the same name. This will overwrite the existing block definition with the clone, however you cannot insert the same clone twice."
                                                    + System.Environment.NewLine +
                                                       System.Environment.NewLine +

                                                    "If you are not inserting a completely new object into the target drawing, you shouldn't use this tool.", "Exact Duplicate Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                                HostApplicationServices.WorkingDatabase = thisDatabase;
                                                return;
                                            }
                                        }
                                    }

                                    if (newReferenceOnDuplicate == true)
                                    {
                                        //ignore will discard the clone and insert a new reference of the existing block already in the target drawing
                                        IdMapping map = new IdMapping();
                                        sourceDatabase.WblockCloneObjects(objectIdCollection, destinationDatabase.BlockTableId, map, DuplicateRecordCloning.Ignore, false);
                                    }

                                    if (doNothingDuplicate == true)
                                    {
                                        //If the user chooses to do nothing....
                                        return;
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    MessageBox.Show(ex.Message);
                                    HostApplicationServices.WorkingDatabase = thisDatabase;
                                    return;
                                }
                                using (destinationDatabase)
                                {
                                    using (Transaction destinationCloneTransaction = destinationDatabase.TransactionManager.StartTransaction())
                                    {
                                        //Get the destination drawings block table, and model space BTR
                                        BlockTable destinationBlockTable = (BlockTable)destinationCloneTransaction.GetObject(destinationDatabase.BlockTableId, OpenMode.ForWrite);
                                        BlockTableRecord modelSpace = (BlockTableRecord)destinationCloneTransaction.GetObject(destinationBlockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                                        //get the correct paper space layout
                                        BlockTableRecord paperSpace = null;
                                        DBDictionary layouts = (DBDictionary)destinationCloneTransaction.GetObject(destinationDatabase.LayoutDictionaryId, OpenMode.ForRead);

                                        //If the user chose model space, place the object in model space
                                        if (placeInModelSpace == true)
                                        {
                                            using (BlockReference newBlockReference = new BlockReference(newBlockInsertPoint, destinationBlockTable[blockName]))
                                            {
                                                modelSpace.UpgradeOpen();
                                                modelSpace.AppendEntity(newBlockReference);
                                                destinationCloneTransaction.AddNewlyCreatedDBObject(newBlockReference, true);
                                            }
                                        }
                                        else
                                        {
                                            //if the user wants to place the block reference in each layout, go through each layout and insert the block...
                                            if (allLayouts == true)
                                            {
                                                foreach (DBDictionaryEntry layout in layouts)
                                                {
                                                    Layout destinationLayout = (Layout)destinationCloneTransaction.GetObject(layout.Value, OpenMode.ForRead);
                                                    if (destinationLayout.LayoutName != "Model")
                                                    {
                                                        paperSpace = (BlockTableRecord)destinationCloneTransaction.GetObject(destinationLayout.BlockTableRecordId, OpenMode.ForRead);

                                                        using (BlockReference newBlockReference = new BlockReference(newBlockInsertPoint, destinationBlockTable[blockName]))
                                                        {
                                                            paperSpace.UpgradeOpen();
                                                            paperSpace.AppendEntity(newBlockReference);
                                                            destinationCloneTransaction.AddNewlyCreatedDBObject(newBlockReference, true);
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                //otherse, find the specified layout and insert it there.
                                                foreach (DBDictionaryEntry layout in layouts)
                                                {
                                                    Layout destinationLayout = (Layout)destinationCloneTransaction.GetObject(layout.Value, OpenMode.ForRead);
                                                    if (destinationLayout.LayoutName != "Model" && destinationLayout.TabOrder == layoutIndex)
                                                    {
                                                        paperSpace = (BlockTableRecord)destinationCloneTransaction.GetObject(destinationLayout.BlockTableRecordId, OpenMode.ForRead);
                                                        using (BlockReference newBlockReference = new BlockReference(newBlockInsertPoint, destinationBlockTable[blockName]))
                                                        {
                                                            paperSpace.UpgradeOpen();
                                                            paperSpace.AppendEntity(newBlockReference);
                                                            destinationCloneTransaction.AddNewlyCreatedDBObject(newBlockReference, true);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        //audit and fix errors
                                        //destinationDatabase.Audit(true, false);
                                        destinationCloneTransaction.Commit();
                                    }
                                    destinationDatabase.SaveAs(destinationFilePath, true, DwgVersion.Current, thisDrawing.Database.SecurityParameters);
                                }
                            }
                            catch (System.Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                        }
                        HostApplicationServices.WorkingDatabase = thisDatabase;
                    }
                }
            }
        }

        public void CloneBlockFromCurrentDrawing(string blockName, string destinationFilePath, bool placeInModelSpace, int layoutIndex, bool allLayouts, Point3d newBlockInsertPoint, bool newReferenceOnDuplicate, bool replaceDuplicate, bool doNothingDuplicate)
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;
            Editor thisEditor = thisDrawing.Editor;

            ObjectId blockId = ObjectId.Null;
            ObjectIdCollection objectIdCollection = new ObjectIdCollection();


            if (File.Exists(destinationFilePath))
            {
                using (thisDatabase)
                {
                    //then use the ReadDwgFile method provided by the API to access a DWG file and get it's database. 
                    //thisDatabase.ReadDwgFile(sourceFilePath, FileOpenMode.OpenForReadAndWriteNoShare, false, null);
                    //thisDatabase.CloseInput(true);
                    //HostApplicationServices.WorkingDatabase = thisDatabase;

                    using (DocumentLock doclock = thisDrawing.LockDocument())
                    {
                        using (Transaction thisTransaction = thisDatabase.TransactionManager.StartTransaction())
                        {
                            BlockTable thisBlockTable = thisTransaction.GetObject(thisDatabase.BlockTableId, OpenMode.ForRead) as BlockTable;
                            if (thisBlockTable.Has(blockName))
                            {
                                BlockTableRecord blockToCopy = thisTransaction.GetObject(thisBlockTable[blockName], OpenMode.ForRead) as BlockTableRecord;
                                blockId = blockToCopy.ObjectId;
                                objectIdCollection.Add(blockId);
                            }
                            else
                            {
                                MessageBox.Show("There is no block in the source drawing with the given name. Double check your block name and try again.",
                                "No block found",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information);
                            }
                            thisTransaction.Commit();
                        }

                        if (objectIdCollection != null && objectIdCollection.Count > 0)
                        {
                            using (Database destinationDatabase = new Database(true, true))
                            {
                                destinationDatabase.ReadDwgFile(destinationFilePath, FileOpenMode.OpenForReadAndWriteNoShare, false, null);
                                destinationDatabase.CloseInput(true);
                                HostApplicationServices.WorkingDatabase = destinationDatabase;

                                using (Transaction destinationTransaction1 = destinationDatabase.TransactionManager.StartTransaction())
                                {
                                    BlockTable destinationBlockTable = destinationTransaction1.GetObject(destinationDatabase.BlockTableId, OpenMode.ForWrite) as BlockTable;
                                    if (destinationBlockTable.Has(blockName))
                                    {
                                        if (replaceDuplicate == true)
                                        {
                                            try
                                            {
                                                IdMapping map = new IdMapping();
                                                thisDatabase.WblockCloneObjects(objectIdCollection, destinationDatabase.BlockTableId, map, DuplicateRecordCloning.Replace, false);
                                            }
                                            catch (System.Exception ex)
                                            {
                                                if (ex.Message == "eHandleExists")
                                                {
                                                    MessageBox.Show($"You attempted to insert an EXACT copy of an object that already exists in {destinationFilePath}."
                                                        + System.Environment.NewLine +
                                                           System.Environment.NewLine +

                                                        "This cannot be done because AutoCAD requires every object in the database to have a unique identifier called a Handle."
                                                        + System.Environment.NewLine +
                                                           System.Environment.NewLine +

                                                        "It is likely that you used this tool to copy the object over already."
                                                        + System.Environment.NewLine +
                                                           System.Environment.NewLine +

                                                        "You can insert a new block with the same name. This will overwrite the existing block definition with the clone, however you cannot insert the same clone twice."
                                                        + System.Environment.NewLine +
                                                           System.Environment.NewLine +

                                                        "If you are not inserting a completely new object into the target drawing, you shouldn't use this tool.", "Exact Duplicate Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                                    HostApplicationServices.WorkingDatabase = thisDatabase;
                                                    return;
                                                }
                                            }
                                        }

                                        if (newReferenceOnDuplicate == true)
                                        {
                                            IdMapping map = new IdMapping();
                                            thisDatabase.WblockCloneObjects(objectIdCollection, destinationDatabase.BlockTableId, map, DuplicateRecordCloning.Ignore, false);
                                        }

                                        if (doNothingDuplicate == true)
                                        {
                                            //If the user chooses to do nothing....
                                            return;
                                        }
                                    }
                                    else
                                    {
                                        try
                                        {
                                            IdMapping map = new IdMapping();
                                            thisDatabase.WblockCloneObjects(objectIdCollection, destinationDatabase.BlockTableId, map, DuplicateRecordCloning.Replace, false);
                                        }
                                        catch (System.Exception ex)
                                        {
                                            if (ex.Message == "eHandleExists")
                                            {
                                                MessageBox.Show($"You attempted to insert an EXACT copy of an object that already exists in {destinationFilePath}."
                                                    + System.Environment.NewLine +
                                                       System.Environment.NewLine +

                                                    "This cannot be done because AutoCAD requires every object in the database to have a unique identifier called a Handle."
                                                    + System.Environment.NewLine +
                                                       System.Environment.NewLine +

                                                    "It is likely that you used this tool to copy the object over already."
                                                    + System.Environment.NewLine +
                                                       System.Environment.NewLine +

                                                    "You can insert a new block with the same name. This will overwrite the existing block definition with the clone, however you cannot insert the same clone twice."
                                                    + System.Environment.NewLine +
                                                       System.Environment.NewLine +

                                                    "If you are not inserting a completely new object into the target drawing, you shouldn't use this tool.", "Exact Duplicate Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                                HostApplicationServices.WorkingDatabase = thisDatabase;
                                                return;
                                            }
                                        }
                                        destinationTransaction1.Commit();
                                    }
                                }

                                #region Block Creation

                                //Create a reference of the block in the destination drawing
                                using (Transaction destinationTransaction2 = destinationDatabase.TransactionManager.StartTransaction())
                                {
                                    BlockTable destinationBlockTable = destinationTransaction2.GetObject(destinationDatabase.BlockTableId, OpenMode.ForRead) as BlockTable;
                                    if (destinationBlockTable.Has(blockName))
                                    {
                                        BlockTableRecord destinationBlock = destinationBlockTable[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;

                                        try
                                        {

                                            #region Insert In Model Space
                                            //If the users chooses model space, put the block in model space
                                            if (placeInModelSpace == true)
                                            {
                                                BlockTableRecord modelSpace = destinationTransaction2.GetObject(destinationBlockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                                                using (BlockReference newBlockReference = new BlockReference(newBlockInsertPoint, destinationBlock.ObjectId))
                                                {
                                                    //change this to be optional
                                                    modelSpace.AppendEntity(newBlockReference);
                                                    destinationTransaction2.AddNewlyCreatedDBObject(newBlockReference, true);
                                                }
                                            }
                                            #endregion

                                            #region Insert in a Layout or All Layouts
                                            //if the user chose paper space, put it in the correct paper space
                                            else
                                            {
                                                DBDictionary destinationDBDictionary = destinationTransaction2.GetObject(destinationDatabase.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                                                //if the user speceified a layout to put it in....
                                                if (allLayouts == false)
                                                {
                                                    foreach (DBDictionaryEntry dBDictionaryEntry in destinationDBDictionary)
                                                    {
                                                        Layout layout = dBDictionaryEntry.Value.GetObject(OpenMode.ForWrite) as Layout;
                                                        if (layout.LayoutName != "Model" && layout.TabOrder == layoutIndex)
                                                        {
                                                            BlockTableRecord blockTablePaperSpace = destinationTransaction2.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                                                            using (BlockReference newBlockReference = new BlockReference(newBlockInsertPoint, destinationBlock.ObjectId))
                                                            {
                                                                //change this to be optional
                                                                blockTablePaperSpace.AppendEntity(newBlockReference);
                                                                destinationTransaction2.AddNewlyCreatedDBObject(newBlockReference, true);
                                                            }
                                                        }

                                                    }
                                                }
                                                //or if the user said to put into all layouts
                                                else
                                                {
                                                    foreach (DBDictionaryEntry dBDictionaryEntry in destinationDBDictionary)
                                                    {
                                                        Layout layout = dBDictionaryEntry.Value.GetObject(OpenMode.ForWrite) as Layout;
                                                        if (layout.LayoutName != "Model")
                                                        {
                                                            BlockTableRecord blockTablePaperSpace = destinationTransaction2.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                                                            using (BlockReference newBlockReference = new BlockReference(newBlockInsertPoint, destinationBlock.ObjectId))
                                                            {
                                                                blockTablePaperSpace.AppendEntity(newBlockReference);
                                                                destinationTransaction2.AddNewlyCreatedDBObject(newBlockReference, true);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion
                                        }
                                        catch (System.Exception ex)
                                        {
                                            MessageBox.Show(ex.Message);
                                        }

                                        destinationTransaction2.Commit();
                                    }
                                    else
                                    {
                                        MessageBox.Show("The block definition does not exist in the target drawing.",
                                        "Something went wrong.",
                                        MessageBoxButton.OK,
                                        MessageBoxImage.Error);
                                        return;
                                    }
                                }
                                #endregion

                                destinationDatabase.SaveAs(destinationFilePath, true, DwgVersion.Current, thisDrawing.Database.SecurityParameters);
                            }
                        }
                        else
                        {
                            MessageBox.Show("A block was not added to the collection to be cloned.",
                                        "Something went wrong.",
                                        MessageBoxButton.OK,
                                        MessageBoxImage.Error);
                        }
                    }

                    HostApplicationServices.WorkingDatabase = thisDatabase;
                }
            }

        }

        #endregion

        #region Helpers

        public void AuditDrawing(string drawingPath)
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;

            using(DocumentLock docLock =  thisDrawing.LockDocument())
            {
                if (File.Exists(drawingPath))
                {
                    if (File.Exists(drawingPath))
                    {
                        using (Database sideDatabase = new Database(false, true))
                        {
                            //then use the ReadDwgFile method provided by the API to access a DWG file and get it's database. 
                            sideDatabase.ReadDwgFile(drawingPath, FileOpenMode.OpenForReadAndWriteNoShare, false, null);
                            sideDatabase.CloseInput(true);
                            HostApplicationServices.WorkingDatabase = sideDatabase;

                            //Get the object to be cloned from the source
                            using (Transaction sourceTransaction = sideDatabase.TransactionManager.StartTransaction())
                            {
                                try
                                {
                                    sideDatabase.Audit(true, false);
                                }
                                catch (System.Exception ex)
                                {
                                    MessageBox.Show(ex.Message);
                                }
                                sourceTransaction.Commit();
                            }
                            sideDatabase.SaveAs(drawingPath, true, DwgVersion.Current, thisDrawing.Database.SecurityParameters);
                        }
                        HostApplicationServices.WorkingDatabase = thisDatabase;
                    }
                }
            }
        }

        #endregion
    }
}
