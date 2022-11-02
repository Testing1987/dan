using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using CADTechnologiesSource.All.Models;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;

namespace CADTechnologiesSource.All.AutoCADHelpers
{
    public class DatabaseHelpers
    {
        #region File Access
        /// <summary>
        /// Attempts to open a .dwg file in AutoCAD.
        /// </summary>
        /// <param name="s">The path of the .dwg file you wish to open.</param>
        public void OpenDrawing(string s)
        {
            DocumentCollection documentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            if (File.Exists(s) && s != null)
            {
                documentCollection.Open(s, false);
            }
            else
            {
                MessageBox.Show($"{s} could not be found. Please verify that the file exists and try again.");
                return;
            }
        }

        /// <summary>
        /// Attempts to open a .dwg file in AutoCAD.
        /// </summary>
        /// <param name="s">The path of the .dwg file you wish to open.</param>
        public void OpenDrawingAndZoomTo(string s, Handle handle, string objectLayout)
        {
            bool isOpen = false;
            //Get the document collection to check to see if the drawing is open
            DocumentCollection documentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            foreach (Document document in documentCollection)
            {
                //Document is found inside the collection so activate it and zoom to the object
                if (document.Name == s)
                {
                    isOpen = true;
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = document;

                    //Zoom To
                    Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Database thisDatabase = thisDrawing.Database;
                    Editor editor = thisDrawing.Editor;
                    using (Transaction transaction1 = thisDatabase.TransactionManager.StartTransaction())
                    {
                        try
                        {
                            //Check to see if our object's layout is in model space....
                            if (objectLayout == "Model")
                            {
                                //if our current layout is not Model
                                if (objectLayout != LayoutManager.Current.CurrentLayout)
                                {
                                    //Switch to Model
                                    LayoutManager.Current.CurrentLayout = objectLayout;
                                }
                            }

                            //Check to see if our object's layout is in a paper space layout
                            if (objectLayout != "Model")
                            {
                                //Check to make sure it's not our current layout and...
                                if (objectLayout != LayoutManager.Current.CurrentLayout)
                                {
                                    //switch to it if it isn't
                                    LayoutManager.Current.CurrentLayout = objectLayout;
                                }
                            }
                            using (ViewTableRecord view = editor.GetCurrentView())
                            {
                                Matrix3d WCS2DCS =
                                    (Matrix3d.Rotation(-view.ViewTwist, view.ViewDirection, view.Target) *
                                    Matrix3d.Displacement(view.Target - Point3d.Origin) *
                                    Matrix3d.PlaneToWorld(view.ViewDirection))
                                    .Inverse();
                                ObjectId objectId = thisDatabase.GetObjectId(false, handle, 0);
                                Entity entity = (Entity)transaction1.GetObject(objectId, OpenMode.ForRead);
                                Extents3d extents3D = entity.GeometricExtents;
                                extents3D.AddExtents(entity.GeometricExtents);
                                extents3D.TransformBy(WCS2DCS);
                                view.Width = extents3D.MaxPoint.X - extents3D.MinPoint.X;
                                view.Height = extents3D.MaxPoint.Y - extents3D.MinPoint.Y;
                                view.CenterPoint =
                                    new Point2d((extents3D.MaxPoint.X + extents3D.MinPoint.X) / 2.0, (extents3D.MaxPoint.Y + extents3D.MinPoint.Y) / 2.0);
                                editor.SetCurrentView(view);
                                transaction1.Commit();
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }

            //Drawing is not open, so open it, then zoom.
            if (File.Exists(s) && s != null && isOpen == false)
            {
                documentCollection.Open(s, false);
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument();

                //Zoom To
                Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Database thisDatabase = thisDrawing.Database;
                Editor editor = thisDrawing.Editor;
                using (Transaction transaction1 = thisDatabase.TransactionManager.StartTransaction())
                {
                    try
                    {
                        //Check to see if our object's layout is in model space....
                        if (objectLayout == "Model")
                        {
                            //if our current layout is not Model
                            if (objectLayout != LayoutManager.Current.CurrentLayout)
                            {
                                //Switch to Model
                                LayoutManager.Current.CurrentLayout = objectLayout;
                            }
                        }

                        //Check to see if our object's layout is in a paper space layout
                        if (objectLayout != "Model")
                        {
                            //Check to make sure it's not our current layout and...
                            if (objectLayout != LayoutManager.Current.CurrentLayout)
                            {
                                //switch to it if it isn't
                                LayoutManager.Current.CurrentLayout = objectLayout;
                            }
                        }

                        using (ViewTableRecord view = editor.GetCurrentView())
                        {
                            Matrix3d WCS2DCS =
                                (Matrix3d.Rotation(-view.ViewTwist, view.ViewDirection, view.Target) *
                                Matrix3d.Displacement(view.Target - Point3d.Origin) *
                                Matrix3d.PlaneToWorld(view.ViewDirection))
                                .Inverse();
                            ObjectId objectId = thisDatabase.GetObjectId(false, handle, 0);
                            Entity entity = (Entity)transaction1.GetObject(objectId, OpenMode.ForRead);
                            Extents3d extents3D = entity.GeometricExtents;
                            extents3D.AddExtents(entity.GeometricExtents);
                            extents3D.TransformBy(WCS2DCS);
                            view.Width = extents3D.MaxPoint.X - extents3D.MinPoint.X;
                            view.Height = extents3D.MaxPoint.Y - extents3D.MinPoint.Y;
                            view.CenterPoint =
                                new Point2d((extents3D.MaxPoint.X + extents3D.MinPoint.X) / 2.0, (extents3D.MaxPoint.Y + extents3D.MinPoint.Y) / 2.0);
                            editor.SetCurrentView(view);
                            transaction1.Commit();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

            }
            else
            {
                return;
            }
        }
        public string GetCurrentDrawingPath()
        {
            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            string Drawingpath = Path.GetFileName(thisDrawing.Name);

            return Drawingpath;
        }
        #endregion

        #region Get Drawing Layers
        /// <summary>
        /// Creates an observable collection of type <see cref="LayerModel"/> and returns it to the caller.
        /// </summary>
        /// <param name="s">The file path of the drawing to build the collection from.</param>
        /// <returns></returns>
        public ObservableCollection<LayerModel> BuildCurrentDrawingLayersCollection()
        {
            ObservableCollection<LayerModel> output = new ObservableCollection<LayerModel>();

            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database thisDatabase = thisDrawing.Database;
            Editor editor = thisDrawing.Editor;
            using (DocumentLock lock1 = thisDrawing.LockDocument())
            {
                using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    try
                    {
                        LayerTable layerTable = trans1.GetObject(thisDatabase.LayerTableId, OpenMode.ForRead) as LayerTable;

                        foreach (ObjectId id in layerTable)
                        {
                            LayerTableRecord layerTableRecord = trans1.GetObject(id, OpenMode.ForRead) as LayerTableRecord;
                            LinetypeTableRecord lineTypeRecord = trans1.GetObject(layerTableRecord.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;

                            if (layerTableRecord != null && lineTypeRecord != null)
                            {
                                LayerModel layer = new LayerModel();

                                layer.DrawingPath = thisDrawing.Name;
                                layer.Name = layerTableRecord.Name;
                                layer.OnOff = layerTableRecord.IsOff ? "Off" : "On";
                                layer.Freeze = layerTableRecord.IsFrozen ? "Frozen" : "Thawed";
                                #region Color
                                //Get the layer color by creating an autocad color, setting it to the value of the color found on the layer, and converting it as necessary.
                                Autodesk.AutoCAD.Colors.Color color1 = layerTableRecord.Color;
                                string color_string = Convert.ToString(color1);
                                //Check for TrueColor by looking for a comma in the value and returning an RGB.
                                if (color_string.ToLower().Contains(",") == true)
                                {
                                    int idx1 = color_string.IndexOf(",", 0);
                                    byte R = Convert.ToByte(color_string.Substring(0, idx1));
                                    int idx2 = color_string.IndexOf(",", idx1 + 1);
                                    byte G = Convert.ToByte(color_string.Substring(idx1 + 1, idx2 - idx1 - 1));
                                    byte B = Convert.ToByte(color_string.Substring(idx2 + 1, color_string.Length - idx2 - 1));
                                    color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(R, G, B);

                                    //Return the truecolor value.
                                    layer.Color = color_string;
                                }
                                else
                                {
                                    //Return the index color.
                                    layer.Color = color_string;
                                }
                                #endregion
                                layer.Linetype = lineTypeRecord.Name;
                                layer.Lineweight = layerTableRecord.LineWeight.ToString();
                                layer.Plot = layerTableRecord.IsPlottable ? "Yes" : "No";
                                #region Transparency
                                if (layerTableRecord.Transparency.IsByAlpha)
                                {
                                    int percentage = (int)(((255 - layerTableRecord.Transparency.Alpha) * 100) / 255);
                                    layer.Transparency = percentage.ToString();
                                }
                                else
                                {
                                    layer.Transparency = "0";
                                }
                                #endregion

                                output.Add(layer);
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }

            //Return the list of AutoCAD models to the ObversableCollection
            return output;
        }

        /// <summary>
        /// Creates an observable collection of type <see cref="LayerModel"/> and returns it to the caller.
        /// </summary>
        /// <param name="s">The file path of the drawing to build the collection from.</param>
        /// <returns></returns>
        public ObservableCollection<LayerModel> BuildExternalDrawingLayersCollection(string s)
        {
            ObservableCollection<LayerModel> output = new ObservableCollection<LayerModel>();

            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor editor = thisDrawing.Editor;
            using (DocumentLock lock1 = thisDrawing.LockDocument())
            {
                using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    if (File.Exists(s) && s != null)
                        try
                        {
                            using (Database database2 = new Database(false, true))
                            {
                                database2.ReadDwgFile(s, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                database2.CloseInput(true);
                                database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);

                                HostApplicationServices.WorkingDatabase = database2;

                                using (Transaction trans2 = database2.TransactionManager.StartTransaction())
                                {
                                    LayerTable layerTable2 = trans2.GetObject(database2.LayerTableId, OpenMode.ForRead) as LayerTable;

                                    foreach (ObjectId id in layerTable2)
                                    {
                                        LayerTableRecord layerTableRecord = trans2.GetObject(id, OpenMode.ForRead) as LayerTableRecord;
                                        LinetypeTableRecord lineTypeRecord = trans2.GetObject(layerTableRecord.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;

                                        if (layerTableRecord != null && lineTypeRecord != null)
                                        {
                                            LayerModel layer = new LayerModel();

                                            layer.DrawingPath = s;
                                            layer.Name = layerTableRecord.Name;
                                            layer.OnOff = layerTableRecord.IsOff ? "Off" : "On";
                                            layer.Freeze = layerTableRecord.IsFrozen ? "Frozen" : "Thawed";
                                            #region Color
                                            //Get the layer color by creating an autocad color, setting it to the value of the color found on the layer, and converting it as necessary.
                                            Autodesk.AutoCAD.Colors.Color color1 = layerTableRecord.Color;
                                            string color_string = Convert.ToString(color1);
                                            //Check for TrueColor by looking for a comma in the value and returning an RGB.
                                            if (color_string.ToLower().Contains(",") == true)
                                            {
                                                int idx1 = color_string.IndexOf(",", 0);
                                                byte R = Convert.ToByte(color_string.Substring(0, idx1));
                                                int idx2 = color_string.IndexOf(",", idx1 + 1);
                                                byte G = Convert.ToByte(color_string.Substring(idx1 + 1, idx2 - idx1 - 1));
                                                byte B = Convert.ToByte(color_string.Substring(idx2 + 1, color_string.Length - idx2 - 1));
                                                color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(R, G, B);

                                                //Return the truecolor value.
                                                layer.Color = color_string;
                                            }
                                            else
                                            {
                                                //Return the index color.
                                                layer.Color = color_string;
                                            }
                                            #endregion
                                            layer.Linetype = lineTypeRecord.Name;
                                            layer.Lineweight = layerTableRecord.LineWeight.ToString();
                                            layer.Plot = layerTableRecord.IsPlottable ? "Yes" : "No";
                                            #region Transparency
                                            if (layerTableRecord.Transparency.IsByAlpha)
                                            {
                                                int percentage = (int)(((255 - layerTableRecord.Transparency.Alpha) * 100) / 255);
                                                layer.Transparency = percentage.ToString();
                                            }
                                            else
                                            {
                                                layer.Transparency = "0";
                                            }
                                            #endregion

                                            output.Add(layer);
                                        }
                                    }
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

            //Return the list of AutoCAD models to the ObversableCollection
            return output;
        }

        /// <summary>
        /// Creates an observable collection of type <see cref="ViewportLayerModel"/> and returns it to the caller.
        /// </summary>
        /// <param name="s">The file path of the drawing to build the collection from.</param>
        /// <returns></returns>
        public ObservableCollection<ViewportLayerModel> BuildDrawingViewportLayersCollection(string s)
        {
            ObservableCollection<ViewportLayerModel> output = new ObservableCollection<ViewportLayerModel>();
            ObjectIdCollection viewportCollection = new ObjectIdCollection();

            Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor editor = thisDrawing.Editor;
            using (DocumentLock lock1 = thisDrawing.LockDocument())
            {
                using (Transaction trans1 = thisDrawing.TransactionManager.StartTransaction())
                {
                    if (File.Exists(s) && s != null)
                    {
                        try
                        {
                            using (Database database2 = new Database(false, true))
                            {
                                database2.ReadDwgFile(s, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                database2.CloseInput(true);
                                database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);

                                HostApplicationServices.WorkingDatabase = database2;

                                using (Transaction trans2 = database2.TransactionManager.StartTransaction())
                                {
                                    //access the layer table of transaction2
                                    LayerTable layerTable2 = trans2.GetObject(database2.LayerTableId, OpenMode.ForRead) as LayerTable;

                                    //access the Layout DBDictionary of database 2 to build a collection of viewports
                                    DBDictionary dBLayoutDictionary = trans2.GetObject(database2.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                                    foreach (DBDictionaryEntry entry in dBLayoutDictionary)
                                    {
                                        Layout layout = trans2.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                                        if (entry.Key != "Model")
                                        {
                                            viewportCollection = layout.GetViewports();
                                            viewportCollection.RemoveAt((0));
                                        }
                                    }

                                    foreach (ObjectId viewportID in viewportCollection)
                                    {
                                        Viewport viewport = trans2.GetObject(viewportID, OpenMode.ForRead, false, true) as Viewport;

                                        foreach (ObjectId id2 in layerTable2)
                                        {
                                            //Set up a LayerTableRecord and a LinetypeTableRecord for each id2
                                            LayerTableRecord vpLayer = trans2.GetObject(id2, OpenMode.ForRead) as LayerTableRecord;
                                            LinetypeTableRecord vpLineTypeRecord = trans2.GetObject(vpLayer.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;

                                            if (vpLayer != null && vpLineTypeRecord != null)
                                            {
                                                //Get the viewport layer overrides
                                                LayerViewportProperties layerViewportProperties = vpLayer.GetViewportOverrides(viewport.ObjectId);
                                                //Get the name of the linetype, based on ObjectId. This will be used later to set the value on the layer model.
                                                LinetypeTableRecord linetypeTableRecord = trans2.GetObject(layerViewportProperties.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;
                                                string viewportLinetypeName = linetypeTableRecord.Name;

                                                //Create a model to contain the information
                                                ViewportLayerModel viewportLayerModel = new ViewportLayerModel();

                                                //Set the properties in the model
                                                viewportLayerModel.ViewportLayer = viewport.Layer;
                                                viewportLayerModel.DrawingPath = s;
                                                viewportLayerModel.Name = vpLayer.Name;
                                                viewportLayerModel.ViewportPosition = viewport.CenterPoint;
                                                viewportLayerModel.ViewportFreeze = viewport.IsLayerFrozenInViewport(vpLayer.ObjectId) ? "Frozen" : "Thawed";
                                                viewportLayerModel.ViewportLinetype = viewportLinetypeName;
                                                viewportLayerModel.ViewportLineweight = layerViewportProperties.LineWeight.ToString();
                                                #region Color
                                                //Get the layer color by creating an autocad color, setting it to the value of the color found on the layer, and converting it as necessary.
                                                Autodesk.AutoCAD.Colors.Color color1 = layerViewportProperties.Color;
                                                string color_string = Convert.ToString(color1);
                                                //Check for TrueColor by looking for a comma in the value and returning an RGB.
                                                if (color_string.ToLower().Contains(",") == true)
                                                {
                                                    int idx1 = color_string.IndexOf(",", 0);
                                                    byte R = Convert.ToByte(color_string.Substring(0, idx1));
                                                    int idx2 = color_string.IndexOf(",", idx1 + 1);
                                                    byte G = Convert.ToByte(color_string.Substring(idx1 + 1, idx2 - idx1 - 1));
                                                    byte B = Convert.ToByte(color_string.Substring(idx2 + 1, color_string.Length - idx2 - 1));
                                                    color1 = Autodesk.AutoCAD.Colors.Color.FromRgb(R, G, B);

                                                    //Return the truecolor value.
                                                    viewportLayerModel.ViewportColor = color_string;
                                                }
                                                else
                                                {
                                                    //Return the index color.
                                                    viewportLayerModel.ViewportColor = color_string;
                                                }
                                                #endregion
                                                #region Transparency
                                                if (layerViewportProperties.Transparency.IsByAlpha)
                                                {
                                                    int percentage = (int)(((255 - layerViewportProperties.Transparency.Alpha) * 100) / 255);
                                                    viewportLayerModel.ViewportTransparency = percentage.ToString();
                                                }
                                                else
                                                {
                                                    viewportLayerModel.ViewportTransparency = "0";
                                                }
                                                #endregion
                                                output.Add(viewportLayerModel);
                                            }
                                        }
                                    }

                                    HostApplicationServices.WorkingDatabase = thisDrawing.Database;
                                    return output;
                                }
                            }
                        }
                        catch (System.Exception)
                        {
                            MessageBox.Show("The drawing could not be accessed.", $"Problem accessing database of {s}", MessageBoxButton.OK, MessageBoxImage.Error);
                            throw;
                        }
                    }
                    return output;
                }
            }
        }
        #endregion

        #region CreateLayers

        static public void CreateLayerInCurrentDatabase(string layerName, short color, bool plot)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document thisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = thisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock docLock = thisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction thisTransaction = thisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.LayerTable layerTable;
                        layerTable = (Autodesk.AutoCAD.DatabaseServices.LayerTable)thisTransaction.GetObject(thisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        if (layerTable.Has(layerName) == true)
                        {
                            layerTable.UpgradeOpen();
                            LayerTableRecord new_layer = thisTransaction.GetObject(layerTable[layerName], OpenMode.ForWrite) as LayerTableRecord;
                            if (new_layer != null)
                            {
                                new_layer.IsPlottable = plot;
                            }
                        }

                        if (layerTable.Has(layerName) == false)
                        {
                            LayerTableRecord newLayer = new Autodesk.AutoCAD.DatabaseServices.LayerTableRecord();
                            newLayer.Name = layerName;
                            newLayer.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, color);
                            newLayer.IsPlottable = plot;
                            layerTable.Add(newLayer);
                            thisTransaction.AddNewlyCreatedDBObject(newLayer, true);
                        }
                        thisTransaction.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        static public void CreateLayerInSideDatabase(Database thisDatabase, string layerName, short color, bool plot)
        {
            try
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction thisTransaction = thisDatabase.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.LayerTable layerTable;
                    layerTable = (Autodesk.AutoCAD.DatabaseServices.LayerTable)thisTransaction.GetObject(thisDatabase.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    if (layerTable.Has(layerName) == true)
                    {
                        layerTable.UpgradeOpen();
                        LayerTableRecord newLayer = thisTransaction.GetObject(layerTable[layerName], OpenMode.ForWrite) as LayerTableRecord;
                        if (newLayer != null)
                        {
                            newLayer.IsPlottable = plot;
                            thisTransaction.Commit();
                        }
                    }

                    if (layerTable.Has(layerName) == false)
                    {
                        layerTable.UpgradeOpen();
                        LayerTableRecord new_layer = new Autodesk.AutoCAD.DatabaseServices.LayerTableRecord();
                        new_layer.Name = layerName;
                        new_layer.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, color);
                        new_layer.IsPlottable = plot;
                        layerTable.Add(new_layer);
                        thisTransaction.AddNewlyCreatedDBObject(new_layer, true);
                        thisTransaction.Commit();
                    }
                    thisTransaction.Dispose();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region Blocks
        /// <summary>
        /// Looks into a <see cref="Document"/> to find the given <see cref="ObjectId"/> and checks if it's a <see cref="BlockReference"/>, and if so, returns it's name.
        /// </summary>
        /// <returns></returns>
        public string GetBlockName(Document thisDrawing, ObjectId objectId)
        {
            string blockName = "";
            try
            {
                // Start a transaction
                using (Transaction transaction = thisDrawing.TransactionManager.StartTransaction())
                {
                    RXClass blockClass = RXObject.GetClass(typeof(BlockReference));
                    if (objectId.ObjectClass != blockClass)
                    {
                        MessageBox.Show("The selected object is not a block.");
                        return blockName;
                    }
                    else
                    {
                        // Create a block out of the user's selection
                        BlockReference selectedBlock = transaction.GetObject(objectId, OpenMode.ForRead) as BlockReference;
                        blockName = selectedBlock.IsDynamicBlock ? ((BlockTableRecord)selectedBlock.DynamicBlockTableRecord.GetObject(OpenMode.ForRead)).Name : selectedBlock.Name;
                        return blockName;
                    }
                }
            }
            catch (System.Exception ex)
            {
                // If an exception is hit display the exception...
                MessageBox.Show(ex.Message);
                return blockName;
            }
        }
        #endregion

        #region Tables

        public static LayerTable GetLayerTableForWrite(Transaction transaction, Database database)
        {
            LayerTable layerTable = transaction.GetObject(database.LayerTableId, OpenMode.ForWrite) as LayerTable;

            return layerTable;
        }

        /// <summary>
        /// Returns a linetype from an accessed side database.
        /// </summary>
        /// /// <param name="fileName">The drawing to access.</param>
        /// <param name="transaction">The transasction with the accessed side database.</param>
        /// <param name="database">The database of the drawing you are looking in.</param>
        /// <param name="linetypeName">The name of the linetype that you want to get.</param>
        /// <returns></returns>
        public static LinetypeTableRecord GetLinetypeFromSideDatabase(string fileName, Transaction transaction, Database database, string linetypeName)
        {
            LinetypeTableRecord linetype = null;
            LinetypeTable linetypeTable = transaction.GetObject(database.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;

            if (linetypeTable.Has(linetypeName))
            {
                linetype = transaction.GetObject(linetypeTable[linetypeName],OpenMode.ForRead) as LinetypeTableRecord;
            }

            return linetype;
        }

        #endregion
    }
}
